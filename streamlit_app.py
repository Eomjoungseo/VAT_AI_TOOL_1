"""
영세율첨부서류제출명세서 자동화 — Streamlit 웹 앱
(주)비나우

실행: streamlit run streamlit_app.py
"""

import streamlit as st
import json, os, tempfile, threading
from pathlib import Path
from datetime import datetime
import pandas as pd

from logic import (
    load_매입매출장, generate_rows, create_excel, create_excel_omni,
    parse_환급PDF, parse_수기전표PDF, parse_면세물품명세서PDF,
    fill_외화, apply_외화_to_rows,
    update_검증요약_step1, update_검증요약_step2, update_검증요약_외화,
    parse_거래기간, parse_기수명,
)

# ── 페이지 기본 설정 ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="(주)비나우 영세율첨부서류제출명세서 자동화 - Streamlit 웹앱",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* 사이드바 완전 숨김 */
[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }

/* 상단 배너 */
.main-banner {
    background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 20px;
    color: white;
}
.main-banner h1 {
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 8px 0;
    color: white;
}
.main-banner p {
    font-size: 0.95rem;
    margin: 0;
    opacity: 0.85;
    color: white;
}

/* 탭 스타일 */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #F0F4F8;
    padding: 6px;
    border-radius: 10px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: 600;
    color: #1F3864;
}
.stTabs [aria-selected="true"] {
    background: #1F3864 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

CONFIG_FILE = Path(__file__).parent / "config.json"

DEFAULT_서류명_목록 = [
    "소포수령증",
    "명세서-온라인매출증빙",
    "명세서-간주공급",
    "구매확인서",
    "외국인관광객 즉시환급 물품 판매 실적명세서",
    "외국인관광객 면세물품 판매 및 환급실적명세서",
]

DEFAULT_통화_목록 = [
    "KRW", "JPY", "USD", "EUR", "GBP", "AUD",
    "TWD", "VND", "SGD", "THB", "PHP", "MYR", "AED", "CNY",
]

DEFAULT_MAPPING = {
    "쇼피_대만":             ["소포수령증", "입금증명서 포함", "TWD"],
    "쇼피_베트남":           ["소포수령증", "입금증명서 포함", "VND"],
    "쇼피_싱가폴":           ["소포수령증", "입금증명서 포함", "SGD"],
    "쇼피_태국":             ["소포수령증", "입금증명서 포함", "THB"],
    "쇼피_필리핀":           ["소포수령증", "입금증명서 포함", "PHP"],
    "쇼피_말레이시아":       ["소포수령증", "입금증명서 포함", "MYR"],
    "큐텐":                  ["소포수령증", "입금증명서 포함", "JPY"],
    "자사몰_일본":           ["소포수령증", "입금증명서 포함", "JPY"],
    "라쿠텐":                ["소포수령증", "입금증명서 포함", "JPY"],
    "K Brands":              ["소포수령증", "인보이스 포함",   "KRW"],
    "아마존_미국":           ["소포수령증", "입금증명서 포함", "USD"],
    "아마존_영국":           ["소포수령증", "입금증명서 포함", "GBP"],
    "아마존_유럽 독일":      ["소포수령증", "입금증명서 포함", "EUR"],
    "아마존_유럽 이탈리아":  ["소포수령증", "입금증명서 포함", "EUR"],
    "아마존_유럽 프랑스":    ["소포수령증", "입금증명서 포함", "EUR"],
    "아마존_유럽 스페인":    ["소포수령증", "입금증명서 포함", "EUR"],
    "아마존_유럽 아일랜드":  ["소포수령증", "입금증명서 포함", "EUR"],
    "아마존_호주":           ["소포수령증", "입금증명서 포함", "AUD"],
    "아마존_아랍에미레이트": ["소포수령증", "입금증명서 포함", "AED"],
    "틱톡샵_태국":           ["소포수령증", "인보이스 포함",   "THB"],
    "아마존_일본":           ["명세서-온라인매출증빙", "입금증명서 포함", "JPY"],
    "티몰글로벌 중국":       ["명세서-온라인매출증빙", "입금증명서 포함", "CNY"],
    "BENOW JAPAN":           ["명세서-온라인매출증빙", "인보이스 포함",   "JPY"],
    "BENOW BEAUTY INC.":     ["명세서-온라인매출증빙", "인보이스 포함",   "USD"],
    "간주공급(사업상증여)":  ["명세서-간주공급", "", "KRW"],
}

DEFAULT_ISSUER_CORRECTIONS = {
    "스킨스퀘어드코리아 유한회사": "스킨스퀘어드코리아",
}


# ── Config 로드/저장 ─────────────────────────────────────────────────────────
def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "기수명": "25년 2기 예정",
        "거래기간": "2025.07.01 ~ 2025.09.30",
        "년도": 2025,
        "사업자등록번호": "833-87-01017",
        "상호": "(주)비나우",
        "대표자": "이일주, 김대영",
        "사업장소재지": "서울특별시 서초구 서초대로 411 (GT TOWER)",
        "업태": "제조업 (화장품)",
        "제출사유": "전자무역기반사업자를 통한 전자문서 제출",
        "작성일자_공란": True,
        "mapping": DEFAULT_MAPPING,
        "issuer_corrections": DEFAULT_ISSUER_CORRECTIONS,
        "커스텀_서류명": [],
        "커스텀_통화": [],
        "면세판매장_코드": {
            "21401131": "퓌 아지트 성수",
            "21401129": "퓌 아지트 연남",
            "21401130": "퓌 아지트 부산",
            "21401175": "노크 아카이브 성수",
        },
    }

def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── Session state 초기화 ─────────────────────────────────────────────────────
if 'config' not in st.session_state:
    st.session_state.config = load_config()
if 'shared_xlsx' not in st.session_state:
    st.session_state.shared_xlsx = None      # 1단계 생성 파일 bytes
if 'shared_xlsx_name' not in st.session_state:
    st.session_state.shared_xlsx_name = None
if 'shared_환급_files' not in st.session_state:
    st.session_state.shared_환급_files = []  # [(name, bytes), ...]
if 'shared_수기전표_files' not in st.session_state:
    st.session_state.shared_수기전표_files = []
if 'shared_매입매출장' not in st.session_state:
    st.session_state.shared_매입매출장 = None
if 'run_error' not in st.session_state:
    st.session_state.run_error = None

# config의 면세판매장_코드를 logic 모듈에 반영 (매 실행마다)
import logic as _logic_module
_logic_module.면세판매장_코드.update(
    st.session_state.config.get('면세판매장_코드', {}))


def save_uploaded_to_tmp(uploaded_file) -> str:
    """업로드 파일을 임시 경로에 저장 후 경로 반환"""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()
    uploaded_file.seek(0)
    return tmp.name


def save_bytes_to_tmp(data: bytes, suffix: str) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(data)
    tmp.close()
    return tmp.name


# ── 사이드바 네비게이션 ──────────────────────────────────────────────────────
cfg = st.session_state.config

# ── 상단 배너 ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-banner">
    <h1>📄 ㈜비나우 영세율첨부서류제출명세서 자동화</h1>
    <p>매입매출장 · 환급실적명세서 · 세금계산서현황을 업로드하면 영세율첨부서류제출명세서를 자동 작성합니다.</p>
</div>
""", unsafe_allow_html=True)

# ── 탭 네비게이션 ─────────────────────────────────────────────────────────────
tab_guide, tab_settings, tab_mapping, tab_run = st.tabs([
    "📖 사용 가이드",
    "🏠 기본 설정",
    "⚙️ 거래처 매핑",
    "🚀 명세서 생성",
])


# ════════════════════════════════════════════════════════════════════════════
# 📖 가이드라인
# ════════════════════════════════════════════════════════════════════════════
with tab_guide:
    st.title("📖 가이드라인")
    st.markdown("""
---
## 개요
이 프로그램은 아래 세 가지 파일을 이용해 **영세율첨부서류제출명세서**를
자동 작성합니다.

| 파일 | 출처 |
|---|---|
| 매입매출장 (엑셀) | ERP 매입매출장 메뉴에서 영세매출·기타영세 필터링 후 추출 |
| 환급실적명세서 PDF |  |
| 세금계산서현황 (엑셀) | ERP 세금계산서현황 메뉴에서 영세매출·기타영세 필터링 후 추출 |

---
## 작업 순서

### 1️⃣ 기본 설정
- 기수명, 거래기간, 사업자 정보를 입력하고 **저장**합니다.
- 매 기수 시작 시 가장 먼저 설정해야 합니다.

### 2️⃣ 거래처 매핑
- 거래처별 서류명, 비고, 통화를 설정합니다.
- **새 거래처**가 생기면 행을 추가하고 저장합니다.
- 서류명과 통화는 목록에 없으면 직접 입력할 수 있습니다.

### 3️⃣ 명세서 생성 (통합)
필요한 파일을 한 화면에서 모두 올린 뒤 **[명세서 생성]** 한 번이면 끝납니다.

| 입력 | 필수 | 설명 |
|---|---|---|
| 매입매출장 (엑셀) | ✅ 필수 | 영세매출·기타영세만 필터링된 것 |
| 세금계산서현황 (엑셀) | 외화 있으면 필수 | 외화금액·환율 자동 입력용 |
| 즉시환급 실적명세서 PDF | 선택 | 매장별 여러 파일 한번에 업로드 가능 |
| 사후환급 실적명세서 PDF | 선택 | 매장별 여러 파일 한번에 업로드 가능 |
| 수기전표 (직접 입력) | 선택 | 사업장별 건수·환급액을 화면에서 입력 |

실행하면 내부적으로 **① 서식 행 생성 → ② 외화금액 채우기 → ③ 환급 검증**이 순차로 처리됩니다.

- 결과는 국세청 업로드용 양식 **VATVTZ02100** 형식(시트 1개)으로 생성됩니다.
- 환급 검증(Step 1·Step 2)과 외화 통화별 대조 결과는 **화면 로그**로 표시됩니다. (별도 검증 시트는 생성하지 않습니다.)
- 발급자명은 양식 규칙에 맞춰 브랜드 괄호 없이 거래처명만 기재됩니다.
- 과세기간(시작/종료 년월)은 **기본 설정의 거래기간**에서 자동으로 채워집니다.

### 📌 면세판매장 코드 관리
새 매장이 오픈하면 **기본 설정 탭** 하단의 면세판매장 코드 관리 테이블에서 직접 추가하세요.

| 항목 | 설명 |
|---|---|
| 면세판매장 지정번호 | 환급실적명세서 PDF 첫 페이지 ⑥번 항목에서 확인 (8자리 숫자) |
| 매장명 | 프로그램 내에서 사용할 매장 표시명 |

저장하면 즉시 반영되며, 이후 환급실적명세서 파싱 시 자동으로 사용됩니다.

---

### 5️⃣ 3단계: 외화금액
- 세금계산서현황 엑셀을 업로드하면 서식의 외화금액·환율이 자동 입력됩니다.
- 간주공급 행은 외화 불필요 → 자동 제외됩니다.
- 매핑 실패 셀은 빨간 배경으로 표시 → 직접 입력 필요합니다.

---
## 주의사항

- **수기전표 PDF**는 이미지 스캔본이라 자동 파싱이 안 됩니다.
  2단계에서 건수·환급액을 직접 입력해주세요.
- **원화 차이가 나는 경우**: 신규 거래처가 매핑에 없는 경우입니다.
  거래처 매핑 탭에서 추가 후 1단계를 재실행하세요.
- **3단계에서 빨간 셀이 있는 경우**: CSV에서 해당 거래처의 외화 정보를 찾지 못한 것입니다.
  엑셀 파일을 직접 열어서 입력해주세요.

---
## 파일 구조
```
streamlit_app.py   ← 메인 앱
logic.py           ← 핵심 비즈니스 로직
config.json        ← 설정 저장 (자동 생성)
requirements.txt   ← 패키지 목록
packages.txt       ← 시스템 패키지 목록
```
    """)


# ════════════════════════════════════════════════════════════════════════════
# 🏠 기본 설정
# ════════════════════════════════════════════════════════════════════════════
with tab_settings:
    st.title("🏠 기본 설정")
    st.caption("매 기수마다 이 화면에서 먼저 설정하세요. 기수명을 입력하면 거래기간·년도가 자동으로 채워집니다.")

    # ── 기수명 입력 (form 밖) → 거래기간·년도 자동계산 ──
    기수명 = st.text_input(
        "기수명",
        value=cfg.get("기수명", ""),
        key="settings_기수명",
        help="예: '25년 2기 예정', '26년 1기 확정' — 입력하면 아래 거래기간·년도가 자동 설정됩니다.",
    )
    파싱_년도, 파싱_거래기간 = parse_기수명(기수명)

    # 자동계산 성공 시 그 값을, 실패 시 기존 저장값을 사용
    auto_거래기간 = 파싱_거래기간 if 파싱_거래기간 else cfg.get("거래기간", "")
    auto_년도     = 파싱_년도 if 파싱_년도 else int(cfg.get("년도", 2025))

    if 기수명:
        if 파싱_거래기간:
            st.success(f"✅ 자동 설정 → 거래기간 **{auto_거래기간}** · 년도 **{auto_년도}**")
        elif 파싱_년도:
            st.warning("⚠️ 연도는 인식했지만 분기(1기/2기·예정/확정)를 못 읽었습니다. 거래기간을 직접 확인하세요.")
        else:
            st.warning("⚠️ 기수명 형식을 인식하지 못했습니다. 예: '25년 2기 예정'. 거래기간·년도를 직접 입력하세요.")

    with st.form("settings_form"):
        col1, col2 = st.columns(2)
        with col1:
            거래기간      = st.text_input("거래기간 (자동 설정 / 필요시 수정)",
                                          value=auto_거래기간)
            년도          = st.number_input("년도 (자동 설정 / 필요시 수정)",
                                            value=int(auto_년도),
                                            min_value=2000, max_value=2099, step=1)
            사업자등록번호 = st.text_input("사업자등록번호", value=cfg.get("사업자등록번호",""))
            상호          = st.text_input("상호(법인명)",   value=cfg.get("상호",""))
        with col2:
            대표자        = st.text_input("대표자",         value=cfg.get("대표자",""))
            사업장소재지  = st.text_input("사업장소재지",   value=cfg.get("사업장소재지",""))
            업태          = st.text_input("업태(종목)",     value=cfg.get("업태",""))
            제출사유      = st.text_input("제출사유",       value=cfg.get("제출사유",""))
            작성일자_공란 = st.checkbox("⑦ 작성일자를 공란으로 처리 (권장)",
                                        value=cfg.get("작성일자_공란", True))

        submitted = st.form_submit_button("💾 설정 저장", type="primary")

    if submitted:
        cfg.update({
            "기수명": 기수명, "거래기간": 거래기간, "년도": int(년도),
            "사업자등록번호": 사업자등록번호, "상호": 상호, "대표자": 대표자,
            "사업장소재지": 사업장소재지, "업태": 업태, "제출사유": 제출사유,
            "작성일자_공란": 작성일자_공란,
        })
        st.session_state.config = cfg
        save_config(cfg)
        st.success("✅ 기본 설정이 저장되었습니다.")

    st.markdown("---")
    st.markdown("#### 🏪 면세판매장 코드 관리")
    st.info(
        "**새 매장이 오픈하면 여기서 직접 추가하세요.**\n\n"
        "- **면세판매장 지정번호**: 환급실적명세서 PDF 첫 페이지 **⑥ 면세판매장 지정번호** 항목에서 확인 (8자리 숫자)\n"
        "- **매장명**: 프로그램 내에서 사용할 매장 표시명을 자유롭게 입력\n"
        "- 저장하면 즉시 반영되며, 환급실적명세서 파싱 시 자동으로 사용됩니다."
    )

    코드_dict = cfg.get("면세판매장_코드", {})
    코드_rows = [{"면세판매장 지정번호 (8자리)": k, "매장명": v} for k, v in 코드_dict.items()]
    코드_df = pd.DataFrame(코드_rows) if 코드_rows else pd.DataFrame(
        columns=["면세판매장 지정번호 (8자리)", "매장명"])

    edited_코드 = st.data_editor(
        코드_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "면세판매장 지정번호 (8자리)": st.column_config.TextColumn(
                "면세판매장 지정번호 (8자리)", width="medium"),
            "매장명": st.column_config.TextColumn("매장명", width="medium"),
        },
        key="shop_code_editor",
        height=220,
    )
    st.caption("셀을 **더블클릭**하면 코드와 매장명을 입력하거나 수정할 수 있습니다. 내용을 지우면 해당 행이 저장에서 제외됩니다.")

    if st.button("💾 면세판매장 코드 저장", type="primary"):
        new_코드 = {}
        for _, row in edited_코드.iterrows():
            k = str(row["면세판매장 지정번호 (8자리)"]).strip()
            v = str(row["매장명"]).strip()
            if k and v and k != "nan" and v != "nan":
                new_코드[k] = v
        cfg["면세판매장_코드"] = new_코드
        st.session_state.config = cfg
        save_config(cfg)
        # logic 모듈에도 즉시 반영
        import logic as _lm
        _lm.면세판매장_코드.clear()
        _lm.면세판매장_코드.update(new_코드)
        st.success(f"✅ 면세판매장 코드 {len(new_코드)}개 저장됨.")
        st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# ⚙️ 거래처 매핑
# ════════════════════════════════════════════════════════════════════════════
with tab_mapping:
    st.title("⚙️ 거래처 매핑")
    st.caption("거래처별 서류명·비고·통화를 설정합니다. 목록에 없는 서류명·통화는 직접 입력 가능합니다.")

    tab_map, tab_cor, tab_custom = st.tabs(["거래처 → 서류명·비고·통화", "발급자명 정제", "서류명·통화 목록 관리"])

    # ── 커스텀 목록 (저장된 것 + 기본 목록 합산) ──
    커스텀_서류명 = cfg.get("커스텀_서류명", [])
    커스텀_통화   = cfg.get("커스텀_통화", [])
    전체_서류명   = DEFAULT_서류명_목록 + [s for s in 커스텀_서류명 if s not in DEFAULT_서류명_목록]
    전체_통화     = DEFAULT_통화_목록   + [t for t in 커스텀_통화   if t not in DEFAULT_통화_목록]

    # ── 거래처 매핑 탭 ──
    with tab_map:
        st.markdown("#### 거래처 매핑 테이블")
        st.caption(
            "셀을 **더블클릭**하면 서류명·비고·통화를 수정할 수 있습니다. "
            "새 거래처 추가는 테이블 **오른쪽 상단 ＋ 버튼**을 클릭하세요. "
            "행 삭제는 행을 선택한 후 `Delete` 키를 누르세요.\n\n"
            "⚠️ 수정 후 반드시 **매핑 저장** 버튼을 눌러주세요."
        )

        mapping = cfg.get("mapping", DEFAULT_MAPPING)

        # DataFrame으로 변환
        rows = []
        for 거래처, v in mapping.items():
            rows.append({
                "거래처명": 거래처,
                "서류명": v[0],
                "비고": v[1] if len(v) > 1 else "",
                "통화": v[2] if len(v) > 2 else "",
            })
        df = pd.DataFrame(rows)

        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "거래처명": st.column_config.TextColumn("거래처명", width="medium"),
                "서류명": st.column_config.SelectboxColumn(
                    "서류명", options=전체_서류명, width="large", required=False),
                "비고": st.column_config.TextColumn("비고", width="medium"),
                "통화": st.column_config.SelectboxColumn(
                    "통화", options=전체_통화, width="small", required=False),
            },
            key="mapping_editor",
            height=500,
        )

        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("💾 매핑 저장", type="primary"):
                new_mapping = {}
                for _, row in edited_df.iterrows():
                    k = str(row["거래처명"]).strip()
                    if k:
                        # 선택박스에 없는 값도 직접 입력된 텍스트로 저장
                        서류명 = str(row["서류명"]) if pd.notna(row["서류명"]) else ""
                        비고   = str(row["비고"])   if pd.notna(row["비고"])   else ""
                        통화   = str(row["통화"])   if pd.notna(row["통화"])   else ""
                        new_mapping[k] = [서류명, 비고, 통화]
                cfg["mapping"] = new_mapping
                st.session_state.config = cfg
                save_config(cfg)
                st.success(f"✅ 거래처 매핑 {len(new_mapping)}건 저장됨.")
                st.rerun()
        with col2:
            if st.button("🔄 기본값으로 초기화"):
                cfg["mapping"] = DEFAULT_MAPPING
                st.session_state.config = cfg
                save_config(cfg)
                st.success("기본 매핑으로 초기화됨.")
                st.rerun()

    # ── 발급자명 정제 탭 ──
    with tab_cor:
        st.markdown("#### 발급자명 정제")
        st.caption("매입매출장 거래처명과 실제 서식 발급자명이 다를 때 보정합니다.")

        corrections = cfg.get("issuer_corrections", DEFAULT_ISSUER_CORRECTIONS)
        cor_rows = [{"원본 거래처명": k, "정제 후 발급자명": v} for k, v in corrections.items()]
        cor_df = pd.DataFrame(cor_rows) if cor_rows else pd.DataFrame(
            columns=["원본 거래처명", "정제 후 발급자명"])

        edited_cor = st.data_editor(
            cor_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "원본 거래처명":  st.column_config.TextColumn(width="large"),
                "정제 후 발급자명": st.column_config.TextColumn(width="large"),
            },
            key="correction_editor",
            height=300,
        )

        if st.button("💾 정제 규칙 저장", type="primary"):
            new_cor = {}
            for _, row in edited_cor.iterrows():
                k = str(row["원본 거래처명"]).strip()
                v = str(row["정제 후 발급자명"]).strip()
                if k and v:
                    new_cor[k] = v
            cfg["issuer_corrections"] = new_cor
            st.session_state.config = cfg
            save_config(cfg)
            st.success(f"✅ 발급자명 정제 {len(new_cor)}건 저장됨.")

    # ── 서류명·통화 목록 관리 탭 ──
    with tab_custom:
        st.markdown("#### 서류명·통화 목록에 항목 추가")
        st.caption("여기서 추가하면 서류명·통화 선택 목록에 영구적으로 추가됩니다.")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**현재 서류명 목록**")
            for s in 전체_서류명:
                st.markdown(f"- {s}")
            st.markdown("---")
            new_서류명 = st.text_input("새 서류명 추가", placeholder="예: 수출신고필증")
            if st.button("서류명 추가"):
                if new_서류명 and new_서류명 not in 전체_서류명:
                    커스텀_서류명.append(new_서류명)
                    cfg["커스텀_서류명"] = 커스텀_서류명
                    st.session_state.config = cfg
                    save_config(cfg)
                    st.success(f"'{new_서류명}' 추가됨.")
                    st.rerun()
                elif new_서류명 in 전체_서류명:
                    st.warning("이미 목록에 있습니다.")

        with col2:
            st.markdown("**현재 통화 목록**")
            for t in 전체_통화:
                st.markdown(f"- {t}")
            st.markdown("---")
            new_통화 = st.text_input("새 통화 추가", placeholder="예: HKD")
            if st.button("통화 추가"):
                if new_통화 and new_통화.upper() not in [t.upper() for t in 전체_통화]:
                    커스텀_통화.append(new_통화.upper())
                    cfg["커스텀_통화"] = 커스텀_통화
                    st.session_state.config = cfg
                    save_config(cfg)
                    st.success(f"'{new_통화.upper()}' 추가됨.")
                    st.rerun()
                elif new_통화:
                    st.warning("이미 목록에 있습니다.")


# ════════════════════════════════════════════════════════════════════════════
# 🚀 명세서 생성 (1·2·3단계 통합)
# ════════════════════════════════════════════════════════════════════════════

def 해석_오류(e: Exception) -> str:
    """예외를 사용자가 이해할 수 있는 한글 설명으로 변환. 모르면 일반 안내."""
    name = type(e).__name__
    msg  = str(e)

    # 컬럼 누락 (KeyError: '컬럼명')
    if name == 'KeyError':
        col = msg.strip().strip("'\"")
        return (f"업로드한 파일에 필요한 컬럼 **'{col}'** 이(가) 없습니다.\n\n"
                f"매입매출장/세금계산서현황의 컬럼명이 바뀌었거나, 다른 형식의 파일을 "
                f"올렸을 수 있습니다. 원본 RAW 파일인지, 컬럼명이 정확한지 확인해 주세요.")

    # 숫자 변환 실패 (콤마·문자 섞인 금액 등)
    if name == 'ValueError' and 'invalid literal for int' in msg:
        return ("금액 컬럼에 숫자로 바꿀 수 없는 값이 들어 있습니다.\n\n"
                "공급가액 등에 숫자·콤마 외의 문자가 섞여 있는지 확인해 주세요.")

    # 인코딩 문제
    if name in ('UnicodeDecodeError', 'UnicodeError') or 'codec' in msg or '인코딩' in msg:
        return ("CSV 파일의 인코딩을 인식하지 못했습니다.\n\n"
                "엑셀에서 다시 저장하거나, 'CSV UTF-8' 형식으로 내보낸 뒤 다시 올려 주세요.")

    # 파일 형식 문제
    if name in ('BadZipFile', 'InvalidFileException') or 'not a zip' in msg.lower() \
       or 'Excel file format' in msg:
        return ("엑셀 파일을 열 수 없습니다.\n\n"
                "파일이 손상되었거나, 실제로는 .xlsx가 아닐 수 있습니다. "
                "원본을 다시 받아 올려 주세요.")

    # 빈 데이터
    if name == 'EmptyDataError' or 'No columns to parse' in msg or 'empty' in msg.lower():
        return ("파일에 읽을 데이터가 없습니다.\n\n빈 파일이 아닌지 확인해 주세요.")

    # 매입매출장에 대상 데이터 없음 (직접 raise 한 메시지 등)
    if '기타영세' in msg or '영세매출' in msg:
        return ("매입매출장에서 '기타영세' 또는 '영세매출' 데이터를 찾지 못했습니다.\n\n"
                "세무 구분 컬럼이 올바른지, 해당 거래가 있는 기간인지 확인해 주세요.")

    # 그 외
    return ("처리 중 예기치 못한 오류가 발생했습니다.\n\n"
            "아래 원문과 상세 내용을 확인하시고, 반복되면 파일과 함께 문의해 주세요.")


@st.dialog("⚠️ 명세서 생성 오류")
def show_error_dialog(쉬운설명: str, 원문: str, 상세: str):
    st.markdown(f"### 무엇이 잘못됐나요?\n\n{쉬운설명}")
    st.markdown("---")
    st.markdown("**오류 원문**")
    st.code(원문, language=None)
    with st.expander("🔧 개발자용 상세 내용 (Traceback)"):
        st.code(상세, language=None)
    if st.button("닫기", type="primary"):
        st.rerun()


with tab_run:
    st.title("🚀 영세율첨부서류제출명세서 생성")
    st.caption(
        "필요한 파일을 한 번에 올리고 [명세서 생성]을 누르면, 서식 생성 → 외화금액 채우기 → "
        "환급 검증까지 한 번에 처리되어 국세청 업로드용 양식(VATVTZ02100)이 만들어집니다."
    )

    # ── 입력 영역 ──
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown("#### 📁 입력 파일")
        매입매출장_file = st.file_uploader(
            "① 매입매출장 — 영세매출·기타영세 (필수)",
            type=["xlsx", "xls", "csv"], key="run_매입매출장"
        )
        세금계산서_file = st.file_uploader(
            "② 세금계산서현황 — 외화금액용 (외화 거래 있으면 필수)",
            type=["xlsx", "xls", "csv"], key="run_세금계산서"
        )
        즉시환급_files = st.file_uploader(
            "③ 즉시환급 실적명세서 PDF — 파일명에 사업장명 포함 (예: 퓌 아지트 성수)",
            type=["pdf"], accept_multiple_files=True, key="run_즉시"
        )
        사후환급_files = st.file_uploader(
            "④ 사후환급 실적명세서 PDF — 파일명에 사업장명 포함",
            type=["pdf"], accept_multiple_files=True, key="run_사후"
        )

    with col_right:
        st.markdown("#### ✏️ 수기전표 직접 입력 (환급 검증 Step 1)")
        st.caption("수기전표 PDF는 이미지 스캔본이라 자동 파싱 불가. 직접 확인 후 입력하세요. (없으면 0)")
        사업장목록 = ['퓌 아지트 성수', '퓌 아지트 부산', '퓌 아지트 연남', '노크 아카이브 성수']
        수기_입력 = {}
        h1, h2, h3 = st.columns([2, 1, 2])
        h1.markdown("**사업장**"); h2.markdown("**건수**"); h3.markdown("**환급액 (원)**")
        for 사업장 in 사업장목록:
            c1, c2, c3 = st.columns([2, 1, 2])
            c1.markdown(f"&nbsp;&nbsp;{사업장}")
            건수 = c2.number_input("건수", min_value=0, value=0, step=1,
                                   key=f"run_수기건수_{사업장}", label_visibility="collapsed")
            액   = c3.number_input("환급액", min_value=0, value=0, step=1000,
                                   key=f"run_수기액_{사업장}", label_visibility="collapsed")
            if 건수 > 0 or 액 > 0:
                수기_입력[사업장] = {'건수': int(건수), '환급액': int(액)}

        st.markdown("#### ⚙️ 출력")
        기수명 = cfg.get("기수명", "")
        출력파일명 = st.text_input("출력 파일명",
                                   value=f"영세율첨부서류제출명세서_{기수명}.xlsx")

    st.markdown("---")
    run_btn = st.button("🚀 명세서 생성", type="primary",
                        disabled=(매입매출장_file is None))
    log_area = st.empty()

    if run_btn and 매입매출장_file:
        st.session_state.run_error = None   # 새 실행 시작 — 이전 오류 초기화
        logs = []
        def log(msg):
            logs.append(msg)
            log_area.code("\n".join(logs), language=None)

        tmp_paths = []   # 정리용
        try:
            log("════════ 명세서 생성 시작 ════════\n")

            # ── 1) 매입매출장 로드 ──
            log(f"📂 매입매출장 로드: {매입매출장_file.name}")
            tmp_매입 = save_uploaded_to_tmp(매입매출장_file); tmp_paths.append(tmp_매입)
            기타, 영세 = load_매입매출장(tmp_매입)
            log(f"   기타영세: {len(기타)}건 / 영세매출: {len(영세)}건\n")

            # ── 2) 환급 PDF 파싱 (행 생성용 월별 배분) ──
            환급_월별 = {}
            환급_원본 = []   # (파일명, 임시경로) — 검증에서 재사용

            def parse_환급_list(files, 구분):
                if not files: return
                log(f"📑 {구분}환급 PDF {len(files)}개 파싱...")
                for uf in files:
                    tmp = save_uploaded_to_tmp(uf); tmp_paths.append(tmp)
                    환급_원본.append((uf.name, tmp, 구분))
                    사업장, 합계, 취소, err = parse_환급PDF(tmp)
                    if err:
                        log(f"  ⚠️  {uf.name}: {err}"); continue
                    취소txt = f" (취소차감 {취소:,})" if 취소 else ""
                    log(f"  ✅ {uf.name} → {사업장}: {합계:,}원{취소txt}")
                    if 사업장:
                        sp_df = 기타[기타['거래처'] == 사업장]
                        months = sp_df['month'].unique()
                        if len(months):
                            per = 합계 // len(months)
                            for m in months:
                                key = (사업장, m)
                                환급_월별.setdefault(key, {})
                                환급_월별[key][구분] = 환급_월별[key].get(구분, 0) + per

            parse_환급_list(즉시환급_files, '즉시')
            parse_환급_list(사후환급_files, '사후')
            if not 즉시환급_files and not 사후환급_files:
                log("⚠️  환급 PDF 미업로드 — 환급 행이 비어있게 생성됩니다.")
            log("")

            # ── 3) 행 생성 ──
            log("⚙️  명세서 행 생성 중...")
            mapping            = cfg.get("mapping", DEFAULT_MAPPING)
            issuer_corrections = cfg.get("issuer_corrections", DEFAULT_ISSUER_CORRECTIONS)
            year               = cfg.get("년도", 2025)

            (rows, 신규거래처, total_원화, 매입매출_원화,
             영세_final, exclude_idx, 간주df, 환급df) = generate_rows(
                기타, 영세, 환급_월별, mapping, issuer_corrections, year)

            log(f"   총 {len(rows)}행 / 신규거래처 {len(신규거래처)}건")
            log(f"   엑셀 원화합계:       {total_원화:>18,}원")
            log(f"   매입매출장 원화합계: {매입매출_원화:>18,}원")
            diff = total_원화 - 매입매출_원화
            log(f"   차이: {diff:,}원 {'✅' if diff == 0 else '❌'}")
            if 신규거래처:
                log(f"\n⚠️  신규 거래처 {len(신규거래처)}건 — 거래처 매핑 탭에서 추가 후 재실행:")
                for nc in 신규거래처:
                    log(f"    • {nc['거래처']} ({nc['브랜드']}) {nc['원화']:,}원")
            log("")

            # ── 4) 외화금액 채우기 (세금계산서현황 있을 때만) ──
            if 세금계산서_file:
                log(f"💱 외화금액 채우기: {세금계산서_file.name}")
                tmp_fx = save_uploaded_to_tmp(세금계산서_file); tmp_paths.append(tmp_fx)
                성공, 실패, csv_합계, 엑셀_합계 = apply_외화_to_rows(rows, tmp_fx)
                log(f"   매핑 성공 {성공}건 / 실패 {len(실패)}건")
                for f_ in 실패:
                    log(f"    ❌ {f_}")
                # 통화별 합계 대조
                모든통화 = sorted(set(csv_합계) | set(엑셀_합계))
                if 모든통화:
                    log(f"\n   {'통화':>5}  {'현황합계':>16}  {'명세서합계':>16}  판정")
                    all_ok = True
                    for 통화 in 모든통화:
                        c_v = csv_합계.get(통화, 0); e_v = 엑셀_합계.get(통화, 0)
                        d = round(e_v - c_v, 2); ok = abs(d) < 0.01
                        all_ok = all_ok and ok
                        log(f"   {통화:>5}  {c_v:>16,.2f}  {e_v:>16,.2f}  {'✅' if ok else f'❌ {d:+,.2f}'}")
                    log(f"   {'✅ 외화 전체 일치' if all_ok else '⚠️ 외화 불일치 — 직접 확인 필요'}")
            else:
                log("💱 세금계산서현황 미업로드 — 외화 행은 환율·외화금액이 빈 채로 생성됩니다.")
            log("")

            # ── 5) 엑셀 생성 (신규 양식) ──
            log("📄 국세청 양식(VATVTZ02100) 생성 중...")
            tmp_out = tempfile.mktemp(suffix=".xlsx")
            create_excel_omni(rows, cfg, tmp_out)
            with open(tmp_out, 'rb') as f:
                xlsx_bytes = f.read()
            os.unlink(tmp_out)
            from_ym, to_ym = parse_거래기간(cfg.get('거래기간', ''))
            log(f"   과세기간: {from_ym or '?'} ~ {to_ym or '?'}  (기본 설정의 거래기간 기준)")
            log(f"   데이터 {len(rows)}행 작성 완료\n")

            # ── 6) 환급 검증 (화면 표시 전용) ──
            log("────────── 환급 검증 ──────────")

            # Step 1: 수기전표 vs 반출승인번호 공란
            log("\n[Step 1] 수기전표 ↔ 반출승인번호 공란")
            수기결과 = dict(수기_입력)
            if 수기결과:
                for sp, vv in 수기결과.items():
                    log(f"  ✏️  {sp}: {vv['건수']}건 / {vv['환급액']:,}원")
            else:
                log("  ⚠️  수기전표 미입력")

            면세결과 = {}
            for fname, tmp_path, 구분 in 환급_원본:
                결과, err = parse_면세물품명세서PDF(tmp_path)
                if err:
                    continue
                for sp, vv in (결과 or {}).items():
                    면세결과.setdefault(sp, {'건수': 0, '환급액': 0})
                    면세결과[sp]['건수']  += vv['건수']
                    면세결과[sp]['환급액'] += vv['환급액']

            for sp in sorted(set(list(수기결과) + list(면세결과))):
                s = 수기결과.get(sp, {'건수': 0, '환급액': 0})
                m = 면세결과.get(sp, {'건수': 0, '환급액': 0})
                일치 = s['건수'] == m['건수'] and s['환급액'] == m['환급액']
                log(f"  {'✅' if 일치 else '❌'} {sp}")
                log(f"       수기전표: {s['건수']}건 / {s['환급액']:,}원")
                log(f"       반출공란: {m['건수']}건 / {m['환급액']:,}원")

            # Step 2: 매입매출장 기타영세 vs 환급실적 합계
            log("\n[Step 2] 매입매출장 ↔ 환급실적명세서 합계")
            환급_검증 = {}
            for 거래처 in 사업장목록:
                df_sub = 기타[기타['거래처'] == 거래처]
                if not df_sub.empty:
                    환급_검증[거래처] = {'매입매출장': int(df_sub['공급가액'].sum()), '즉시': 0, '사후': 0}
            for fname, tmp_path, 구분 in 환급_원본:
                사업장, 합계, 취소, err = parse_환급PDF(tmp_path)
                if err or not 사업장: continue
                환급_검증.setdefault(사업장, {'매입매출장': 0, '즉시': 0, '사후': 0})
                환급_검증[사업장][구분] += 합계
            if 환급_검증:
                for sp in sorted(환급_검증):
                    v = 환급_검증[sp]; 명세합계 = v['즉시'] + v['사후']; 매입 = v['매입매출장']
                    일치 = 명세합계 == 매입
                    log(f"  {'✅' if 일치 else f'❌ 차이 {명세합계-매입:+,}'} {sp}")
                    log(f"       매입매출장 {매입:,} / 즉시 {v['즉시']:,} / 사후 {v['사후']:,}")
            else:
                log("  ⚠️  환급 대상 없음")

            # ── 세션 저장 (다운로드용) ──
            st.session_state.shared_xlsx = xlsx_bytes
            st.session_state.shared_xlsx_name = 출력파일명

            log("\n🎉 완료! 아래에서 파일을 다운로드하세요.")

        except Exception as e:
            import traceback
            상세 = traceback.format_exc()
            원문 = f"{type(e).__name__}: {e}"
            쉬운설명 = 해석_오류(e)
            log(f"\n❌ 오류: {원문}")
            log(상세)
            # 배너용으로 세션에 저장 (모달을 닫아도 화면에 남도록)
            st.session_state.run_error = {
                '쉬운설명': 쉬운설명, '원문': 원문, '상세': 상세,
            }
            # 빨간 배너 (즉시 표시)
            st.error(f"**명세서 생성 실패**\n\n{쉬운설명}\n\n---\n오류 원문: `{원문}`")
            # 모달 팝업
            show_error_dialog(쉬운설명, 원문, 상세)
        finally:
            for p in tmp_paths:
                try: os.unlink(p)
                except: pass

        if st.session_state.shared_xlsx:
            st.download_button(
                label="⬇️ 명세서 다운로드",
                data=st.session_state.shared_xlsx,
                file_name=출력파일명,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

    else:
        # 버튼을 누르지 않은 렌더(모달 닫은 직후 rerun 포함):
        # 직전 실행에서 오류가 있었으면 배너를 다시 표시
        err = st.session_state.get('run_error')
        if err:
            st.error(f"**명세서 생성 실패**\n\n{err['쉬운설명']}\n\n---\n오류 원문: `{err['원문']}`")
            with st.expander("🔧 개발자용 상세 내용 (Traceback)"):
                st.code(err['상세'], language=None)

        if st.session_state.shared_xlsx:
            st.info(f"✅ 직전에 생성된 파일이 있습니다: {st.session_state.shared_xlsx_name}")
            st.download_button(
                label="⬇️ 명세서 다운로드",
                data=st.session_state.shared_xlsx,
                file_name=st.session_state.shared_xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
