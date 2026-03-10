import streamlit as st
import sys, os

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.theme import init_theme, render_theme_css, render_theme_toggle, accent_color, muted_color
from utils.excel_gen import (
    generate_excel, get_price,
    load_pricelist as _load_pricelist,
    build_product_catalog as _build_catalog,
    build_template_index as _build_template_index,
    select_template,
    BU_FILES, DATA_DIR,
)
from datetime import date

st.set_page_config(page_title="Ansys 견적서 생성기", page_icon="📊", layout="centered")
init_theme()
render_theme_css()
render_theme_toggle()

# ─── 캐싱 래퍼 ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_pricelist():
    return _load_pricelist()

@st.cache_data(show_spinner=False)
def build_product_catalog():
    return _build_catalog()

@st.cache_data(show_spinner=False)
def build_template_index():
    return _build_template_index()


# ════════════════════════════════════════════════════════════════
#  PAGE UI
# ════════════════════════════════════════════════════════════════
st.markdown('<div class="page-tag">Ansys License Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="page-title">견적서 <em>자동 생성기</em></div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Pricelist 기반으로 모듈을 추가하고 엑셀 견적서를 즉시 다운로드하세요.</div>', unsafe_allow_html=True)
st.divider()

with st.spinner("단가표 및 제품 카탈로그 로딩 중..."):
    pricelist_df   = load_pricelist()
    product_cat    = build_product_catalog()
    template_index = build_template_index()
    all_products   = list(product_cat.keys())

if "quote_items" not in st.session_state:
    st.session_state.quote_items = []

# ─── STEP 1: 기본 정보 ──────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">01</span> 기본 정보 입력</div>', unsafe_allow_html=True)

left, right = st.columns(2)
with left:
    customer = st.text_input("회 사 명", placeholder="예) (주) 한국 컴퍼니")
    contact  = st.text_input("수      신", placeholder="예) 홍길동 부장")
    cust_tel = st.text_input("연 락 처", placeholder="02-0000-0000")

with right:
    issue_date = st.date_input("견 적 일 자", value=date.today())
    our_name   = st.text_input("견 적 담 당", placeholder="예) 김담당")
    our_tel    = st.text_input("담당자 TEL", placeholder="010-0000-0000")
    our_email  = st.text_input("E - M a i l", placeholder="sales@tsne.co.kr")

# ─── STEP 2: 라이선스 유형 ──────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">02</span> 라이선스 유형</div>', unsafe_allow_html=True)

lic_mode = st.radio(
    "라이선스 유형",
    ["구매 (영구 라이선스)", "임대 (연간 라이선스)"],
    horizontal=True,
    label_visibility="collapsed",
)
lic_key = "구매" if "구매" in lic_mode else "임대"

# ─── STEP 3: 모듈 추가 ──────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">03</span> 모듈 추가</div>', unsafe_allow_html=True)

# Pricelist 기반 제품 목록 생성 (구매/임대 열에 가격이 있는 제품만)
price_col = "구매(영구)" if lic_key == "구매" else "임대(연간)"
filtered_products = sorted(
    pricelist_df[pricelist_df[price_col] > 0]["제품명"].tolist()
)

col_search, col_qty, col_btn = st.columns([5, 1, 1])
with col_search:
    search_query = st.text_input(
        "제품명 입력",
        placeholder="예) Mechanical, Fluent, HFSS, CFX...",
        label_visibility="collapsed",
    )
    # 검색 필터
    if search_query.strip():
        matches = [p for p in filtered_products
                   if search_query.lower() in p.lower()]
    else:
        matches = filtered_products

    sel_product = st.selectbox(
        "검색 결과",
        matches if matches else ["(일치하는 제품 없음)"],
        label_visibility="collapsed",
    )

with col_qty:
    add_qty = st.number_input("수량", min_value=1, value=1, label_visibility="collapsed")

with col_btn:
    st.markdown("<div style='padding-top:0.3rem'></div>", unsafe_allow_html=True)
    add_clicked = st.button("＋ 추가", use_container_width=True)

if add_clicked and sel_product and sel_product != "(일치하는 제품 없음)":
    # BU 카탈로그에서 bu/sheet/desc 조회, 해당 유형 없으면 반대쪽에서 fallback
    cat_entry = product_cat.get(sel_product, {})
    alt_key   = "임대" if lic_key == "구매" else "구매"
    cat_info  = cat_entry.get(lic_key) or cat_entry.get(alt_key) or {}
    price     = get_price(pricelist_df, sel_product, lic_key)
    st.session_state.quote_items.append({
        "name":  sel_product,
        "bu":    cat_info.get("bu", "MBU"),
        "sheet": cat_info.get("sheet", ""),
        "desc":  cat_info.get("desc", []),
        "qty":   add_qty,
        "price": price,
    })
    st.rerun()

# ─── STEP 4: 견적 내역 ──────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">04</span> 견적 내역</div>', unsafe_allow_html=True)

items = st.session_state.quote_items

if not items:
    st.info("③ 위에서 모듈을 추가하면 여기에 견적 내역이 표시됩니다.")
else:
    # 테이블 헤더
    h1, h2, h3, h4, h5 = st.columns([5, 1, 2, 2, 1])
    for col, txt in zip([h1,h2,h3,h4], ["제 품 명","수 량","단  가","금  액"]):
        col.markdown(
            f"<div style='font-family:IBM Plex Mono,monospace;font-size:.7rem;"
            f"letter-spacing:.1em;text-transform:uppercase;color:{accent_color()};'>{txt}</div>",
            unsafe_allow_html=True,
        )
    st.markdown(f"<hr style='margin:.3rem 0;border-color:{accent_color()};opacity:.3'>",
                unsafe_allow_html=True)

    updated = []
    for i, item in enumerate(items):
        c1, c2, c3, c4, c5 = st.columns([5, 1, 2, 2, 1])
        with c1:
            st.markdown(
                f"<div style='font-size:.9rem;font-weight:700;color:{accent_color()}'>"
                f"{item['name']}</div>",
                unsafe_allow_html=True,
            )
            # 설명 행 표시
            for d in item["desc"][:8]:
                if d.strip():
                    st.markdown(
                        f"<div style='font-size:.75rem;color:{muted_color()};padding-left:.8rem'>"
                        f"{d}</div>",
                        unsafe_allow_html=True,
                    )
        with c2:
            new_qty = st.number_input(f"qty_{i}", value=item["qty"], min_value=1,
                                      label_visibility="collapsed")
        with c3:
            new_price = st.number_input(f"price_{i}", value=item["price"], step=100000,
                                        label_visibility="collapsed", format="%d")
        with c4:
            st.markdown(
                f"<div style='text-align:right;padding-top:.45rem;font-size:.88rem'>"
                f"{new_qty*new_price:,}원</div>",
                unsafe_allow_html=True,
            )
        with c5:
            if st.button("✕", key=f"del_{i}"):
                st.session_state.quote_items.pop(i)
                st.rerun()
        updated.append({**item, "qty": new_qty, "price": new_price})

    st.session_state.quote_items = updated

    # 합계
    st.markdown(f"<hr style='margin:.5rem 0;border-color:{muted_color()};opacity:.3'>",
                unsafe_allow_html=True)
    disc_col, sum_col = st.columns([2, 3])
    with disc_col:
        st.markdown(
            f"<div style='font-family:IBM Plex Mono,monospace;font-size:.72rem;"
            f"letter-spacing:.1em;text-transform:uppercase;color:{accent_color()};"
            f"margin-bottom:.4rem'>할인율 (%)</div>",
            unsafe_allow_html=True,
        )
        disc_pct = st.number_input(
            "할인율",
            min_value=0, max_value=100, value=0, step=1,
            label_visibility="collapsed",
        )
    with sum_col:
        subtotal = sum(it["qty"] * it["price"] for it in updated)
        disc_amt = int(subtotal * disc_pct / 100)
        final    = subtotal - disc_amt
        st.markdown(
            f"<div style='text-align:right;line-height:2;padding-top:.2rem'>"
            f"소계: {subtotal:,}원<br>"
            f"할인: -{disc_amt:,}원 ({disc_pct}%)<br>"
            f"<span style='font-size:1.1rem;font-weight:900;color:{accent_color()}'>"
            f"합계 (부가세 별도): {final:,}원</span></div>",
            unsafe_allow_html=True,
        )

# ─── STEP 5: 다운로드 ────────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">05</span> 견적서 생성</div>', unsafe_allow_html=True)

if st.button("📊  견적서 엑셀 생성", use_container_width=True):
    if not customer.strip():
        st.error("❌ 회사명을 입력해주세요.")
    elif not items:
        st.error("❌ 모듈을 1개 이상 추가해주세요.")
    else:
        info = {
            "customer":   customer,
            "contact":    contact,
            "tel":        cust_tel,
            "our_name":   our_name,
            "our_tel":    our_tel,
            "our_email":  our_email,
            "issue_date": issue_date,
        }
        with st.spinner("엑셀 생성 중..."):
            try:
                bu_key, sheet_name = select_template(template_index, len(items), lic_key)
                buf   = generate_excel(bu_key, sheet_name, info, items,
                                      disc_pct=disc_pct)
                fname = f"Ansys_견적서_{customer}_{issue_date.strftime('%Y%m%d')}.xlsx"
                st.success("✅ 견적서가 생성되었습니다!")
                st.download_button(
                    "⬇️  엑셀 견적서 다운로드",
                    data=buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"생성 실패: {e}")

st.markdown("---")
st.markdown(
    f"<center style='color:{muted_color()};font-size:.75rem;font-family:monospace'>"
    "Ansys Quotation Generator · 2026 Edition"
    "</center>",
    unsafe_allow_html=True,
)
