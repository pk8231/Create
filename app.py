"""
app.py  ─  오프라인 판매 업로드 자동화 웹앱
Streamlit Cloud 에 무료 배포하여 사용합니다.
"""

import streamlit as st
from processor import run

# ── 페이지 기본 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="판매 업로드 자동화",
    page_icon="📊",
    layout="centered",
)

# ── 제목 ─────────────────────────────────────────────────────
st.title("📊 오프라인 판매 업로드 자동화")
st.markdown(
    "출고현황과 반품현황 파일을 올리면 "
    "**오프라인 판매 업로드 엑셀 파일**을 자동으로 만들어 드립니다."
)
st.divider()

# ── 파일 업로드 ──────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("① 출고현황")
    out_file = st.file_uploader(
        "출고현황 파일 (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="out",
    )

with col2:
    st.subheader("② 반품현황")
    ret_file = st.file_uploader(
        "반품현황 파일 (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="ret",
    )

st.divider()

# ── 실행 버튼 ────────────────────────────────────────────────
if out_file and ret_file:
    if st.button("🚀 파일 생성하기", use_container_width=True, type="primary"):
        with st.spinner("계산 중입니다. 잠시만 기다려 주세요..."):
            try:
                excel_bytes, summary = run(out_file, ret_file)

                # 결과 요약
                st.success("✅ 완료! 아래 버튼을 눌러 다운로드 하세요.")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("기준 날짜",    summary["날짜"])
                c2.metric("항목 수",      f"{summary['항목수']:,} 건")
                c3.metric("총 판매수량",  f"{summary['총판매수량']:,} 개")
                c4.metric("총 실판매액",  f"{summary['총실판매액']:,} 원")

                # 다운로드 버튼
                st.download_button(
                    label="⬇️  엑셀 파일 다운로드",
                    data=excel_bytes,
                    file_name=f"오프라인_판매_업로드_{summary['날짜']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ 오류가 발생했습니다:\n\n{e}")
                st.info("파일 형식이 올바른지 확인해 주세요. (출고현황 / 반품현황 순서)")

else:
    st.info("⬆️  위에서 두 파일을 모두 올려주세요.")

# ── 사용 안내 ────────────────────────────────────────────────
with st.expander("📌 사용 방법 안내"):
    st.markdown("""
    1. **출고현황** 파일을 왼쪽에 업로드
    2. **반품현황** 파일을 오른쪽에 업로드
    3. **파일 생성하기** 버튼 클릭
    4. 완료 후 **엑셀 파일 다운로드** 클릭

    ---
    **처리 규칙**
    - 날짜: 반품 파일의 마지막 날짜로 전체 적용
    - 상품코드(바코드): 품번 + 컬러 숫자코드(01~99) + 사이즈
    - 스타일: 품번
    - 판매수량: 출고수량 − 반품수량
    - 실판매액: 출고금액 − 반품금액
    - 3행: 수량·금액 합계 자동 계산
    """)
