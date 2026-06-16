# 📊 오프라인 판매 업로드 자동화

출고현황, 반품현황 엑셀 파일을 업로드하면  
오프라인 판매 업로드 양식을 자동으로 만들어주는 웹 앱입니다.

---

## 🌐 인터넷에 무료로 올리는 방법 (Streamlit Cloud)

### 1단계 — GitHub 가입 (이미 있으면 건너뜀)
1. [https://github.com](https://github.com) 접속
2. **Sign up** → 이메일, 비밀번호 입력 → 가입 완료

---

### 2단계 — 새 저장소(Repository) 만들기
1. GitHub 로그인 후 오른쪽 위 **+** 버튼 → **New repository**
2. Repository name: `sales-upload-app` (아무 이름 가능)
3. **Public** 선택 → **Create repository** 클릭

---

### 3단계 — 파일 올리기
1. 방금 만든 저장소 페이지에서 **Add file → Upload files** 클릭
2. 아래 4개 파일을 드래그해서 올리기:
   - `app.py`
   - `processor.py`
   - `requirements.txt`
   - `README.md`
3. **Commit changes** 클릭

---

### 4단계 — Streamlit Cloud 배포
1. [https://streamlit.io/cloud](https://streamlit.io/cloud) 접속
2. **Get started free** → GitHub 계정으로 로그인
3. **New app** 클릭
4. 아래와 같이 입력:
   - Repository: `sales-upload-app`
   - Branch: `main`
   - Main file path: `app.py`
5. **Deploy!** 클릭
6. 1~2분 후 웹 주소가 생성됨 (예: `https://sales-upload-app.streamlit.app`)

> ✅ 이 주소를 북마크해두면 매달 그냥 접속해서 사용하면 됩니다!

---

## 💻 내 컴퓨터에서 실행하는 방법 (로컬)

```bash
# 1. 필요 패키지 설치
pip install streamlit pandas openpyxl xlrd

# 2. 앱 실행
streamlit run app.py
```

브라우저에서 자동으로 열립니다.

---

## 📁 파일 구성

| 파일 | 설명 |
|------|------|
| `app.py` | 웹 화면 (Streamlit UI) |
| `processor.py` | 데이터 처리 핵심 로직 |
| `requirements.txt` | 필요 패키지 목록 |

---

## 📋 처리 규칙

| 항목 | 처리 방법 |
|------|----------|
| 날짜 | 반품 파일의 마지막 날짜로 전체 적용 |
| 상품코드(바코드) | 품번 + 컬러숫자코드(01~99) + 사이즈 |
| 스타일 | 품번 |
| 판매수량 | 출고수량 − 반품수량 |
| 실판매액 | 출고금액 − 반품금액 |
| 3행 합계 | 수량·금액 자동 합산 |
