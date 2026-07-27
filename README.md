# 학원 단어시험 진도 관리 API (1단계: 백엔드 뼈대)

## 폴더 구조
```
academy_vocab_app/
├── requirements.txt
└── app/
    ├── __init__.py
    ├── database.py   # DB 연결 설정 (지금은 SQLite)
    ├── models.py      # 테이블 정의 (teachers, students, books, words, exam_records)
    ├── schemas.py     # API가 주고받는 데이터 형식
    └── main.py        # 실제 API 엔드포인트들
```

## 로컬에서 실행하는 법

1. 이 폴더를 VSCode로 열기

2. 가상환경 만들고 켜기
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Mac / Linux
source venv/bin/activate
```

3. 패키지 설치
```bash
pip install -r requirements.txt
```

4. 서버 실행
```bash
uvicorn app.main:app --reload
```

5. 브라우저에서 확인
- http://127.0.0.1:8000/docs 로 들어가면 Swagger UI가 뜹니다.
- 여기서 각 API를 클릭해서 "Try it out" 버튼으로 실제 데이터를 넣고 테스트해볼 수 있습니다.
- 처음 실행하면 같은 폴더에 `vocab_app.db` 라는 SQLite 파일이 자동으로 생깁니다. 이게 실제 DB 파일입니다.

## 지금 가능한 것 (테스트 완료)

- 선생님 / 단어책 / 학생 / 단어 등록 (`POST /teachers`, `/books`, `/students`, `/words`)
- 특정 단어책의 Day 범위로 단어 조회 → `GET /words?book_id=1&start_day=15&end_day=17`
  (이 API가 나중에 시험지 자동생성이 데이터를 가져오는 통로가 됩니다)
- 시험 기록 등록 (`POST /exam-records`) — 점수, 합/불합, 재시험 여부까지 기록
- 재시험 대상자 자동 조회 (`GET /exam-records/retest-candidates`)

## 아직 안 된 것 (다음 단계에서 추가할 부분)

- 로그인/인증 (지금은 비밀번호를 암호화 없이 저장 중 — 실서비스 전 반드시 교체)
- 프론트엔드 화면 (지금은 API만 있고 화면 없음 — Swagger UI로만 테스트 가능)
- 시험지 PDF 자동 생성 (템플릿 + WeasyPrint)
- 단어/학생 데이터 대량 입력 (엑셀 업로드 등)
- 실제 서버(Railway/Render) 배포

## 수정/삭제 API가 없는 이유

지금은 등록(POST)과 조회(GET)만 있습니다. 수정(PUT/PATCH)과 삭제(DELETE)는 다음 단계에서
프론트엔드 화면을 만들 때 같이 추가하는 게 자연스러워서 일부러 뺐습니다.
