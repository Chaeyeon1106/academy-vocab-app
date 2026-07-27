"""
DB 연결 설정.
지금은 SQLite(파일 기반 DB)를 쓰고, 나중에 Railway/Render에 배포할 때
DATABASE_URL 환경변수만 PostgreSQL 주소로 바꿔주면 코드 수정 없이 전환됩니다.
"""
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# 로컬 개발 중에는 이 폴더에 vocab_app.db 라는 SQLite 파일이 생깁니다.
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./vocab_app.db")

# SQLite는 여러 스레드에서 접근할 때 이 옵션이 필요합니다. (PostgreSQL이면 자동 무시됨)
connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}

engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()


def get_db():
    """각 API 요청마다 DB 세션을 하나씩 만들어주고, 끝나면 자동으로 닫아주는 함수."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
