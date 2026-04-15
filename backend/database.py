import os
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.pool import StaticPool

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://postgres:change_me@db:5432/credit_analyzer")


def _create_engine():
    url = DATABASE_URL
    if url.startswith("sqlite"):
        connect_args = {"check_same_thread": False}
        # Single shared in-memory DB for all connections (tests / local sqlite)
        if ":memory:" in url:
            return create_engine(
                url,
                connect_args=connect_args,
                poolclass=StaticPool,
            )
        return create_engine(url, connect_args=connect_args)
    return create_engine(url, pool_pre_ping=True)


engine = _create_engine()
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()
