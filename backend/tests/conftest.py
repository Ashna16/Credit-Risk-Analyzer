import os

# Force in-memory SQLite for tests so a local .env Postgres URL cannot break CI.
os.environ["DATABASE_URL"] = "sqlite:///:memory:"

import pytest
from fastapi.testclient import TestClient

from main import app


@pytest.fixture()
def client() -> TestClient:
    return TestClient(app)
