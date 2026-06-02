"""MongoDB connection and collection helpers.

Behavior:
- If MONGODB_URI is set (via st.secrets or env), connect and persist.
- If MongoDB is unreachable or no URI provided, fall back to session-only mode.
  Callers should check `is_connected()` and gracefully degrade.

Collections:
- profiles         — named mapping profiles (one per client/year)
- mappings         — per-profile account → target_line overrides
- entity_aliases   — per-profile QB-entity → target-column-header mapping
- runs             — audit log of every generated workbook
"""

import os
from datetime import datetime, timezone
from typing import Optional

try:
    import streamlit as st
except ImportError:
    st = None

try:
    from pymongo import MongoClient, ASCENDING
    from pymongo.errors import ConnectionFailure, ServerSelectionTimeoutError
    HAVE_PYMONGO = True
except ImportError:
    HAVE_PYMONGO = False


_client: Optional["MongoClient"] = None
_db = None
_connection_error: Optional[str] = None


def _get_uri() -> Optional[str]:
    """Read MONGODB_URI from st.secrets, then env. Return None if unset."""
    if st is not None:
        try:
            uri = st.secrets.get("MONGODB_URI", None)
            if uri:
                return uri
        except Exception:
            pass
    return os.environ.get("MONGODB_URI")


def _get_db_name() -> str:
    if st is not None:
        try:
            name = st.secrets.get("MONGODB_DB", None)
            if name:
                return name
        except Exception:
            pass
    return os.environ.get("MONGODB_DB", "qb_combiner")


def get_client() -> Optional["MongoClient"]:
    """Lazy-initialize the MongoDB client. Returns None if unavailable."""
    global _client, _db, _connection_error
    if _client is not None:
        return _client
    if not HAVE_PYMONGO:
        _connection_error = "pymongo is not installed"
        return None
    uri = _get_uri()
    if not uri:
        _connection_error = "MONGODB_URI not set"
        return None
    try:
        _client = MongoClient(uri, serverSelectionTimeoutMS=2500)
        _client.admin.command("ping")
        _db = _client[_get_db_name()]
        _ensure_indexes(_db)
        return _client
    except (ConnectionFailure, ServerSelectionTimeoutError) as e:
        _connection_error = f"Could not connect to MongoDB: {e}"
        _client = None
        return None
    except Exception as e:
        _connection_error = f"Mongo init error: {e}"
        _client = None
        return None


def get_db():
    """Return the active mongo database, or None if unavailable."""
    if get_client() is None:
        return None
    return _db


def is_connected() -> bool:
    return get_db() is not None


def connection_error() -> Optional[str]:
    return _connection_error


def _ensure_indexes(db):
    db.profiles.create_index([("name", ASCENDING)], unique=True)
    db.mappings.create_index([
        ("profile_id", ASCENDING),
        ("statement", ASCENDING),
        ("breadcrumb", ASCENDING),
        ("qb_account", ASCENDING),
    ], unique=True)
    db.entity_mappings.create_index([
        ("profile_id", ASCENDING),
        ("entity", ASCENDING),
        ("statement", ASCENDING),
        ("breadcrumb", ASCENDING),
        ("qb_account", ASCENDING),
    ], unique=True)
    db.entity_aliases.create_index([
        ("profile_id", ASCENDING),
        ("qb_entity_name", ASCENDING),
    ], unique=True)
    db.runs.create_index([("profile_id", ASCENDING), ("ran_at", ASCENDING)])


def now() -> datetime:
    return datetime.now(timezone.utc)
