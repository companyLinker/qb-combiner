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
import uuid
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


def _safe_index_op(label: str, fn) -> None:
    """Run one index-maintenance step in isolation. A single bad index (most
    commonly: a unique index that can't build because legacy data already
    violates it) must never take down the whole MongoDB connection — that
    would silently drop the entire app into session-only mode over what's
    usually just stale/duplicate data in one collection."""
    try:
        fn()
    except Exception as e:
        print(f"[qb_combiner] index maintenance warning — {label}: {e}")


def _migrate_entity_mappings(db):
    """Add dup_id support to entity_mappings (per-account 'duplicate row'
    fan-out to a second Target Line), self-healing any legacy duplicate
    documents instead of failing to build the new index."""
    # Backfill legacy docs first so dup_id:"" queries/upserts match them.
    db.entity_mappings.update_many(
        {"dup_id": {"$exists": False}}, {"$set": {"dup_id": ""}}
    )

    # Any pre-existing documents that already share the same (profile_id,
    # entity, statement, breadcrumb, qb_account) — e.g. old data saved before
    # a unique index was ever successfully enforced — would block the new
    # index build below. Rather than deleting that data, re-tag the extras
    # with a fresh dup_id so they become legitimate "duplicate" mapping
    # entries the app already understands.
    pipeline = [
        {"$group": {
            "_id": {
                "profile_id": "$profile_id", "entity": "$entity",
                "statement": "$statement", "breadcrumb": "$breadcrumb",
                "qb_account": "$qb_account",
            },
            "docs": {"$push": {"id": "$_id", "dup_id": "$dup_id"}},
            "count": {"$sum": 1},
        }},
        {"$match": {"count": {"$gt": 1}}},
    ]
    for grp in db.entity_mappings.aggregate(pipeline):
        seen_dup_ids = set()
        for d in grp["docs"]:
            did = d.get("dup_id") or ""
            if did in seen_dup_ids:
                db.entity_mappings.update_one(
                    {"_id": d["id"]}, {"$set": {"dup_id": uuid.uuid4().hex[:8]}}
                )
            else:
                seen_dup_ids.add(did)

    # The old 5-field unique index no longer matches the new document shape —
    # drop it before creating the dup_id-aware replacement.
    old_keys = {"profile_id", "entity", "statement", "breadcrumb", "qb_account"}
    for idx_name, idx_info in db.entity_mappings.index_information().items():
        if set(k for k, _ in idx_info.get("key", [])) == old_keys:
            db.entity_mappings.drop_index(idx_name)

    db.entity_mappings.create_index([
        ("profile_id", ASCENDING),
        ("entity", ASCENDING),
        ("statement", ASCENDING),
        ("breadcrumb", ASCENDING),
        ("qb_account", ASCENDING),
        ("dup_id", ASCENDING),
    ], unique=True)


def _ensure_indexes(db):
    _safe_index_op("profiles.name", lambda: db.profiles.create_index(
        [("name", ASCENDING)], unique=True))

    _safe_index_op("mappings unique key", lambda: db.mappings.create_index([
        ("profile_id", ASCENDING),
        ("statement", ASCENDING),
        ("breadcrumb", ASCENDING),
        ("qb_account", ASCENDING),
    ], unique=True))

    _safe_index_op("entity_mappings dup_id migration",
                   lambda: _migrate_entity_mappings(db))

    _safe_index_op("entity_aliases unique key", lambda: db.entity_aliases.create_index([
        ("profile_id", ASCENDING),
        ("qb_entity_name", ASCENDING),
    ], unique=True))

    _safe_index_op("runs index", lambda: db.runs.create_index(
        [("profile_id", ASCENDING), ("ran_at", ASCENDING)]))


def now() -> datetime:
    return datetime.now(timezone.utc)
