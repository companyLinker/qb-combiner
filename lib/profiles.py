"""Mapping-profile CRUD on top of MongoDB.

A "profile" is a named bucket of mapping overrides — typically one profile per
client × fiscal year. Profiles let users save mapping decisions once and reuse
them next year (or for another client) without redoing the review step.
"""

import json
import time
from typing import Optional
from bson.objectid import ObjectId
from . import db as dblib

# ── In-process mapping cache (invalidated on every write) ────────────────────
# Key: profile_id str  →  (timestamp_float, dict)
_MAPPING_CACHE: dict = {}
_ENTITY_MAPPING_CACHE: dict = {}
_CACHE_TTL_SECONDS = 60  # max staleness if another process wrote to Mongo


def _invalidate_mapping_cache(profile_id: str):
    """Call after any write so the next read is fresh from MongoDB."""
    _MAPPING_CACHE.pop(profile_id, None)
    _ENTITY_MAPPING_CACHE.pop(profile_id, None)


def _ensure_db():
    d = dblib.get_db()
    if d is None:
        raise RuntimeError(f"MongoDB not available: {dblib.connection_error()}")
    return d


# ---------------- Profiles ----------------

def list_profiles():
    """Return all profiles, sorted by most-recently-updated."""
    d = dblib.get_db()
    if d is None:
        return []
    return list(d.profiles.find().sort("updated_at", -1))


def get_profile(profile_id: str):
    d = _ensure_db()
    return d.profiles.find_one({"_id": ObjectId(profile_id)})


def get_profile_by_name(name: str):
    d = dblib.get_db()
    if d is None:
        return None
    return d.profiles.find_one({"name": name})


def create_profile(name: str, description: str = "", target_template_name: str = "",
                   created_by: str = "") -> str:
    d = _ensure_db()
    now = dblib.now()
    doc = {
        "name": name,
        "description": description,
        "target_template_name": target_template_name,
        "created_at": now,
        "updated_at": now,
        "created_by": created_by,
    }
    res = d.profiles.insert_one(doc)
    return str(res.inserted_id)


def rename_profile(profile_id: str, new_name: str):
    d = _ensure_db()
    d.profiles.update_one(
        {"_id": ObjectId(profile_id)},
        {"$set": {"name": new_name, "updated_at": dblib.now()}},
    )


def delete_profile(profile_id: str):
    d = _ensure_db()
    pid = ObjectId(profile_id)
    d.mappings.delete_many({"profile_id": pid})
    d.entity_aliases.delete_many({"profile_id": pid})
    d.runs.delete_many({"profile_id": pid})
    d.profiles.delete_one({"_id": pid})


def touch_profile(profile_id: str):
    d = _ensure_db()
    d.profiles.update_one(
        {"_id": ObjectId(profile_id)},
        {"$set": {"updated_at": dblib.now()}},
    )


# ---------------- Mappings ----------------

def upsert_mapping(profile_id: str, statement: str, breadcrumb: str,
                   qb_account: str, target_line: str,
                   source: str = "manual", confidence: Optional[float] = None,
                   updated_by: str = ""):
    """Insert or update a single account → target-line mapping."""
    d = _ensure_db()
    d.mappings.update_one(
        {
            "profile_id": ObjectId(profile_id),
            "statement": statement,
            "breadcrumb": breadcrumb,
            "qb_account": qb_account,
        },
        {
            "$set": {
                "target_line": target_line,
                "source": source,
                "confidence": confidence,
                "updated_at": dblib.now(),
                "updated_by": updated_by,
            }
        },
        upsert=True,
    )
    _invalidate_mapping_cache(profile_id)


def delete_mapping(profile_id: str, statement: str, breadcrumb: str, qb_account: str):
    d = _ensure_db()
    d.mappings.delete_one({
        "profile_id": ObjectId(profile_id),
        "statement": statement,
        "breadcrumb": breadcrumb,
        "qb_account": qb_account,
    })
    _invalidate_mapping_cache(profile_id)


def list_mappings(profile_id: str):
    d = dblib.get_db()
    if d is None:
        return []
    return list(d.mappings.find({"profile_id": ObjectId(profile_id)}))


def mapping_lookup(profile_id: str) -> dict:
    """Return a dict keyed by 'P&L|<bc>|<account>' → target_line for fast lookup.
    Cached in-process with 60s TTL; invalidated on any write in this process."""
    now = time.monotonic()
    cached = _MAPPING_CACHE.get(profile_id)
    if cached and (now - cached[0]) < _CACHE_TTL_SECONDS:
        return cached[1]
    out = {}
    for m in list_mappings(profile_id):
        stmt = m.get("statement")
        bc = m.get("breadcrumb", "")
        qb_acc = m.get("qb_account")
        target_line = m.get("target_line")
        if stmt and qb_acc and target_line is not None:
            key = f"{stmt}|{bc}|{qb_acc}"
            out[key] = target_line
    _MAPPING_CACHE[profile_id] = (now, out)
    return out


def mapping_count(profile_id: str) -> int:
    d = dblib.get_db()
    if d is None:
        return 0
    return d.mappings.count_documents({
        "profile_id": ObjectId(profile_id),
        "statement": {"$ne": "__template_row_override__"}
    })


# ---------------- Entity-specific Mappings ----------------
# Key format in session/lookup: "E|{statement}|{entity}|{breadcrumb}|{qb_account}"

ENTITY_KEY_PREFIX = "E|"


def upsert_entity_mapping(profile_id: str, entity: str, statement: str,
                           breadcrumb: str, qb_account: str, target_line: str,
                           source: str = "manual", updated_by: str = ""):
    """Save a per-company account → target-line override."""
    d = _ensure_db()
    d.entity_mappings.update_one(
        {
            "profile_id": ObjectId(profile_id),
            "entity": entity,
            "statement": statement,
            "breadcrumb": breadcrumb,
            "qb_account": qb_account,
        },
        {"$set": {
            "target_line": target_line,
            "source": source,
            "updated_at": dblib.now(),
            "updated_by": updated_by,
        }},
        upsert=True,
    )
    _invalidate_mapping_cache(profile_id)


def delete_entity_mapping(profile_id: str, entity: str, statement: str,
                           breadcrumb: str, qb_account: str):
    d = _ensure_db()
    d.entity_mappings.delete_one({
        "profile_id": ObjectId(profile_id),
        "entity": entity,
        "statement": statement,
        "breadcrumb": breadcrumb,
        "qb_account": qb_account,
    })
    _invalidate_mapping_cache(profile_id)


def entity_mapping_lookup(profile_id: str) -> dict:
    """Return dict keyed by 'E|{stmt}|{entity}|{bc}|{lbl}' → target_line.
    Cached in-process with 60s TTL; invalidated on any write in this process."""
    now = time.monotonic()
    cached = _ENTITY_MAPPING_CACHE.get(profile_id)
    if cached and (now - cached[0]) < _CACHE_TTL_SECONDS:
        return cached[1]
    d = dblib.get_db()
    if d is None:
        return {}
    out = {}
    for m in d.entity_mappings.find({"profile_id": ObjectId(profile_id)}):
        stmt = m.get("statement")
        ent = m.get("entity")
        bc = m.get("breadcrumb", "")
        qb_acc = m.get("qb_account")
        target_line = m.get("target_line")
        if stmt and ent and qb_acc and target_line is not None:
            key = f"{ENTITY_KEY_PREFIX}{stmt}|{ent}|{bc}|{qb_acc}"
            out[key] = target_line
    _ENTITY_MAPPING_CACHE[profile_id] = (now, out)
    return out


# ---------------- Runs (audit log) ----------------

def log_run(profile_id: str, *, ran_by: str = "",
            entities_count: int = 0, auto_count: int = 0,
            manual_count: int = 0, review_count: int = 0,
            output_filename: str = ""):
    d = _ensure_db()
    d.runs.insert_one({
        "profile_id": ObjectId(profile_id),
        "ran_at": dblib.now(),
        "ran_by": ran_by,
        "entities_count": entities_count,
        "auto_count": auto_count,
        "manual_count": manual_count,
        "review_count": review_count,
        "output_filename": output_filename,
    })


def list_runs(profile_id: str, limit: int = 20):
    d = dblib.get_db()
    if d is None:
        return []
    return list(d.runs.find({"profile_id": ObjectId(profile_id)})
                .sort("ran_at", -1).limit(limit))


# ---------------- Import / Export ----------------

def export_profile_json(profile_id: str) -> str:
    """Serialize a profile (metadata + mappings + entity aliases) to JSON."""
    d = _ensure_db()
    pid = ObjectId(profile_id)
    profile = d.profiles.find_one({"_id": pid})
    if not profile:
        raise ValueError("Profile not found")
    mappings = list(d.mappings.find({"profile_id": pid}))
    aliases = list(d.entity_aliases.find({"profile_id": pid}))

    def jsonify(doc):
        out = {}
        for k, v in doc.items():
            if k == "_id" or k == "profile_id":
                continue
            if hasattr(v, "isoformat"):
                out[k] = v.isoformat()
            else:
                out[k] = v
        return out

    payload = {
        "schema_version": 1,
        "profile": jsonify(profile),
        "mappings": [jsonify(m) for m in mappings],
        "entity_aliases": [jsonify(a) for a in aliases],
    }
    return json.dumps(payload, indent=2)


def import_profile_json(json_str: str, *, created_by: str = "") -> str:
    """Import a profile from JSON. Returns the new profile_id."""
    d = _ensure_db()
    payload = json.loads(json_str)
    p = payload["profile"]
    pid = create_profile(
        name=p["name"] + " (imported)",
        description=p.get("description", ""),
        target_template_name=p.get("target_template_name", ""),
        created_by=created_by,
    )
    pid_obj = ObjectId(pid)
    for m in payload.get("mappings", []):
        stmt = m.get("statement")
        qb_acc = m.get("qb_account")
        target_line = m.get("target_line")
        if stmt and qb_acc and target_line is not None:
            d.mappings.insert_one({
                "profile_id": pid_obj,
                "statement": stmt,
                "breadcrumb": m.get("breadcrumb", ""),
                "qb_account": qb_acc,
                "target_line": target_line,
                "source": m.get("source", "imported"),
                "confidence": m.get("confidence"),
                "updated_at": dblib.now(),
                "updated_by": created_by,
            })
    for a in payload.get("entity_aliases", []):
        d.entity_aliases.insert_one({
            "profile_id": pid_obj,
            "qb_entity_name": a["qb_entity_name"],
            "target_pnl_header": a.get("target_pnl_header"),
            "target_bs_header": a.get("target_bs_header"),
            "position": a.get("position"),
        })
    return pid
