from __future__ import annotations

from contextlib import contextmanager
from hashlib import sha256
import json
from pathlib import Path
import sqlite3
import os
import fcntl
from typing import Any, Iterator


class Ledger:
    def __init__(self, path: Path, *, busy_timeout_seconds: float | None = None):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        init_lock = (path.parent / f".{path.name}.init.lock").open("a+b")
        try:
            fcntl.flock(init_lock.fileno(), fcntl.LOCK_EX)
            self.db = sqlite3.connect(path, **({"timeout": busy_timeout_seconds} if busy_timeout_seconds is not None else {}))
            self._transaction_depth = 0
            if busy_timeout_seconds is not None:
                self.db.execute(f"PRAGMA busy_timeout={int(busy_timeout_seconds * 1000)}")
            self.db.execute("PRAGMA journal_mode=WAL")
            self.db.execute("PRAGMA synchronous=FULL")
            self.db.executescript("""
            CREATE TABLE IF NOT EXISTS records(seq INTEGER PRIMARY KEY AUTOINCREMENT, id TEXT NOT NULL, document_id TEXT NOT NULL, kind TEXT NOT NULL, payload TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, UNIQUE(id, kind));
            CREATE TABLE IF NOT EXISTS events(seq INTEGER PRIMARY KEY AUTOINCREMENT, record_id TEXT NOT NULL, event TEXT NOT NULL, payload TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS dependencies(record_id TEXT NOT NULL, dependency_id TEXT NOT NULL, axis TEXT NOT NULL, PRIMARY KEY(record_id, dependency_id, axis));
            CREATE TABLE IF NOT EXISTS checkpoints(document_id TEXT NOT NULL, stage TEXT NOT NULL, payload TEXT NOT NULL, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, PRIMARY KEY(document_id, stage));
            CREATE TABLE IF NOT EXISTS cache(cache_key TEXT PRIMARY KEY, payload TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
            """)
        finally:
            fcntl.flock(init_lock.fileno(), fcntl.LOCK_UN)
            init_lock.close()

    def append(self, record_id: str, document_id: str, kind: str, payload: dict[str, Any]) -> None:
        encoded = json.dumps(payload, sort_keys=True)
        self.db.execute("INSERT OR IGNORE INTO records(id,document_id,kind,payload) VALUES(?,?,?,?)", (record_id, document_id, kind, encoded))
        existing = self.db.execute("SELECT payload FROM records WHERE id=? AND kind=?", (record_id, kind)).fetchone()
        if existing is None or existing[0] != encoded:
            raise ValueError(f"immutable record collision:{kind}:{record_id}")
        self._commit_if_needed()

    def event(self, record_id: str, event: str, payload: dict[str, Any] | None = None) -> None:
        self.db.execute("INSERT INTO events(record_id,event,payload) VALUES(?,?,?)", (record_id, event, json.dumps(payload or {}, sort_keys=True)))
        self._commit_if_needed()

    def depend(self, record_id: str, dependency_id: str, axis: str) -> None:
        self.db.execute("INSERT OR IGNORE INTO dependencies VALUES(?,?,?)", (record_id, dependency_id, axis)); self._commit_if_needed()

    def invalidate(self, dependency_id: str, axis: str) -> set[str]:
        direct = {row[0] for row in self.db.execute("SELECT record_id FROM dependencies WHERE dependency_id=? AND axis=?", (dependency_id, axis)).fetchall()}
        ids = set(direct)
        frontier = set(direct)
        while frontier:
            placeholders = ",".join("?" for _ in frontier)
            downstream = {row[0] for row in self.db.execute(f"SELECT record_id FROM dependencies WHERE dependency_id IN ({placeholders})", tuple(frontier)).fetchall()}
            frontier = downstream - ids
            ids.update(frontier)
        with self.transaction():
            for record_id in sorted(ids):
                self.event(record_id, "invalidated", {"dependency_id": dependency_id, "axis": axis})
        return ids

    def checkpoint(self, document_id: str, stage: str, payload: dict[str, Any]) -> None:
        self.db.execute("INSERT INTO checkpoints VALUES(?,?,?,CURRENT_TIMESTAMP) ON CONFLICT(document_id,stage) DO UPDATE SET payload=excluded.payload,updated_at=CURRENT_TIMESTAMP", (document_id, stage, json.dumps(payload, sort_keys=True))); self._commit_if_needed()

    def get_checkpoint(self, document_id: str, stage: str) -> dict[str, Any] | None:
        row = self.db.execute("SELECT payload FROM checkpoints WHERE document_id=? AND stage=?", (document_id, stage)).fetchone()
        return json.loads(row[0]) if row else None

    def cache_put(self, cache_key: str, payload: dict[str, Any]) -> None:
        self.db.execute("INSERT OR IGNORE INTO cache(cache_key,payload) VALUES(?,?)", (cache_key, json.dumps(payload, sort_keys=True)))
        self._commit_if_needed()

    def cache_get(self, cache_key: str) -> dict[str, Any] | None:
        row = self.db.execute("SELECT payload FROM cache WHERE cache_key=?", (cache_key,)).fetchone()
        return json.loads(row[0]) if row else None

    @staticmethod
    def cache_key(source_hash: str, region: str, config: dict[str, Any], versions: dict[str, str]) -> str:
        return sha256(json.dumps([source_hash, region, config, versions], sort_keys=True).encode()).hexdigest()

    def close(self) -> None:
        self.db.close()

    def is_invalidated(self, record_id: str) -> bool:
        row = self.db.execute("SELECT event FROM events WHERE record_id=? AND event IN ('invalidated','revalidated') ORDER BY seq DESC LIMIT 1", (record_id,)).fetchone()
        return bool(row and row[0] == "invalidated")

    @contextmanager
    def transaction(self) -> Iterator[None]:
        outer = self._transaction_depth == 0
        if outer:
            self.db.execute("BEGIN IMMEDIATE")
        self._transaction_depth += 1
        try:
            yield
            self._transaction_depth -= 1
            if outer:
                self.db.commit()
        except BaseException:
            self._transaction_depth -= 1
            if outer:
                self.db.rollback()
            raise

    def _commit_if_needed(self) -> None:
        if self._transaction_depth == 0:
            self.db.commit()


def content_hash(path: Path) -> str:
    h = sha256()
    with path.open("rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def preserve_source(path: Path, root: Path, expected_hash: str) -> Path:
    """Copy source bytes once into a content-addressed, read-only local store."""
    root.mkdir(parents=True, exist_ok=True)
    destination = root / expected_hash
    if destination.exists():
        if content_hash(destination) != expected_hash:
            raise ValueError("immutable source store hash mismatch")
        return destination
    temporary = root / f".{expected_hash}.{os.getpid()}.tmp"
    try:
        with path.open("rb") as source, temporary.open("xb") as target:
            for block in iter(lambda: source.read(1024 * 1024), b""):
                target.write(block)
            target.flush()
            os.fsync(target.fileno())
        if content_hash(temporary) != expected_hash:
            raise ValueError("source changed while preserving bytes")
        os.chmod(temporary, 0o444)
        try:
            os.link(temporary, destination)
        except FileExistsError:
            if content_hash(destination) != expected_hash:
                raise ValueError("immutable source store collision")
        return destination
    finally:
        temporary.unlink(missing_ok=True)
