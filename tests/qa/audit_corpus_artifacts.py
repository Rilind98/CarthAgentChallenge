"""Audit a completed corpus output without reading development answer keys."""
from __future__ import annotations

from collections import Counter
from hashlib import sha256
import argparse
import json
from pathlib import Path
import sqlite3


SUPPORTED = {".pdf", ".xlsx"}
FORBIDDEN_MARKERS = (".key.json", "/_generator/", "/_intern/", "/.env")


def digest(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def source_files(root: Path) -> list[Path]:
    return sorted(
        path
        for path in root.rglob("*")
        if path.is_file()
        and path.suffix.lower() in SUPPORTED
        and not any(part.startswith(("_", ".")) for part in path.relative_to(root).parts[:-1])
    )


def audit(corpus: Path, output: Path) -> dict[str, object]:
    sources = source_files(corpus)
    keys = sorted(corpus.glob("f*/*.key.json"))
    paired_sources: list[Path] = []
    for key in keys:
        stem = key.name.removesuffix(".key.json")
        matches = [candidate for suffix in (".pdf", ".xlsx") if (candidate := key.with_name(stem + suffix)).exists()]
        assert len(matches) == 1, f"key must pair with exactly one source: {key} -> {matches}"
        paired_sources.extend(matches)
    assert set(paired_sources) == set(sources), "key/source pairing differs from supported corpus inventory"
    source_digests = [digest(path) for path in sources]
    assert len(source_digests) == len(set(source_digests)), "duplicate source bytes would collapse hash-based evaluator pairing"
    source_hashes = dict(zip(source_digests, sources, strict=True))
    manifests = sorted(output.glob("*/manifest.v1.json"))
    assert manifests, "no published manifests"
    staging = sorted(path for path in output.glob(".*.staging-*") if path.is_dir())
    assert not staging, f"publication staging directories remain: {staging[:5]}"
    evidence_directories = {path.parent for path in output.rglob("evidence.v1.json") if ".carthagent" not in path.parts}
    manifest_directories = {path.parent for path in manifests}
    assert evidence_directories == manifest_directories, "unmanifested or evidence-free publication directory"
    seen_hashes: list[str] = []
    status_counts: Counter[str] = Counter()
    fact_count = abstention_count = binding_count = 0
    published_ids: set[str] = set()
    published_dependencies: set[str] = set()

    for manifest_path in manifests:
        directory = manifest_path.parent
        manifest = json.loads(manifest_path.read_text())
        assert manifest.get("readback_verified") is True, manifest_path
        assert (directory / ".complete").read_text().strip() == manifest["snapshot_digest"], manifest_path
        assert manifest["source_hash"] in source_hashes, manifest_path
        seen_hashes.append(manifest["source_hash"])
        for name, expected in manifest["files"].items():
            artifact = directory / name
            assert artifact.is_file(), artifact
            assert digest(artifact) == expected, artifact

        evidence = json.loads((directory / "evidence.v1.json").read_text())
        assert evidence["run_id"] == manifest["run_id"], manifest_path
        assert evidence["document"]["source_hash"] == manifest["source_hash"], manifest_path
        assert evidence["snapshot"]["id"] == manifest["snapshot_id"], manifest_path
        assert evidence["snapshot"]["digest"] == manifest["snapshot_digest"], manifest_path
        serialized = json.dumps(evidence, sort_keys=True)
        assert not any(marker in serialized for marker in FORBIDDEN_MARKERS), manifest_path
        status_counts[manifest["status"]] += 1
        fact_count += len(evidence["facts"])
        abstention_count += len(evidence["abstentions"])
        binding_count += len(evidence["output_bindings"])
        published_ids.add(evidence["snapshot"]["id"])
        for fact in evidence["facts"]:
            published_ids.add(fact["id"])
            published_dependencies.update(fact["dependency_ids"])
        for binding in evidence["output_bindings"]:
            published_ids.add(binding["id"])
            published_dependencies.add(binding["record_id"])
            published_dependencies.add(binding["snapshot_id"])

    assert len(seen_hashes) == len(set(seen_hashes)), "one source was published more than once"
    assert set(seen_hashes) == set(source_hashes), "published source set differs from corpus source set"

    database = output / ".carthagent" / "ledger.sqlite3"
    assert database.is_file(), database
    connection = sqlite3.connect(f"file:{database}?mode=ro", uri=True)
    try:
        records = {row[0]: (row[1], row[2]) for row in connection.execute("SELECT id, kind, payload FROM records")}
        record_ids = set(records)
        dependencies = list(connection.execute("SELECT record_id, dependency_id, axis FROM dependencies"))
        dependency_pairs = {(record, dependency) for record, dependency, _ in dependencies}
        orphan_edges = [(record, dependency) for record, dependency, _ in dependencies if record not in record_ids or dependency not in record_ids]
        assert not orphan_edges, orphan_edges[:5]
        assert published_ids <= record_ids, sorted(published_ids - record_ids)[:5]
        assert published_dependencies <= record_ids, sorted(published_dependencies - record_ids)[:5]
        export_by_snapshot: dict[str, list[str]] = {}
        for record_id, (kind, payload) in records.items():
            if kind == "export":
                snapshot_id = json.loads(payload).get("snapshot_id")
                export_by_snapshot.setdefault(snapshot_id, []).append(record_id)
        for manifest_path in manifests:
            evidence = json.loads((manifest_path.parent / "evidence.v1.json").read_text())
            snapshot_id = evidence["snapshot"]["id"]
            fact_ids = {fact["id"] for fact in evidence["facts"]}
            binding_ids = {binding["id"] for binding in evidence["output_bindings"]}
            for fact in evidence["facts"]:
                assert all((fact["id"], dependency) in dependency_pairs for dependency in fact["dependency_ids"]), fact["id"]
                assert (snapshot_id, fact["id"]) in dependency_pairs, fact["id"]
            for binding in evidence["output_bindings"]:
                assert (binding["id"], binding["record_id"]) in dependency_pairs, binding["id"]
                assert (binding["id"], snapshot_id) in dependency_pairs, binding["id"]
            exports = export_by_snapshot.get(snapshot_id, [])
            assert len(exports) == 1, (snapshot_id, exports)
            assert all((exports[0], binding_id) in dependency_pairs for binding_id in binding_ids), exports[0]
            assert {binding["record_id"] for binding in evidence["output_bindings"]} == fact_ids
        ledger = {"records": len(record_ids), "dependency_edges": len(dependencies), "orphan_edges": 0}
    finally:
        connection.close()

    return {
        "key_files": len(keys),
        "source_documents": len(sources),
        "published_documents": len(manifests),
        "statuses": dict(sorted(status_counts.items())),
        "facts": fact_count,
        "abstentions": abstention_count,
        "output_bindings": binding_count,
        "ledger": ledger,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    arguments = parser.parse_args()
    print(json.dumps(audit(arguments.corpus, arguments.output), sort_keys=True))


if __name__ == "__main__":
    main()
