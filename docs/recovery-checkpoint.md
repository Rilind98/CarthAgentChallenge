# Recovery checkpoint — 18 October 2026

This branch is an implementation checkpoint requested by the user before a
session limit. It is not a release approval. Continue on
`codex/robust-extraction`; do not merge into the team's branch without approval.

## Current objective

Run the application on all 72 supplied PDF/XLSX documents and compare its
output against the accompanying `.key.json` files. Fix demonstrated failures.
The user needs a presentation immediately and authorized pushing code to this
isolated branch. Private inputs, keys, templates, credentials and generated
reports must remain local.

## What is implemented

- Native Excel/PDF acquisition, local OCR geometry, optional Mistral OCR and
  OpenAI semantic mapping.
- Source-backed facts, explicit abstentions, versioned evidence ledger and
  JSON/workbook publication with source pointers.
- Arbeitsversion layout-20 projection, with unresolved output bindings marked.
- CLI, independent fixtures, adversarial tests and a separate development
  evaluator. Runtime must never receive answer keys or generator modules.

## Validation evidence and limits

The frozen v6 build was app 0.12.0 / acquisition 12. It passed 162 existing
tests before the independent review added further probes. Five new acceptance
probes failed on that frozen build: stacked Excel schedules with conflicting
years/units, and four formula-prefixed sheet-name locators.

The last completed full 72-document local-only baseline (v5) produced 2,824
accepted facts: 2,031 proxy-correct and 793 wrong. Accepted precision was
71.92%; correct coverage was 5.65%. This is a failed baseline, not the result of
the current AI-assisted run.

A full v6 run with OpenAI gpt-6-astra/max and Mistral OCR began at
14:30:08 UTC. At the 10-document evaluated checkpoint it had 689 accepted facts,
679 proxy matches and 10 wrong, out of 4,831 expected values. That is 98.55%
proxy precision and 14.06% correct coverage on finished documents only. These
are provisional development metrics, not full-corpus or official scores.

The current evaluator checks values, fields, source rows/cells and available
asset IDs/subnumbers/years. It does not verify every business-identity axis,
uses a loose year comparison, and excludes 455 printed-dash slots from
precision if accepted. A separate QA audit identified these gaps. Preserve
historical proxy results and version stricter scoring separately. Never call
this full-tuple accuracy or real-client validation.

## Confirmed review findings requiring closure

The current source is app 0.13.1 / acquisition 13 / structure 8 / semantics 12.
The implementation owner reports all eight repairs below plus a local footer
veto are present. Focused checks passed: the five new stacked-table/locator
acceptance probes, five focused unit probes, and 60 owned unit tests. The root's
independent full run completed with 172 passed, 4 skipped and 2 failed in
102.93 seconds. One failure came from a private OCR fixture-count assumption;
that test is now isolated to explicit known responses and its focused suite
passes 25/25. The other is a repeatable tampered-evidence rebuild failure:
`ValueError:immutable record collision:export`. App 0.13.1 fixes this by binding
each immutable export-record ID to its manifest hash. Eight focused checks
passed after that fix; the complete suite is being rerun with the private
official template enabled. A complete passing result has not yet been recorded.

Source fingerprint at app 0.13.1 edit freeze:
`c91cf0bc86bd496317028913dbd51b4e3e4c38449f4026e6e65a6f0701d1c708`.

1. Scope Excel header caches to each table/amount row.
2. Persist mapper dependencies so accepted facts have a closed replay graph.
3. Use complete business identity when clearing template formulas.
4. Bind manifest status to verified evidence during resume.
5. Escape untrusted source locators in review workbooks.
6. Make immutable ledger append atomic for identical-byte concurrent inputs.
7. Bound mapper HTTP responses (chosen default: 8 MiB).
8. Inventory pages/sheets skipped by acquisition deadlines.

Repairs may be present in this checkpoint; only an explicit later test/review
record closes a finding. In addition, several layout families currently produce
no facts. A live Mistral response for one register was rejected for duplicate
table references. Keep strict rejection until a valid protocol interpretation
is demonstrated. Do not loosen ownership checks just to raise coverage.

## Local continuation paths

The user's checkout is `/Users/wassimbensalem/Desktop/CarthAgentChallenge`.
Everything below `hackathon-local/` is deliberately excluded from Git:

- Corpus: `hackathon-local/ElabHackathon/A3-Hackathon-Corpus`
- Template: `hackathon-local/templates/Arbeitsversion-merged-layout-20.xlsx`
- Frozen v6 source: `hackathon-local/benchmark-build-v6`
- In-progress outputs: `hackathon-local/corpus-full-v6`
- Run progress: `hackathon-local/corpus-full-v6/benchmark-logs/progress.jsonl`
- Finished marker: `hackathon-local/corpus-full-v6/benchmark-logs/completed.json`
- Evaluation: `hackathon-local/corpus-full-v6-evaluation.json`
- Private runner/evaluator: `hackathon-local/benchmark-tools/`
- Credentials: `.env` (never display or commit its contents).

Read progress before restarting anything. The frozen run is independent of
current source changes; preserve its source/config fingerprints and outputs.
Review and QA working reports are under the local chat workspace
`/Users/wassimbensalem/Documents/Codex/2026-10-18/thi/work/`.

## Reproduce checks

```sh
uv sync --locked
uv run pytest tests/unit tests/acceptance tests/qa -q
uv run python tests/acceptance/evaluate_corpus.py \
  hackathon-local/ElabHackathon/A3-Hackathon-Corpus \
  hackathon-local/corpus-full-v6 --report /tmp/carthagent-evaluation.json
```

Use the README for model configuration and CLI usage. Keep source truth,
accepted precision, correct coverage, workbook destination coverage, and
artifact integrity as separate checks. Passing small fixtures alone is not
sufficient to declare the application robust.
