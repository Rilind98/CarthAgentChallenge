# CarthAgent — Claude continuation handoff

Prepared 18 October 2026, 14:54 UTC. This is a sanitized session handoff, not a
raw chat transcript. It contains no credentials. Read this file and the repo's
`AGENTS.md`, inspect current Git/process state, and continue implementation.

## User's objective and latest instruction

Build a robust extractor for German fixed-asset schedules in PDF and Excel.
Extract values with their correct asset/class, field, reporting year, scale and
source location, then populate the supplied Arbeitsversion workbook and emit
auditable JSON. The supplied corpus is synthetic; organizers confirmed that the
hidden test uses real client documents. Do not assume supplied layouts cover
the complexity of the hidden test.

The user is under severe presentation time pressure. They explicitly authorized
parallel fixes/tests, pushing our code, and most recently **implementing a VLM
fallback for complicated layouts**. They now want Claude to continue because
this chat may hit its limit. Proceed without asking for the same permissions.
Cost is not the current constraint: use the configured strongest OpenAI model.

A motivating ambiguity: an amount can sit between two labels with blank rows
between them. Do not assign it to the nearest label unless record boundaries or
other source evidence establish ownership.

## Repository, branch, authorization

- Local checkout: `/Users/wassimbensalem/Desktop/CarthAgentChallenge`
- GitHub: `https://github.com/Rilind98/CarthAgentChallenge`
- Work only on `codex/robust-extraction`; preserve the team's branches.
- Code and authored documentation may be committed/pushed to that branch.
- Do not merge or open a release PR merely because tests pass.
- Private documents, template, API keys and generated evaluation outputs must
  remain local. `hackathon-local/` and `.env` are Git-ignored.
- Existing unrelated `.DS_Store` is untracked; leave it alone.

Last code/document checkpoint before this handoff: `1802f46`, already pushed.
Earlier commits are `0ec1290` (export-recovery fix), `eaff33d` (OCR test isolation),
and `2d3ce71` (full implementation checkpoint). This handoff may be a later
documentation-only commit; inspect HEAD rather than assuming its hash.

At 14:53 UTC the working tree had no tracked changes. **No VLM implementation
had been written yet.** All coding agents confirmed they have stopped without writing any VLM code.
Verify Git status before editing; the existing benchmark jobs may still run.

## What is verified now

Current implemented version: **app 0.13.1**, acquisition 13, structure 8,
semantics 12. The model prompt/schema before vision are
`compact-region-plan-v8` / `semantic_region_plan_v7`.

The independent complete test run with the actual workbook template enabled
finished with **180 passed, no failures or skips**, five third-party
deprecation warnings, in 131.94 seconds:

```sh
cd /Users/wassimbensalem/Desktop/CarthAgentChallenge
CARTHAGENT_TEST_TEMPLATE=hackathon-local/templates/Arbeitsversion-merged-layout-20.xlsx \
  .venv/bin/python -m pytest tests/unit tests/acceptance tests/qa -q
```

App0.13.1 source fingerprint:
`c91cf0bc86bd496317028913dbd51b4e3e4c38449f4026e6e65a6f0701d1c708`.
Its last change includes the manifest hash in immutable export-record IDs, so
rebuilding tampered output creates a new publication record rather than an
immutable-record collision. This preserves prior history.

**The application is not broadly ready.** Fresh independent review still
requests changes, and corpus quality is poor on complex layouts. Test passes
must not be presented as extraction accuracy or production readiness.

## Actual extraction results

### Completed representative check: frozen app 0.13.0, eight documents

This used gpt-6-astra/max, Mistral OCR and a 300-second document budget. The
0.13.1 patch changes publication recovery, not extraction behavior.

| Source | Accepted | Proxy matches | Wrong | Expected slots |
|---|---:|---:|---:|---:|
| f00_demo_01.xlsx | 415 | 415 | 0 | 444 |
| f00_demo_02.pdf | 220 | 220 | 0 | 444 |
| f00_demo_03.pdf | 197 | 185 | 12 | 444 |
| f01_net_spiegel_01.pdf | 44 | 44 | 0 | 264 |
| f03_block_spiegel_02.xlsx | 17 | 17 | 0 | 443 |
| f05_sap_class_gitter_02.xlsx | 0 | 0 | 0 | 396 |
| f07_sap_class_summary_02.xlsx | 0 | 0 | 0 | 90 |
| f08_excel_spiegel_2y_01.xlsx | 0 | 0 | 0 | 220 |

Total: 893 accepted, 881 proxy matches, 12 wrong asset identities, 2,745 expected
slots. Proxy precision 98.66%; correct coverage 32.09%. All eight outputs are
partial. An independent artifact audit passed: 24,529 ledger records, 17,107
dependency edges, zero orphan edges, matched workbook/JSON hashes.

This set includes easy examples and is not representative of all 72 documents.
The three zero-output examples are native Excel: their problem is layout and
identity interpretation, not OCR recognition. VLM alone will not fix that.

### Full corpus run still active: frozen v6, older app 0.12.0

At 14:53:42 UTC, 54/72 documents had finished. One command failed, 53 published
partial outputs. On finished documents: 2,088 accepted, 1,800 proxy matches,
287 wrong, 1 unevaluable acceptance, 26,661 expected slots. Proxy precision 86.25%
excludes the unevaluable acceptance; correct coverage 6.75%.

Do not substitute these older-build results for the repaired build. Preserve
the frozen source and outputs so measurements remain attributable.

The last completed full 72 local-only baseline was v5: 2,031 correct of 2,824
accepted, 793 wrong, 35,975 expected slots; 71.92% precision and 5.65% coverage.
It failed. We have not completed a full 72 run of the latest source.

## Evaluation caveats — essential

`tests/acceptance/evaluate_corpus.py` is a separate development evaluator. It
reads `.key.json`; runtime must never read keys or generator code. The official
JSON contract/scorer is still unavailable.

The current proxy checks signed printed amounts, scale, semantic field, exact
XLSX cell or PDF row ownership, available asset IDs/subnumbers and year. It does
not grade every class/name/account/entity/book/granularity axis, uses a year
substring check, and does not reject unexpected subnumbers on blank-key slots.
PDF gold contains row boxes, not exact amount boxes. A proxy match therefore
does not prove complete business-identity correctness.

There are 35,975 source slots after excluding 412 XLSX fields with no physical
cell. 455 printed `-` slots have core-key zero but the current parser marks them
unevaluable; 194 flagged slots also need an explicit policy. Do not silently
change historical metrics. QA wrote a proposed `strict_tuple_v2` policy with
source-visible identity rules and explicit dash-zero provenance.

Read these local reports:

- `/Users/wassimbensalem/Documents/Codex/2026-10-18/thi/work/evaluator-audit-v6.md`
- `/Users/wassimbensalem/Documents/Codex/2026-10-18/thi/work/code-review-v7.md`
- `/Users/wassimbensalem/Documents/Codex/2026-10-18/thi/work/qa-carthagent.md`

## Remaining independent-review findings

Latest verdict is **CHANGES_REQUESTED**, not approval. Reproduce and fix:

1. Destination-cell collisions: `exporter.py:209-255` can assign more than one
   fact to the same workbook cell. Protect against overwriting/aliasing.
2. Region evidence closure: `mapper.py:271-369` validates references against
   chunk-wide IDs, allowing context from a different table. Every assignment
   must belong to its claimed region or an explicitly valid context ancestry.
3. Publication reuse: `exporter.py:135-150` still needs stronger completion
   marker and exact file-membership checks before reusing existing output.
4. Mapper-plan cardinality/expansion: bound work and check deadlines during
   validation/expansion, not only around HTTP.

Line numbers refer to the pre-vision review snapshot and may shift. The 0.13.1
export-ID fix was outside that review. Broader errors also remain: native-PDF
sign mistakes, scanned asset-ID mistakes, class/position/two-year layouts with
no accepted facts, and an oversized coherent region that caused
`mapper_minimum_coherent_region_exceeds_request_budget` on
`f09_excel_register_02.pdf` instead of publishing useful partial results.

## Latest authorized task: add grounded PDF vision

Current OpenAI input is text plus geometry. **No page images are sent yet.**
Add optional rendered PDF crop images for locally unresolved regions to the
existing Responses request. Do not spend the entire time budget on text-only
model calls before attaching images. Keep native Excel structure primary.

Exact agreed configuration, not implemented yet:

```toml
[mapper.vision]
enabled = false
detail = "original"
include_page_overview = true
max_long_edge_pixels = 2048
max_total_encoded_image_bytes = 4194304
max_images_per_request = 2
max_requests_per_document = 4
```

Implementation contract:

- Nested `VisionConfig`; opt-in, no new service or key.
- Render bounded crop plus optional overview from original PDF bytes.
- Responses user content includes existing `input_text` plus `input_image`
  data URLs. Use gpt-6-astra/max and configurable validated detail enum.
- Preserve the same typed region-plan response schema: existing observation
  IDs and semantic enums only. Models cannot author amounts or coordinates.
- Fix region evidence closure as a prerequisite; adding images must not widen
  what a plan is allowed to associate.
- Share the existing document, mapper-phase and HTTP-attempt budgets. Vision
  attempts count within the 16 total HTTP attempts; no unbounded extra retry.
- Persist image hash, source page/crop coordinates and audit metadata, not
  image base64 in evidence, logs or Git.
- On render/budget failure, retain local/text extraction with a diagnostic.
- Include image identity/config in cache fingerprints. Bump relevant versions.
- Native Excel remains text/geometry; do not rasterize it in this first patch.

OpenAI's current official vision documentation supports `input_image` data
URLs and `detail="original"` on gpt-6-astra. It recommends original detail for
dense/spatial images and warns that vision can misread small/rotated text:
https://developers.openai.com/api/docs/guides/images-vision

Planned independent acceptance tests were **not yet written**:
loopback server inspects actual multimodal payload; enabled PDF sends bounded
images, disabled mode does not; Excel stays text-only; invalid/empty model plans
preserve local facts; foreign-region IDs are rejected; evidence contains hashes
rather than base64. Then run a small live text-versus-vision comparison on the
same difficult PDFs, preserving separate build/config/output fingerprints.
Do not claim improved accuracy merely because an API image request succeeds.

## Current architecture and ownership rules

Immutable source bytes → observations → layout/association hypotheses →
validated facts → frozen snapshot → output bindings → paired JSON/workbook.
Unknown ownership/period/unit/field remains explicit. Printed signs and source
arithmetic inconsistencies must be preserved. True native amount-grid blanks
may be zero only under the established grid policy; unreadable/N/A is not zero.

Native XLSX/PDF first; Tesseract provides local word geometry on scans/outlines;
Mistral provides additional recognition candidates. Mistral table HTML and
paragraph/table boxes are not authoritative cell coordinates. Its f04 register
response currently fails strict duplicate-table-reference validation; do not
weaken that gate without proving valid protocol semantics.

`src/carthagent/`: acquisition, structure, mapper, semantics, policies, ledger,
exporter and CLI. `tests/fixtures` and `tests/acceptance`: independent source
truth and black-box tests. `tests/qa`: adversarial/evaluator/integrity checks.
Preserve provenance and use separate verification before declaring readiness.

## Credentials and private files

Working OpenAI and Mistral credentials are in the local `.env`, already ignored.
Load them with `uv run --env-file .env`; never print or paste their values.
A credential was exposed earlier and replaced: do not reconstruct or retain it
from chat history. No Azure credentials are required for the current path.

Under `/Users/wassimbensalem/Desktop/CarthAgentChallenge/`:

- Corpus: `hackathon-local/ElabHackathon/A3-Hackathon-Corpus`
- Template: `hackathon-local/templates/Arbeitsversion-merged-layout-20.xlsx`
- Full run: `hackathon-local/corpus-full-v6`
- Full frozen build: `hackathon-local/benchmark-build-v6`
- Full config: `hackathon-local/full-pipeline-quality.toml`
- Representative run: `hackathon-local/presentation-quality-v7`
- Representative evaluation: `hackathon-local/presentation-quality-v7-evaluation.json`
- Representative frozen build: `hackathon-local/benchmark-build-v7-presentation`
- Representative config: `hackathon-local/presentation-quality-v7.toml`
- Live Mistral payloads: `hackathon-local/mistral-benchmark-v1/`
- Private orchestration: `hackathon-local/benchmark-tools/`

The official workbook is layout 20, 27 sheets, 500 asset rows and 20 class-grid
rows. Its SHA256 is
`0449e028926cbe6e32c32fabc6da27a2693c36147c581c8d3b008e9543239668`.
Total depreciation is not automatically planned depreciation; net-book-value
disposals are not gross-cost disposals. Preserve unresolved bindings.

## Running jobs — inspect before starting duplicates

At 14:53 UTC the full v6 benchmark was still running: wrapper PID 54607,
runner PID 54627. It uses eight workers. The finite local finalizer PID 62390
waits for its completion marker, evaluates it and refreshes the local report.
PIDs can change or finish; verify command lines before interacting with them.
Neither job edits source or pushes Git.

Progress: `hackathon-local/corpus-full-v6/benchmark-logs/progress.jsonl`.
Completion: `hackathon-local/corpus-full-v6/benchmark-logs/completed.json`.
The finalizer times out after two hours rather than claiming completion.
Do not mistake the older frozen benchmark for a test of new vision code.

## Useful commands

```sh
cd /Users/wassimbensalem/Desktop/CarthAgentChallenge
git status --short --branch
uv sync --locked

# Local extraction, no model calls by default:
uv run carthagent run /path/to/input --output /path/to/output

# Explicit configured model/OCR run:
uv run --env-file .env carthagent run /path/to/input \
  --output /path/to/new-output --config /path/to/config.toml \
  --template hackathon-local/templates/Arbeitsversion-merged-layout-20.xlsx

# Separate evaluator; never feed keys to runtime:
uv run python tests/acceptance/evaluate_corpus.py \
  hackathon-local/ElabHackathon/A3-Hackathon-Corpus \
  hackathon-local/corpus-full-v6 --report /tmp/carthagent-evaluation.json

# Artifact/ledger audit requires a completed source/output set:
uv run python tests/qa/audit_corpus_artifacts.py \
  hackathon-local/presentation-subset-v7 hackathon-local/presentation-quality-v7
```

Use explicit staging allowlists and scan for secrets/private files before
commit. Normal pushes to `origin/codex/robust-extraction` are authorized; do not
force-push or merge. The benchmark runner intentionally never reads keys.

## Presentation files already prepared

Directory:
`/Users/wassimbensalem/Documents/Codex/2026-10-18/thi/outputs/carthagent-presentation-checkpoint/`

- `status.md` and `results.csv`: completed eight-document result.
- `working-example/review.v1.xlsx`: clean native Excel demonstration.
- Matching evidence, manifest and completion marker are alongside it.

That example has 415 proxy-correct accepted values out of 444 expected slots;
330 values bind to workbook destinations and 85 remain contract-pending.
Present it as a limited partial example, not proof of broad performance.

## Recommended first actions

1. Inspect Git and running jobs; confirm no other agent is editing.
2. Read the fresh review and implement region-scoped evidence plus bounded PDF
   vision, maintaining native extraction and strict source validation.
3. Independently test the payload, limits, cache behavior and foreign-region
   rejection; run a small live comparison on the same documents.
4. Fix the remaining review findings and the native-Excel zero-output causes.
5. Re-run appropriate tests, push the coherent checkpoint, and report actual
   accuracy/coverage separately. Preserve unresolved failures honestly.
