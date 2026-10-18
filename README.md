# CarthAgent

Extract fixed-asset schedules from PDF and Excel into a source-backed evidence
ledger and a workbook. The application preserves uncertainty: a value must have
an evidenced row, field, context, and source pointer before it becomes a fact.

**Development status:** validation is in progress. Complex table mappings still
have known failures; this checkout is not ready for unattended client use or
submission. Passing the small fixtures does not establish corpus accuracy.
See `docs/premortem-current.md` for the last completed failed checkpoint and the
remaining release gates.
See `docs/recovery-checkpoint.md` for the latest handoff and known failures.

The supplied Arbeitsversion **layout 20** is supported through a separate
template adapter. The official per-value JSON schema and leaderboard scorer are
not included in this repository. Evidence JSON and local evaluation results
must not be described as official submissions or leaderboard scores.

## Install

Requires Python 3.11+ and [uv](https://docs.astral.sh/uv/).

```sh
uv sync --locked
```

For scanned or outline PDFs, install Tesseract with German and English language
data. Native PDF and Excel processing do not require a remote model provider.
The default configuration makes no model-service requests.

## Enable OpenAI assistance

Store `OPENAI_API_KEY` in a local `.env` file, excluded from Git, and restrict its
permissions with `chmod 600 .env`. Keep the value out of chat, logs, and config
files. Create a TOML configuration such as `hackathon-local/openai-quality.toml`:

```toml
[mapper]
enabled = true
model = "gpt-6-astra"
reasoning_effort = "max"
timeout_seconds = 300
max_output_tokens = 32768

[bounds]
document_timeout_seconds = 600
transient_retries = 1
```

Then run with credentials loaded into the process environment:

```sh
uv run --env-file .env carthagent run /path/to/documents \
  --output /path/to/results \
  --config hackathon-local/openai-quality.toml \
  --template /path/to/Arbeitsversion-merged-layout-20.xlsx
```

Model assistance proposes compact mappings for unresolved regions. Local code
still validates evidence and parses amounts. This sends document observations
to OpenAI; use the default local configuration when remote processing is not
authorized. An authentication or model-access failure is reported in the result;
it does not silently select a cheaper model.

See `docs/model-and-ocr-selection.md` for the model choice, OCR options, and their
validation limits. Tesseract is the local OCR baseline. An optional Azure
Document Intelligence adapter requires a separate Azure resource and key;
an OpenAI API key does not enable it.

## Run

```sh
uv run carthagent run /path/to/documents --output /path/to/results
```

Use the private template when available:

```sh
uv run carthagent run /path/to/documents \
  --output /path/to/results \
  --template /path/to/Arbeitsversion-merged-layout-20.xlsx
```

The input can be a single file or a folder. The CLI emits one JSON result per
document. A partial result contains useful accepted facts plus explicit unresolved
work. A failed or unsupported input is reported; it is not silently discarded.
The folder runner excludes answer keys and generator directories.

Each successfully published result directory contains:

- `evidence.v1.json`: source observations, accepted facts, abstentions, coverage,
  dependencies, output bindings, and the frozen snapshot identity.
- `review.v1.xlsx`: the supplied template when selected, otherwise a clearly
  provisional review workbook.
- `manifest.v1.json`: artifact hashes, snapshot identity, status, and readback
  verification.
- `.complete`: publication marker for the matched artifact pair.

Consumers must require the completion marker and verify the manifest. A file's
existence alone does not establish successful publication. The output directory
also contains a local `.carthagent` ledger for processing state and provenance.

## Template behavior

The actual layout-20 targets are `Anlagenverzeichnis`,
`Gitter nach Anlagenklasse`, and `Gitter nach Bilanzposition`. The workbook's
field catalogs define stable field IDs. Its other audit worksheets are preserved;
their presence does not expand this extractor into an invoice or general-ledger
application.

Source amounts remain distinct from target representations. For example, a SAP
credit may print with a trailing minus while the template requires a positive
depreciation magnitude. Any conversion belongs to the output binding. A printed
source closing balance takes precedence over a template formula; a formula must
never replace an unreadable or ambiguously mapped source value.

See `docs/template-contract.md` for field mappings, row limits, source/derived
value separation, and unresolved contract details.

## Verify

```sh
uv run pytest tests/unit tests/acceptance
```

Enable the private-template acceptance check locally:

```sh
CARTHAGENT_TEST_TEMPLATE=/path/to/Arbeitsversion-merged-layout-20.xlsx \
  uv run pytest tests/acceptance
```

Independent fixture builders are under `tests/fixtures`; no private challenge
files are required for those tests. The private-template test is skipped when
its environment variable is absent. Skipping it is not a template-validation
pass.

For development-only corpus evaluation, first run extraction on the document
folder, then evaluate the published artifacts separately:

```sh
uv run python tests/acceptance/evaluate_corpus.py \
  /path/to/A3-Hackathon-Corpus /path/to/results \
  --report /path/to/development-evaluation.json
```

This evaluator reads the supplied keys outside runtime extraction. It compares
signed printed values, fields, source locations, and available identities.
It is a local proxy: the supplied PDF keys contain row boxes rather than exact
value boxes, and the official scorer is still unavailable. Synthetic development
performance does not establish accuracy on hidden real-client documents.

## Local data and project boundaries

Keep challenge inputs, templates, run outputs, and evaluation reports under
`hackathon-local/` for this checkout. That directory is excluded from Git.
Do not commit client documents, credentials, generated evidence, or private
templates. Remote model use must be configured explicitly and must comply with
the data-handling rules for the documents being processed.

`docs/architecture.md` records the approved architecture. Implementation checks
and measured limitations are reported separately; design approval is not a test
result.
