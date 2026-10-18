# Model and OCR selection

Decision date: 2026-10-18. The user prioritizes accuracy over model cost.

## OpenAI reasoning and vision

The quality-first target is `gpt-6-astra` with `reasoning.effort=max`, using the
Responses API and strict structured outputs. OpenAI currently documents Astra
as its most capable model and lists image input and structured outputs as
supported. Account access and end-to-end extraction quality require live tests;
model catalog availability is not proof of either.

Model choice must be explicit in the run configuration and provenance. Do not
silently downgrade on an access error. A provider failure must remain visible in
the result. Keep requests bounded even when cost is not the limiting factor.

The model proposes relationships using supplied observation IDs. It must not
invent amounts, coordinates, source evidence, or certainty. Local code validates
those proposals against the source structure and parses monetary lexemes.
The current mapper receives text and structural/geometry observations. Direct
page-image input to the reasoning model is not implemented at this checkpoint.
It remains a possible escalation for unresolved layouts; model-generated
coordinates would not be accepted as source pointers.

Sources:
- https://developers.openai.com/api/docs/models/gpt-6-astra
- https://developers.openai.com/api/docs/guides/latest-model
- https://developers.openai.com/api/docs/guides/structured-outputs
- https://developers.openai.com/api/docs/guides/images-vision

## OCR selection

Native Excel cells and readable PDF text remain the preferred acquisition
channels. OCR is needed for raster content, including scanned regions on pages
that also contain native text. A page-level native-word count cannot establish
that all relevant regions have been read.

Tesseract with German and English language data is the locally available
baseline. Its word boxes and confidence scores support source traceability.
This is a reproducible baseline, not evidence that it is the most accurate OCR
engine for the hidden client documents.

Azure Document Intelligence `prebuilt-layout` is an additional cloud candidate
because its documented response contains word polygons and confidence,
plus table structure. It requires a separate Azure resource and credentials.
Its superiority on this workload is unverified until the same difficult pages
are evaluated against independent transcription and coordinate checks.

Mistral OCR is the available cloud candidate being integrated. Live calls using
`mistral-ocr-latest`, HTML tables, block output and word confidence succeeded on
four synthetic source documents (19 pages). The adapter also passes replay of
those actual responses. Its block boxes and word confidence do not establish
exact per-value source coordinates. Tables remain interpretation evidence;
native text or separately validated OCR geometry supplies precise pointers.

In one scanned sample Mistral recognized all 30 asset IDs, correcting five
Tesseract ID errors. This does not establish complete extraction: its HTML
omitted 19 of 220 printed monetary occurrences on that same document. A separate
scanned block-layout document contained OCR number substitutions. These findings
support native-first acquisition and preserving provider disagreements. They
do not establish Mistral as the best engine for real client documents.

Compare candidates on monetary character accuracy, sign/decimal errors, missing
regions, row/column association, rotation/crop correctness, and false acceptance.
A provider's overall document benchmark is not a substitute for these tests.
Provider disagreement must be retained and resolved or abstained from rather
than hidden by majority voting.

Sources:
- https://tesseract-ocr.github.io/tessdoc/Command-Line-Usage.html
- https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/prebuilt/layout?view=doc-intel-4.0.0
- https://docs.mistral.ai/studio/document-processing/basic_ocr

### Optional Azure configuration

The adapter is implemented but has not been validated against a live Azure
resource in this project. Its protocol and failure cases have local tests;
that is not a comparative OCR accuracy result.

To configure a resource, place its key in the environment variable
`AZURE_DOCUMENT_INTELLIGENCE_KEY`, then add this section to the run TOML. Replace
the example endpoint with the endpoint of that resource:

```toml
[ocr]
enabled = true
provider = "azure_document_intelligence"

[ocr.azure]
endpoint = "https://your-resource.cognitiveservices.azure.com"
api_version = "2024-11-30"
key_env = "AZURE_DOCUMENT_INTELLIGENCE_KEY"
timeout_seconds = 180
transient_retries = 1
retry_delay_seconds = 1
max_input_bytes = 209715200
max_response_bytes = 52428800
poll_interval_seconds = 1
```

This enables remote processing of relevant PDF content. Validate the data-handling
arrangement and run the same difficult-page checks before selecting it for client
documents. To return to local OCR, omit these sections or select
`provider = "tesseract"` without an Azure subsection.

## Credentials

Use environment variables loaded from the private, Git-ignored `.env` file for
local development. Keep file permissions at `0600`. Never include secrets in
configuration examples, prompts, test fixtures, error output, or published
artifacts. The app should report authentication status without echoing provider
error messages that may contain a key fragment.

A promotional redemption code is not an API credential. Redeem the credits in
the intended API organization, then create a project API key in that organization.
