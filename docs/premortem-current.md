# Implementation premortem — 2026-10-18

This review concerns the in-progress implementation, not a release approval.
The failures below are concrete checks to close before final QA.
The incident headlines are hypothetical; their root causes identify observed
implementation gaps, not claims that an external security incident occurred.

The last completed full-corpus checkpoint is frozen build v5 (application
0.11.0), with remote model assistance disabled. Across 72 synthetic documents,
2,824 values were accepted: 2,031 passed the development proxy and 793 failed.
Accepted precision was 71.92%; correct coverage was 5.65% of 35,975 source slots.
Five of ten layout families yielded no facts. This is a failed checkpoint.
The proxy does not grade every account, class and entity identity axis; even
these results must not be described as complete tuple accuracy.

The same checkpoint had passed 70 unit, 39 acceptance and 19 adversarial tests.
The corpus failures show why fixture passes alone are insufficient. Subsequent
work addresses row-measure blocks, period and identity admission, bounded model
requests, and hybrid OCR. New independent fixtures exposed zero extraction for
two recoverable block layouts and unsafe acceptance without a reporting period.
The implementation is still changing; all release gates require a fresh source
fingerprint and rerun. No real-client validation or official scoring has occurred.

<premortem>
  <headlines>
    <headline category="data-corruption" probability="high">
      <text>Asset schedules exported section headings as zero-value assets and assigned depreciation disposals to acquisition-cost disposals.</text>
      <root_cause>structure.py hierarchical header resolution and semantics.py empty-zero admission initially treated any named row as a monetary record. The independent fiscal-year fixture and initial corpus run reproduced this.</root_cause>
      <matrix_row_missed>4: structure and context ambiguity</matrix_row_missed>
      <fix>Require positive record-role evidence, combine parent and leaf headers, separate account/class/asset axes, and test fiscal-year and local-unit overrides. Model proposals must satisfy the same evidence rules.</fix>
    </headline>
    <headline category="security" probability="medium">
      <text>A mistyped provider URL sent a development API credential over plaintext HTTP.</text>
      <root_cause>mapper.py attached Authorization whenever the hostname matched api.openai.com, without first requiring HTTPS. Default redirect behavior also lacked an explicit credential boundary.</root_cause>
      <matrix_row_missed>8: credential and network boundaries</matrix_row_missed>
      <fix>Require HTTPS for credential-bearing endpoints, disallow redirects, allow explicit loopback HTTP only for credential-free test adapters, and test that invalid endpoints make no network request. Keep provider errors sanitized.</fix>
    </headline>
    <headline category="performance" probability="high">
      <text>One large spreadsheet blocked every later client document despite a configured document timeout.</text>
      <root_cause>pipeline.py checked elapsed time only after acquisition and interpretation. structure.py repeatedly rebuilt document indexes inside each value's resolution. These checks cannot interrupt a blocked parser or unbounded CPU loop.</root_cause>
      <matrix_row_missed>9: resource exhaustion and batch isolation</matrix_row_missed>
      <fix>Execute documents in bounded workers with enforceable deadlines, index observations once per document or region, validate resource settings, and demonstrate that a timed-out document does not stop the next one.</fix>
    </headline>
  </headlines>
  <verdict>REVISE</verdict>
  <revise_actions>Keep the regression evidence for the three issues below. Also close the observed block ownership, identity, period and coverage failures, complete hybrid OCR integration, then repeat the full corpus and independent QA against the final source revision.</revise_actions>
</premortem>

Additional closure checks: acquisition must consume the preserved source bytes,
not re-read a mutable input after preservation; cache keys must isolate document
occurrences and relevant configuration; provenance invalidation must traverse
all dependent facts and output bindings; resuming must verify actual artifact
hashes; model assistance must not erase independently supported mappings merely
because its target-ID shape differs from the local validator's expectation.
