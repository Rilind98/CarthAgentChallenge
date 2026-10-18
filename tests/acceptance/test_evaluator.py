"""Independent checks for the development oracle's location matching."""
from evaluate_corpus import matching_candidates


def test_unique_original_row_containment_resolves_overlapping_tolerances():
    fact = {"field": "bw_prev", "pointer": {"page_index": 0, "bbox": [700, 134.64, 710, 144]}}
    expected = [
        {"field": "bw_prev", "pointer": {"page": 1, "bbox": [30, 132.7, 810, 140]}},
        {"field": "bw_prev", "pointer": {"page": 1, "bbox": [30, 142.1, 810, 149.5]}},
    ]
    assert [index for index, _ in matching_candidates(fact, expected)] == [0]


def test_genuinely_overlapping_original_rows_remain_ambiguous():
    fact = {"field": "bw_prev", "pointer": {"page_index": 0, "bbox": [700, 134, 710, 144]}}
    expected = [
        {"field": "bw_prev", "pointer": {"page": 1, "bbox": [30, 132, 810, 140]}},
        {"field": "bw_prev", "pointer": {"page": 1, "bbox": [30, 138, 810, 146]}},
    ]
    assert len(matching_candidates(fact, expected)) == 2
