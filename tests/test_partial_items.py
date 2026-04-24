"""Unit tests for _extract_partial_items — the truncated-JSON rescue parser."""
import json
import pytest


@pytest.fixture
def fn():
    import app
    return app._extract_partial_items


def test_clean_json(fn):
    text = json.dumps({
        "company_name": "NOAH Marine",
        "items": [
            {"product_name": "Phi Phi | Adult", "net_rate": 1000, "selling_rate": 1500, "notes": ""},
            {"product_name": "Phi Phi | Child", "net_rate": 800, "selling_rate": 1200, "notes": ""},
        ]
    })
    items, company = fn(text)
    assert company == "NOAH Marine"
    assert len(items) == 2
    assert items[0]["product_name"] == "Phi Phi | Adult"


def test_truncated_mid_item(fn):
    # The final item is cut off — should recover the earlier complete ones
    text = (
        '{"company_name": "Charter X", "items": ['
        '{"product_name": "A", "net_rate": 1},'
        '{"product_name": "B", "net_rate": 2},'
        '{"product_name": "C'  # cut off
    )
    items, company = fn(text)
    assert company == "Charter X"
    assert [i["product_name"] for i in items] == ["A", "B"]


def test_escaped_quotes_inside_strings(fn):
    text = (
        '{"company_name": "Has \\"Quotes\\"", "items": ['
        '{"product_name": "X \\"Y\\" Z", "net_rate": 1},'
        '{"product_name": "P2", "net_rate": 2}'
        ']}'
    )
    items, company = fn(text)
    assert company == 'Has "Quotes"'
    assert items[0]["product_name"] == 'X "Y" Z'
    assert len(items) == 2


def test_no_items_returns_empty(fn):
    items, company = fn('garbage text with no json at all')
    assert items == []
    assert company == ""


def test_missing_product_name_is_skipped(fn):
    # Objects without product_name must not be included (guards against false positives
    # when the bracket walker hits some other nested object)
    text = (
        '{"company_name": "CO", "items": ['
        '{"unrelated": 1},'
        '{"product_name": "Real", "net_rate": 500}'
        ']}'
    )
    items, _ = fn(text)
    assert [i["product_name"] for i in items] == ["Real"]
