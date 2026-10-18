import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import build_dashboard as bd  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from make_images import load_summary  # noqa: E402


@pytest.fixture(scope="module")
def orders():
    return bd.load_orders()


@pytest.fixture(scope="module")
def page(orders):
    return bd.build(orders)


def embedded_rows(page):
    match = re.search(r"const ROWS = (\[.*?\]);\n", page, re.S)
    return json.loads(match.group(1))


def test_every_order_line_is_embedded(orders, page):
    rows = embedded_rows(page)
    assert len(rows) == len(orders) == 797
    assert {"id", "date", "month", "product", "qty", "total", "status", "state"} <= set(rows[0])


def test_net_sales_match_the_sales_report(page):
    rows = embedded_rows(page)
    net = sum(r["total"] for r in rows if r["status"] in bd.SOLD)
    report = load_summary(bd.SOURCE)
    assert net == pytest.approx(report["Net sales (shipped and delivered)"], abs=0.01)


def test_period_and_placeholders(page):
    assert "April to September 2026" in page
    assert "__DATA__" not in page and "__PERIOD__" not in page and "__SOLD__" not in page


def test_chart_library_is_pinned(page):
    assert "chart.js@4.4.1" in page
