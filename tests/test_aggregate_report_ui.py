from pathlib import Path

from flask import Flask, render_template


def test_summary_settings_offer_project_and_date_range_fields():
    app = Flask(__name__, template_folder=str(Path(__file__).resolve().parents[1] / "templates"))
    with app.test_request_context("/"):
        html = render_template("index.html")

    assert 'id="project-name"' in html
    assert 'id="report-from-date"' in html
    assert 'id="report-to-date"' in html
    assert 'id="export-month"' not in html


def test_summary_tab_has_aggregate_report_list_and_auto_refresh_hook():
    project_root = Path(__file__).resolve().parents[1]
    app_js = (project_root / "static" / "js" / "app.js").read_text(encoding="utf-8")
    template = (project_root / "templates" / "index.html").read_text(encoding="utf-8")

    assert 'id="aggregate-reports-table-body"' in template
    assert "File Word Giải Trình Tổng Hợp" in template
    assert "async function loadAggregateReports" in app_js
    assert "apiGet('/api/aggregate-reports')" in app_js
    assert "loadAggregateReports();" in app_js
    assert app_js.count("navigateToResults('summary')") >= 2
