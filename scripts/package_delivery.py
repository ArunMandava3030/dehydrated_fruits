"""Validate and package only the current canonical prospect delivery."""
from __future__ import annotations

import ast
import csv
import hashlib
import json
import posixpath
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
REVIEW_DATE = "2026-10-24"
QUALIFIED = [
    "qualified_local_revision.json", "qualified_hyderabad_revision.json",
    "qualified_south_revision.json", "qualified_regional_revision.json",
]
DECISIONS = [
    "local_revision_decisions.json", "hyderabad_revision_decisions.json",
    "regional_revision_decisions.json",
]
CANONICAL_OUTPUTS = [
    "dehydrated_fruit_b2b_leads.xlsx", "dehydrated_fruit_leads.csv",
    "leads.json", "summary.json", "validation_report.json",
]
CURRENT_SCRIPTS = ["build_workbook.py", "package_delivery.py", "publish_github.py"]
SHEETS = [
    "StartHere", "MasterLeads", "FirstCalls", "Analytics", "ApproachPlans",
    "Changes", "FruitPlan", "BusinessFit", "EmailDrafts", "ContactLog",
    "PilotCosting", "Sources",
]


def read_json(path):
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(value, dict):
        for key in ("leads", "records", "decisions"):
            if isinstance(value.get(key), list):
                return value[key]
    return value


def text(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(text(v) for v in value)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return str(value)


def csv_value(value):
    result = text(value)
    return "'" + result if result.startswith(("=", "+", "-", "@")) else result


def present(value):
    return text(value).strip().casefold() not in {"", "not publicly found", "not found", "unknown", "n/a", "none"}


def excel_col(reference):
    letters = re.match(r"[A-Z]+", reference).group()
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter) - 64
    return value - 1


def worksheet_values(archive, name, shared):
    root = ET.fromstring(archive.read(name))
    result = {}
    for row in root.findall("s:sheetData/s:row", NS):
        values = {}
        for cell in row.findall("s:c", NS):
            kind = cell.get("t")
            raw = cell.find("s:v", NS)
            if kind == "s":
                val = shared[int(raw.text)] if raw is not None else ""
            elif kind == "inlineStr":
                val = "".join(node.text or "" for node in cell.findall(".//s:t", NS))
            else:
                val = raw.text or "" if raw is not None else ""
            values[excel_col(cell.get("r"))] = val
        result[int(row.get("r"))] = values
    return root, result


def table_rows(values, first_data_row, count, columns):
    return [[values.get(i, {}).get(c, "") for c in range(columns)]
            for i in range(first_data_row, first_data_row + count)]


def validate():
    baseline_path = ROOT / "research" / "original_50_baseline.json"
    baseline = read_json(baseline_path)
    leads = read_json(OUT / "leads.json")
    summary = read_json(OUT / "summary.json")
    assert len(baseline) == 50
    assert [r["lead_id"] for r in leads] == [f"L{i:03}" for i in range(1, len(leads) + 1)]
    assert summary["lead_count"] == len(leads)
    assert summary["original_baseline_count"] == 50
    assert summary["qualified_additions_count"] == len(leads) - 50
    assert summary["baseline_sha256"] == hashlib.sha256(baseline_path.read_bytes()).hexdigest()
    ids = {r["lead_id"]: r for r in leads}
    preserved_fields = 0
    for original in baseline:
        current = ids[original["lead_id"]]
        for key, value in original.items():
            assert current.get(key) == value, ("Baseline field changed", original["lead_id"], key)
            preserved_fields += 1
        assert current["cohort"] == "Original 50 - preserved baseline"
        assert current["qualification_review_date"] == ""
        assert current["qualification_format_points"] == 0
    required = [
        "commercial_format", "format_evidence", "format_source_url", "product_evidence",
        "product_source_url", "contact_access", "sales_friction", "qualification_gap",
        "review_date", "opportunity_type",
    ]
    for row in leads[50:]:
        assert all(present(row.get(key)) for key in required), ("Missing addition evidence", row["lead_id"])
        assert row["cohort"] == "Qualified addition - format and product reviewed"
        assert row["qualification_review_date"] == row["review_date"]
        assert date.fromisoformat(row["review_date"]) <= date.fromisoformat(REVIEW_DATE)
        assert row["contact_access"] in {"Official page reviewed", "Directory page reviewed", "Search index only", "Mixed"}
        for key in ("format_source_url", "product_source_url"):
            assert row[key].startswith(("https://", "http://"))
        assert row.get("contact_source_url") or row.get("contact_source_urls")
        route_text = text(row.get("contact_notes")) + " " + text(row.get("recommended_first_contact"))
        assert (present(row.get("phone")) or present(row.get("email")) or
                row.get("contact_form_url") or "form" in route_text.casefold())
        assert row["qualification_demand_status"] == "Unconfirmed; no buyer contacted"
    for row in leads:
        if row.get("verified_on"):
            assert date.fromisoformat(row["verified_on"]) <= date.fromisoformat(REVIEW_DATE)
        assert row["qualification_outreach_status"] == "Not contacted"
        assert row["qualification_duplicate_review"]
        points = (row["qualification_product_points"] * 5 +
                  row["qualification_format_points"] * 4 +
                  row["qualification_contact_points"] * 3 +
                  row["qualification_geography_points"] * 2)
        if row["cohort"] == "Original 50 - preserved baseline":
            assert row["qualification_score"] == "Not scored - preserved baseline"
        else:
            assert row["qualification_score"] == round(100 * points / 51)
            assert 0 <= row["qualification_score"] <= 100
        assert row["qualification_next_action"] and row["qualification_question_to_ask"]
    identities = [(re.sub(r"[^a-z0-9]", "", r["business"].casefold()),
                   re.sub(r"[^a-z0-9]", "", r["city"].casefold())) for r in leads]
    assert len(identities) == len(set(identities)), "Exact duplicate business/city"
    decisions = []
    for filename in DECISIONS:
        decisions.extend(read_json(ROOT / "research" / filename))
    def prior_id(row):
        return row.get("existing_lead_id") or row.get("lead_id") or row.get("prior_expansion_id")
    assert len(decisions) == 159
    assert {prior_id(d) for d in decisions} == {f"N{i:03}" for i in range(1, 160)}
    assert all(d.get("decision") in {"include", "exclude"} and present(d.get("reason")) for d in decisions)
    retained_ids = {prior_id(d) for d in decisions if d["decision"] == "include"}
    addition_prior_ids = [r["prior_expansion_id"] for r in leads[50:] if r.get("prior_expansion_id")]
    assert set(addition_prior_ids) == retained_ids
    assert len(addition_prior_ids) == len(set(addition_prior_ids))
    assert summary["prior_expansion_retained_count"] == len(retained_ids)
    assert summary["prior_expansion_removed_count"] == 159 - len(retained_ids)
    assert summary["newly_researched_additions_count"] == sum(not r.get("prior_expansion_id") for r in leads[50:])
    assert summary["city_counts"] == dict(sorted(Counter(r["city"] for r in leads).items()))
    for output, field in [
        ("cohort_counts", "cohort"), ("business_type_counts", "qualification_business_type"),
        ("contact_access_counts", "qualification_contact_access"),
        ("lane_counts", "qualification_procurement_lane"),
    ]:
        assert summary[output] == dict(sorted(Counter(r[field] for r in leads).items()))
    assert summary["published_email_count"] == sum(present(r.get("email")) for r in leads)
    assert summary["published_phone_count"] == sum(present(r.get("phone")) for r in leads)
    assert summary["confirmed_buyers"] == 0 and summary["outreach_sent"] == 0

    columns = summary["csv_columns"]
    with (OUT / "dehydrated_fruit_leads.csv").open(encoding="utf-8-sig", newline="") as stream:
        csv_rows = list(csv.reader(stream))
    assert len(csv_rows) == len(leads) + 1
    assert csv_rows[0] == [c["label"] for c in columns]
    assert len(set(csv_rows[0])) == len(columns)
    for row, record in zip(csv_rows[1:], leads):
        assert row == [csv_value(record.get(c["key"], "")) for c in columns], record["lead_id"]

    xlsx_path = OUT / "dehydrated_fruit_b2b_leads.xlsx"
    with ZipFile(xlsx_path) as archive:
        assert archive.testzip() is None
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find("s:sheets", NS)
        assert [s.get("name") for s in sheets] == SHEETS
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels = {r.get("Id"): posixpath.normpath(posixpath.join("xl", r.get("Target").lstrip("/")))
                for r in relationships}
        sheet_paths = {s.get("name"): rels[s.get("{" + NS["r"] + "}id")] for s in sheets}
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(node.text or "" for node in item.iter("{" + NS["s"] + "}t"))
                      for item in strings]
        formulas = []
        tables = {}
        for path in archive.namelist():
            if path.startswith("xl/tables/table") and path.endswith(".xml"):
                node = ET.fromstring(archive.read(path))
                tables[node.get("name")] = node.get("ref")
        assert "MasterLeadsTable" in tables and "ChangesTable" in tables
        assert int(re.search(r"(\d+)$", tables["MasterLeadsTable"]).group()) == len(leads) + 5
        assert int(re.search(r"(\d+)$", tables["ChangesTable"]).group()) == 164
        all_values = {}
        for sheet_name, path in sheet_paths.items():
            root, values = worksheet_values(archive, path, shared)
            all_values[sheet_name] = values
            for cell in root.findall(".//s:c", NS):
                formula = cell.find("s:f", NS)
                if formula is not None:
                    formulas.append((sheet_name, cell.get("r"), formula.text))
        assert len(formulas) == 7
        assert all(name == "PilotCosting" for name, _, _ in formulas)
        assert [ref for _, ref, _ in formulas] == [f"B{i}" for i in range(20, 27)]
        master = all_values["MasterLeads"]
        assert [master[5].get(i, "") for i in range(len(columns))] == csv_rows[0]
        master_rows = table_rows(master, 6, len(leads), len(columns))
        for row, record in zip(master_rows, leads):
            assert row == [text(record.get(c["key"], "")) for c in columns], ("XLSX mismatch", record["lead_id"])
        assert not master.get(len(leads) + 6), "Extra master row"
        assert len({r[0] for r in master_rows}) == len(leads)
        first_count = summary["first_calls_count"]
        first_rows = table_rows(all_values["FirstCalls"], 6, first_count, 12)
        first_ids = [r[1] for r in first_rows]
        assert len(set(first_ids)) == first_count
        assert all(i in ids and ids[i]["qualification_first_call_eligible"] for i in first_ids)
        assert "L005" not in first_ids and "L006" not in first_ids
        assert [all_values["ApproachPlans"][i].get(0) for i in range(6, len(leads) + 6)] == list(ids)
        assert [all_values["ContactLog"][i].get(0) for i in range(6, len(leads) + 6)] == list(ids)
        change_rows = table_rows(all_values["Changes"], 6, 159, 8)
        assert {r[0] for r in change_rows} == {f"N{i:03}" for i in range(1, 160)}
        for row in change_rows:
            expected = next(d for d in decisions if prior_id(d) == row[0])
            assert row[3] == expected["decision"] and row[5] == expected["reason"]
            if row[3] == "include":
                assert row[4] in ids and ids[row[4]].get("prior_expansion_id") == row[0]
            else:
                assert row[4] == ""
        assert len([n for n in archive.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")]) >= 1
        source_count = summary["sources_count"]
        assert len([r for r in all_values["Sources"] if r >= 6]) == source_count
        source_rows = table_rows(all_values["Sources"], 6, source_count, 9)
        for row in source_rows:
            if row[1]:
                lead = ids[row[1]]
                expected_date = lead["qualification_review_date"] or lead.get("verified_on", "")
                assert row[6] == expected_date
        # Actual chart/city data are static research counts, not Excel formulas.
        expected_city = [
            [city, str(v["original"]), str(v["qualified_additions"]), str(v["original"] + v["qualified_additions"])]
            for city, v in summary["city_by_cohort"].items()
        ]
        city_start = int(re.search(r"\d+", tables["CityCohortTable"]).group()) + 1
        assert table_rows(all_values["Analytics"], city_start, len(expected_city), 4) == expected_city
    for filename in CURRENT_SCRIPTS:
        path = ROOT / "scripts" / filename
        assert path.is_file(), filename
        ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    report = {
        "status": "passed", "review_date": REVIEW_DATE,
        "lead_count": len(leads), "original_baseline_count": 50,
        "qualified_additions_count": len(leads) - 50,
        "preserved_original_field_values_checked": preserved_fields,
        "prior_expansion_decisions_checked": 159,
        "checks": [
            "Every existing field/value of original L001-L050 exactly matches preserved baseline",
            "Unique sequential master IDs; deterministic qualified addition IDs; prior expansion IDs retained",
            "Every addition has qualifying format/product evidence, source URLs, dated review and a public contact route",
            "All 159 expansion decisions accounted for; included prior IDs match retained additions",
            "Exact business/city duplicates rejected; shared routes flagged for manual buyer-account review",
            "CSV and every master XLSX value match the JSON records and column mapping",
            "FirstCalls uses eligible records and separates practical from formal procurement lanes",
            "ApproachPlans and ContactLog contain every master ID",
            "Actual cohort/city/type/contact counts checked; city chart has a static count table",
            "All source-sheet dates follow their corresponding record dates",
            "XLSX archive integrity, twelve expected worksheets and seven formulas confined to PilotCosting",
            "Current Python scripts parse successfully",
        ],
        "limits": [
            "Original 50 preserved, not all requalified under the revised format criteria",
            "Contacts not called or inbox-tested; indexed-only and mixed evidence remain labelled",
            "Buyer interest, independent purchasing authority, volumes and conversion are unconfirmed",
            "Research priority score is not a probability or sales forecast",
            "Automatic shared-route flags do not replace manual buying-account coordination",
            "Workbook not visually tested in desktop Excel",
            "GitHub publication status is reported separately",
        ],
    }
    (OUT / "validation_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def package(report):
    required = [ROOT / name for name in ("README.md", ".gitignore", "requirements.txt")]
    required += [ROOT / "research" / name for name in
                 ["original_50_baseline.json", "fruit_evidence.json"] + QUALIFIED + DECISIONS]
    required += [ROOT / "scripts" / name for name in CURRENT_SCRIPTS]
    required += [OUT / name for name in CANONICAL_OUTPUTS]
    required += sorted(path for path in (ROOT / "docs").rglob("*")
                       if path.is_file() and "archive" not in {part.casefold() for part in path.relative_to(ROOT).parts}
                       and "__pycache__" not in path.parts and path.suffix.lower() in {".md", ".txt", ".pdf"})
    assert all(path.is_file() for path in required), "A current required delivery file is missing"
    files = sorted(set(required), key=lambda p: p.relative_to(ROOT).as_posix())
    manifest = {path.relative_to(ROOT).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest() for path in files}
    delivery = ROOT / "delivery"
    delivery.mkdir(exist_ok=True)
    target = delivery / "dehydrated_fruit_business_pack.zip"
    with ZipFile(target, "w", compression=ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, path.relative_to(ROOT).as_posix())
        archive.writestr("sha256_manifest.json", json.dumps(manifest, indent=2) + "\n")
    with ZipFile(target) as archive:
        assert archive.testzip() is None
        names = archive.namelist()
        assert set(names) == set(manifest) | {"sha256_manifest.json"}
        assert [n for n in names if n.endswith(".xlsx")] == ["deliverables/dehydrated_fruit_b2b_leads.xlsx"]
        assert not any(n.startswith((".tools/", ".git/", ".codex/", "archive/", "delivery/")) for n in names)
        assert not any("additional_" in n or "__pycache__" in n or "expansion_outreach_playbook" in n for n in names)
        assert all(hashlib.sha256(archive.read(name)).hexdigest() == digest for name, digest in manifest.items())
    print(json.dumps({
        "validation": report["status"], "leads": report["lead_count"],
        "qualified_additions": report["qualified_additions_count"],
        "zip": str(target), "packaged_files": len(files) + 1, "bytes": target.stat().st_size,
    }, indent=2))


def main():
    package(validate())


if __name__ == "__main__":
    main()
