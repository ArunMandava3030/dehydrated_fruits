"""Build one canonical prospect workbook; preserve the original 50 verbatim in JSON."""
from __future__ import annotations

import copy
import csv
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".tools" / "python"))
import xlsxwriter

OUT = ROOT / "deliverables"
REVIEW_DATE = "2026-10-24"
BASELINE = "original_50_baseline.json"
QUALIFIED = [
    "qualified_local_revision.json", "qualified_hyderabad_revision.json",
    "qualified_south_revision.json", "qualified_regional_revision.json",
]
DECISIONS = [
    "local_revision_decisions.json", "hyderabad_revision_decisions.json",
    "regional_revision_decisions.json",
]
REQUIRED = [
    "business", "city", "commercial_format", "format_evidence", "format_source_url",
    "product_evidence", "product_source_url", "review_date", "contact_access",
    "sales_friction", "qualification_gap", "opportunity_type",
]
EXPECTED_SHEETS = [
    "StartHere", "MasterLeads", "FirstCalls", "Analytics", "ApproachPlans",
    "Changes", "FruitPlan", "BusinessFit", "EmailDrafts", "ContactLog",
    "PilotCosting", "Sources",
]
LOCAL_CITIES = {
    "poranki", "vijayawada", "gannavaram", "guntur", "mangalagiri", "tadepalli",
    "tadigadapa", "kanuru", "penamaluru",
}
AP_CITIES = LOCAL_CITIES | {
    "visakhapatnam", "vizag", "kakinada", "rajahmundry", "rajamahendravaram",
    "tirupati", "nellore", "ongole", "kurnool", "anantapur", "bhimavaram",
    "machilipatnam", "gudivada", "nuzvid", "tanuku", "tadepalligudem",
    "palakollu", "tenali", "vijayanagaram", "vizianagaram",
}
MISSING = {"", "not publicly found", "not found", "unknown", "n/a", "none"}
COHORT_ORIGINAL = "Original 50 - preserved baseline"
COHORT_NEW = "Qualified addition - format and product reviewed"


def read_json(name):
    value = json.loads((ROOT / "research" / name).read_text(encoding="utf-8-sig"))
    if isinstance(value, dict):
        for key in ("leads", "records", "decisions"):
            if isinstance(value.get(key), list):
                return value[key]
    return value


def as_text(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(as_text(v) for v in value)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return str(value)


def csv_safe(value):
    text = as_text(value)
    return "'" + text if text.startswith(("=", "+", "-", "@")) else text


def present(value):
    return as_text(value).strip().casefold() not in MISSING


def valid_url(value):
    return isinstance(value, str) and value.startswith(("https://", "http://"))


def urls(value):
    if isinstance(value, list):
        return [v for v in value if valid_url(v)]
    if valid_url(value):
        return [value]
    return []


def identity(row):
    return (
        re.sub(r"[^a-z0-9]", "", row["business"].casefold()),
        re.sub(r"[^a-z0-9]", "", row["city"].casefold()),
    )


def city_key(row):
    return row.get("city", "").casefold().strip()


def geography(row):
    city = city_key(row)
    tier = row.get("geographic_tier", "").casefold()
    if tier.startswith("1 ") or city in {"poranki", "kanuru", "tadigadapa", "penamaluru"}:
        return 1, 5, "Poranki belt"
    if city == "vijayawada":
        return 2, 4, "Vijayawada"
    if city in LOCAL_CITIES:
        return 3, 4, "Nearby Andhra cities"
    if city == "hyderabad":
        return 4, 3, "Hyderabad"
    if city in AP_CITIES or "andhra" in row.get("state", "").casefold():
        return 5, 3, "Other Andhra Pradesh"
    return 6, 2, "Wider southern India"


def state_for(row):
    city = city_key(row)
    if city in AP_CITIES:
        return "Andhra Pradesh"
    mapping = {
        "hyderabad": "Telangana", "bengaluru": "Karnataka", "bangalore": "Karnataka",
        "mysuru": "Karnataka", "mangalore": "Karnataka", "mangaluru": "Karnataka",
        "chennai": "Tamil Nadu", "coimbatore": "Tamil Nadu", "madurai": "Tamil Nadu",
        "kanchipuram": "Tamil Nadu", "tiruchirappalli": "Tamil Nadu",
        "tirunelveli": "Tamil Nadu", "salem": "Tamil Nadu", "vellore": "Tamil Nadu",
        "kochi": "Kerala", "cochin": "Kerala", "thiruvananthapuram": "Kerala",
        "trivandrum": "Kerala", "kozhikode": "Kerala", "thrissur": "Kerala",
        "puducherry": "Puducherry", "pondicherry": "Puducherry",
    }
    return mapping.get(city, row.get("state", ""))


def access_for(row, original=False):
    raw = row.get("contact_access", "")
    if raw in {"Official page reviewed", "Directory page reviewed", "Search index only", "Mixed"}:
        return raw
    source = as_text(row.get("source_access", "")).casefold()
    if "index only" in source:
        return "Search index only"
    if "mixed" in source:
        return "Mixed"
    confidence = row.get("contact_confidence", "").casefold()
    if original:
        return "Baseline source classification; not re-reviewed"
    if "official" in confidence and "reviewed" in source:
        return "Official page reviewed"
    if "directory" in confidence and "reviewed" in source:
        return "Directory page reviewed"
    return "Access not specified; confirm"


def contact_points(row, original):
    access = access_for(row, original)
    if access == "Official page reviewed":
        return 3
    if access in {"Directory page reviewed", "Mixed"}:
        return 2
    if access == "Search index only":
        return 1
    if original:
        return 2 if "official" in row.get("contact_confidence", "").casefold() else 1
    return 0


def product_points(row, original):
    evidence = as_text(row.get("evidence_note" if original else "product_evidence", "")).casefold()
    explicit = " ".join(as_text(row.get(k, "")) for k in ("evidence_strength", "fit_signal")).casefold()
    if explicit.startswith("explicit dehydrated") or row.get("fit_signal", "").casefold() == "explicit dried garnish":
        return 4, "Explicit dried-fruit menu evidence"
    named_fruit = re.search(
        r"\b(orange|lemon|lemonade|lime|pineapple|mango|guava|cranberry|raspberry|"
        r"strawberry|berry|berries|coconut|grapefruit|passionfruit|fruitcake|plum|"
        r"apple|apricot|peach|cherry|mandarin|citrus|jamun|muskmelon|watermelon|kokum)\b",
        evidence,
    )
    if named_fruit:
        return 3, "Relevant published fruit ingredient or named fruit recipe; dried use unconfirmed"
    if any(term in evidence for term in
           ("cocktail", "mocktail", "mojito", "beverage", "bakery", "dessert", "pastry",
            "patisserie", "coffee", "cake", "bar", "cold brew", "mixed drink",
            "daiquiri", "negroni", "colada", "margarita", "highball", "martini", "spiked tea")):
        return 2, "Relevant programme or product category; exact application needs agreement"
    if row.get("fit_signal", "").casefold() in {"cocktail or mocktail menu", "cafe beverages", "bakery or dessert menu"}:
        return 2, "Relevant programme or product category; exact application needs agreement"
    if "garnish" in evidence or "training" in evidence:
        return 1, "Professional-service or garnish cue; purchasing application remains uncertain"
    if original or "segment" in evidence:
        return 1, "Segment or general offering only"
    return 0, "Insufficient product signal"


def business_type(row):
    text = (row.get("segment", "") + " " + row.get("commercial_format", "")).casefold()
    if any(x in text for x in ("hotel", "resort")):
        return "Hotel / resort"
    if any(x in text for x in ("restobar", "brewery", "cocktail bar", "restro bar", "lounge", "bar")):
        return "Bar / brewery / restobar"
    if any(x in text for x in ("bakery", "baking", "patisserie", "pasticceria", "production")):
        return "Bakery / patisserie"
    if any(x in text for x in ("cafe", "café", "coffee")):
        return "Cafe / coffee"
    return "Restaurant / other original venue"


def derive(row, original, source_file):
    # Original values must never be normalised or replaced.
    row["cohort"] = COHORT_ORIGINAL if original else COHORT_NEW
    row["qualification_status"] = (
        "Preserved original record; not all requalified in this revision" if original
        else "Commercial format and relevant product evidence reviewed; demand unconfirmed"
    )
    row["qualification_source_file"] = source_file
    row["qualification_review_date"] = "" if original else row["review_date"]
    row["qualification_contact_access"] = access_for(row, original)
    row["qualification_business_type"] = business_type(row)
    order, geo_points, geo_label = geography(row)
    product, product_label = product_points(row, original)
    format_score = 0 if original else (2 if row.get("source_access", "").casefold() == "search index only" else 3)
    contact_score = contact_points(row, original)
    row["qualification_product_points"] = product
    row["qualification_product_basis"] = product_label
    row["qualification_format_points"] = format_score
    row["qualification_format_basis"] = (
        "Not re-reviewed; zero denotes missing revision assessment, not an unsuitable business"
        if original else "Qualifying commercial format documented; does not establish revenue or buying volume"
    )
    row["qualification_contact_points"] = contact_score
    row["qualification_geography_points"] = geo_points
    row["qualification_geography_band"] = geo_label
    row["qualification_score"] = (
        "Not scored - preserved baseline" if original else
        round(100 * (product * 5 + format_score * 4 + contact_score * 3 + geo_points * 2) / 51)
    )
    row["qualification_priority"] = (
        "A - Local application discussion" if geo_points >= 4 and product >= 2
        else "B - Remote application qualification" if product >= 2
        else "C - Confirm product application first"
    )
    formal = row["qualification_business_type"] == "Hotel / resort"
    row["qualification_procurement_lane"] = (
        ("Local formal hotel / group route" if formal else "Local practical introduction")
        if geo_points >= 4 else
        ("Remote formal hotel / group route" if formal else "Remote qualification before samples")
    )
    row["qualification_procurement_friction"] = row.get("sales_friction") or (
        "Ask whether F&B approval and vendor registration are required; actual purchasing process unknown."
        if formal else
        "Buyer authority, current supplier and interest in outsourcing remain unknown; qualify before samples."
    )
    row["qualification_product_caveat"] = (
        "Suggested fruits are proposals, not confirmed demand. Dried garnish does not replace fresh juice or puree."
    )
    row["qualification_contact_status"] = "Public routing contact; not called, inbox-tested or confirmed as a buyer"
    row["qualification_local_authority"] = row.get("local_purchase_authority") or "Unknown; ask who approves ingredients"
    row["qualification_question_to_ask"] = row.get("qualification_question") or (
        "Which current drink or recipe could use these fruits; who approves it, do you buy or make garnishes, "
        "and what specification, weekly usage and delivered-price limit would a paid trial need to meet?"
    )
    row["qualification_next_action"] = row.get("recommended_first_contact") or (
        f"Contact the published business route and ask for {row.get('buyer_role', 'the beverage or kitchen buyer')}. "
        "Confirm one suitable menu application and approval process before agreeing a sample."
    )
    row["qualification_trial_offer"] = row.get("sample_offer") or (
        "Agree a small trial for one buyer-selected drink or recipe, using the suggested fruits only after "
        "the buyer confirms a useful application. Agree sample quantity, charges, delivery and acceptance criteria."
    )
    row["qualification_state"] = row.get("state") or state_for(row)
    row["qualification_group_note"] = (
        row.get("group_coordination") or
        "Rows represent venue routes, not proven independent purchasing accounts; coordinate any shared group before outreach."
    )
    row["qualification_outreach_status"] = "Not contacted"
    row["qualification_demand_status"] = "Unconfirmed; no buyer contacted"
    baseline_route = (row.get("segment", "") + " " + row.get("business", "")).casefold()
    row["qualification_first_call_eligible"] = product >= 2 and (
        not original or any(term in baseline_route for term in ("brewery", "bar", "hotel", "coffee roasters"))
    )


def decision_id(row):
    return row.get("existing_lead_id") or row.get("lead_id") or row.get("prior_expansion_id")


def collect():
    missing = [f for f in [BASELINE] + QUALIFIED + DECISIONS if not (ROOT / "research" / f).is_file()]
    if missing:
        raise ValueError("Required inputs not ready: " + ", ".join(missing))
    baseline = read_json(BASELINE)
    assert len(baseline) == 50, "Original baseline must contain exactly 50 records"
    assert {r["lead_id"] for r in baseline} == {f"L{i:03}" for i in range(1, 51)}
    decisions = []
    for filename in DECISIONS:
        for row in read_json(filename):
            item = copy.deepcopy(row)
            item["existing_lead_id"] = decision_id(row)
            item["decision_source_file"] = filename
            decisions.append(item)
    assert len(decisions) == 159, f"All 159 prior expansion decisions required; got {len(decisions)}"
    assert {d["existing_lead_id"] for d in decisions} == {f"N{i:03}" for i in range(1, 160)}
    assert all(d.get("decision") in {"include", "exclude"} and present(d.get("reason")) for d in decisions)
    additions = []
    for filename in QUALIFIED:
        for raw in read_json(filename):
            assert raw.get("decision", "include") == "include", (filename, raw.get("business"))
            assert all(present(raw.get(k)) for k in REQUIRED), ("Missing qualification evidence", filename, raw.get("business"))
            assert all(valid_url(raw[k]) for k in ("format_source_url", "product_source_url")), raw["business"]
            assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw["review_date"]), raw["business"]
            assert raw["contact_access"] in {"Official page reviewed", "Directory page reviewed", "Search index only", "Mixed"}, raw["business"]
            contact_urls = urls(raw.get("contact_source_urls")) + urls(raw.get("contact_source_url"))
            assert contact_urls, ("Missing public contact provenance", raw["business"])
            has_route = (present(raw.get("phone")) or present(raw.get("email")) or
                         valid_url(raw.get("contact_form_url")) or
                         ("form" in (as_text(raw.get("recommended_first_contact")) + " " +
                                     as_text(raw.get("contact_notes"))).casefold() and contact_urls))
            assert has_route, ("No public phone/email/form route", raw["business"])
            row = copy.deepcopy(raw)
            prior = raw.get("prior_expansion_id") or raw.get("existing_lead_id") or raw.get("lead_id", "")
            if prior and re.fullmatch(r"N\d{3}", prior):
                row["prior_expansion_id"] = prior
            else:
                row["prior_expansion_id"] = ""
            row["qualification_source_file"] = filename
            additions.append(row)
    additions.sort(key=lambda r: (geography(r)[0], city_key(r), r["business"].casefold(), r.get("address", "")))
    selected = copy.deepcopy(baseline)
    for row in selected:
        derive(row, True, BASELINE)
    for i, row in enumerate(additions, 51):
        row["lead_id"] = f"L{i:03}"
        derive(row, False, row["qualification_source_file"])
        selected.append(row)
    assert len({r["lead_id"] for r in selected}) == len(selected)
    identities = defaultdict(list)
    for row in selected:
        identities[identity(row)].append(row["lead_id"])
    exact_duplicates = {str(k): v for k, v in identities.items() if len(v) > 1}
    assert not exact_duplicates, ("Exact business/city duplicate", exact_duplicates)
    prior_ids = [r["prior_expansion_id"] for r in additions if r["prior_expansion_id"]]
    assert len(prior_ids) == len(set(prior_ids)), "Duplicate retained prior expansion ID"
    retained_decisions = {d["existing_lead_id"] for d in decisions if d["decision"] == "include"}
    assert set(prior_ids) == retained_decisions, ("Decision / retained record mismatch", set(prior_ids) ^ retained_decisions)
    by_original_id = {r["lead_id"]: r for r in selected[:50]}
    for original in baseline:
        for key, value in original.items():
            assert by_original_id[original["lead_id"]].get(key) == value, ("Original field changed", original["lead_id"], key)
    flag_shared_routes(selected)
    for row in decisions:
        replacement = next((r["lead_id"] for r in additions if r["prior_expansion_id"] == row["existing_lead_id"]), "")
        row["replacement_lead_id"] = replacement
    return baseline, selected, decisions


def flag_shared_routes(selected):
    routes = defaultdict(set)
    for row in selected:
        for address in re.findall(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", as_text(row.get("email"))):
            routes[("email", address.casefold())].add(row["lead_id"])
        for chunk in re.split(r"[;/\n]", as_text(row.get("phone"))):
            digits = re.sub(r"\D", "", chunk.split("(")[0])
            if 10 <= len(digits) <= 13:
                routes[("phone", digits[-10:])].add(row["lead_id"])
        group = as_text(row.get("purchasing_group")).strip().casefold()
        if group and group not in MISSING and "unknown" not in group and "unconfirmed" not in group and len(group) >= 5:
            routes[("group description", group)].add(row["lead_id"])
    notes = defaultdict(list)
    for (kind, token), ids in routes.items():
        if len(ids) > 1:
            for lead_id in ids:
                other = ", ".join(sorted(ids - {lead_id}))
                notes[lead_id].append(f"Shared {kind} with {other}; confirm buying-account overlap")
    for row in selected:
        row["qualification_duplicate_review"] = "; ".join(notes[row["lead_id"]]) or (
            "No exact shared route detected automatically; check group and branch notes before contact"
        )


CORE_COLS = [
    ("lead_id", "Lead ID", 11), ("business", "Business / venue", 42),
    ("cohort", "Cohort / review status", 43),
    ("qualification_priority", "Qualification priority", 39),
    ("qualification_procurement_lane", "Contact lane", 39),
    ("city", "City", 21), ("area", "Locality", 34),
    ("recommended_fruits", "Suggested fruits - inferred", 44),
    ("phone", "Public business phone", 37), ("email", "Public business email", 43),
    ("qualification_next_action", "Personalized next action", 75),
    ("buyer_role", "Ask for", 34),
    ("qualification_question_to_ask", "Qualification question", 80),
    ("commercial_format", "Reviewed commercial format - additions", 43),
    ("product_evidence", "Reviewed product evidence - additions", 82),
    ("qualification_contact_access", "Contact source access", 39),
    ("qualification_procurement_friction", "Procurement friction - unconfirmed", 75),
    ("qualification_score", "Research priority score / 100", 20),
    ("qualification_product_points", "Product evidence points / 4", 19),
    ("qualification_format_points", "Format evidence points / 3", 19),
    ("qualification_contact_points", "Contact evidence points / 3", 19),
    ("qualification_geography_points", "Geography points / 5", 19),
    ("qualification_duplicate_review", "Shared route / duplicate review", 80),
    ("prior_expansion_id", "Prior expansion ID", 19),
    ("qualification_review_date", "Addition review date", 19),
    ("verified_on", "Original source research date", 21),
    ("contact_notes", "Contact caveats", 90),
    ("purchasing_group", "Buying group / coordination", 65),
    ("local_purchase_authority", "Local purchase authority", 60),
    ("qualification_gap", "Unresolved qualification gap", 75),
    ("format_evidence", "Commercial format evidence", 85),
    ("format_source_url", "Format source URL", 70),
    ("product_source_url", "Product source URL", 70),
    ("contact_source_urls", "Contact source URLs", 75),
    ("address", "Published address", 70), ("website", "Website", 55),
    ("segment", "Business type - source", 35),
    ("use_case", "Proposed use - inferred", 72),
    ("evidence_strength", "Original product evidence level", 42),
    ("evidence_note", "Original menu / business evidence", 85),
    ("evidence_source_url", "Original menu / fit source URL", 70),
    ("contact_confidence", "Original contact classification", 33),
    ("priority", "Original priority - preserved", 25),
    ("geographic_tier", "Original geographic tier", 33),
    ("state", "Original state", 24),
    ("demand_status", "Original demand status", 45),
    ("outreach_status", "Original outreach status", 30),
]


def columns_for(selected):
    columns = list(CORE_COLS)
    used = {c[0] for c in columns}
    for key in dict.fromkeys(k for row in selected for k in row):
        if key not in used:
            label = key.replace("_", " ").capitalize()
            if label in {c[1] for c in columns}:
                label += " - source field"
            columns.append((key, label, 45))
            used.add(key)
    assert len({c[1] for c in columns}) == len(columns), "Duplicate Excel/CSV column labels"
    return columns


def source_rows_for(selected, evidence):
    result = []
    for row in selected:
        date = row.get("qualification_review_date") or row.get("verified_on", "")
        source_specs = [
            ("Contact", urls(row.get("contact_source_urls")) + urls(row.get("contact_source_url")),
             row["qualification_contact_access"], row.get("contact_notes", "")),
            ("Original menu / fit", urls(row.get("evidence_source_url")),
             row.get("evidence_strength", ""), row.get("evidence_note", "")),
            ("Commercial format", urls(row.get("format_source_url")),
             row.get("commercial_format", ""), row.get("format_evidence", "")),
            ("Product application", urls(row.get("product_source_url")),
             row.get("opportunity_type", ""), row.get("product_evidence", "")),
            ("Operating status", urls(row.get("operating_status_source_urls")),
             row.get("source_access", ""), row.get("contact_notes", "")),
        ]
        n = 0
        for purpose, links, kind, note in source_specs:
            for url in dict.fromkeys(links):
                n += 1
                result.append([f"{row['lead_id']}-S{n}", row["lead_id"], row["business"],
                               purpose, url, kind, date, row["cohort"], note])
    for row in evidence.get("sources", []):
        result.append([row["id"], "", row["title"], "Product research", row["url"],
                       row["source_type"], "2026-10-23", "Original product research",
                       row.get("facts", "") + " " + row.get("implication", "")])
    return result


def first_calls(selected):
    limits = {
        "Local practical introduction": 12, "Local formal hotel / group route": 8,
        "Remote qualification before samples": 6, "Remote formal hotel / group route": 6,
    }
    result = []
    for lane, limit in limits.items():
        rows = [r for r in selected if r["qualification_procurement_lane"] == lane and r["qualification_first_call_eligible"]]
        rows.sort(key=lambda r: (
            -r["qualification_geography_points"], -r["qualification_product_points"],
            -r["qualification_contact_points"], r["business"].casefold()
        ))
        result.extend(rows[:limit])
    return result


def summary_for(selected, decisions, sources, first):
    original = [r for r in selected if r["cohort"] == COHORT_ORIGINAL]
    new = [r for r in selected if r["cohort"] == COHORT_NEW]
    def count(key, rows=selected):
        return dict(sorted(Counter(r.get(key, "Unspecified") for r in rows).items()))
    return {
        "lead_count": len(selected), "original_baseline_count": len(original),
        "qualified_additions_count": len(new),
        "prior_expansion_audited_count": len(decisions),
        "prior_expansion_retained_count": sum(d["decision"] == "include" for d in decisions),
        "prior_expansion_removed_count": sum(d["decision"] == "exclude" for d in decisions),
        "newly_researched_additions_count": sum(not r["prior_expansion_id"] for r in new),
        "cohort_counts": count("cohort"), "city_counts": count("city"),
        "business_type_counts": count("qualification_business_type"),
        "contact_access_counts": count("qualification_contact_access"),
        "addition_contact_access_counts": count("qualification_contact_access", new),
        "contact_classification_counts": count("contact_confidence"),
        "lane_counts": count("qualification_procurement_lane"),
        "city_by_cohort": {
            city: {"original": sum(r["city"] == city for r in original),
                   "qualified_additions": sum(r["city"] == city for r in new)}
            for city in sorted({r["city"] for r in selected})
        },
        "published_email_count": sum(present(r.get("email")) for r in selected),
        "published_phone_count": sum(present(r.get("phone")) for r in selected),
        "first_calls_count": len(first), "sources_count": len(sources),
        "review_date": REVIEW_DATE, "original_review_dates": sorted({r.get("verified_on", "") for r in original}),
        "confirmed_buyers": 0, "outreach_sent": 0,
        "priority_formula": "Round(100 * (product*5 + format*4 + contact*3 + geography*2) / 51)",
        "priority_limit": "Research sorting aid for qualified additions, not conversion probability, revenue, expected volume or confirmed demand. Original baseline has no overall score and its unreviewed format dimension is 0.",
        "baseline_sha256": hashlib.sha256((ROOT / "research" / BASELINE).read_bytes()).hexdigest(),
        "original_preservation": "Every existing field and value on L001-L050 preserved; derived qualification fields added.",
        "deduplication_limit": "Unique business/city identities checked; shared contact/group routes flagged for manual coordination. Venue rows are not independent buying-account counts.",
    }


def build_xlsx(selected, decisions, columns, evidence, sources, summary, first):
    wb = xlsxwriter.Workbook(OUT / "dehydrated_fruit_b2b_leads.xlsx",
                            {"strings_to_formulas": False, "strings_to_urls": False})
    wb.set_properties({"title": "Dehydrated fruit B2B master prospects",
                       "subject": "Original 50 preserved; qualified additions and approach plans",
                       "author": "Business research",
                       "comments": "Public-source research only. Demand and buying volumes unconfirmed."})
    navy, green = "#173B35", "#E8F3EC"
    fmt = {
        "title": wb.add_format({"bold": True, "font_size": 19, "font_color": "#FFFFFF", "bg_color": navy, "valign": "vcenter"}),
        "note": wb.add_format({"font_color": "#44544B", "bg_color": green, "text_wrap": True, "valign": "vcenter"}),
        "body": wb.add_format({"text_wrap": True, "valign": "top", "font_size": 10}),
        "head": wb.add_format({"bold": True, "bg_color": navy, "font_color": "#FFFFFF", "text_wrap": True, "valign": "vcenter"}),
        "label": wb.add_format({"bold": True, "font_color": navy, "text_wrap": True, "valign": "top"}),
        "input": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "num_format": "0.00"}),
        "money": wb.add_format({"num_format": '"₹"#,##0.00', "bg_color": green, "border": 1}),
        "pct": wb.add_format({"num_format": "0.0%", "bg_color": green, "border": 1}),
        "link": wb.add_format({"font_color": "#0563C1", "underline": 1, "text_wrap": True, "valign": "top"}),
        "warning": wb.add_format({"bg_color": "#FFF0D9", "text_wrap": True, "valign": "top"}),
    }

    def page(name, title, note, widths):
        ws = wb.add_worksheet(name)
        ws.hide_gridlines(2)
        for i, width in enumerate(widths):
            ws.set_column(i, i, width, fmt["body"])
        last = max(len(widths) - 1, 1)
        # A huge master table still has a readable title across the first few columns.
        title_last = min(last, 7)
        ws.merge_range(0, 0, 0, title_last, title, fmt["title"])
        ws.set_row(0, 35)
        ws.merge_range(1, 0, 2, title_last, note, fmt["note"])
        ws.set_row(1, 26)
        ws.set_row(2, 26)
        ws.freeze_panes(5, 2 if len(widths) > 7 else 0)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.set_paper(9)
        ws.repeat_rows(0, 4)
        ws.set_footer("&LDehydrated fruit prospect research &RPage &P of &N")
        return ws

    def table(ws, headers, data, name, height=76, start=4):
        assert data, f"Empty worksheet table: {name}"
        clean = [[as_text(v) if isinstance(v, (list, dict, bool)) or v is None else v for v in row] for row in data]
        ws.add_table(start, 0, start + len(data), len(headers) - 1,
                     {"name": name, "style": "Table Style Medium 4",
                      "columns": [{"header": h, "header_format": fmt["head"]} for h in headers], "data": clean})
        ws.set_row(start, 36)
        for i in range(start + 1, start + 1 + len(data)):
            ws.set_row(i, height)

    ws = page("StartHere", "One master list: original 50 + qualified additions",
              f"{len(selected)} venue routes | 50 preserved original records + {summary['qualified_additions_count']} qualified additions | Revision {REVIEW_DATE}. No buyer contacted.", [34, 31, 31, 31, 26, 26])
    stats = [
        ("Master prospects", len(selected)), ("Original 50 - preserved, not all requalified", 50),
        ("Qualified additions", summary["qualified_additions_count"]),
        ("Prior expansion retained / audited", f"{summary['prior_expansion_retained_count']} / 159"),
        ("Prior expansion removed", summary["prior_expansion_removed_count"]),
        ("Newly researched additions", summary["newly_researched_additions_count"]),
        ("Public business phone / email records", f"{summary['published_phone_count']} / {summary['published_email_count']}"),
        ("Confirmed buyers / messages sent", "0 / 0"),
    ]
    for i, (label, val) in enumerate(stats, 4):
        ws.merge_range(i, 0, i, 2, label, fmt["label"])
        ws.write(i, 3, val)
        ws.set_row(i, 28)
    navigation = [
        ("MasterLeads", "Filter by city, cohort, contact access and procurement lane; original fields remain available."),
        ("FirstCalls", "A manageable first round, separated into local practical, local formal and remote routes."),
        ("Analytics", "Actual cohort, city, business-type and contact-access counts; no sales forecast."),
        ("ApproachPlans", "Personalized next action, buyer question, evidence cue and proposed trial for every venue."),
        ("Changes", "Decision trail for all 159 superseded expansion records, mapped to retained master IDs."),
        ("FruitPlan", "Product trial order and buyer questions; suggested demand remains unconfirmed."),
        ("EmailDrafts", "Editable drafts; no outreach has been sent."),
        ("ContactLog", "Blank tracker for actual replies, sample acceptance, volume and price."),
        ("PilotCosting", "Enter measured batch costs and usable output before quoting."),
        ("Sources", "Separate contact, format and product URLs with record-specific review dates and caveats."),
    ]
    for i, (name, note) in enumerate(navigation, 14):
        ws.write_url(i, 0, f"internal:'{name}'!A1", fmt["link"], name)
        ws.merge_range(i, 1, i, 5, note, fmt["body"])
        ws.set_row(i, 36)
    notes = [
        "Qualification confirms documented commercial format and a plausible product application; it does not establish buying interest, order size, revenue, independent purchasing authority or willingness to pay.",
        "Original L001-L050 are preserved from the first delivery. Their contact and review dates remain unchanged; this revision does not imply that all 50 passed the new format review.",
        "Addition score = 100 × (product×5 + format×4 + contact×3 + geography×2) / 51, rounded. Product 0–4, format 0–3, contact 0–3, geography 2–5. It is a research sorting aid, not conversion probability.",
        "Original records have no overall score. Their format points are 0 because that dimension was not re-reviewed, not because they failed. FirstCalls uses geography, product cue and contact access within each lane.",
        "Product: 4 explicit dried-fruit evidence; 3 published fruit ingredient or named fruit recipe; 2 relevant cocktail/mocktail programme or product category; 1 general segment. Format: 3 reviewed qualifying format, 2 index-only, 0 original not re-reviewed.",
        "Contact: 3 official page reviewed; 2 directory or mixed; 1 index-only. Original contact scores use the older source classification, capped at 2. Geography: 5 Poranki belt; 4 Vijayawada/nearby cities; 3 Hyderabad/other Andhra; 2 wider South.",
        "Procurement friction is shown separately. Hotel and chain prospects may require F&B approval and vendor onboarding; those processes and timelines are unconfirmed.",
        "Search index only means the result text was available but the contact page could not be fully reviewed. No phone was called, email inbox tested or site visit made.",
        "Begin orange and lemon trials only after a buyer chooses an application. Keep lime specifications separate. Dried fruit does not replace fresh juice or puree.",
        "Rows are venue routes, not independent accounts. Review shared-route flags and purchasing-group notes before approaching branches. There is no target padding or guaranteed 150-addition claim.",
        "MasterLeads places the practical action/contact columns first. Supplementary fields after the main columns are grouped and hidden initially; use the column-outline + control to expand them. JSON and CSV contain every field.",
    ]
    for i, note in enumerate(notes):
        row = 26 + i * 2
        ws.merge_range(row, 0, row + 1, 5, note, fmt["note"])
        ws.set_row(row, 29)
        ws.set_row(row + 1, 29)

    ws = page("MasterLeads", f"{len(selected)} master venue prospects",
              "Original records preserved; additions reviewed for format and product fit. Suggested fruits are inferred. Filter cohort and contact access before comparing scores.", [c[2] for c in columns])
    table(ws, [c[1] for c in columns], [[as_text(r.get(c[0], "")) for c in columns] for r in selected], "MasterLeadsTable", 108)
    if len(columns) > 47:
        ws.set_column(47, len(columns) - 1, 45, fmt["body"], {"hidden": True, "level": 1})
        ws.set_column(len(columns), len(columns), 3, None, {"collapsed": True})
    access_col = next(i for i, c in enumerate(columns) if c[0] == "qualification_contact_access")
    ws.conditional_format(5, access_col, 4 + len(selected), access_col,
                          {"type": "text", "criteria": "containing", "value": "Search index only", "format": fmt["warning"]})
    for col, (key, _, _) in enumerate(columns):
        if key in {"website", "format_source_url", "product_source_url", "evidence_source_url", "contact_source_url"}:
            for row_num, row in enumerate(selected, 5):
                if valid_url(row.get(key)):
                    ws.write_url(row_num, col, row[key], fmt["link"], row[key])

    fc = [
        ("qualification_procurement_lane", "Contact lane", 40), ("lead_id", "Lead ID", 11),
        ("business", "Business", 43), ("city", "City", 20), ("cohort", "Cohort", 40),
        ("phone", "Phone", 35), ("email", "Email", 40),
        ("qualification_next_action", "Personalized first action", 85),
        ("qualification_question_to_ask", "Ask before samples", 85),
        ("qualification_procurement_friction", "Procurement friction", 75),
        ("qualification_contact_access", "Contact access", 34),
        ("qualification_duplicate_review", "Coordinate shared routes", 70),
    ]
    ws = page("FirstCalls", f"{len(first)} routes for the first outreach round",
              "Up to 12 local practical + 8 local formal + 6 remote practical + 6 remote formal routes. Originals enter this short round only with a bar/brewery, hotel or coffee-roaster cue; all50 remain in MasterLeads. Sequence is proposed, not buyer interest.", [c[2] for c in fc])
    table(ws, [c[1] for c in fc], [[r.get(c[0], "") for c in fc] for r in first], "FirstCallsTable", 128)

    ws = page("Analytics", "Actual list composition", "Counts describe this research list, not market size, purchasing volume or conversion. Original 50 are a preserved cohort; additions were screened under the revised format standard.", [40, 21, 23, 20, 20, 20, 22, 22])
    analytics = [("Cohort", k, v) for k, v in summary["cohort_counts"].items()]
    analytics += [("Business type", k, v) for k, v in summary["business_type_counts"].items()]
    analytics += [("Contact access", k, v) for k, v in summary["contact_access_counts"].items()]
    analytics += [("Procurement lane", k, v) for k, v in summary["lane_counts"].items()]
    table(ws, ["Dimension", "Category", "Actual records"], analytics, "AnalyticsTable", 48)
    ws.set_column(1, 1, 62)
    start = len(analytics) + 8
    city_rows = [[city, values["original"], values["qualified_additions"], values["original"] + values["qualified_additions"]]
                 for city, values in summary["city_by_cohort"].items()]
    table(ws, ["City", "Original 50", "Qualified additions", "Total"], city_rows, "CityCohortTable", 28, start)
    chart = wb.add_chart({"type": "bar", "subtype": "stacked"})
    for col, label, color in [(1, "Original 50", "#9EBFAE"), (2, "Qualified additions", "#173B35")]:
        chart.add_series({"name": label,
                          "categories": ["Analytics", start + 1, 0, start + len(city_rows), 0],
                          "values": ["Analytics", start + 1, col, start + len(city_rows), col],
                          "fill": {"color": color}, "border": {"none": True}})
    chart.set_title({"name": "Venue routes by city and cohort"})
    chart.set_x_axis({"name": "Record count", "major_unit": 10})
    chart.set_y_axis({"reverse": True})
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 820, "height": max(430, len(city_rows) * 29)})
    ws.insert_chart("F5", chart)

    ap = [
        ("lead_id", "Lead ID", 11), ("business", "Business", 42), ("city", "City", 20),
        ("qualification_procurement_lane", "Lane", 40), ("buyer_role", "Ask for", 38),
        ("qualification_next_action", "Personalized approach", 90),
        ("qualification_question_to_ask", "Qualification question", 88),
        ("qualification_trial_offer", "Trial proposal - agree first", 88),
        ("recommended_fruits", "Suggested fruits - inferred", 42),
        ("product_evidence", "Reviewed product cue - additions", 80),
        ("evidence_note", "Original product cue", 80),
        ("qualification_procurement_friction", "Approval / procurement friction", 75),
        ("qualification_gap", "Remaining qualification gap", 75),
        ("qualification_local_authority", "Purchasing authority", 60),
        ("qualification_duplicate_review", "Shared contact/account check", 80),
    ]
    ws = page("ApproachPlans", "How to approach every venue",
              "Use the cited menu cue, ask for the right buyer and qualify one application. Confirm the contact and permission to send samples. No messages have been sent.", [c[2] for c in ap])
    table(ws, [c[1] for c in ap], [[r.get(c[0], "") for c in ap] for r in selected], "ApproachPlansTable", 135)

    change_rows = [[r["existing_lead_id"], r.get("business", ""), r.get("city", ""),
                    r["decision"], r["replacement_lead_id"], r["reason"],
                    r.get("review_date", REVIEW_DATE), r["decision_source_file"]]
                   for r in sorted(decisions, key=lambda r: r["existing_lead_id"])]
    ws = page("Changes", "Audit trail for the superseded 159 additions",
              "All 159 previous expansion records have an explicit keep/remove decision. Exclusion means insufficient fit or evidence for the revised benchmark, not a judgement about business quality. Original 50 were preserved.", [18, 45, 22, 16, 19, 110, 20, 45])
    table(ws, ["Prior expansion ID", "Business", "City", "Decision", "Current master ID",
               "Reason / evidence", "Review date", "Decision source"], change_rows, "ChangesTable", 108)

    ws = page("FruitPlan", "Fruit trial priorities",
              "Trial priorities, not a measured ranking of market demand. Separate lemon and lime. Dried garnish does not replace fresh juice or puree.", [12, 30, 40, 44, 80, 85, 20])
    table(ws, ["Trial rank", "Fruit", "Launch role", "Best fit", "Reasoning / limits", "Ask the buyer", "Source IDs"],
          [[r["rank"], r["fruit"], r["launch_role"], r["strongest_customer_fit"], r["reasoning"], r["ask_buyer"], ", ".join(r["source_ids"])]
           for r in evidence["trial_priority_by_fruit"]], "FruitPlanTable", 125)
    ws = page("BusinessFit", "What to offer each buyer type",
              "Obtain agreement on a specific application and evaluate a paid trial before interpreting interest as repeat demand.", [38, 68, 39, 95])
    table(ws, ["Business type", "Suggested first offer", "Decision maker", "Evidence to collect"],
          [[r["business_type"], r["first_offer"], r["decision_maker"], r["proof_to_collect"]]
           for r in evidence["priority_by_business_type"]], "BusinessFitTable", 110)

    draft_path = ROOT / "docs" / "outreach_templates.md"
    draft_text = draft_path.read_text(encoding="utf-8") if draft_path.is_file() else ""
    drafts = []
    for block in draft_text.split("\n## ")[1:]:
        title, _, body = block.partition("\n")
        match = re.search(r"\*\*Subject:\*\* (.+)\n\n([\s\S]+)", body)
        if match:
            drafts.append([title, match.group(1), match.group(2).strip(), "DRAFT - replace placeholders; not sent"])
    if not drafts:
        drafts = [["Buyer introduction", "Small dehydrated-fruit trial for [business]",
                   "Hello [business] team, I am [name] from Poranki. Could I speak with [buyer role]? "
                   "Your published [specific menu item] suggests a possible application for [fruit]. "
                   "Would a small agreed trial help compare appearance and cost per usable piece? "
                   "Please advise your specification and supplier approval process. Thank you, [name / phone].",
                   "DRAFT - personalize with cited evidence; not sent"]]
    ws = page("EmailDrafts", "Editable outreach drafts",
              "Use the corresponding ApproachPlans cue and replace placeholders. Do not claim production capacity, certifications, stock, shelf life or customer results that have not been established.", [33, 65, 125, 40])
    table(ws, ["Use case", "Subject", "Email body", "Status"], drafts, "EmailDraftsTable", 300)

    ws = page("ContactLog", "Record actual conversations and trial outcomes",
              "Blank quantities mean unknown. Record responses only after contact. Coordinate shared buying groups and honour requests to stop contact.", [12, 44, 20, 25, 30, 38, 33, 40, 22, 25, 32, 22, 25, 55])
    log_headers = ["Lead ID", "Business", "Last contact date", "Status", "Buyer name / role", "Confirmed buyer email",
                   "Current garnish method", "Approved fruit / specification", "Weekly use - pieces", "Target INR / piece",
                   "Sample outcome", "Next follow-up date", "Permission / opt-out", "Notes"]
    table(ws, log_headers, [[r["lead_id"], r["business"], "", "Not contacted"] + [""] * 10 for r in selected], "ContactLogTable", 62)
    ws.data_validation(5, 3, 4 + len(selected), 3, {"validate": "list", "source": [
        "Not contacted", "Called", "Email sent", "Sample agreed", "Sample delivered",
        "Paid trial", "Reorder", "Not interested", "Do not contact", "Recheck / closed"]})
    for col in (8, 9):
        ws.data_validation(5, col, 4 + len(selected), col, {"validate": "decimal", "criteria": ">=", "value": 0, "ignore_blank": True})

    ws = page("PilotCosting", "Measure before quoting a price",
              "Yellow inputs require actual batch data. Enter zero for genuinely nil costs. No example prices, yield or profit promises are assumed. One product/batch at a time.", [44, 24, 21, 98])
    ws.write_row(4, 0, ["Input / output", "Value", "Unit", "How to use"], fmt["head"])
    cost_rows = [
        (5, "Fresh fruit input", "kg", "Whole fruit purchased for this batch."),
        (6, "Fresh fruit purchase price", "INR / kg", "Actual invoice price."),
        (7, "Measured electricity used", "kWh", "Metered batch consumption, including standby if relevant."),
        (8, "Electricity tariff", "INR / kWh", "Applicable actual rate."),
        (9, "Labor time", "hours", "Preparation, loading, drying checks, packing and cleaning."),
        (10, "Labor cost", "INR / hour", "Include an allowance for your own work."),
        (11, "Packaging cost", "INR / batch", "Bags, labels and actual packaging."),
        (12, "Delivery cost", "INR / batch", "Allocate actual delivered-order cost."),
        (13, "Testing and overhead allocation", "INR / batch", "Allocate testing, depreciation, cleaning and other overhead."),
        (14, "Usable dried output", "kg", "After rejected or damaged product is removed."),
        (15, "Saleable garnish pieces", "pieces", "Count usable pieces; zero only for a weight-sold product."),
        (16, "Proposed selling price", "INR / piece", "For piece sales; enter zero if evaluating weight only."),
    ]
    for row, label, unit, helptext in cost_rows:
        ws.write(row, 0, label, fmt["label"])
        ws.write_blank(row, 1, None, fmt["input"])
        ws.write(row, 2, unit)
        ws.write(row, 3, helptext)
        ws.set_row(row, 38)
    ws.data_validation("B6:B17", {"validate": "decimal", "criteria": ">=", "value": 0, "ignore_blank": True})
    formulas = [
        (19, "Total batch cost", '=IF(COUNT(B6:B14)<9,"",B6*B7+B8*B9+B10*B11+SUM(B12:B14))', "INR", "Enter all cost inputs, including appropriate zeros.", "money"),
        (20, "Usable dry yield", '=IF(OR(COUNT(B6,B15)<2,B6<=0),"",B15/B6)', "%", "Usable dried output divided by purchased fresh weight.", "pct"),
        (21, "Cost per kg usable product", '=IF(OR(B20="",COUNT(B15)=0,B15<=0),"",B20/B15)', "INR / kg", "Includes listed costs and rejects through usable output.", "money"),
        (22, "Cost per usable garnish piece", '=IF(OR(B20="",COUNT(B16)=0,B16<=0),"",B20/B16)', "INR / piece", "Compare against the buyer-approved specification.", "money"),
        (23, "Batch revenue at proposed price", '=IF(OR(COUNT(B16:B17)<2,B16<=0,B17<=0),"",B16*B17)', "INR", "Assumes all usable pieces sell; not a sales forecast.", "money"),
        (24, "Surplus over listed costs", '=IF(OR(B24="",B20=""),"",B24-B20)', "INR", "Excludes taxes or costs not entered.", "money"),
        (25, "Margin over listed costs", '=IF(OR(B24="",B24<=0,B25=""),"",B25/B24)', "%", "Batch surplus divided by assumed revenue.", "pct"),
    ]
    for row, label, formula, unit, helptext, style in formulas:
        ws.write(row, 0, label, fmt["label"])
        ws.write_formula(row, 1, formula, fmt[style], "")
        ws.write(row, 2, unit)
        ws.write(row, 3, helptext)
        ws.set_row(row, 40)
    ws.merge_range(28, 0, 29, 3, "Use observed sample acceptance and paid reorders for production decisions. This is a batch-cost worksheet, not a forecast or formal accounting advice.", fmt["note"])

    ws = page("Sources", "Contact, commercial-format and product provenance",
              "Dates belong to each record: original source review dates are preserved. Search-index-only and conflicting evidence remain labelled. A public route is not a confirmed buyer contact.", [19, 12, 46, 26, 100, 52, 20, 45, 110])
    table(ws, ["Source ID", "Lead ID", "Business / title", "Purpose", "URL", "Source / evidence type",
               "Reviewed on", "Cohort", "What it supports / caveats"], sources, "SourcesTable", 86)
    for i, row in enumerate(sources, 5):
        if valid_url(row[4]):
            ws.write_url(i, 4, row[4], fmt["link"], row[4])
    wb.close()


def main():
    baseline, selected, decisions = collect()
    evidence = read_json("fruit_evidence.json")
    columns = columns_for(selected)
    sources = source_rows_for(selected, evidence)
    first = first_calls(selected)
    summary = summary_for(selected, decisions, sources, first)
    OUT.mkdir(exist_ok=True)
    build_xlsx(selected, decisions, columns, evidence, sources, summary, first)
    (OUT / "leads.json").write_text(json.dumps(selected, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    with (OUT / "dehydrated_fruit_leads.csv").open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.writer(stream)
        writer.writerow([c[1] for c in columns])
        writer.writerows([[csv_safe(r.get(c[0], "")) for c in columns] for r in selected])
    summary["csv_columns"] = [{"key": c[0], "label": c[1]} for c in columns]
    summary["expected_worksheets"] = EXPECTED_SHEETS
    (OUT / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: summary[k] for k in (
        "lead_count", "original_baseline_count", "qualified_additions_count",
        "prior_expansion_retained_count", "prior_expansion_removed_count",
        "newly_researched_additions_count")}, indent=2))


if __name__ == "__main__":
    main()
