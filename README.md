# Dehydrated fruit business prospects — one current master list

Use the [current Excel workbook](deliverables/dehydrated_fruit_b2b_leads.xlsx). It preserves the original 50 prospects and replaces the broad expansion with businesses assessed against stricter commercial-format, product-fit and contact criteria. The original IDs and source facts remain intact; each addition includes its evidence, suggested fruit application, approach and unanswered buying questions.

The earlier expansion workbook has been retired from active deliverables. Its files are preserved under [archive](archive/superseded_2026-10-24/README_ARCHIVE.md) for history only. You do not need to combine two lists.

The master contains **186 prospects: the original 50 plus 136 selected additions**. The review removed **95** of the previous 159 expansion entries, retained **64**, and added **72** newly researched prospects. It stops at the evidence-supported count rather than filling a quota with weak matches.

| Coverage | Prospect routes |
| --- | ---: |
| Vijayawada / Poranki belt | 37 |
| Guntur and Mangalagiri | 6 |
| Other Andhra Pradesh | 18 |
| Hyderabad | 66 |
| Other southern cities | 59 |
| Total | 186 |

There are **112 published business emails and 185 published phone contacts**. Of the 136 additions, 92 use reviewed official contact pages, 29 reviewed directory pages, eight mixed sources and seven search-index-only contacts. These are source-quality counts, not tested contacts or independent purchasing-account counts.

## Start here

- [Excel master workbook](deliverables/dehydrated_fruit_b2b_leads.xlsx): contacts, local first calls, approach plans, analytics, review decisions, source links, outreach drafts, contact log and batch costing.
- [Original 50-lead CSV](deliverables/original_50_dehydrated_fruit_leads.csv): the original file, preserved exactly for separate reference. These 50 records are also included in the current master workbook.
- [Approach and qualification guide](docs/qualification_strategy.md): how to reach the buyer, agree a useful sample and progress to a paid trial.
- [Email drafts](docs/outreach_templates.md) and [launch plan](docs/launch_plan.md).
- [CSV](deliverables/dehydrated_fruit_leads.csv), [structured records](deliverables/leads.json), [coverage summary](deliverables/summary.json) and [validation report](deliverables/validation_report.json).
- [Agent Reach assessment](docs/agent_reach_review.md): why public web research was sufficient for this work.

## What makes this list different

The replacement excludes weak generic cafe / shake / delivery listings, venues with closure signals, and unresolved operating-status conflicts. Additions need a documented dine-in, bar, hotel, event or production-bakery format and a relevant drink or recipe application. Establishment size and online ratings are not treated as proof of demand.

The analysis keeps evidence quality, contact access, geography and purchasing friction visible. A hotel may be a strong cocktail fit yet need formal vendor onboarding. A nearby established bar may offer a more practical first sample conversation. Scores help order research and calls; they are **not conversion probabilities**.

Start with **orange and lemon wheel trials**. Keep lime specifications separate, and offer pineapple or bakery ingredients only after the buyer confirms a recipe. Fresh fruit, candied fruit, ordinary dehydrated fruit and freeze-dried fruit are different products. A fruit-flavoured menu item does not prove the venue wants dried garnish.

## Read the evidence correctly

Revision research: **24 October 2026**. The original 50 retain their **23 October 2026** source dates and are labelled as the preserved baseline, rather than being presented as fully requalified under the new rules. Contact access distinguishes reviewed pages, directories, mixed sources and search-index-only records.

All demand remains unconfirmed. No business was called or emailed, no inbox was tested, and no sample was delivered. Confirm operation, current menu, buyer authority, delivery address and pricing requirements before spending on a sample. Missing contacts are marked explicitly. Venue rows are not a count of independent purchasing accounts; coordinate sister outlets and shared hotel routes.

## Rebuild and validate

Install the package in requirements.txt, then run:

```powershell
python scripts/build_workbook.py
python scripts/package_delivery.py
```

The package script produces `delivery/dehydrated_fruit_business_pack.zip` with current files only. Keep your own working copy before entering real contact-log notes; rebuilding recreates the research workbook. Historical lists are excluded from the current ZIP.

## Publication

Destination: [ArunMandava3030/dehydrated_fruits](https://github.com/ArunMandava3030/dehydrated_fruits).

`scripts/publish_github.py` publishes a reviewed manifest containing exact local hashes and an expected remote commit. It verifies historical copies before retiring paths, refuses unexpected remote or local changes, uses a non-forced branch update, and checks the final tree. A fresh manifest must be prepared for later revisions. Provide credentials through the local environment or hidden prompt, never through chat. The script reports the verified commit after publication.
