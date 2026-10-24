# Dehydrated fruit business research

Start with the new [159 additional prospects workbook](deliverables/additional_dehydrated_fruit_leads.xlsx). It adds new venue prospects beyond the [original 50-lead workbook](deliverables/dehydrated_fruit_b2b_leads.xlsx), with a specific approach, buyer question, sample suggestion and editable opening message for every business. Both lists prioritize a Poranki delivery base. The initial product recommendation remains **orange and lemon wheels**, followed by buyer-requested trials.

## New expansion: 159 additional prospects

| Geography | New prospects |
| --- | ---: |
| Poranki, Kanuru and Tadigadapa belt | 6 |
| Other Vijayawada | 30 |
| Guntur, Mangalagiri and Tenali | 25 |
| Other Andhra towns and Visakhapatnam | 43 |
| Hyderabad | 55 |
| Total additional | 159 |

All 159 have a published business phone; **52 have a published email**. Missing emails are marked explicitly. Research date: **24 October 2026**. Sources are linked in every row: 42 records use reviewed pages, 32 combine reviewed pages and search-index evidence, and 85 are search-index-only. None of the contacts has been tested by calling or emailing.

- [New Excel workbook](deliverables/additional_dehydrated_fruit_leads.xlsx): **First 30**, **New Leads**, **Approach Plans**, **Contact Log**, **Priority Rules**, **Sources** and **Start Here**.
- [How to approach and convert prospects](docs/expansion_outreach_playbook.md): buyer introductions, useful sample agreements, qualification questions, paid trials and follow-up.
- [New CSV](deliverables/additional_dehydrated_fruit_leads.csv), [structured JSON](deliverables/additional_leads.json), [coverage summary](deliverables/additional_summary.json) and [validation report](deliverables/additional_validation_report.json).

Begin with **First 30**, which selects nearby prospects for initial conversations. The main list is ordered by geography. Priority labels describe the next action, not a conversion probability: 1 nearby prospect has a reviewed specific drinks menu and clear official contact; 29 more are nearby product-trial candidates; 40 are farther-market prospects; **89 require contact or menu verification before sample spending**. A high-priority entry is still unqualified until a buyer responds.

Known sibling brands are grouped where evidence supports it. Distinct hotel properties and the uncertain One Stop Cafe relationship carry coordination notes. The combined **209 venue rows** are not a claim of 209 independent buying accounts. No purchase interest, volumes, budgets or supplier approval have been confirmed. The original files remain unchanged.

## Original 50: coverage

| Geography | Prospects |
| --- | ---: |
| Poranki, Kanuru and Tadigadapa belt | 6 |
| Other Vijayawada | 19 |
| Hyderabad | 17 |
| Visakhapatnam, Bengaluru, Chennai and Kochi | 8 |
| Total | 50 |

The list contains **33 published business emails** and **49 published phone contacts**. Missing contacts are marked `Not publicly found`. Thirty records use official website contacts; other records include directory or mixed evidence and carry recheck notes. Review date: **23 October 2026**. Contacts have not been tested by phone or email.

These are prospects, not confirmed customers. Fruit matches indicate potential applications; they do not establish local demand volumes or purchase commitments. Only Lush's reviewed menu explicitly names dehydrated orange. Shared buying groups are flagged so venue rows are not mistaken for independent purchasing accounts. No outreach has been sent.

## Files

- [Excel workbook](deliverables/dehydrated_fruit_b2b_leads.xlsx): Start Here, Leads, Local First, Fruit Plan, Business Fit, Email Drafts, Contact Log, Pilot Costing and Sources.
- [CSV lead list](deliverables/dehydrated_fruit_leads.csv): portable export. Phone values have a leading apostrophe to stop Excel interpreting `+91` as a formula; the workbook and JSON preserve the original text.
- [Structured lead list](deliverables/leads.json): the same 50 selected prospects.
- [Launch plan](docs/launch_plan.md): fruit priorities by buyer type, a four-week validation plan and machine comparison criteria.
- [Outreach drafts](docs/outreach_templates.md): café, bar, bakery, hotel, expansion and follow-up emails, plus a phone opening.
- [Agent Reach assessment](docs/agent_reach_review.md): practical fit, limitations and a repeatable research process.
- [Source research](research/): regional records and fruit evidence. Includes two additional lower-fit Vizag candidates excluded from the first 50 to prioritize local leads.

## First calls

Start with Bratt Brewery + Bratt Cafe, Mr Mocktails, Fat Biddy, Besta, Zamindar Haus, Cafe Ice Magic Patamata, Ironhill Vijayawada and Verandah. Read each row's caveats before calling. Ask who buys beverage ingredients, which drink could use the sample, and whether the venue already buys or makes dried garnish.

Use **Contact Log** to record actual buyer preferences, outreach status and paid trials. **Leads** is the initial research snapshot and does not synchronize with that log. Keep estimated weekly use blank until the buyer provides it. Replace every bracketed field in an email draft before sending.

## Rebuild

Use Python 3.10 or newer with dependencies from `requirements.txt`:

```powershell
python -m pip install --target .tools/python -r requirements.txt
python scripts/build_workbook.py
```

To rebuild and validate only the expansion:

```powershell
python scripts/build_expansion_workbook.py
python scripts/package_expansion.py
```

The second command checks every exported lead cell against the source data and creates `delivery/additional_leads_business_pack.zip`. It does not rebuild the original workbook. Keep a separate working copy of either workbook before entering sales notes.

The workbook is regenerated from source JSON and the email template document. **Save a copy before rebuilding if you have entered outreach notes or costs into Excel**; a rebuild replaces the workbook. Edit source JSON when updating the research snapshot.

## Publish when authenticated

The expansion uses `python scripts/publish_expansion.py`. It adds an explicit file list and updates README against the reviewed original commit, retains other repository files, and verifies the published hashes. It refuses an unexpected branch change or a different existing expansion file. Re-running it on identical published files is a no-op; later edits require a fresh review of the baseline.

The requested destination is [ArunMandava3030/dehydrated_fruits](https://github.com/ArunMandava3030/dehydrated_fruits). Publishing requires GitHub Contents read/write access for this repository and no Git installation. The original `publish_github.py` is retained for historical reproducibility; use the expansion publisher for this delivery.

```powershell
python scripts/publish_expansion.py
```

If using this fallback, enter a repository-scoped token only at the local hidden prompt; do not paste it into chat. The script does not store it. This script is a publishing option, not evidence that an upload has succeeded. Check the reported commit URL after a successful run.
