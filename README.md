# Dehydrated fruit business research

Start with the [Excel lead workbook](deliverables/dehydrated_fruit_b2b_leads.xlsx). It contains **50 business prospects**, prioritizing Poranki and Vijayawada, with suggested fruit applications, public contacts, source links and outreach drafts. The initial product recommendation is **orange and lemon wheels**, followed by buyer-requested lime and pineapple trials.

## Coverage

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

The workbook is regenerated from source JSON and the email template document. **Save a copy before rebuilding if you have entered outreach notes or costs into Excel**; a rebuild replaces the workbook. Edit source JSON when updating the research snapshot.

## Publish when authenticated

The requested destination is [ArunMandava3030/dehydrated_fruits](https://github.com/ArunMandava3030/dehydrated_fruits). The included publisher uploads only the listed project files, preserves unrelated repository content, and refuses to replace an existing file with different content. It requires GitHub Contents read/write access for this repository. It needs no Git installation.

```powershell
python scripts/publish_github.py
```

If using this fallback, enter a repository-scoped token only at the local hidden prompt; do not paste it into chat. The script does not store it. This script is a publishing option, not evidence that an upload has succeeded. Check the reported commit URL after a successful run.
