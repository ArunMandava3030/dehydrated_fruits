"""Build the researched lead workbook and CSV without sending outreach."""
from __future__ import annotations

import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.tools' / 'python'))
import xlsxwriter

DATE = '2026-10-23'
OUT = ROOT / 'deliverables'
OUT.mkdir(exist_ok=True)


def read_json(name):
    return json.loads((ROOT / 'research' / name).read_text(encoding='utf-8-sig'))


local = read_json('local_leads.json')
hyd = read_json('hyderabad_leads.json')
south = read_json('south_leads.json')
# Retain two lower-fit expansion prospects in source research, not the first 50.
excluded = {'Binge - Cafe & Private Theater', 'Bean Board - Vishalakshi Nagar (Iti Ha Asa Foods)'}
selected = local + hyd + [r for r in south if r['business'] not in excluded]
selected.sort(key=lambda r: (r['geographic_tier'], r['priority'], r['business'].casefold()))
assert len(selected) == 50, f'Expected 50 curated leads, got {len(selected)}'
assert len({(r['business'].casefold(), r['city']) for r in selected}) == 50
for i, row in enumerate(selected, 1):
    row['lead_id'] = f'L{i:03d}'
    row['demand_status'] = 'Unconfirmed; no buyer contacted'
    row['outreach_status'] = 'Not contacted'
    row['contact_source_urls'] = list(dict.fromkeys(row.get('contact_source_urls', [row['contact_source_url']])))
    row['state'] = {'Vijayawada': 'Andhra Pradesh', 'Hyderabad': 'Telangana', 'Visakhapatnam': 'Andhra Pradesh', 'Bengaluru': 'Karnataka', 'Chennai': 'Tamil Nadu', 'Kochi': 'Kerala'}.get(row['city'], '')

(OUT / 'leads.json').write_text(json.dumps(selected, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

cols = [
    ('lead_id', 'Lead ID', 11), ('priority', 'Priority', 11),
    ('business', 'Business / venue', 39), ('segment', 'Business type', 25),
    ('geographic_tier', 'Geographic order', 23), ('city', 'City', 17),
    ('area', 'Locality', 31), ('recommended_fruits', 'Suggested fruits (inferred)', 42),
    ('email', 'Public business email', 39), ('phone', 'Public business phone', 37),
    ('recommended_first_contact', 'Next action', 53), ('buyer_role', 'Ask for', 33),
    ('address', 'Published address', 60), ('state', 'State', 19),
    ('website', 'Website', 43), ('use_case', 'Proposed use (inferred)', 61),
    ('evidence_strength', 'Evidence level', 32), ('evidence_note', 'Menu / business evidence', 70),
    ('contact_confidence', 'Contact source quality', 25), ('contact_notes', 'Contact caveats', 80),
    ('contact_source_urls', 'Contact source URLs', 70), ('evidence_source_url', 'Menu / fit source URL', 65),
    ('verified_on', 'Source review date', 18), ('purchasing_group', 'Buying group / coordination', 37),
    ('demand_status', 'Buyer demand status', 38), ('outreach_status', 'Initial outreach status', 23),
]


def value(row, key):
    v = row.get(key, '')
    return '\n'.join(v) if isinstance(v, list) else v


with (OUT / 'dehydrated_fruit_leads.csv').open('w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow([c[1] for c in cols])
    for row in selected:
        vals = [value(row, c[0]) for c in cols]
        # Excel may interpret a leading + in a CSV phone as a formula. XLSX uses strings.
        writer.writerow(["'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v for v in vals])

wb = xlsxwriter.Workbook(OUT / 'dehydrated_fruit_b2b_leads.xlsx', {'strings_to_formulas': False, 'strings_to_urls': False})
wb.set_properties({'title': 'Dehydrated fruit B2B leads from Poranki', 'subject': '50 researched prospects, product trials and outreach', 'author': 'Business research', 'comments': 'Public-source research reviewed 23 October 2026. Demand unconfirmed.'})
NAVY = '#173B35'
GREEN = '#E8F3EC'
ORANGE = '#E69D45'
fmt = {
    'title': wb.add_format({'bold': True, 'font_size': 20, 'font_color': '#FFFFFF', 'bg_color': NAVY, 'valign': 'vcenter'}),
    'note': wb.add_format({'font_color': '#44544B', 'bg_color': GREEN, 'text_wrap': True, 'valign': 'vcenter'}),
    'body': wb.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 10}),
    'head': wb.add_format({'bold': True, 'bg_color': NAVY, 'font_color': '#FFFFFF', 'text_wrap': True, 'valign': 'vcenter'}),
    'label': wb.add_format({'bold': True, 'font_color': NAVY, 'text_wrap': True, 'valign': 'top'}),
    'input': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'num_format': '0.00'}),
    'money': wb.add_format({'num_format': '"₹"#,##0.00', 'bg_color': GREEN, 'border': 1}),
    'number': wb.add_format({'num_format': '0.00', 'bg_color': GREEN, 'border': 1}),
    'pct': wb.add_format({'num_format': '0.0%', 'bg_color': GREEN, 'border': 1}),
    'link': wb.add_format({'font_color': '#0563C1', 'underline': 1, 'text_wrap': True, 'valign': 'top'}),
    'warning': wb.add_format({'bg_color': '#FFF0D9', 'text_wrap': True, 'valign': 'top'}),
}


def page(name, title, note, widths):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    for i, width in enumerate(widths):
        ws.set_column(i, i, width, fmt['body'])
    last = max(len(widths) - 1, 1)
    ws.merge_range(0, 0, 0, last, title, fmt['title'])
    ws.set_row(0, 34)
    ws.merge_range(1, 0, 2, last, note, fmt['note'])
    ws.set_row(1, 24)
    ws.set_row(2, 24)
    ws.freeze_panes(5, 3 if len(widths) > 8 else 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.set_paper(9)
    ws.repeat_rows(0, 4)
    ws.set_footer('&LPoranki dehydrated fruit research &RPage &P of &N')
    return ws


def table(ws, headers, data, name, height=62):
    ws.add_table(4, 0, 4 + len(data), len(headers) - 1, {
        'name': name, 'style': 'Table Style Medium 4',
        'columns': [{'header': h, 'header_format': fmt['head']} for h in headers],
        'data': data,
    })
    ws.set_row(4, 32)
    for i in range(5, 5 + len(data)):
        ws.set_row(i, height)


emails = sum(r['email'] != 'Not publicly found' for r in selected)
phones = sum(r['phone'] != 'Not publicly found' for r in selected)
official = sum(r['contact_confidence'] == 'Official website' for r in selected)
regions = Counter(r['geographic_tier'] for r in selected)
readme = page('Start Here', 'Dehydrated fruit B2B prospecting', '50 prospects | Poranki first | Research date: 23 October 2026 | Public contacts checked against web sources; not phone-verified.', [27, 35, 35, 35, 22, 22])
summary = [
    ('Total prospects', 50), ('Poranki / Kanuru / Tadigadapa belt', regions['1 Poranki belt']),
    ('Other Vijayawada', regions['2 Vijayawada']), ('Hyderabad', len(hyd)),
    ('Other South India', 8), ('Published business emails', emails),
    ('Published business phones', phones), ('Official website contact records', official),
    ('Confirmed buyers', 0), ('Messages sent', 0),
]
for i, (label, count) in enumerate(summary, 4):
    readme.merge_range(i, 0, i, 2, label, fmt['label'])
    readme.write(i, 3, count)
readme.merge_range(15, 0, 16, 5, 'Start with orange and lemon wheel trials. Offer lime separately; test pineapple next. These are trial priorities, not a measured list of the most purchased fruits.', fmt['note'])
nav = [
    ('Local First', '25 nearby prospects with contacts and next actions; begin calls here.'),
    ('Leads', 'All 50 prospects, source links, evidence, contact caveats and suggested fruit uses.'),
    ('Fruit Plan', 'Suggested product order and buyer questions; actual demand must be validated.'),
    ('Business Fit', 'What to offer each business type and who to approach.'),
    ('Email Drafts', 'Seven editable email drafts; replace bracketed fields before sending.'),
    ('Contact Log', 'Editable outreach tracker. No messages have been sent.'),
    ('Pilot Costing', 'Enter measured batch costs and output before setting a price.'),
    ('Sources', 'Public contact, menu and research URLs; date is research date, not publication date.'),
]
for i, (name, description) in enumerate(nav, 18):
    readme.write_url(i, 0, f"internal:'{name}'!A1", fmt['link'], name)
    readme.merge_range(i, 1, i, 5, description, fmt['body'])
    readme.set_row(i, 34)
notes = [
    'Source review means the page or indexed result was reviewed. No number was called, inbox tested, site visit made or buyer contacted.',
    'Not publicly found means no usable public contact was found in this research; it does not prove no business email exists.',
    'P1 = local early outreach; P2 = qualified follow-up / Hyderabad; P3 = expansion or extra qualification. Geographic tier determines the starting route.',
    'The 50 rows represent venue prospects, not 50 independent purchasing accounts. Read the buying-group and contact notes before outreach.',
    'Only Lush explicitly lists a dehydrated fruit in the reviewed menu. Other fruit matches are inferred. No source establishes local purchase volumes.',
    'Source JSON retains two extra lower-fit Vizag candidates excluded from the first 50. General inboxes and reservation lines need routing to a buyer.',
    'Workbook values are a dated snapshot. Recheck operation, menu, contacts and delivery address before sending samples. Contact sources include official and directory evidence.',
]
for i, note in enumerate(notes, 28):
    readme.merge_range(i * 2 - 28, 0, i * 2 - 27, 5, note, fmt['note'])

ws = page('Leads', '50 researched business prospects', 'Dated research snapshot. Suggested fruits are inferred; read caveats and sources before outreach. Record later calls, replies and status changes in Contact Log.', [c[2] for c in cols])
table(ws, [c[1] for c in cols], [[value(r, c[0]) for c in cols] for r in selected], 'LeadsTable', 100)
ws.conditional_format(5, 1, 54, 1, {'type': 'text', 'criteria': 'containing', 'value': 'P1', 'format': wb.add_format({'bg_color': '#CFE9D4', 'bold': True})})
ws.conditional_format(5, 8, 54, 8, {'type': 'text', 'criteria': 'containing', 'value': 'Not publicly', 'format': fmt['warning']})
for i, r in enumerate(selected, 5):
    if r['website'].startswith('https://'):
        ws.write_url(i, 14, r['website'], fmt['link'])
    ws.write_url(i, 21, r['evidence_source_url'], fmt['link'])

nearby = [r for r in selected if r['city'] == 'Vijayawada']
lc = [('lead_id','Lead ID',10),('business','Business',38),('priority','Priority',10),('area','Area',33),('recommended_fruits','Suggested fruits',42),('phone','Phone',38),('email','Email',40),('recommended_first_contact','Next action',55),('contact_confidence','Contact quality',25),('contact_notes','Read before contact',80)]
ws = page('Local First', '25 nearby prospects to start with', 'Six prospects in the Poranki / Kanuru / Tadigadapa belt, followed by Vijayawada. Call to qualify the buyer and arrange a visit; check directory contacts before traveling.', [c[2] for c in lc])
table(ws, [c[1] for c in lc], [[value(r,c[0]) for c in lc] for r in nearby], 'LocalTable', 96)

evidence = read_json('fruit_evidence.json')
ws = page('Fruit Plan', 'Fruit trial priorities', 'Order for buyer trials, not measured market demand. Keep lemon and lime specifications separate. Dried garnish does not replace fresh juice or puree.', [10, 29, 38, 43, 74, 80, 18])
table(ws, ['Trial rank','Fruit','Launch role','Best fit','Reasoning / limits','Ask the buyer','Source IDs'], [[r['rank'],r['fruit'],r['launch_role'],r['strongest_customer_fit'],r['reasoning'],r['ask_buyer'],', '.join(r['source_ids'])] for r in evidence['trial_priority_by_fruit']], 'FruitTable', 120)
ws = page('Business Fit', 'What to offer each buyer type', 'Qualify a specific drink or recipe and collect a paid trial outcome before interpreting a lead as demand.', [37, 65, 37, 90])
table(ws, ['Business type','Suggested first offer','Decision maker','Evidence to collect'], [[r['business_type'],r['first_offer'],r['decision_maker'],r['proof_to_collect']] for r in evidence['priority_by_business_type']], 'FitTable', 102)

draft_text = (ROOT/'docs'/'outreach_templates.md').read_text(encoding='utf-8')
drafts = []
for block in draft_text.split('\n## ')[1:]:
    title, _, body = block.partition('\n')
    match = re.search(r'\*\*Subject:\*\* (.+)\n\n([\s\S]+)', body)
    if match:
        drafts.append([title, match.group(1), match.group(2).strip(), 'DRAFT - replace placeholders; not sent'])
ws = page('Email Drafts', 'Outreach email drafts', 'Use the relevant segment draft and personalize only with verified menu information. Do not claim certifications, stock, shelf life or customer results that you have not established.', [31, 62, 125, 36])
table(ws, ['Use case','Subject','Email body','Status'], drafts, 'DraftTable', 290)

ws = page('Contact Log', 'Your outreach and validation tracker', 'Fill after actual calls, visits or replies. Blank quantities mean unknown. Coordinate shared buying groups and respect requests to stop contact.', [11, 40, 20, 25, 28, 25, 29, 35, 22, 23, 27, 25, 22, 48])
log_headers = ['Lead ID','Business','Last contact date','Status','Buyer name / role','Confirmed buyer email','Current garnish method','Approved fruit / specification','Weekly use (pieces)','Target price / piece INR','Sample outcome','Next follow-up date','Permission / opt-out','Notes']
table(ws, log_headers, [[r['lead_id'],r['business'],'','Not contacted','','','','','','','','','',''] for r in selected], 'ContactLogTable', 62)
ws.data_validation(5, 3, 54, 3, {'validate': 'list', 'source': ['Not contacted','Called','Email sent','Sample agreed','Sample delivered','Paid trial','Reorder','Not interested','Do not contact','Recheck / closed']})
ws.data_validation(5, 8, 54, 8, {'validate': 'decimal','criteria': '>=','value': 0,'ignore_blank': True})
ws.data_validation(5, 9, 54, 9, {'validate': 'decimal','criteria': '>=','value': 0,'ignore_blank': True})

ws = page('Pilot Costing', 'Measure before quoting a price', 'Blank yellow inputs require your actual batch data. Enter zero for genuinely nil costs. No sample prices, yields or profit promises are assumed. One product / batch at a time.', [43, 22, 18, 95])
ws.write_row(4,0,['Input / output','Value','Unit','How to use'],fmt['head'])
cost_rows = [
    (5,'Fresh fruit input','kg','Whole fruit purchased for this batch.'),
    (6,'Fresh fruit purchase price','INR / kg','Actual invoice price.'),
    (7,'Measured electricity used','kWh','Metered batch consumption, including standby if relevant.'),
    (8,'Electricity tariff','INR / kWh','Applicable actual rate.'),
    (9,'Labor time','hours','Prep, loading, drying checks, packing and cleaning.'),
    (10,'Labor cost','INR / hour','Include an allowance for your own work.'),
    (11,'Packaging cost','INR / batch','Bags, labels and other actual packaging.'),
    (12,'Delivery cost','INR / batch','Allocate actual delivered-order cost.'),
    (13,'Testing and overhead allocation','INR / batch','Allocate testing, machine depreciation, cleaning and other overhead.'),
    (14,'Usable dried output','kg','After removing rejected / damaged product.'),
    (15,'Saleable garnish pieces','pieces','Count usable pieces; enter zero only for a weight-sold product.'),
    (16,'Proposed selling price','INR / piece','For garnish sold by pieces. Enter zero if evaluating weight only.'),
]
for row,label,unit,helptext in cost_rows:
    ws.write(row,0,label,fmt['label']); ws.write_blank(row,1,None,fmt['input']); ws.write(row,2,unit); ws.write(row,3,helptext); ws.set_row(row,35)
ws.data_validation('B6:B17',{'validate':'decimal','criteria':'>=','value':0,'ignore_blank':True})
formulas = [
    (19,'Total batch cost','=IF(COUNT(B6:B14)<9,"",B6*B7+B8*B9+B10*B11+SUM(B12:B14))','INR','Enter all cost inputs, including zeros where appropriate.','money'),
    (20,'Usable dry yield','=IF(OR(COUNT(B6,B15)<2,B6<=0),"",B15/B6)','%','Measured usable dry output divided by wet purchased weight.','pct'),
    (21,'Cost per kg of usable product','=IF(OR(B20="",COUNT(B15)=0,B15<=0),"",B20/B15)','INR / kg','Includes the listed costs and rejects through usable output.','money'),
    (22,'Cost per usable garnish piece','=IF(OR(B20="",COUNT(B16)=0,B16<=0),"",B20/B16)','INR / piece','Compare against the buyer-approved specification.','money'),
    (23,'Batch revenue at proposed price','=IF(OR(COUNT(B16:B17)<2,B16<=0,B17<=0),"",B16*B17)','INR','Assumes all usable pieces sell at this price; not a sales forecast.','money'),
    (24,'Surplus over listed costs','=IF(OR(B24="",B20=""),"",B24-B20)','INR','Excludes tax or any cost you have not entered.','money'),
    (25,'Margin over listed costs','=IF(OR(B24="",B24<=0,B25=""),"",B25/B24)','%','Batch surplus divided by assumed revenue.','pct'),
]
for row,label,formula,unit,helptext,style in formulas:
    ws.write(row,0,label,fmt['label']);ws.write_formula(row,1,formula,fmt[style],'');ws.write(row,2,unit);ws.write(row,3,helptext);ws.set_row(row,38)
ws.merge_range(28,0,29,3,'The first production decision should use observed sample acceptance and paid reorders. Discuss taxes and formal accounting separately; this sheet is a batch-cost worksheet.',fmt['note'])

source_rows = []
for r in selected:
    for j,url in enumerate(r['contact_source_urls'],1):
        source_rows.append([f"{r['lead_id']}-C{j}",r['lead_id'],r['business'],'Contact',url,r['contact_confidence'],DATE,'See Leads contact caveats for mixed or indexed-only evidence.'])
    source_rows.append([f"{r['lead_id']}-E",r['lead_id'],r['business'],'Menu / fit',r['evidence_source_url'],r['evidence_strength'],DATE,r['evidence_note']])
for r in evidence['sources']:
    source_rows.append([r['id'],'',r['title'],'Product research',r['url'],r['source_type'],DATE,r['facts']+' '+r['implication']])
extra_sources = [
    ('R1','Agent Reach documentation','https://github.com/Panniantong/Agent-Reach/blob/main/docs/README_en.md','Project documentation'),
    ('R2','Agent Reach install guide','https://github.com/Panniantong/Agent-Reach/blob/main/docs/install.md','Project documentation'),
    ('R3','FSSAI turnover thresholds order March 2026','https://www.fssai.gov.in/upload/advisories/2026/03/69b4054bb6cd6Order%20dated%2013thMarch2026_Revised%20Turnover%20threshold.pdf','Official regulator'),
    ('R4','FSSAI amendment FAQs March 2026','https://fssai.gov.in/upload/advisories/2026/03/69c6a23234827order_27032026.pdf','Official regulator'),
]
for sid,title,url,typ in extra_sources:
    source_rows.append([sid,'',title,'Reference',url,typ,DATE,'See corresponding written guide.'])
ws = page('Sources', 'Research sources and provenance', 'Use contact source IDs Lxxx-Cn and evidence IDs Lxxx-E with the lead ID. D and Q sources support product planning. Search-index-only or conflicting sources are identified in contact notes.', [17, 12, 48, 22, 95, 45, 19, 100])
table(ws,['Source ID','Lead ID','Business / title','Purpose','URL','Source / evidence type','Reviewed on','What it supports / caveats'],source_rows,'SourceTable',82)
for i,r in enumerate(source_rows,5):
    ws.write_url(i,4,r[4],fmt['link'])
wb.close()

counts = {'lead_count':len(selected),'local_count':len(nearby),'poranki_belt_count':regions['1 Poranki belt'],'hyderabad_count':len(hyd),'south_count':8,'published_email_count':emails,'published_phone_count':phones,'official_contact_records':official,'research_date':DATE,'outreach_sent':0,'confirmed_buyers':0,'sources':len(source_rows)}
(OUT/'summary.json').write_text(json.dumps(counts,indent=2)+'\n',encoding='utf-8')
print(json.dumps(counts,indent=2))
