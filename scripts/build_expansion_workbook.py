"""Create a separate workbook of additional prospects and individual approach plans.

Does not regenerate or edit the original 50-lead workbook.
"""
from __future__ import annotations

import csv
import json
import math
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.tools' / 'python'))
import xlsxwriter

DATE = '2026-10-24'
OUT = ROOT / 'deliverables'
MISSING = 'Not publicly found'
SOURCES = ['expansion_vijayawada.json', 'expansion_andhra.json', 'expansion_regional.json', 'expansion_hyderabad.json']
FIELDS = ['business','segment','city','area','address','email','phone','website','contact_source_url',
          'contact_source_urls','evidence_source_url','evidence_note','recommended_fruits','use_case',
          'contact_confidence','contact_notes','geographic_tier','recommended_first_contact','buyer_role',
          'verified_on','evidence_strength','purchasing_group','fit_signal','source_access',
          'local_purchase_authority','qualification_question']

# Reviewed official cocktail/mocktail menu section; other local cues require
# a current recipe check or have unresolved contact conflicts.
NEARBY_REVIEWED_DRINK_MENU = {'Element E7 Restro Bar - Vijayawada'}


def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def norm(text):
    return re.sub(r'[^a-z0-9]', '', text.casefold())


def available(text):
    return bool(text and text.strip() and text != MISSING)


def classify(row):
    tier = int(row['geographic_tier'][0])
    fit = row['fit_signal']
    caution = row['contact_notes'].lower()
    if fit == 'Segment only' or row['source_access'] == 'Search index only':
        return 'D - Verify before sample', 'Confirm the current menu and published contact before investing in a sample.'
    if any(word in caution for word in ['conflicting', 'phone conflict', 'different phone', 'temporarily closed', 'malformed']):
        return 'D - Verify before sample', 'Resolve the contact or operating-status caveat first.'
    if tier <= 2 and row['business'] in NEARBY_REVIEWED_DRINK_MENU:
        return 'A - Nearby drink fit', 'Nearby venue with a relevant drinks cue; call the beverage buyer to test the application.'
    if tier <= 3:
        return 'B - Nearby product trial', 'Reachable regional prospect; qualify one drink or recipe and the person who can approve it.'
    return 'C - Expand selectively', 'Relevant farther-market prospect; confirm interest and sample logistics before shipping.'


def sample_for(row):
    fit = row['fit_signal']
    if fit in ['Explicit dried garnish','Cocktail or mocktail menu']:
        return 'Propose a small orange-and-lemon wheel comparison in one buyer-selected drink. Offer lime or pineapple only if requested. Agree quantity and any sample charge first.'
    if fit == 'Bakery or dessert menu':
        return 'Ask the chef to choose one fruit from the suggested range and specify pieces, texture and a recipe. Agree a small recipe trial before producing it.'
    if fit == 'Cafe beverages':
        if 'shake' in (row['evidence_note'] + row['segment']).lower():
            return 'Ask whether one buyer-selected dried-fruit topping suits an existing shake or dessert. Agree a tiny portion trial in the actual serving pack; do not assume dried fruit replaces syrup or fresh fruit.'
        return 'First confirm a suitable iced tea, cooler or coffee-tonic application. Then agree a small orange or lemon garnish trial in that drink.'
    return 'Confirm a suitable drink or recipe first. Select the fruit and sample specification with the buyer before making or dispatching a sample.'


def prepare():
    original = load(OUT / 'leads.json')
    original_keys = {(norm(r['business']), norm(r['city'])) for r in original}
    records = []
    for filename in SOURCES:
        batch = load(ROOT / 'research' / filename)
        assert isinstance(batch, list), filename
        records.extend(batch)
    assert len(records) >= 150, f'Need at least 150 additional prospects; currently {len(records)}'
    seen = set()
    for row in records:
        missing = [f for f in FIELDS if f not in row]
        assert not missing, (row.get('business'), missing)
        key = (norm(row['business']), norm(row['city']))
        assert key not in original_keys, f"Original-list duplicate: {row['business']}"
        assert key not in seen, f"New-list duplicate: {row['business']}"
        seen.add(key)
        assert row['verified_on'] == DATE
        assert available(row['phone']) or available(row['email']) or row.get('contact_form_url'), row['business']
        assert row['fit_signal'] in ['Explicit dried garnish','Cocktail or mocktail menu','Cafe beverages','Bakery or dessert menu','Segment only'], row['business']
        assert row['source_access'] in ['Page reviewed','Search index only','Mixed'], row['business']
        assert isinstance(row['contact_source_urls'],list) and row['contact_source_urls'], row['business']
        for url in row['contact_source_urls'] + [row['evidence_source_url']]:
            assert url.startswith(('http://','https://')), (row['business'],url)
        row['priority_band'], row['priority_reason'] = classify(row)
        row['sample_offer'] = sample_for(row)
        email_first = row['recommended_first_contact'].lower().startswith('email') and available(row['email'])
        row['contact_route'] = 'Email first' if email_first else ('Phone first' if available(row['phone']) else ('Email first' if available(row['email']) else 'Published contact form'))
        row['email_subject'] = f"Small dried-fruit trial for {row['business']}"
        row['opening_message'] = (
            f"Hi {row['business']} team, I'm [your name] from Poranki, and I'm setting up a small dehydrated-fruit business. "
            f"Could I speak with the {row['buyer_role']}? {row['qualification_question']} "
            f"If there is a useful application, I'd like to agree a small trial of {row['recommended_fruits'].rstrip('.')} "
            "with your team before discussing regular supply. Please let me know the right buyer and a convenient time. "
            "Thank you, [your name], [business phone]."
        )
        row['demand_status'] = 'Unconfirmed; no buyer contacted'
        row['first_list_overlap'] = 'New venue; excluded exact original-list duplicates'
    records.sort(key=lambda r:(int(r['geographic_tier'][0]),r['priority_band'],0 if r['fit_signal']=='Cocktail or mocktail menu' else 1,r['city'].casefold(),r['business'].casefold()))
    for i,row in enumerate(records,1):
        row['lead_id'] = f'N{i:03d}'
    # A group label is only a coordination cue, not evidence of independent buyers.
    groups = Counter(norm(r['purchasing_group']) for r in records if r['purchasing_group'].strip().lower() not in ['unknown','not publicly found',''])
    for row in records:
        g = norm(row['purchasing_group'])
        row['group_coordination'] = 'Shared group appears in new list; confirm one buyer before duplicate outreach.' if groups.get(g,0)>1 else 'Confirm outlet purchasing authority; one row is not proof of an independent buyer.'
        name = row['business']
        if 'One Stop Cafe' in name:
            row['group_coordination'] = 'Another One Stop Cafe appears in PNT Colony/Mangalagiri. Ownership relationship is unknown; ask whether the buyers are connected before approaching both.'
        elif 'Welcomhotel' in name or 'Fortune Murali' in name:
            row['group_coordination'] = 'Other ITC/Welcomhotel/Fortune property routes appear in this list. Shared brand or reservations contact does not prove shared purchasing; request the relevant property buyer and coordinate introductions.'
        elif 'Hyatt' in name:
            row['group_coordination'] = 'Hyatt-branded property routes appear in both Vijayawada and Hyderabad in this list. Confirm whether suppliers are approved by the property or centrally before duplicate outreach.'
        elif 'Bheemli' in name:
            row['group_coordination'] = 'Ask about Varun group procurement and its relationship to the original-list Novotel Vijayawada route; treat the property as a separate venue, not a proven separate buyer.'
    return records, original


COLUMNS = [
 ('lead_id','Lead ID',10),('priority_band','Outreach priority',28),('business','Business / venue',42),
 ('city','City',20),('area','Locality',36),('segment','Business type',26),
 ('recommended_fruits','Suggested fruits (inferred)',42),('phone','Public business phone',37),
 ('email','Public business email',42),('contact_route','First channel',20),
 ('buyer_role','Ask for',38),('recommended_first_contact','Specific approach',76),
 ('qualification_question','Ask this first',72),('priority_reason','Why this priority',68),
 ('fit_signal','Product fit signal',32),('evidence_note','Published evidence and limits',90),
 ('use_case','Proposed use (inferred)',70),('address','Published address',70),
 ('geographic_tier','Geographic tier',28),('website','Website',55),
 ('contact_confidence','Contact source quality',28),('source_access','How source was reviewed',24),
 ('contact_notes','Contact caveats',95),('contact_source_urls','Contact source URLs',85),
 ('evidence_source_url','Menu / fit source URL',85),('verified_on','Research date',18),
 ('purchasing_group','Purchasing group',45),('local_purchase_authority','Outlet buying authority',60),
 ('group_coordination','Group coordination',70),('demand_status','Buyer demand status',38),
]


def val(row,key):
    value=row.get(key,'')
    return '\n'.join(value) if isinstance(value,list) else value


def main():
    records,original=prepare()
    OUT.mkdir(exist_ok=True)
    n=len(records)
    (OUT/'additional_leads.json').write_text(json.dumps(records,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    with (OUT/'additional_dehydrated_fruit_leads.csv').open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.writer(stream);writer.writerow([c[1] for c in COLUMNS])
        for row in records:
            values=[val(row,c[0]) for c in COLUMNS]
            writer.writerow(["'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v for v in values])
    book=xlsxwriter.Workbook(OUT/'additional_dehydrated_fruit_leads.xlsx',{'strings_to_formulas':False,'strings_to_urls':False})
    book.set_properties({'title':f'{n} additional dehydrated fruit prospects','author':'Business research','comments':'Public-source prospect research; no buyer demand confirmed.'})
    formats={
        'title':book.add_format({'bg_color':'#173B35','font_color':'white','bold':True,'font_size':20,'valign':'vcenter'}),
        'note':book.add_format({'bg_color':'#E8F3EC','font_color':'#29463C','text_wrap':True,'valign':'vcenter'}),
        'head':book.add_format({'bg_color':'#173B35','font_color':'white','bold':True,'text_wrap':True,'valign':'vcenter'}),
        'text':book.add_format({'font_size':10,'text_wrap':True,'valign':'top'}),
        'label':book.add_format({'bold':True,'font_color':'#173B35','text_wrap':True}),
        'link':book.add_format({'font_color':'#0563C1','underline':1,'text_wrap':True,'valign':'top'}),
        'warn':book.add_format({'bg_color':'#FFF2CC','text_wrap':True,'valign':'top'}),
    }

    def sheet(name,title,note,widths):
        ws=book.add_worksheet(name);ws.hide_gridlines(2)
        for i,width in enumerate(widths):ws.set_column(i,i,width,formats['text'])
        ws.merge_range(0,0,0,len(widths)-1,title,formats['title']);ws.set_row(0,34)
        ws.merge_range(1,0,2,len(widths)-1,note,formats['note']);ws.set_row(1,24);ws.set_row(2,24)
        ws.freeze_panes(5,3 if len(widths)>7 else 0)
        ws.set_landscape();ws.fit_to_pages(1,0);ws.repeat_rows(0,4)
        ws.set_footer('&LAdditional prospect research | 24 October 2026 &R&P of &N')
        return ws

    def write_table(ws,headers,data,name,widths,min_height=65):
        assert data, name
        ws.add_table(4,0,4+len(data),len(headers)-1,{'name':name,'style':'Table Style Medium 4','columns':[{'header':h,'header_format':formats['head']} for h in headers],'data':data})
        ws.set_row(4,34)
        for i,row in enumerate(data,5):
            lines=max(sum(max(1,math.ceil(len(part)/max(widths[j]-2,8))) for part in str(v).split('\n')) for j,v in enumerate(row))
            ws.set_row(i,min(409,max(min_height,lines*13+10)))

    stats={'additional_prospects':n,'original_prospects':len(original),'combined_venue_rows':n+len(original),
           'public_emails':sum(available(r['email']) for r in records),'public_phones':sum(available(r['phone']) for r in records),
           'cities':dict(Counter(r['city'] for r in records)),'geographic_tiers':dict(Counter(r['geographic_tier'] for r in records)),
           'priority_bands':dict(Counter(r['priority_band'] for r in records)),
           'fit_signals':dict(Counter(r['fit_signal'] for r in records)),
           'source_access':dict(Counter(r['source_access'] for r in records)),
           'research_date':DATE,'confirmed_buyers':0,'messages_sent':0}
    ws=sheet('Start Here',f'{n} additional business prospects','New venues beyond the original 50 | Poranki delivery base | 24 October 2026 | Priorities are outreach judgments, not conversion probabilities.',[30,32,32,32,27,27])
    for i,(label,count) in enumerate([('Additional prospects',n),('Original list retained',len(original)),('Published business emails',stats['public_emails']),('Published business phones',stats['public_phones']),('Confirmed buyers',0)],4):
        ws.merge_range(i,0,i,2,label,formats['label']);ws.write(i,3,count)
    nav=[('First 30','Start with this manageable calling shortlist; confirm the buyer before making samples.'),
         ('New Leads','All new venues, public contacts, evidence, fruit suggestions and contact caveats.'),
         ('Approach Plans','A specific contact route, buyer question, sample idea and editable opening for every venue.'),
         ('Contact Log','Record actual calls, replies, sample decisions, paid trials and follow-up here.'),
         ('Priority Rules','Transparent outreach criteria. A nearby venue with a menu cue is still an unqualified prospect.'),
         ('Sources','Source links for every contact and menu/fit cue; indexed-only and directory caveats remain visible.')]
    for i,(name,description) in enumerate(nav,11):
        ws.write_url(i,0,f"internal:'{name}'!A1",formats['link'],name);ws.merge_range(i,1,i,5,description,formats['text']);ws.set_row(i,37)
    notes=[
        'No business has been called or emailed for this research. Missing public emails are explicit. Published contact details and operating status need confirmation before a visit or sample dispatch.',
        'Suggested fruit needs are inferred. Fresh-fruit drinks and dessert menus do not prove demand for the dried form. No purchase volumes, budgets or conversion rates are assumed.',
        'Each row is a venue prospect, not proof of an independent purchasing account. Use purchasing-group notes to coordinate outreach and ask whether the outlet can approve suppliers.',
        'New Leads is a dated snapshot. Contact Log is your working sales record; the sheets do not synchronize automatically. Save a working copy before rebuilding.',
        'Use docs/expansion_outreach_playbook.md for calls, sample agreements, follow-up and paid-trial steps. The original workbook and its costs/drafts remain separate.',
    ]
    for i,note in enumerate(notes):ws.merge_range(19+i*3,0,21+i*3,5,note,formats['note'])
    widths=[c[2] for c in COLUMNS]
    ws=sheet('New Leads',f'{n} additional prospects','Read the specific approach and source caveats. A priority band tells you what to do next; it does not establish interest, buying authority or likely revenue.',widths)
    write_table(ws,[c[1] for c in COLUMNS],[[val(r,c[0]) for c in COLUMNS] for r in records],'AdditionalLeads',widths,95)
    for i,row in enumerate(records,5):
        if row['website'].startswith(('https://','http://')):ws.write_url(i,19,row['website'],formats['link'])
        ws.write_url(i,24,row['evidence_source_url'],formats['link'])
    ws.conditional_format(5,8,n+4,8,{'type':'text','criteria':'containing','value':MISSING,'format':formats['warn']})
    first=sorted(records,key=lambda r:(r['priority_band'].startswith('D'),int(r['geographic_tier'][0]),r['priority_band'],0 if r['fit_signal']=='Cocktail or mocktail menu' else 1,0 if 'official' in r['contact_confidence'].lower() else 1,r['business'].casefold()))[:30]
    short=[('lead_id','Lead ID',10),('business','Business',40),('priority_band','Priority',27),('city','City',20),('area','Area',32),('phone','Phone',38),('email','Email',42),('buyer_role','Ask for',37),('recommended_first_contact','Next action',85),('qualification_question','First question',75)]
    widths=[c[2] for c in short]
    ws=sheet('First 30','Your first 30 conversations','Work in small calling batches. Confirm the branch, buyer, current drink or recipe, and a useful next step. This is a research shortlist; no appointments are booked.',widths)
    write_table(ws,[c[1] for c in short],[[val(r,c[0]) for c in short] for r in first],'FirstThirty',widths,92)
    approach=[('lead_id','Lead ID',10),('business','Business',40),('contact_route','Channel',20),('buyer_role','Ask for',36),('evidence_note','Why approach this venue',85),('recommended_first_contact','How to approach',85),('qualification_question','Qualify this need',75),('sample_offer','Proposed sample',85),('email_subject','Email subject',62),('opening_message','Editable opening message',115),('local_purchase_authority','Buying authority',55),('contact_notes','Check before contact',90)]
    widths=[c[2] for c in approach]
    ws=sheet('Approach Plans','A practical approach for every prospect','Drafts are unsent. Replace [your name] and [business phone]. Confirm indexed menu cues by asking a question; agree sample size, any charge and delivery before dispatch.',widths)
    write_table(ws,[c[1] for c in approach],[[val(r,c[0]) for c in approach] for r in records],'ApproachPlans',widths,135)
    widths=[10,40,25,20,28,34,30,60,25,25,30,24,40,24,48]
    ws=sheet('Contact Log','Convert research into a buyer conversation','Enter only actual results. This is the single working status log; source sheets stay as the original research snapshot. Blank buying volumes and budgets mean unknown.',widths)
    headers=['Lead ID','Business','Stage','Last contact date','Buyer / role confirmed','Buyer business contact','Outlet can approve supplier?','Confirmed application / specification','Weekly use (pieces or g)','Buyer target price / unit','Sample result','Next action date','Agreed next action','Permission / opt-out','Notes']
    write_table(ws,headers,[[r['lead_id'],r['business'],'Not contacted']+['']*12 for r in records],'ExpansionContactLog',widths,62)
    ws.data_validation(5,2,n+4,2,{'validate':'list','source':['Not contacted','Contact verified','Buyer identified','Interest confirmed','Sample agreed','Sample tested','Paid trial','Reorder','Not interested','Do not contact','Closed / stale']})
    widths=[32,85,90]
    ws=sheet('Priority Rules','How the outreach priorities work','These are qualitative rules for effort allocation. They are not statistical predictions, measured demand, or verified purchasing decisions.',widths)
    rules=[
        ['A - Nearby drink fit','Poranki/Vijayawada-area venue with a manually reviewed official cocktail/mocktail menu section and no unresolved contact caveat. Broad beverage labels and customer mentions do not qualify.','Call the beverage buyer, confirm the actual application, then agree a small trial.'],
        ['B - Nearby product trial','Other reviewed local/regional fit in the Poranki, Vijayawada or Guntur-area tiers. Bakery and coffee-led applications need recipe qualification.','Confirm the menu and buying authority; let the buyer choose the drink or recipe before making samples.'],
        ['C - Expand selectively','Reviewed prospects in farther Andhra markets or Hyderabad. Local delivery economics have not been measured.','Qualify interest and specification remotely before paying for samples or travel.'],
        ['D - Verify before sample','Segment-only fit, search-index-only evidence or a flagged contact conflict.','Verify operation, contact and current menu. Upgrade only after you obtain better evidence.'],
        ['Contact quality','Official website describes where the contact was published. Directory; recheck identifies weaker or mixed contact evidence. Neither means the number was called.','Use the source links and record a confirmed business contact after the first call.'],
        ['Buying group','Shared group labels and local-purchase-authority notes identify possible routing or central approval. Brand counts are not independent buyer counts.','Coordinate one introduction per known buyer; ask who handles the relevant outlet.'],
        ['Suggested fruits','Fruit use is an application proposal unless explicitly observed on a source menu; even an observed menu item does not show a supplier opportunity.','Ask about existing supplier/in-house drying, size, texture, accepted cost and a paid-trial decision.'],
    ]
    write_table(ws,['Rule','Basis','Action'],rules,'PriorityRules',widths,105)
    source_rows=[]
    for row in records:
        for j,url in enumerate(dict.fromkeys(row['contact_source_urls']),1):
            source_rows.append([f"{row['lead_id']}-C{j}",row['lead_id'],row['business'],'Contact',url,row['contact_confidence'],row['source_access'],DATE,row['contact_notes']])
        source_rows.append([f"{row['lead_id']}-E",row['lead_id'],row['business'],'Menu / fit',row['evidence_source_url'],row['evidence_strength'],row['source_access'],DATE,row['evidence_note']])
    widths=[17,10,45,20,100,35,24,18,100]
    ws=sheet('Sources','Source links and limitations','Date means research date. Each contact or menu cue is attributed to a public URL; source access may be page review, search index only, or a mixture.',widths)
    write_table(ws,['Source ID','Lead ID','Business','Purpose','URL','Source / evidence type','Access mode','Reviewed on','Evidence or caveat'],source_rows,'ExpansionSources',widths,82)
    for i,row in enumerate(source_rows,5):ws.write_url(i,4,row[4],formats['link'])
    book.close()
    stats['source_rows']=len(source_rows)
    (OUT/'additional_summary.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(stats,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
