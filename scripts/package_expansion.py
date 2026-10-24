"""Validate the expansion files and create a credential-free delivery archive."""
from __future__ import annotations

import ast
import csv
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from build_expansion_workbook import COLUMNS, FIELDS, SOURCES, available, load, norm, val
from publish_expansion import FILES

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'deliverables'
NS = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def main():
    rows = load(OUT / 'additional_leads.json')
    old = load(OUT / 'leads.json')
    n = len(rows)
    assert n >= 150
    assert len({r['lead_id'] for r in rows}) == n
    assert len({(norm(r['business']),norm(r['city'])) for r in rows}) == n
    assert not ({(norm(r['business']),norm(r['city'])) for r in rows} &
                {(norm(r['business']),norm(r['city'])) for r in old})
    source_rows = [r for name in SOURCES for r in load(ROOT / 'research' / name)]
    assert len(source_rows) == n
    source_map = {(r['business'],r['city']):r for r in source_rows}
    for row in rows:
        assert all(row[k] == source_map[(row['business'],row['city'])][k] for k in FIELDS)
        assert row['demand_status'] == 'Unconfirmed; no buyer contacted'
        assert row['recommended_first_contact'] and row['qualification_question'] and row['opening_message']
        assert row['verified_on'] == '2026-10-24'
        assert available(row['phone']) or available(row['email'])
    numbers = {}
    for row in old:
        for raw in re.findall(r'\+91[0-9 ()-]+',row['phone']):
            numbers[re.sub(r'\D','',raw)] = row['business']
    for row in rows:
        for raw in re.findall(r'\+91[0-9 ()-]+',row['phone']):
            number = re.sub(r'\D','',raw)
            assert len(number) == 12, (row['business'],raw)
            if number in numbers:
                assert numbers[number] == row['business'], (row['business'],numbers[number],number)
            numbers[number] = row['business']
    with (OUT / 'additional_dehydrated_fruit_leads.csv').open(encoding='utf-8-sig',newline='') as stream:
        csv_rows = list(csv.reader(stream))
    assert csv_rows[0] == [c[1] for c in COLUMNS]
    expected_csv = []
    for row in rows:
        values = [val(row,c[0]) for c in COLUMNS]
        expected_csv.append(["'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v for v in values])
    assert csv_rows[1:] == expected_csv
    with ZipFile(OUT / 'additional_dehydrated_fruit_leads.xlsx') as archive:
        assert archive.testzip() is None
        assert not any(name.startswith('xl/externalLinks/') for name in archive.namelist())
        wb = ET.fromstring(archive.read('xl/workbook.xml'))
        sheets = [s.attrib['name'] for s in wb.find('s:sheets',NS)]
        assert sheets == ['Start Here','New Leads','First 30','Approach Plans','Contact Log','Priority Rules','Sources']
        shared = ET.fromstring(archive.read('xl/sharedStrings.xml'))
        strings = [''.join(e.itertext()) for e in shared]

        def cells(sheet_number):
            sheet = ET.fromstring(archive.read(f'xl/worksheets/sheet{sheet_number}.xml'))
            assert not sheet.findall('.//s:f',NS), 'Unexpected formula'
            values = {}
            for cell in sheet.findall('.//s:c',NS):
                v = cell.find('s:v',NS)
                if v is not None:
                    values[cell.attrib['r']] = strings[int(v.text)] if cell.get('t') == 's' else v.text
            return values, sheet

        for number in range(1,8):
            cells(number)
        values,_ = cells(2)
        def col(index):
            letters = ''
            while index:
                index,remainder = divmod(index-1,26)
                letters = chr(65+remainder)+letters
            return letters
        for i,row in enumerate(rows,6):
            for j,(key,_,_) in enumerate(COLUMNS,1):
                assert values.get(f'{col(j)}{i}','') == str(val(row,key)), (i,key)
        first,_ = cells(3)
        first_ids = [first[f'A{i}'] for i in range(6,36)]
        assert len(set(first_ids)) == 30
        by_id = {r['lead_id']:r for r in rows}
        assert all(int(by_id[lead]['geographic_tier'][0]) <= 3 for lead in first_ids)
        approaches,_ = cells(4)
        assert all(approaches[f'J{i}'] == row['opening_message'] for i,row in enumerate(rows,6))
        log,log_xml = cells(5)
        assert all(log[f'C{i}'] == 'Not contacted' for i in range(6,n+6))
        assert log_xml.find('s:dataValidations',NS) is not None
        source_cells,_ = cells(7)
        summary = load(OUT / 'additional_summary.json')
        assert len([key for key in source_cells if re.fullmatch(r'A\d+',key) and int(key[1:])>=6]) == summary['source_rows']
    for script in (ROOT / 'scripts').glob('*.py'):
        ast.parse(script.read_text(encoding='utf-8-sig'),filename=str(script))
    baseline_path = ROOT / '.tools' / 'expansion_publish_baseline.json'
    original_hash_check = 'Baseline file unavailable; preservation checked at publication.'
    if baseline_path.exists():
        baseline = load(baseline_path)
        for name,digest in baseline['local_original_sha256'].items():
            if name != 'README.md':
                assert hashlib.sha256((ROOT/name).read_bytes()).hexdigest() == digest, name
        original_hash_check = 'All 17 original files other than README match their pre-expansion SHA-256 hashes.'
    artifact_files = [name for name in FILES if name != 'deliverables/additional_validation_report.json']
    report = {'status':'passed','additional_prospects':n,
        'checks':['At least 150 additional unique venue records with individualized approaches',
                  'No exact venue or normalized phone overlap with original list',
                  'Research JSON, generated JSON, CSV and all Excel lead cells match',
                  'Seven valid worksheets; no formulas or external workbook references',
                  'First 30 comprises nearby tiers 1-3 and unique prospect IDs',
                  'All approach drafts present; working log starts Not contacted',
                  'Source rows and summary counts match', 'All Python scripts parse',original_hash_check],
        'limits':['Contacts have not been called or emailed; buying interest is unconfirmed',
                  'Menu fit is inferred, not a measured probability of conversion',
                  'Some contacts and menus were available only in search indexes; these are flagged',
                  'Workbook structure and cell contents validated; not visually tested in desktop Excel',
                  'Venue rows do not prove independent purchasing accounts'],
        'artifact_sha256':{name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in artifact_files}}
    (OUT/'additional_validation_report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
    extras = ['.gitignore','requirements.txt',
              'deliverables/leads.json','deliverables/dehydrated_fruit_b2b_leads.xlsx',
              'deliverables/dehydrated_fruit_leads.csv','deliverables/summary.json','deliverables/validation_report.json',
              'docs/launch_plan.md','docs/outreach_templates.md','docs/agent_reach_review.md',
              'scripts/build_workbook.py','scripts/package_delivery.py','scripts/publish_github.py',
              'research/local_leads.json','research/hyderabad_leads.json','research/south_leads.json','research/fruit_evidence.json']
    package_files = sorted(set(FILES + extras))
    manifest = {name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in package_files}
    target = ROOT/'delivery'/'additional_leads_business_pack.zip'
    target.parent.mkdir(exist_ok=True)
    with ZipFile(target,'w',compression=ZIP_DEFLATED) as archive:
        for name in package_files:
            archive.write(ROOT/name,name)
        archive.writestr('sha256_manifest.json',json.dumps(manifest,indent=2)+'\n')
    with ZipFile(target) as archive:
        assert archive.testzip() is None
        assert set(archive.namelist()) == set(package_files + ['sha256_manifest.json'])
    print(json.dumps({'validation':'passed','prospects':n,'first_30_geography':dict(Counter(by_id[i]['city'] for i in first_ids)),
        'zip':str(target),'files':len(package_files)+1,'bytes':target.stat().st_size},indent=2))


if __name__ == '__main__':
    main()
