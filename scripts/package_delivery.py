"""Validate the delivery's structure and package only intended project files."""
import ast
import csv
import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'deliverables'
leads = json.loads((OUT / 'leads.json').read_text(encoding='utf-8'))
assert len(leads) == 50
assert len({row['lead_id'] for row in leads}) == 50
assert sum(row['city'] == 'Vijayawada' for row in leads) == 25
assert sum(row['city'] == 'Hyderabad' for row in leads) == 17
assert sum(row['email'] != 'Not publicly found' for row in leads) == 33
assert sum(row['phone'] != 'Not publicly found' for row in leads) == 49
for row in leads:
    assert row['contact_source_urls'] and row['evidence_source_url'].startswith('https://')
    assert row['demand_status'] == 'Unconfirmed; no buyer contacted'
with (OUT / 'dehydrated_fruit_leads.csv').open(encoding='utf-8-sig', newline='') as stream:
    csv_rows = list(csv.reader(stream))
assert len(csv_rows) == 51
assert all(len(row) == len(csv_rows[0]) for row in csv_rows)
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(OUT / 'dehydrated_fruit_b2b_leads.xlsx') as archive:
    assert archive.testzip() is None
    workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    names = [sheet.attrib['name'] for sheet in workbook.find('s:sheets', ns)]
    assert names == ['Start Here','Leads','Local First','Fruit Plan','Business Fit','Email Drafts','Contact Log','Pilot Costing','Sources']
    formulas = []
    for name in archive.namelist():
        if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
            sheet = ET.fromstring(archive.read(name))
            for cell in sheet.findall('.//s:c', ns):
                formula = cell.find('s:f', ns)
                if formula is not None:
                    formulas.append((name, cell.attrib['r'], formula.text))
    assert len(formulas) == 7
    assert all(name == 'xl/worksheets/sheet8.xml' for name, _, _ in formulas)
    assert [cell for _,cell,_ in formulas] == [f'B{i}' for i in range(20,27)]
for script in (ROOT / 'scripts').glob('*.py'):
    ast.parse(script.read_text(encoding='utf-8'), filename=str(script))
report = {
    'status': 'passed',
    'checks': ['50 unique lead IDs', '25 local / 17 Hyderabad / 8 other southern prospects', '33 public emails / 49 public phones', 'All leads have source URLs and unconfirmed demand status', 'CSV has 50 consistent records', 'XLSX archive integrity and nine worksheets', 'Seven costing formulas in the expected cells', 'Python scripts parse successfully'],
    'limits': ['Contacts were not called or emailed', 'Buyer demand and purchasing volumes are unconfirmed', 'Workbook visually untested in desktop Excel', 'GitHub upload status is reported separately after authentication'],
}
(OUT / 'validation_report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
files = [ROOT / name for name in ('README.md', '.gitignore', 'requirements.txt')]
for folder in ('docs','research','scripts','deliverables'):
    files.extend(path for path in sorted((ROOT / folder).glob('*')) if path.is_file() and path.suffix != '.zip')
manifest = {path.relative_to(ROOT).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest() for path in files}
delivery = ROOT / 'delivery'
delivery.mkdir(exist_ok=True)
target = delivery / 'dehydrated_fruit_business_pack.zip'
with ZipFile(target,'w',compression=ZIP_DEFLATED) as archive:
    for path in files:
        archive.write(path,path.relative_to(ROOT).as_posix())
    archive.writestr('sha256_manifest.json',json.dumps(manifest,indent=2)+'\n')
with ZipFile(target) as archive:
    assert archive.testzip() is None
    assert not any(name.startswith(('.tools/','.sf/','.git/','.codex/')) for name in archive.namelist())
print(json.dumps({'validation':'passed','zip':str(target),'files':len(files)+1,'bytes':target.stat().st_size},indent=2))
