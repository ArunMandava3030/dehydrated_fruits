"""Publish this reviewed expansion while retaining the original research files.

Uses GitHub's Git database API: https://docs.github.com/en/rest/git/trees
Only README can replace a baseline file. New files must be absent or identical.
"""
from __future__ import annotations

import base64
import getpass
import hashlib
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO = 'ArunMandava3030/dehydrated_fruits'
API = f'https://api.github.com/repos/{REPO}'
EXPECTED_HEAD = 'a287d62282de66ed48faf1c411202c5910aa8a6e'
EXPECTED_README = 'e57a791476b89f1ec0dad6c71d8d3f8e86c1b740'
FILES = [
    'README.md',
    'docs/expansion_outreach_playbook.md',
    'research/expansion_vijayawada.json', 'research/expansion_andhra.json',
    'research/expansion_regional.json', 'research/expansion_hyderabad.json',
    'scripts/build_expansion_workbook.py', 'scripts/package_expansion.py',
    'scripts/publish_expansion.py',
    'deliverables/additional_leads.json',
    'deliverables/additional_dehydrated_fruit_leads.csv',
    'deliverables/additional_dehydrated_fruit_leads.xlsx',
    'deliverables/additional_summary.json', 'deliverables/additional_validation_report.json',
]


def blob_sha(content):
    return hashlib.sha1(f'blob {len(content)}\0'.encode() + content).hexdigest()


def main():
    # Complete local validation before any remote mutation.
    contents = {name: (ROOT / name).read_bytes() for name in FILES}
    expected = {name: blob_sha(data) for name, data in contents.items()}
    report = json.loads(contents['deliverables/additional_validation_report.json'])
    if report.get('status') != 'passed':
        raise SystemExit('Run the expansion validation and packaging script first.')
    for name, sha in report['artifact_sha256'].items():
        if hashlib.sha256((ROOT / name).read_bytes()).hexdigest() != sha:
            raise SystemExit(f'{name} changed after validation; validate again.')
    token = os.environ.get('GITHUB_TOKEN') or os.environ.get('GH_TOKEN') or getpass.getpass('GitHub token (hidden): ')
    if not token:
        raise SystemExit('No token supplied; nothing uploaded.')

    def api(path, method='GET', body=None):
        request = urllib.request.Request(API + path,
            data=json.dumps(body).encode() if body is not None else None,
            method=method, headers={'Authorization':'Bearer '+token,
                'Accept':'application/vnd.github+json', 'Content-Type':'application/json',
                'User-Agent':'dehydrated-fruit-expansion', 'X-GitHub-Api-Version':'2026-03-10'})
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.load(response)
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f'GitHub HTTP {exc.code} for {method} {path}; check repository access and branch rules.') from None

    def tree_at(commit_sha):
        commit = api('/git/commits/' + commit_sha)
        tree = api('/git/trees/' + commit['tree']['sha'] + '?recursive=1')
        if tree.get('truncated'):
            raise SystemExit('Truncated repository tree; no branch update attempted.')
        return commit['tree']['sha'], {e['path']:e for e in tree['tree']}

    head = api('/git/ref/heads/main')['object']['sha']
    tree_sha, existing = tree_at(head)
    if all(existing.get(name,{}).get('sha') == sha for name,sha in expected.items()):
        print(f'All expansion files match https://github.com/{REPO}/commit/{head}')
        return
    if head != EXPECTED_HEAD:
        raise SystemExit('Main changed from the reviewed baseline. Reconcile changes before publishing.')
    if existing.get('README.md',{}).get('sha') != EXPECTED_README:
        raise SystemExit('README changed from the reviewed baseline; no branch update attempted.')
    for name in FILES:
        if name != 'README.md' and name in existing and existing[name].get('sha') != expected[name]:
            raise SystemExit(f'Existing file differs: {name}; no overwrite attempted.')
    changes = []
    for name,data in contents.items():
        if existing.get(name,{}).get('sha') == expected[name]:
            continue
        blob = api('/git/blobs','POST',{'content':base64.b64encode(data).decode(),'encoding':'base64'})
        if blob['sha'] != expected[name]:
            raise SystemExit(f'Uploaded blob mismatch for {name}.')
        changes.append({'path':name,'mode':'100644','type':'blob','sha':blob['sha']})
    tree = api('/git/trees','POST',{'base_tree':tree_sha,'tree':changes})
    count = json.loads(contents['deliverables/additional_summary.json'])['additional_prospects']
    commit = api('/git/commits','POST',{'message':f'Add {count} additional prospects with individual outreach plans',
        'tree':tree['sha'],'parents':[head]})
    if api('/git/ref/heads/main')['object']['sha'] != head:
        raise SystemExit('Main changed during upload; no branch update attempted.')
    api('/git/refs/heads/main','PATCH',{'sha':commit['sha'],'force':False})
    actual = api('/git/ref/heads/main')['object']['sha']
    if actual != commit['sha']:
        raise SystemExit('Main moved after publication. Inspect GitHub before claiming the current head.')
    _,published = tree_at(actual)
    for name,sha in expected.items():
        assert published[name]['sha'] == sha, name
    for name,entry in existing.items():
        if entry['type']=='blob' and name not in FILES:
            assert published[name]['sha'] == entry['sha'], f'Unexpected change: {name}'
    print(json.dumps({'status':'published and verified','files':len(changes),'commit':actual,
        'url':f'https://github.com/{REPO}/commit/{actual}',
        'original_files_preserved':True},indent=2))


if __name__ == '__main__':
    try:
        main()
    except (RuntimeError, urllib.error.URLError) as exc:
        print(str(exc),file=sys.stderr)
        raise SystemExit(1)
