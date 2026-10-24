"""Publish a reviewed manifest with an expected GitHub base commit.

Manifest: repo, branch, base_commit, base_blobs, files {path: sha256}, delete_paths.
Tokens are read from environment or a hidden prompt and are never saved.
No force push. Every retired file must have an exact historical copy.
"""
from __future__ import annotations
import argparse
import base64
import getpass
import hashlib
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO = 'ArunMandava3030/dehydrated_fruits'

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--plan', type=Path, default=ROOT / '.tools/revision_publication_plan.json')
    args = parser.parse_args()
    plan = json.loads(args.plan.read_text(encoding='utf-8'))
    if plan['repo'] != REPO:
        raise SystemExit('Manifest repository differs from the authorized destination.')
    contents = {}
    for name, expected in plan['files'].items():
        path = (ROOT / name).resolve()
        if not path.is_relative_to(ROOT) or any(part in {'.tools', '.git', '.codex', '.sf', '__pycache__'} for part in Path(name).parts):
            raise SystemExit(f'Unsafe or private publication path: {name}')
        content = path.read_bytes()
        if hashlib.sha256(content).hexdigest() != expected:
            raise SystemExit(f'File changed after manifest review: {name}')
        contents[name] = content
    deletes = plan.get('delete_paths', [])
    if set(contents).intersection(deletes):
        raise SystemExit('Manifest both writes and deletes the same path.')
    for name in deletes:
        if not (ROOT/name).resolve().is_relative_to(ROOT) or name not in plan['base_blobs']:
            raise SystemExit(f'Unreviewed deletion: {name}')
        archived = 'archive/superseded_2026-10-24/' + name
        if archived not in contents:
            raise SystemExit(f'Missing historical copy for {name}')
        raw = contents[archived]
        if hashlib.sha1(f'blob {len(raw)}\0'.encode()+raw).hexdigest() != plan['base_blobs'][name]:
            raise SystemExit(f'Archive differs from the previous GitHub file: {name}')
    token = os.environ.get('GITHUB_TOKEN') or os.environ.get('GH_TOKEN') or getpass.getpass('GitHub token (hidden): ')
    if not token:
        raise SystemExit('No token supplied.')
    api_root = f'https://api.github.com/repos/{REPO}'

    def api(path, method='GET', body=None):
        request = urllib.request.Request(api_root+path, method=method,
            data=json.dumps(body).encode() if body is not None else None,
            headers={'Authorization':'Bearer '+token,'Accept':'application/vnd.github+json',
                     'Content-Type':'application/json','User-Agent':'dehydrated-fruit-research',
                     'X-GitHub-Api-Version':'2022-11-28'})
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.load(response)
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f'GitHub HTTP {exc.code}: {method} {path}') from None
    branch = urllib.parse.quote(plan['branch'], safe='')
    head = api('/git/ref/heads/'+branch)['object']['sha']
    commit = api('/git/commits/'+head)
    tree = api('/git/trees/'+commit['tree']['sha']+'?recursive=1')
    if tree.get('truncated'):
        raise SystemExit('Truncated repository tree; no update attempted.')
    existing = {r['path']:r['sha'] for r in tree['tree'] if r['type']=='blob'}
    desired = {name:hashlib.sha1(f'blob {len(data)}\0'.encode()+data).hexdigest() for name,data in contents.items()}
    if all(existing.get(p)==sha for p,sha in desired.items()) and not set(deletes).intersection(existing):
        print(json.dumps({'status':'already matches','commit':head,'files_verified':len(desired)}))
        return
    if head != plan['base_commit'] or existing != plan['base_blobs']:
        raise SystemExit('GitHub changed since the reviewed baseline; reconcile before publishing.')
    def upload(name):
        if existing.get(name)==desired[name]: return None
        blob=api('/git/blobs','POST',{'encoding':'base64','content':base64.b64encode(contents[name]).decode()})
        if blob['sha'] != desired[name]: raise RuntimeError(f'Uploaded hash mismatch: {name}')
        return {'path':name,'mode':'100644','type':'blob','sha':blob['sha']}
    with ThreadPoolExecutor(max_workers=4) as pool:
        changes=[entry for entry in pool.map(upload,contents) if entry]
    changes.extend({'path':p,'mode':'100644','type':'blob','sha':None} for p in deletes if p in existing)
    new_tree=api('/git/trees','POST',{'base_tree':commit['tree']['sha'],'tree':changes})
    new_commit=api('/git/commits','POST',{
        'message':'Replace broad expansion with qualified master prospects and approach analytics',
        'tree':new_tree['sha'],'parents':[head]})
    if api('/git/ref/heads/'+branch)['object']['sha'] != head:
        raise SystemExit('Branch changed during upload; branch not updated.')
    api('/git/refs/heads/'+branch,'PATCH',{'sha':new_commit['sha'],'force':False})
    actual_head=api('/git/ref/heads/'+branch)['object']['sha']
    actual_tree=api('/git/trees/'+actual_head+'?recursive=1')
    if actual_tree.get('truncated'): raise RuntimeError('Cannot verify a truncated published tree.')
    actual={r['path']:r['sha'] for r in actual_tree['tree'] if r['type']=='blob'}
    expected={**existing,**desired}
    for name in deletes: expected.pop(name,None)
    if actual_head != new_commit['sha'] or actual != expected:
        raise RuntimeError('Published tree differs from expected files; inspect the branch.')
    print(json.dumps({'status':'published and verified','commit':actual_head,
        'url':f'https://github.com/{REPO}/commit/{actual_head}',
        'files_verified':len(actual),'retired_active_paths':len(deletes)},indent=2))

if __name__=='__main__':
    main()
