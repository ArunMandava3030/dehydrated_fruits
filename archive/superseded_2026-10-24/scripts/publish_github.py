"""Publish the prepared files to the requested GitHub repository.

Uses the standard library. Run locally; enter a repository-scoped GitHub token
at the hidden prompt. The token is neither printed nor saved. A pre-existing
file is never replaced with different content by this script.
"""
from __future__ import annotations

import base64
import getpass
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO = 'ArunMandava3030/dehydrated_fruits'
API = f'https://api.github.com/repos/{REPO}'
FILES = ['README.md', '.gitignore', 'requirements.txt']
for folder in ('docs', 'research', 'scripts', 'deliverables'):
    FILES.extend(p.relative_to(ROOT).as_posix() for p in sorted((ROOT / folder).glob('*')) if p.is_file() and p.suffix != '.zip')


def main():
    token = os.environ.get('GITHUB_TOKEN') or os.environ.get('GH_TOKEN') or getpass.getpass('GitHub token with Contents read/write for this repository (hidden): ')
    if not token:
        raise SystemExit('No token supplied; nothing uploaded.')

    def api(path, method='GET', body=None):
        data = json.dumps(body).encode('utf-8') if body is not None else None
        request = urllib.request.Request(API + path, data=data, method=method, headers={
            'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json',
            'Content-Type': 'application/json', 'User-Agent': 'dehydrated-fruits-research',
            'X-GitHub-Api-Version': '2022-11-28',
        })
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                return json.load(response)
        except urllib.error.HTTPError as exc:
            # Do not print request headers or the token.
            raise RuntimeError(f'GitHub returned HTTP {exc.code} for {method} {path}. Check repository access and branch rules.') from None

    repo = api('')
    branch = repo['default_branch']
    branch_path = urllib.parse.quote(branch, safe='')
    # The branches listing distinguishes an empty repository from a missing ref.
    branches = api('/branches?per_page=100')
    if not branches:
        text = (ROOT / 'README.md').read_bytes()
        api('/contents/README.md', 'PUT', {
            'message': 'Initialize dehydrated fruit business research', 'branch': branch,
            'content': base64.b64encode(text).decode('ascii'),
        })
    head = api('/git/ref/heads/' + branch_path)['object']['sha']
    commit = api('/git/commits/' + head)
    old_tree = commit['tree']['sha']
    existing_tree = api('/git/trees/' + old_tree + '?recursive=1')
    if existing_tree.get('truncated'):
        raise SystemExit('Repository tree is truncated; no upload attempted.')
    existing = {entry['path']: entry for entry in existing_tree['tree']}
    changes = []
    for name in FILES:
        content = (ROOT / name).read_bytes()
        if name in existing:
            if existing[name]['type'] != 'blob':
                raise SystemExit(f'Path conflict at {name}; no branch update attempted.')
            blob = api('/git/blobs/' + existing[name]['sha'])
            previous = base64.b64decode(blob['content'])
            if previous == content:
                continue
            raise SystemExit(f'{name} already exists with different content. Review it before publishing; this script will not overwrite it.')
        blob = api('/git/blobs', 'POST', {'content': base64.b64encode(content).decode('ascii'), 'encoding': 'base64'})
        changes.append({'path': name, 'mode': '100644', 'type': 'blob', 'sha': blob['sha']})
    if not changes:
        print('All prepared files already match GitHub; nothing to update.')
        return
    tree = api('/git/trees', 'POST', {'base_tree': old_tree, 'tree': changes})
    new_commit = api('/git/commits', 'POST', {
        'message': 'Add 50 sourced B2B leads, Excel workbook and launch outreach guides',
        'tree': tree['sha'], 'parents': [head],
    })
    if api('/git/ref/heads/' + branch_path)['object']['sha'] != head:
        raise SystemExit('Branch changed during preparation; no branch update attempted. Reconcile before retrying.')
    api('/git/refs/heads/' + branch_path, 'PATCH', {'sha': new_commit['sha'], 'force': False})
    actual = api('/git/ref/heads/' + branch_path)['object']['sha']
    if actual != new_commit['sha']:
        raise SystemExit('Branch moved after upload. Inspect GitHub before claiming the current head.')
    print(f'Published {len(changes)} files: https://github.com/{REPO}/commit/{actual}')


if __name__ == '__main__':
    try:
        main()
    except (RuntimeError, urllib.error.URLError) as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)

