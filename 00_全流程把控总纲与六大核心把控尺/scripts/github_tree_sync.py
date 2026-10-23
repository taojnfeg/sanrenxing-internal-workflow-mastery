# -*- coding: utf-8 -*-
"""
GitHub API 全量目录树极速镜像同步工具
特点：使用 POST /git/trees 构建全新完整目录树（不带 base_tree），确保本地删除的冗余总集或旧文件在远端被 100% 同步清理。
用法：GITHUB_TOKEN=ghp_xxx GITHUB_REPO=owner/repo python3 github_tree_sync.py "commit message"
"""
import os, sys, json, base64, hashlib, urllib.request

TOKEN = os.environ.get("GITHUB_TOKEN", "")
REPO = os.environ.get("GITHUB_REPO", "")
BRANCH = os.environ.get("GITHUB_BRANCH", "main")

def api_req(method, path, data=None):
    url = f"https://api.github.com/repos/{REPO}{path}"
    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", f"token {TOKEN}")
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "drama-template-sync")
    if body is not None:
        req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))

def git_blob_sha(content_bytes):
    header = f"blob {len(content_bytes)}\0".encode("utf-8")
    return hashlib.sha1(header + content_bytes).hexdigest()

def sync_to_github(repo_root, commit_msg):
    assert TOKEN and REPO, "Please set GITHUB_TOKEN and GITHUB_REPO environment variables!"
    ref_data = api_req("GET", f"/git/refs/heads/{BRANCH}")
    head_sha = ref_data["object"]["sha"]
    commit_data = api_req("GET", f"/git/commits/{head_sha}")
    base_tree_sha = commit_data["tree"]["sha"]
    tree_data = api_req("GET", f"/git/trees/{base_tree_sha}?recursive=1")
    remote_shas = {item["path"]: item["sha"] for item in tree_data.get("tree", []) if item["type"] == "blob"}

    full_tree_items = []
    uploaded = 0
    for r, dirs, files in os.walk(repo_root):
        dirs[:] = [d for d in sorted(dirs) if d != ".git" and d != "__pycache__"]
        for fn in sorted(files):
            if fn.endswith(".pyc") or fn == ".DS_Store": continue
            abs_p = os.path.join(r, fn)
            rel_p = os.path.relpath(abs_p, repo_root).replace("\\", "/")
            with open(abs_p, "rb") as f:
                raw = f.read()
            local_sha = git_blob_sha(raw)
            if remote_shas.get(rel_p) == local_sha:
                full_tree_items.append({"path": rel_p, "mode": "100644", "type": "blob", "sha": local_sha})
            else:
                b64 = base64.b64encode(raw).decode("ascii")
                blob_res = api_req("POST", "/git/blobs", {"content": b64, "encoding": "base64"})
                full_tree_items.append({"path": rel_p, "mode": "100644", "type": "blob", "sha": blob_res["sha"]})
                uploaded += 1

    new_tree = api_req("POST", "/git/trees", {"tree": full_tree_items})
    new_commit = api_req("POST", "/git/commits", {"message": commit_msg, "tree": new_tree["sha"], "parents": [head_sha]})
    api_req("PATCH", f"/git/refs/heads/{BRANCH}", {"sha": new_commit["sha"], "force": True})
    print(f"[OK] Synced {len(full_tree_items)} files (uploaded {uploaded}) to {REPO}@{BRANCH} (Commit: {new_commit['sha'][:7]})")

if __name__ == "__main__":
    msg = sys.argv[1] if len(sys.argv) > 1 else "chore: sync repository"
    sync_to_github(".", msg)
