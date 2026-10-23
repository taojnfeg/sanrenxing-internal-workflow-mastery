# -*- coding: utf-8 -*-
import os, re, sys, json, base64, hashlib, http.client

def get_credentials(repo_root: str):
    token = os.environ.get("GITHUB_TOKEN", "")
    repo_full = "taojnfeg/sanrxing"
    cfg_path = os.path.join(repo_root, ".git", "config")
    if os.path.exists(cfg_path):
        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = f.read()
        m = re.search(r"x-access-token:([^@]+)@github\.com/([^/]+/[^.\s]+)", cfg)
        if m:
            token = token or m.group(1)
            repo_full = m.group(2)
    if not token:
        raise RuntimeError("未检测到 GitHub Token！")
    return token, repo_full

def git_blob_sha(data: bytes) -> str:
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()

def sync_to_github(repo_root: str, commit_msg: str, branch: str = "main"):
    token, repo_full = get_credentials(repo_root)
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "Content-Type": "application/json",
        "User-Agent": "WorkRally-SanRenXing-V26",
        "Connection": "keep-alive"
    }
    conn = http.client.HTTPSConnection("api.github.com", timeout=25)

    def api_req(method, path, body=None):
        nonlocal conn
        full_path = f"/repos/{repo_full}{path}"
        data = json.dumps(body).encode("utf-8") if body is not None else None
        for attempt in range(3):
            try:
                conn.request(method, full_path, body=data, headers=headers)
                resp = conn.getresponse()
                raw_resp = resp.read().decode("utf-8")
                if resp.status >= 400:
                    raise RuntimeError(f"HTTP {resp.status}: {raw_resp[:200]}")
                return resp.status, json.loads(raw_resp)
            except Exception as e:
                conn.close()
                conn = http.client.HTTPSConnection("api.github.com", timeout=25)
                if attempt == 2:
                    raise

    _, ref_data = api_req("GET", f"/git/ref/heads/{branch}")
    base_commit_sha = ref_data["object"]["sha"]
    _, commit_data = api_req("GET", f"/git/commits/{base_commit_sha}")
    base_tree_sha = commit_data["tree"]["sha"]
    _, tree_data = api_req("GET", f"/git/trees/{base_tree_sha}?recursive=1")
    remote_shas = {item["path"]: item["sha"] for item in tree_data.get("tree", []) if item.get("type") == "blob"}

    # Build lookup by sha as well! If a file was renamed without content change, its blob sha ALREADY exists on GitHub!
    existing_blob_shas = set(remote_shas.values())

    to_upload = []
    full_tree_items = []
    for dirpath, _, filenames in os.walk(repo_root):
        if ".git" in dirpath.split(os.sep) or "__pycache__" in dirpath:
            continue
        for fn in filenames:
            if fn.endswith(".pyc"):
                continue
            full_p = os.path.join(dirpath, fn)
            rel_p = os.path.relpath(full_p, repo_root).replace("\\", "/")
            with open(full_p, "rb") as bf:
                raw = bf.read()
            local_sha = git_blob_sha(raw)
            if local_sha in existing_blob_shas:
                full_tree_items.append({
                    "path": rel_p,
                    "mode": "100644",
                    "type": "blob",
                    "sha": local_sha
                })
            else:
                to_upload.append((rel_p, full_p, local_sha))

    for rel_p, full_p, local_sha in to_upload:
        with open(full_p, "rb") as bf:
            b64_str = base64.b64encode(bf.read()).decode("ascii")
        _, blob_res = api_req("POST", "/git/blobs", {
            "content": b64_str,
            "encoding": "base64"
        })
        existing_blob_shas.add(blob_res["sha"])
        full_tree_items.append({
            "path": rel_p,
            "mode": "100644",
            "type": "blob",
            "sha": blob_res["sha"]
        })

    _, new_tree = api_req("POST", "/git/trees", {"tree": full_tree_items})
    if new_tree["sha"] == base_tree_sha:
        print(f"[OK] 远端目录树已完全一致 ({base_commit_sha[:7]})")
        conn.close()
        return

    _, new_commit = api_req("POST", "/git/commits", {
        "message": commit_msg,
        "tree": new_tree["sha"],
        "parents": [base_commit_sha]
    })
    api_req("PATCH", f"/git/refs/heads/{branch}", {"sha": new_commit["sha"], "force": True})
    conn.close()
    print(f"[OK] 极速同步全库 {len(full_tree_items)} 个文件（新上传 {len(to_upload)} 个）至 {repo_full}@{branch} (Commit: {new_commit['sha'][:7]})")

if __name__ == "__main__":
    msg = sys.argv[1] if len(sys.argv) > 1 else "chore(V26.0): 精确同步《三人行》最新资产库"
    repo_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
    sync_to_github(repo_dir, msg)
