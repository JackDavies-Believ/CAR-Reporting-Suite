#!/usr/bin/env python3
"""
upload_to_sharepoint.py
========================
Uploads the dashboard and data files to SharePoint using Microsoft Graph API.
Called automatically by the GitHub Actions workflow after build.py runs.

Credentials read from environment variables (set as GitHub Secrets):
    AZURE_TENANT_ID        Directory (tenant) ID from Azure App registration
    AZURE_CLIENT_ID        Application (client) ID from Azure App registration
    AZURE_CLIENT_SECRET    Client secret value from Azure App registration
    SHAREPOINT_SITE_URL    e.g. https://believ.sharepoint.com/sites/Operations
    SHAREPOINT_FOLDER      Document library + folder, e.g. CAR Reporting Suite
"""

import os
import sys
import requests

# ── Files to upload ────────────────────────────────────────────────────────────
UPLOAD_FILES = {
    "index.html":               "index.html",
    "data/account_ops.xls":     "data/account_ops.xls",
    "data/site_data.xls":       "data/site_data.xls",
    "data/status_movement.xls": "data/status_movement.xls",
    "data/forecast_movement.xls":"data/forecast_movement.xls",
}

# ── Auth ───────────────────────────────────────────────────────────────────────

def get_access_token():
    tenant_id     = os.environ["AZURE_TENANT_ID"]
    client_id     = os.environ["AZURE_CLIENT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]

    url  = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)

    if resp.status_code != 200:
        print(f"  ✗ Auth failed: {resp.status_code} {resp.text[:200]}")
        sys.exit(1)

    print("  ✓ Authenticated with Microsoft Graph")
    return resp.json()["access_token"]


# ── Get SharePoint site ID ─────────────────────────────────────────────────────

def get_site_id(token, site_url):
    # Extract host and path from URL
    # e.g. https://believ.sharepoint.com/sites/Operations
    site_url = site_url.rstrip("/")
    parts    = site_url.replace("https://", "").split("/", 1)
    host     = parts[0]
    path     = "/" + parts[1] if len(parts) > 1 else ""

    url  = f"https://graph.microsoft.com/v1.0/sites/{host}:{path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)

    if resp.status_code != 200:
        print(f"  ✗ Could not find SharePoint site: {resp.status_code} {resp.text[:200]}")
        print(f"    Check SHAREPOINT_SITE_URL is correct: {site_url}")
        sys.exit(1)

    site_id = resp.json()["id"]
    print(f"  ✓ SharePoint site found: {resp.json().get('displayName', site_url)}")
    return site_id


# ── Get or create drive item path ─────────────────────────────────────────────

def get_drive_id(token, site_id):
    url  = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if resp.status_code != 200:
        print(f"  ✗ Could not get drive: {resp.status_code}")
        sys.exit(1)
    return resp.json()["id"]


# ── Upload a file ──────────────────────────────────────────────────────────────

def upload_file(token, site_id, drive_id, folder, local_path, remote_filename):
    """Upload a file using Graph API PUT (handles files up to 4MB inline)."""
    if not os.path.exists(local_path):
        print(f"  ⚠ Skipping {local_path} — file not found")
        return False

    with open(local_path, "rb") as f:
        content = f.read()

    size_kb = len(content) // 1024

    # Build the remote path inside the document library folder
    remote_path = f"{folder}/{remote_filename}".lstrip("/")

    # Use upload session for large files (>4MB), simple PUT for smaller
    if len(content) > 4 * 1024 * 1024:
        ok = upload_large_file(token, site_id, drive_id, remote_path, content)
    else:
        url  = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{remote_path}:/content"
        resp = requests.put(url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"},
            data=content, timeout=120)
        ok = resp.status_code in (200, 201)
        if not ok:
            print(f"  ✗ Upload failed for {remote_filename}: {resp.status_code} {resp.text[:200]}")

    if ok:
        print(f"  ✓ {remote_filename} ({size_kb} KB) → SharePoint/{folder}/")
    return ok


def upload_large_file(token, site_id, drive_id, remote_path, content):
    """Upload large files using an upload session."""
    # Create upload session
    url  = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{remote_path}:/createUploadSession"
    resp = requests.post(url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30)

    if resp.status_code != 200:
        print(f"  ✗ Could not create upload session: {resp.status_code}")
        return False

    upload_url = resp.json()["uploadUrl"]
    chunk_size = 3 * 1024 * 1024  # 3MB chunks
    total      = len(content)

    for i in range(0, total, chunk_size):
        chunk = content[i:i+chunk_size]
        end   = min(i + chunk_size - 1, total - 1)
        resp  = requests.put(upload_url,
            headers={"Content-Range": f"bytes {i}-{end}/{total}",
                     "Content-Length": str(len(chunk))},
            data=chunk, timeout=120)
        if resp.status_code not in (200, 201, 202):
            print(f"  ✗ Chunk upload failed: {resp.status_code}")
            return False

    return True


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("SharePoint Upload")
    print("=" * 40)

    # Check env vars
    required = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET",
                "SHAREPOINT_SITE_URL", "SHAREPOINT_FOLDER"]
    missing  = [v for v in required if not os.environ.get(v)]
    if missing:
        print(f"\n✗ Missing environment variables: {', '.join(missing)}")
        sys.exit(1)

    site_url = os.environ["SHAREPOINT_SITE_URL"].rstrip("/")
    folder   = os.environ["SHAREPOINT_FOLDER"].strip("/")

    token   = get_access_token()
    site_id = get_site_id(token, site_url)
    drive_id= get_drive_id(token, site_id)

    print(f"\nUploading {len(UPLOAD_FILES)} files to SharePoint/{folder}/...")
    results = {}
    for local_path, remote_name in UPLOAD_FILES.items():
        results[remote_name] = upload_file(token, site_id, drive_id, folder, local_path, remote_name)

    failed = [k for k, v in results.items() if not v]
    if failed:
        print(f"\n✗ {len(failed)} file(s) failed: {', '.join(failed)}")
        sys.exit(1)

    print(f"\n✓ All files uploaded to SharePoint successfully.")
    print(f"  Location: {site_url}/{folder}/")


if __name__ == "__main__":
    main()
