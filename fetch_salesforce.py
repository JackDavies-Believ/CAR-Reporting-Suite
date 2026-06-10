#!/usr/bin/env python3
"""
fetch_salesforce.py
====================
Authenticates with Salesforce via OAuth 2.0 (username/password flow)
and downloads the four report exports as XLS files into the data/ folder.

Called automatically by the GitHub Actions workflow.
Can also be run locally:

    python fetch_salesforce.py

Credentials are read from environment variables (set as GitHub Secrets):
    SF_USERNAME        your Salesforce login email
    SF_PASSWORD        your Salesforce password
    SF_TOKEN           your Salesforce security token
    SF_CONSUMER_KEY    Connected App consumer key
    SF_CONSUMER_SECRET Connected App consumer secret
    SF_INSTANCE_URL    e.g. https://believ.my.salesforce.com (optional, defaults to login.salesforce.com)
"""

import os
import sys
import time
import requests

# ── Report IDs ────────────────────────────────────────────────────────────────
REPORTS = {
    "account_ops":       "00ON100000EfAOLMA3",   # Account / Ops Data
    "site_data":         "00ON100000EfAY1MAN",   # Site / Pipeline Data
    "status_movement":   "00ON100000EfAmYMAV",   # Site Status Movement
    "forecast_movement": "00ON100000EfDHNMA3",   # Charger Install Sign Off tracker
}

OUTPUT_FILES = {
    "account_ops":       "data/account_ops.xls",
    "site_data":         "data/site_data.xls",
    "status_movement":   "data/status_movement.xls",
    "forecast_movement": "data/forecast_movement.xls",
}

# ── Auth ───────────────────────────────────────────────────────────────────────

def get_access_token():
    """OAuth 2.0 username/password flow."""
    username      = os.environ["SF_USERNAME"]
    password      = os.environ["SF_PASSWORD"]
    security_token= os.environ.get("SF_TOKEN", "")
    client_id     = os.environ["SF_CONSUMER_KEY"]
    client_secret = os.environ["SF_CONSUMER_SECRET"]
    instance_url  = os.environ.get("SF_INSTANCE_URL", "https://sitetracker-libertycharge.my.salesforce.com")

    # Use the instance URL directly for custom domains
    login_url = instance_url.rstrip("/")
    token_url = f"{login_url}/services/oauth2/token"

    resp = requests.post(token_url, data={
        "grant_type":    "password",
        "client_id":     client_id,
        "client_secret": client_secret,
        "username":      username,
        "password":      password + security_token,
    }, timeout=30)

    if resp.status_code != 200:
        print(f"  ✗ Auth failed: {resp.status_code} {resp.text[:300]}")
        sys.exit(1)

    data = resp.json()
    print(f"  ✓ Authenticated as {data.get('id','').split('/')[-1]}")
    return data["access_token"], data["instance_url"]


# ── Download reports ──────────────────────────────────────────────────────────

def download_report(session, instance_url, report_id, output_path, label):
    """
    Uses the Salesforce Analytics API to export a report in Excel format.
    Falls back to CSV if XLS export is not available.
    """
    # Excel export via /async/reports or direct export URL
    # The simplest reliable method: export URL with ?export=1&enc=UTF-8&xf=xls
    url = f"{instance_url}/00O{report_id[3:]}" if not report_id.startswith("00O") else f"{instance_url}/{report_id}"
    export_url = f"{instance_url}/{report_id}?export=1&enc=UTF-8&xf=xls&apiVersion=66.0"

    # Retry up to 3 times
    for attempt in range(1, 4):
        resp = session.get(export_url, timeout=120)
        if resp.status_code == 200 and len(resp.content) > 100:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            with open(output_path, "wb") as f:
                f.write(resp.content)
            size_kb = len(resp.content) // 1024
            print(f"  ✓ {label}: {size_kb} KB → {output_path}")
            return True
        elif resp.status_code == 200 and len(resp.content) <= 100:
            print(f"  ⚠ {label}: Empty response on attempt {attempt}, retrying...")
            time.sleep(5 * attempt)
        else:
            print(f"  ⚠ {label}: HTTP {resp.status_code} on attempt {attempt}, retrying...")
            time.sleep(5 * attempt)

    print(f"  ✗ {label}: Failed after 3 attempts")
    return False


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Salesforce Report Fetcher")
    print("=" * 40)

    # Check required env vars
    required = ["SF_USERNAME", "SF_PASSWORD", "SF_CONSUMER_KEY", "SF_CONSUMER_SECRET"]
    missing  = [v for v in required if not os.environ.get(v)]
    if missing:
        print(f"\n✗ Missing environment variables: {', '.join(missing)}")
        print("  Set these as GitHub Secrets or in your local environment.")
        sys.exit(1)

    print("Authenticating with Salesforce...")
    access_token, instance_url = get_access_token()
    print(f"  Instance: {instance_url}")

    # Create an authenticated session
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept":        "application/vnd.ms-excel",
    })

    print("\nDownloading reports...")
    results = {}
    for key, report_id in REPORTS.items():
        label = key.replace("_", " ").title()
        ok = download_report(session, instance_url, report_id, OUTPUT_FILES[key], label)
        results[key] = ok

    failed = [k for k, v in results.items() if not v]
    if failed:
        print(f"\n✗ {len(failed)} report(s) failed: {', '.join(failed)}")
        sys.exit(1)

    print("\n✓ All reports downloaded. Run python build.py to rebuild the dashboard.")


if __name__ == "__main__":
    main()
