"""
SOH Daily Report

Pulls current stock-on-hand for a given warehouse from Zoho Inventory and
writes it to an Excel file named SOH_DD_MM_YYYY.xlsx.

Zoho Inventory's API has no direct "Stock Summary" report endpoint, so this
approximates it using the Items API: the items *list* endpoint doesn't
include per-warehouse stock, so for every active item we fetch its full
detail record (which does include a `warehouses` array) and keep the entry
matching WAREHOUSE_NAME. Detail lookups run concurrently, paced by a shared
rate limiter to stay under Zoho's request quota (a live run against ~1850
items crashed with a 429 after being throttled for the better part of 30
minutes, which is what the rate limiter and Retry-After handling below are
for), print progress every 100 items plus a heartbeat every 20 seconds (so
a stalled run is visible in real time rather than going silent), retry
network errors and 429s with backoff, and skip (rather than crash on) any
item that still fails after retries. The access token is refreshed
automatically if a run takes long enough to approach the 1-hour token
expiry.

Required environment variables:
    ZOHO_CLIENT_ID
    ZOHO_CLIENT_SECRET
    ZOHO_REFRESH_TOKEN

Optional environment variables:
    ZOHO_ORGANIZATION_ID   - required if the Zoho account has more than one
                              organization (auto-detected otherwise)
    WAREHOUSE_NAME          - defaults to "Store 1"
    REPORT_TZ               - IANA timezone for naming the output file,
                               defaults to "Africa/Nairobi"
    RATE_LIMIT_PER_MINUTE   - max Zoho API requests/minute, defaults to 60
"""
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook

ACCOUNTS_TOKEN_URL = "https://accounts.zoho.com/oauth/v2/token"
OUTPUT_COLUMNS = ["BAR CODE", "SKU", "Item Name", "SOH", "Sales Price", "Brand"]
BARCODE_FIELDS = ("upc", "ean", "isbn", "part_number")
TOKEN_REFRESH_INTERVAL_SECONDS = 45 * 60  # access tokens expire after 1 hour
DETAIL_FETCH_WORKERS = 6
PROGRESS_EVERY = 100
HEARTBEAT_SECONDS = 20
RETRY_ATTEMPTS = 6
DEFAULT_RATE_LIMIT_PER_MINUTE = 60
MAX_RETRY_AFTER_SECONDS = 90


def get_access_token():
    response = requests.post(
        ACCOUNTS_TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": os.environ["ZOHO_CLIENT_ID"],
            "client_secret": os.environ["ZOHO_CLIENT_SECRET"],
            "refresh_token": os.environ["ZOHO_REFRESH_TOKEN"],
        },
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    if "access_token" not in data:
        raise RuntimeError(f"Zoho did not return an access token: {data}")
    return data["access_token"], data["api_domain"]


class TokenStore:
    """Thread-safe access token holder that transparently refreshes long-running jobs."""

    def __init__(self, access_token):
        self._lock = threading.Lock()
        self._token = access_token
        self._issued_at = time.monotonic()

    def get(self):
        with self._lock:
            if time.monotonic() - self._issued_at > TOKEN_REFRESH_INTERVAL_SECONDS:
                self._token, _ = get_access_token()
                self._issued_at = time.monotonic()
            return self._token


class RateLimiter:
    """Paces requests across all worker threads to a fixed rate per minute.

    A short per-item exponential backoff isn't enough to recover from a
    sustained rate limit (a live run got 429'd for the better part of 30
    minutes and then crashed) -- this caps how fast requests actually go
    out in the first place.
    """

    def __init__(self, max_per_minute):
        self._lock = threading.Lock()
        self._interval = 60.0 / max_per_minute
        self._next_slot = time.monotonic()

    def wait(self):
        with self._lock:
            now = time.monotonic()
            start = max(now, self._next_slot)
            self._next_slot = start + self._interval
        sleep_for = start - now
        if sleep_for > 0:
            time.sleep(sleep_for)


def resolve_organization_id(session, api_domain):
    org_id = os.environ.get("ZOHO_ORGANIZATION_ID")
    if org_id:
        return org_id

    response = session.get(f"{api_domain}/inventory/v1/organizations", timeout=30)
    response.raise_for_status()
    organizations = response.json().get("organizations", [])
    if len(organizations) == 1:
        return organizations[0]["organization_id"]

    names = ", ".join(f"{o['name']} ({o['organization_id']})" for o in organizations)
    raise RuntimeError(
        "Zoho account has multiple organizations; set ZOHO_ORGANIZATION_ID to "
        f"one of: {names}"
    )


def fetch_active_item_ids(session, api_domain, organization_id):
    item_ids = []
    page = 1
    while True:
        response = session.get(
            f"{api_domain}/inventory/v1/items",
            params={
                "organization_id": organization_id,
                "filter_by": "Status.Active",
                "page": page,
                "per_page": 200,
            },
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        item_ids.extend(item["item_id"] for item in payload.get("items", []))

        if not payload.get("page_context", {}).get("has_more_page"):
            break
        page += 1

    return item_ids


def parse_retry_after(response, attempt):
    header_value = response.headers.get("Retry-After")
    if header_value is not None:
        try:
            return min(float(header_value), MAX_RETRY_AFTER_SECONDS)
        except ValueError:
            pass
    return min(2**attempt, MAX_RETRY_AFTER_SECONDS)


def fetch_item_detail(session, api_domain, organization_id, token_store, rate_limiter, item_id):
    last_error = None
    for attempt in range(RETRY_ATTEMPTS):
        rate_limiter.wait()
        try:
            response = session.get(
                f"{api_domain}/inventory/v1/items/{item_id}",
                params={"organization_id": organization_id},
                headers={"Authorization": f"Zoho-oauthtoken {token_store.get()}"},
                timeout=30,
            )
        except requests.exceptions.RequestException as exc:
            last_error = exc
            print(f"  [item {item_id}] attempt {attempt + 1} network error: {exc}", file=sys.stderr, flush=True)
            time.sleep(min(2**attempt, MAX_RETRY_AFTER_SECONDS))
            continue

        if response.status_code == 429:
            wait_s = parse_retry_after(response, attempt)
            last_error = requests.exceptions.HTTPError(f"429 Too Many Requests for item {item_id}")
            print(f"  [item {item_id}] attempt {attempt + 1} rate-limited (429), waiting {wait_s:.1f}s", file=sys.stderr, flush=True)
            time.sleep(wait_s)
            continue

        try:
            response.raise_for_status()
        except requests.exceptions.HTTPError as exc:
            last_error = exc
            print(f"  [item {item_id}] attempt {attempt + 1} HTTP error: {exc}", file=sys.stderr, flush=True)
            time.sleep(min(2**attempt, MAX_RETRY_AFTER_SECONDS))
            continue

        return response.json()["item"]

    print(f"  [item {item_id}] giving up after {RETRY_ATTEMPTS} attempts: {last_error}", file=sys.stderr, flush=True)
    return None


def fetch_active_items(session, api_domain, organization_id, token_store, rate_limiter):
    item_ids = fetch_active_item_ids(session, api_domain, organization_id)
    print(f"Found {len(item_ids)} active item(s); fetching per-warehouse stock detail...", flush=True)

    items = []
    completed = 0
    failed = 0
    stop_heartbeat = threading.Event()

    def heartbeat():
        while not stop_heartbeat.wait(HEARTBEAT_SECONDS):
            print(
                f"  ...still working: {completed}/{len(item_ids)} fetched ({failed} failed so far)",
                flush=True,
            )

    heartbeat_thread = threading.Thread(target=heartbeat, daemon=True)
    heartbeat_thread.start()

    try:
        with ThreadPoolExecutor(max_workers=DETAIL_FETCH_WORKERS) as executor:
            futures = {
                executor.submit(
                    fetch_item_detail, session, api_domain, organization_id, token_store, rate_limiter, item_id
                ): item_id
                for item_id in item_ids
            }
            for future in as_completed(futures):
                detail = future.result()
                completed += 1
                if detail is None:
                    failed += 1
                else:
                    items.append(detail)
                if completed % PROGRESS_EVERY == 0 or completed == len(item_ids):
                    print(f"  ...{completed}/{len(item_ids)} items fetched ({failed} failed)", flush=True)
    finally:
        stop_heartbeat.set()
        heartbeat_thread.join()

    if failed:
        print(f"Warning: {failed} item(s) could not be fetched after retries and were skipped.", file=sys.stderr, flush=True)

    return items


def find_custom_field(item, *label_fragments):
    for field in item.get("custom_fields") or []:
        label = (field.get("label") or "").strip().lower()
        if any(fragment in label for fragment in label_fragments):
            value = field.get("value")
            if value:
                return value
    return ""


def extract_barcode(item):
    for field_name in BARCODE_FIELDS:
        value = item.get(field_name)
        if value:
            return value
    return find_custom_field(item, "bar code", "barcode")


def extract_brand(item):
    return item.get("brand") or find_custom_field(item, "brand")


def build_rows(items, warehouse_name):
    warehouse_name = warehouse_name.strip().lower()
    rows = []
    for item in items:
        for warehouse in item.get("warehouses") or []:
            if (warehouse.get("warehouse_name") or "").strip().lower() != warehouse_name:
                continue
            rows.append(
                {
                    "BAR CODE": extract_barcode(item),
                    "SKU": item.get("sku", ""),
                    "Item Name": item.get("name", ""),
                    "SOH": warehouse.get("warehouse_stock_on_hand", ""),
                    "Sales Price": item.get("rate", ""),
                    "Brand": extract_brand(item),
                }
            )
            break
    return rows


def write_excel(rows, output_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(OUTPUT_COLUMNS)
    for row in rows:
        sheet.append([row[column] for column in OUTPUT_COLUMNS])
    workbook.save(output_path)


def main():
    warehouse_name = os.environ.get("WAREHOUSE_NAME", "Store 1")
    report_tz = ZoneInfo(os.environ.get("REPORT_TZ", "Africa/Nairobi"))

    # This job runs early morning after Store 1 has closed, so the report
    # date is "yesterday" in the business's local timezone.
    report_date = datetime.now(report_tz) - timedelta(days=1)
    output_path = f"SOH_{report_date.strftime('%d_%m_%Y')}.xlsx"

    rate_limit_per_minute = int(os.environ.get("RATE_LIMIT_PER_MINUTE", DEFAULT_RATE_LIMIT_PER_MINUTE))
    rate_limiter = RateLimiter(rate_limit_per_minute)

    access_token, api_domain = get_access_token()
    token_store = TokenStore(access_token)
    session = requests.Session()
    session.headers["Authorization"] = f"Zoho-oauthtoken {access_token}"

    organization_id = resolve_organization_id(session, api_domain)
    items = fetch_active_items(session, api_domain, organization_id, token_store, rate_limiter)
    rows = build_rows(items, warehouse_name)

    if not rows:
        seen_warehouses = sorted(
            {
                warehouse.get("warehouse_name", "")
                for item in items
                for warehouse in (item.get("warehouses") or [])
            }
        )
        print(
            f"Warning: no active items found stocked at warehouse '{warehouse_name}'. "
            f"Checked {len(items)} active item(s); warehouse names seen: {seen_warehouses}",
            file=sys.stderr,
        )

    write_excel(rows, output_path)
    print(f"Wrote {len(rows)} rows to {output_path}")


if __name__ == "__main__":
    main()
