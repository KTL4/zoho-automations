"""
SOH Daily Report

Pulls current stock-on-hand for a given warehouse from Zoho Inventory and
writes it to an Excel file named SOH_DD_MM_YYYY.xlsx.

Zoho Inventory's API has no direct "Stock Summary" report endpoint, so this
approximates it using the Items API: for every active item, it reads the
per-warehouse stock figures embedded in that item's `warehouses` array and
keeps the entry matching WAREHOUSE_NAME.

Required environment variables:
    ZOHO_CLIENT_ID
    ZOHO_CLIENT_SECRET
    ZOHO_REFRESH_TOKEN

Optional environment variables:
    ZOHO_ORGANIZATION_ID  - required if the Zoho account has more than one
                             organization (auto-detected otherwise)
    WAREHOUSE_NAME         - defaults to "Store 1"
    REPORT_TZ              - IANA timezone for naming the output file,
                              defaults to "Africa/Nairobi"
"""
import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook

ACCOUNTS_TOKEN_URL = "https://accounts.zoho.com/oauth/v2/token"
OUTPUT_COLUMNS = ["BAR CODE", "SKU", "Item Name", "SOH", "Sales Price", "Brand"]
BARCODE_FIELDS = ("upc", "ean", "isbn", "part_number")


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


def fetch_active_items(session, api_domain, organization_id):
    items = []
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
        items.extend(payload.get("items", []))

        if not payload.get("page_context", {}).get("has_more_page"):
            break
        page += 1

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

    access_token, api_domain = get_access_token()
    session = requests.Session()
    session.headers["Authorization"] = f"Zoho-oauthtoken {access_token}"

    organization_id = resolve_organization_id(session, api_domain)
    items = fetch_active_items(session, api_domain, organization_id)
    rows = build_rows(items, warehouse_name)

    if not rows:
        print(
            f"Warning: no active items found stocked at warehouse '{warehouse_name}'.",
            file=sys.stderr,
        )

    write_excel(rows, output_path)
    print(f"Wrote {len(rows)} rows to {output_path}")


if __name__ == "__main__":
    main()
