"""One-off test: does GET /itemdetails return per-warehouse stock in bulk?

Not part of the production pipeline - run manually via a throwaway
workflow_dispatch, inspect the output, then delete this file either way.
"""
import json
import os

import requests

ACCOUNTS_TOKEN_URL = "https://accounts.zoho.com/oauth/v2/token"


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
    payload = response.json()
    return payload["access_token"], payload["api_domain"]


def main():
    access_token, api_domain = get_access_token()
    session = requests.Session()
    session.headers["Authorization"] = f"Zoho-oauthtoken {access_token}"

    org_id = os.environ.get("ZOHO_ORGANIZATION_ID")
    if not org_id:
        orgs = session.get(f"{api_domain}/inventory/v1/organizations", timeout=30).json()
        org_id = orgs["organizations"][0]["organization_id"]

    # Grab 5 real active item ids to test with.
    list_response = session.get(
        f"{api_domain}/inventory/v1/items",
        params={"organization_id": org_id, "filter_by": "Status.Active", "page": 1, "per_page": 5},
        timeout=30,
    )
    list_response.raise_for_status()
    item_ids = [item["item_id"] for item in list_response.json()["items"]]
    print("Testing with item_ids:", item_ids)

    bulk_response = session.get(
        f"{api_domain}/inventory/v1/itemdetails",
        params={"organization_id": org_id, "item_ids": ",".join(item_ids)},
        timeout=30,
    )
    print("Status code:", bulk_response.status_code)
    print("Response headers content-type:", bulk_response.headers.get("content-type"))
    try:
        payload = bulk_response.json()
    except ValueError:
        print("Non-JSON response body (first 1000 chars):")
        print(bulk_response.text[:1000])
        return

    print("Top-level keys:", list(payload.keys()))
    items = payload.get("items") or payload.get("itemdetails") or []
    print(f"Items returned: {len(items)}")
    if items:
        first = items[0]
        print("First item top-level keys:", list(first.keys()))
        print("Has 'warehouses' key:", "warehouses" in first)
        print("First item (truncated JSON):")
        print(json.dumps(first, indent=2)[:2000])


if __name__ == "__main__":
    main()
