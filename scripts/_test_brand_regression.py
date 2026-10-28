"""One-off diagnostic: why is Brand blank with the bulk /itemdetails fetch?

Compares custom_fields (and any 'brand' field) between the single-item
endpoint and the bulk /itemdetails endpoint for the same real item ids.
Not part of the production pipeline - delete after use.
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

    list_response = session.get(
        f"{api_domain}/inventory/v1/items",
        params={"organization_id": org_id, "filter_by": "Status.Active", "page": 1, "per_page": 3},
        timeout=30,
    )
    list_response.raise_for_status()
    item_ids = [item["item_id"] for item in list_response.json()["items"]]
    print("Testing with item_ids:", item_ids)

    for item_id in item_ids:
        print(f"\n=== item_id={item_id} ===")

        single = session.get(
            f"{api_domain}/inventory/v1/items/{item_id}",
            params={"organization_id": org_id},
            timeout=30,
        ).json()["item"]
        print("SINGLE-ITEM custom_fields:")
        print(json.dumps(single.get("custom_fields"), indent=2))
        print("SINGLE-ITEM top-level 'brand' key present:", "brand" in single, "value:", single.get("brand"))

        bulk = session.get(
            f"{api_domain}/inventory/v1/itemdetails",
            params={"organization_id": org_id, "item_ids": item_id},
            timeout=30,
        ).json()
        bulk_items = bulk.get("items", [])
        if not bulk_items:
            print("BULK: no items returned!", bulk)
            continue
        bulk_item = bulk_items[0]
        print("BULK custom_fields:")
        print(json.dumps(bulk_item.get("custom_fields"), indent=2))
        print("BULK top-level 'brand' key present:", "brand" in bulk_item, "value:", bulk_item.get("brand"))


if __name__ == "__main__":
    main()
