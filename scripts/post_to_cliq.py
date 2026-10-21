#!/usr/bin/env python3
"""Posts a generated report file into a Zoho Cliq channel.

Usage: post_to_cliq.py <file_path>

Required environment variables:
  ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET   - same Zoho Self Client used for Inventory
  ZOHO_CLIQ_REFRESH_TOKEN              - refresh token generated with the
                                          ZohoCliq.Webhooks.CREATE scope

Optional environment variables:
  CLIQ_ORG_ID              - defaults to the Kingdom Trading org (941419808)
  CLIQ_CHANNEL_UNIQUE_NAME - defaults to "salesmarketingug" (#Sales & Marketing UG)
"""
import os
import sys

import requests

ACCOUNTS_TOKEN_URL = "https://accounts.zoho.com/oauth/v2/token"
DEFAULT_ORG_ID = "941419808"
DEFAULT_CHANNEL_UNIQUE_NAME = "salesmarketingug"


def get_access_token():
    response = requests.post(
        ACCOUNTS_TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": os.environ["ZOHO_CLIENT_ID"],
            "client_secret": os.environ["ZOHO_CLIENT_SECRET"],
            "refresh_token": os.environ["ZOHO_CLIQ_REFRESH_TOKEN"],
        },
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError(f"No access_token in Cliq token refresh response: {payload}")
    return access_token


def post_file(file_path):
    org_id = os.environ.get("CLIQ_ORG_ID", DEFAULT_ORG_ID)
    channel_unique_name = os.environ.get("CLIQ_CHANNEL_UNIQUE_NAME", DEFAULT_CHANNEL_UNIQUE_NAME)
    access_token = get_access_token()

    url = f"https://cliq.zoho.com/company/{org_id}/api/v2/channelsbyname/{channel_unique_name}/files"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}

    filename = os.path.basename(file_path)
    with open(file_path, "rb") as fh:
        files = {
            "file": (
                filename,
                fh,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        }
        response = requests.post(url, headers=headers, files=files, timeout=60)

    if response.status_code >= 300:
        raise RuntimeError(f"Cliq file upload failed ({response.status_code}): {response.text}")

    print(f"Posted {filename} to Cliq channel '{channel_unique_name}'")


def main():
    if len(sys.argv) < 2:
        print("Usage: post_to_cliq.py <file_path>", file=sys.stderr)
        sys.exit(1)

    post_file(sys.argv[1])


if __name__ == "__main__":
    main()
