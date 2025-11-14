#!/usr/bin/env python3
"""Synchronize MTG Arena The Last Airbender card rarities into a Google Sheet.

This script authenticates with the Google Sheets API using either service account or
OAuth credentials, fetches card rarities from the Scryfall API, and writes mapped
rarity codes next to each card name in the specified sheet.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import requests

try:
    from google.auth.exceptions import GoogleAuthError
except ImportError:  # pragma: no cover - google-auth not installed
    class GoogleAuthError(Exception):
        """Fallback exception if google-auth is not installed."""

try:
    from google.oauth2 import service_account
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError as exc:  # pragma: no cover - libs missing at runtime
    MISSING_GOOGLE_LIBS = exc
else:
    MISSING_GOOGLE_LIBS = None

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SCRYFALL_SEARCH_ENDPOINT = "https://api.scryfall.com/cards/search"
RARITY_MAP = {
    "mythic": "M",
    "rare": "R",
    "uncommon": "U",
    "common": "C",
}


def column_to_index(column: str) -> int:
    column = column.strip().upper()
    if not column:
        raise SheetSyncError("Column reference cannot be empty.")
    index = 0
    for char in column:
        if not ("A" <= char <= "Z"):
            raise SheetSyncError(f"Invalid column reference: {column}")
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def index_to_column(index: int) -> str:
    if index < 1:
        raise SheetSyncError("Column index must be >= 1.")
    chars: List[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        chars.append(chr(65 + remainder))
    return "".join(reversed(chars))


class SheetSyncError(Exception):
    """Base exception for sync failures."""


@dataclass
class CardRecord:
    """Card data fetched from Scryfall."""

    name: str
    rarity: str

    @property
    def rarity_code(self) -> Optional[str]:
        return RARITY_MAP.get(self.rarity.lower())


@dataclass
class SyncResult:
    """Summary of the synchronization run."""

    updated: int
    missing_cards: List[str]


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--sheet-url",
        help="Target Google Sheet URL. If omitted, a new spreadsheet will be created.",
    )
    parser.add_argument(
        "--worksheet",
        default="Sheet1",
        help="Worksheet/tab name inside the spreadsheet (default: Sheet1).",
    )
    parser.add_argument(
        "--column",
        default="A",
        help="Column containing card names (default: column A).",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help="Row number where card names begin (default: 2 to skip header).",
    )
    parser.add_argument(
        "--set-code",
        default="tla",
        help="Scryfall set code to sync (default: 'tla' for The Last Airbender).",
    )
    parser.add_argument(
        "--create-title",
        default="The Last Airbender Sync",
        help="Title to use when creating a new spreadsheet.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch data and show planned updates without modifying the sheet.",
    )
    return parser.parse_args(argv)


def require_google_libs() -> None:
    if MISSING_GOOGLE_LIBS is not None:
        raise SheetSyncError(
            "Required Google libraries are missing. Install google-auth, "
            "google-auth-oauthlib, and google-api-python-client."
        ) from MISSING_GOOGLE_LIBS


def get_credentials() -> Credentials:
    """Load Google credentials from service account or OAuth files."""

    require_google_libs()

    service_account_file = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if service_account_file and os.path.exists(service_account_file):
        return service_account.Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES
        )

    client_secrets_file = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRETS")
    token_path = os.environ.get("GOOGLE_OAUTH_TOKEN", "token.json")

    if client_secrets_file and os.path.exists(client_secrets_file):
        creds: Optional[Credentials] = None
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    client_secrets_file, SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open(token_path, "w", encoding="utf-8") as token_file:
                token_file.write(creds.to_json())
        return creds

    raise SheetSyncError(
        "No valid Google credentials found. Set GOOGLE_APPLICATION_CREDENTIALS to a "
        "service account JSON file or configure OAuth via GOOGLE_OAUTH_CLIENT_SECRETS."
    )


def build_sheets_service(creds: Credentials):
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def parse_sheet_id(sheet_url: str) -> Optional[str]:
    if not sheet_url:
        return None
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    return match.group(1) if match else None


def create_spreadsheet(service, title: str) -> Tuple[str, str]:
    body = {"properties": {"title": title}}
    try:
        response = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
    except HttpError as exc:  # pragma: no cover - network dependent
        raise SheetSyncError(f"Failed to create spreadsheet: {exc}") from exc
    sheet_id = response["spreadsheetId"]
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    return sheet_id, url


def fetch_scryfall_cards(set_code: str) -> Dict[str, CardRecord]:
    params = {"q": f"set:{set_code}"}
    cards: Dict[str, CardRecord] = {}
    session = requests.Session()
    url = SCRYFALL_SEARCH_ENDPOINT
    retries = 0

    while url:
        try:
            response = session.get(url, params=params if url == SCRYFALL_SEARCH_ENDPOINT else None)
        except requests.RequestException as exc:  # pragma: no cover - network dependent
            raise SheetSyncError(f"Failed to contact Scryfall: {exc}") from exc

        if response.status_code == 429:
            wait_time = int(response.headers.get("Retry-After", 2))
            time.sleep(wait_time)
            retries += 1
            if retries > 5:
                raise SheetSyncError("Exceeded retry attempts due to Scryfall rate limits.")
            continue
        elif 500 <= response.status_code < 600:
            retries += 1
            if retries > 5:
                raise SheetSyncError(
                    f"Scryfall returned {response.status_code} repeatedly. Aborting."
                )
            time.sleep(2 ** retries)
            continue
        elif not response.ok:
            raise SheetSyncError(
                f"Scryfall request failed with {response.status_code}: {response.text}"
            )

        retries = 0
        payload = response.json()
        for card in payload.get("data", []):
            name = (card.get("name") or "").strip()
            rarity = card.get("rarity")
            if not name or not rarity:
                continue
            cards[name] = CardRecord(name=name, rarity=rarity)

        if payload.get("has_more"):
            url = payload.get("next_page")
            params = None
        else:
            break

    return cards


def read_sheet_rows(
    service, spreadsheet_id: str, worksheet: str, column: str, start_row: int
) -> List[Tuple[int, str]]:
    if start_row < 1:
        raise SheetSyncError("start_row must be >= 1")
    range_notation = f"{worksheet}!{column}{start_row}:{column}"
    try:
        response = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_notation,
            majorDimension="ROWS",
        ).execute()
    except HttpError as exc:  # pragma: no cover - network dependent
        raise SheetSyncError(f"Failed to read sheet values: {exc}") from exc

    values = response.get("values", [])
    rows: List[Tuple[int, str]] = []
    for index, row in enumerate(values, start=start_row):
        if not row:
            continue
        rows.append((index, row[0]))
    return rows


def write_rarities(
    service,
    spreadsheet_id: str,
    worksheet: str,
    target_column: str,
    updates: List[Tuple[int, str]],
) -> None:
    if not updates:
        return

    data = []
    for row_index, rarity_code in updates:
        range_notation = f"{worksheet}!{target_column}{row_index}"
        data.append({"range": range_notation, "values": [[rarity_code]]})

    body = {"valueInputOption": "RAW", "data": data}
    try:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body
        ).execute()
    except HttpError as exc:  # pragma: no cover - network dependent
        raise SheetSyncError(f"Failed to update sheet values: {exc}") from exc


def sync_sheet(
    spreadsheet_id: str,
    worksheet: str,
    name_column: str,
    start_row: int,
    cards: Dict[str, CardRecord],
    service,
    dry_run: bool = False,
) -> SyncResult:
    rows = read_sheet_rows(service, spreadsheet_id, worksheet, name_column, start_row)
    target_column = index_to_column(column_to_index(name_column) + 1)

    updates: List[Tuple[int, str]] = []
    missing_cards: List[str] = []

    for row_index, card_name in rows:
        normalized_name = card_name.strip()
        if not normalized_name:
            continue
        card = cards.get(normalized_name)
        if not card:
            missing_cards.append(card_name)
            continue
        code = card.rarity_code
        if not code:
            missing_cards.append(card_name)
            continue
        updates.append((row_index, code))

    if dry_run:
        for row_index, code in updates:
            print(f"Would update row {row_index} to {code}")
    else:
        write_rarities(service, spreadsheet_id, worksheet, target_column, updates)

    return SyncResult(updated=len(updates), missing_cards=missing_cards)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    try:
        creds = get_credentials()
    except (SheetSyncError, GoogleAuthError) as exc:
        print(f"Credential error: {exc}", file=sys.stderr)
        return 1

    service = build_sheets_service(creds)

    sheet_id = parse_sheet_id(args.sheet_url) if args.sheet_url else None
    if not sheet_id:
        try:
            sheet_id, sheet_url = create_spreadsheet(service, args.create_title)
        except SheetSyncError as exc:
            print(exc, file=sys.stderr)
            return 1
        print(f"Created spreadsheet: {sheet_url}")

    try:
        cards = fetch_scryfall_cards(args.set_code)
    except SheetSyncError as exc:
        print(exc, file=sys.stderr)
        return 1

    try:
        result = sync_sheet(
            sheet_id,
            args.worksheet,
            args.column,
            args.start_row,
            cards,
            service,
            dry_run=args.dry_run,
        )
    except SheetSyncError as exc:
        print(exc, file=sys.stderr)
        return 1

    if result.missing_cards:
        print("The following cards were not found:")
        for name in result.missing_cards:
            print(f" - {name}")

    print(f"Updated {result.updated} rows.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
