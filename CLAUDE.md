# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

TLA Tierlist is a utility script for synchronizing MTG Arena "The Last Airbender" card data between Scryfall and Google Sheets. The script reads card names from a Google Sheet, fetches rarity information from the Scryfall API, and populates the adjacent column with single-letter rarity codes.

## Commands

### Installation

Install required dependencies:
```bash
pip install google-api-python-client google-auth google-auth-oauthlib requests
```

### Running the Script

Basic usage with existing Google Sheet:
```bash
python scripts/last_airbender_sync.py \
  --sheet-url "https://docs.google.com/spreadsheets/d/<spreadsheet-id>/edit" \
  --worksheet "Sheet1" \
  --column A \
  --start-row 2
```

Create a new spreadsheet:
```bash
python scripts/last_airbender_sync.py \
  --create-title "My Card List"
```

Dry run (preview changes without modifying sheet):
```bash
python scripts/last_airbender_sync.py \
  --sheet-url "<url>" \
  --dry-run
```

Use different Scryfall set:
```bash
python scripts/last_airbender_sync.py \
  --sheet-url "<url>" \
  --set-code "mkm"
```

### Authentication Setup

The script requires Google Sheets API credentials via one of these environment variables:

**Service Account** (recommended for automation):
```bash
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/service-account.json"
```

**OAuth Client** (for interactive use):
```bash
export GOOGLE_OAUTH_CLIENT_SECRETS="/path/to/client_secrets.json"
export GOOGLE_OAUTH_TOKEN="token.json"  # optional, defaults to token.json
```

## Architecture

### Data Flow

1. **Authentication**: Loads Google credentials from service account JSON or OAuth client secrets
2. **Sheet Access**: Parses spreadsheet ID from URL or creates new spreadsheet
3. **Scryfall Fetch**: Queries Scryfall API for all cards in the set, handling pagination and rate limits
4. **Rarity Mapping**: Converts Scryfall rarities (`mythic`, `rare`, `uncommon`, `common`) to single letters (`M`, `R`, `U`, `C`)
5. **Sheet Update**: Writes rarity codes to the column adjacent to card names using batch update API

### Key Components

**Column Conversion**: `column_to_index()` and `index_to_column()` convert between Excel-style column letters (A, B, AA, etc.) and numeric indices. The script automatically writes to the column immediately after the name column.

**Scryfall Pagination**: `fetch_scryfall_cards()` follows Scryfall's pagination using the `has_more` and `next_page` fields in API responses, accumulating all cards from the set.

**Rate Limiting**: Handles Scryfall's 429 responses by sleeping for the duration specified in the `Retry-After` header, with a maximum of 5 retry attempts. Also backs off exponentially on 5xx server errors.

**Batch Updates**: Uses Google Sheets `batchUpdate` API to write all rarity codes in a single request rather than individual cell writes, improving performance and reducing quota usage.

**Error Reporting**: Returns a `SyncResult` containing the count of updated rows and a list of card names that couldn't be found in Scryfall data. Missing cards are printed at the end of execution.

## Important Conventions

- The script expects card names in the spreadsheet to exactly match Scryfall names (case-sensitive)
- `--start-row` defaults to 2, assuming row 1 contains headers
- Dry run mode fetches all data and computes updates but skips the write operation
- When creating a new spreadsheet, the script prints the URL so you can access it
- Empty rows in the sheet are skipped during processing
