# Orion Discord Bot — Initial Build

This first Orion build contains:

- `/accept-cadet`
- `/help`
- Cadet role and nickname management
- CT-number lookup/allocation
- `Cadets` tab entry
- `Game Activity` Cadets-section entry
- Duplicate Discord ID and IGN checks
- Bot-channel logging

## 1. Requirements

- Node.js 20 or newer
- A Discord bot token and application/client ID for Orion
- A Google service account with access to the Galactic Sim spreadsheet

## 2. Complete `.env`

The Discord server, channels and role IDs supplied for Orion are already filled in.

You still need to fill:

```env
DISCORD_TOKEN=
CLIENT_ID=
GOOGLE_SPREADSHEET_ID=
GOOGLE_CREDENTIALS_JSON=
```

Share the spreadsheet with the `client_email` from the Google service-account JSON and give it Editor access.

Keep `GOOGLE_CREDENTIALS_JSON` on one line. Newlines inside `private_key` must remain written as `\\n`.

## 3. Install and register commands

Open a terminal in the Orion-Bot folder:

```bash
npm install
npm run deploy
npm start
```

`npm run deploy` registers the slash commands to guild `1209664377284730920`.

## `/accept-cadet` behaviour

Required options:

- `cadet`
- `in_game_name`
- `timezone`
- `ct_origin`

Optional options:

- `existing_ct_number`
- `allow_existing_ign`

For a mother-group member, select the existing-number route and provide a CT number already listed on the `CT Numbers` tab.

For a new member, Orion assigns the lowest available number between `53000` and `53999` whose adjacent Name cell is blank.

The Cadets `Time Served` formula is written as:

```excel
=TODAY()-G[row]
```

For example, a cadet written to row 16 receives `=TODAY()-G16`.

## Initial-build assumptions

- The `Cadets` tab has an available pre-formatted empty row.
- The CT Numbers tab keeps its repeated `Name | Number` blocks.
- Existing/mother-group CT numbers must already be present on the CT Numbers tab.
- The `Game Activity` tab contains a section header named `Cadets`.
- `Form responses 1` is not read or modified.
