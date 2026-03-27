require('dotenv').config();
const { google } = require('googleapis');

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const TRAINEE_SPREADSHEET_ID = (process.env.TRAINEE_SPREADSHEET_ID || '').trim();
const RATINGS_SPREADSHEET_ID = (
  process.env.RATINGS_SPREADSHEET_ID ||
  process.env.MOS_SPREADSHEET_ID ||
  process.env.TRAINEE_SPREADSHEET_ID ||
  ''
).trim();

const TRAINEE_SHEET_NAME = (process.env.TRAINEE_SHEET_NAME || 'Trainees').trim();
const RATINGS_SHEET_NAME = (process.env.RATINGS_SHEET_NAME || 'Ratings').trim();

async function getSheetsClient() {
  const authClient = await auth.getClient();
  return google.sheets({
    version: 'v4',
    auth: authClient,
  });
}

async function getSpreadsheetMeta(spreadsheetId = TRAINEE_SPREADSHEET_ID) {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  return response.data;
}

/**
 * Normalises names for safe matching between Discord and Sheets.
 * Handles:
 * - rank prefixes (full + abbreviated)
 * - [clan tags]
 * - quoted nicknames like "PMC"
 * - (Ex Skira)
 * - punctuation / symbols
 * - spacing / case differences
 */
function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+\(ex skira\)$/i, '')
    .replace(/\s+\[[^\]]+\]\s*/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(
      /^(tr|pvt|cdt|o\/cdt|tpr|p\/o|lcpl|cpl|sgt|ssgt|lt|f\/o|2lt|wo1|wo2|cpt|maj|col|brig|lieutenant|captain|major|colonel|brigadier|flight lieutenant|squadron leader)\s+/i,
      ''
    )
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

async function getTraineeRows() {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: TRAINEE_SPREADSHEET_ID,
    range: `${TRAINEE_SHEET_NAME}!A:I`,
  });

  return response.data.values || [];
}

async function getRatingsRows() {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
    range: `${RATINGS_SHEET_NAME}!A:S`,
  });

  return response.data.values || [];
}

function isPlaceholderOrEmpty(value) {
  if (!value) return true;

  const cleaned = String(value).trim().toLowerCase();

  return (
    cleaned === '' ||
    cleaned === 'name' ||
    cleaned === 'dd/mm/yyyy' ||
    cleaned === 'steamid64' ||
    cleaned === 'notes'
  );
}

/* ---------------- TRAINEE ---------------- */

async function findTraineeRowByDiscordId(discordId) {
  const rows = await getTraineeRows();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowDiscordId = (row[8] || '').toString().trim();

    if (rowDiscordId === discordId) {
      return {
        rowNumber: i + 1,
        rowValues: row,
      };
    }
  }

  return null;
}

async function findTraineeRowBySteamId64(steamId64) {
  const rows = await getTraineeRows();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowSteamId64 = (row[3] || '').toString().trim();

    if (rowSteamId64 === steamId64) {
      return {
        rowNumber: i + 1,
        rowValues: row,
      };
    }
  }

  return null;
}

async function findTraineeRowsByName(name) {
  const rows = await getTraineeRows();
  const target = normalizeName(name);
  const matches = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowName = normalizeName(row[0] || '');

    if (rowName && rowName === target) {
      matches.push({
        rowNumber: i + 1,
        rowValues: row,
      });
    }
  }

  return matches;
}

async function writeTraineeRow(values) {
  const sheets = await getSheetsClient();
  const rows = await getTraineeRows();

  let targetRowNumber = null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const nameCell = row[0] || '';
    const dateCell = row[1] || '';
    const steamCell = row[3] || '';
    const notesCell = row[7] || '';

    if (
      isPlaceholderOrEmpty(nameCell) &&
      isPlaceholderOrEmpty(dateCell) &&
      isPlaceholderOrEmpty(steamCell) &&
      isPlaceholderOrEmpty(notesCell)
    ) {
      targetRowNumber = i + 1;
      break;
    }
  }

  if (!targetRowNumber) {
    targetRowNumber = rows.length + 1;
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: TRAINEE_SPREADSHEET_ID,
    range: `${TRAINEE_SHEET_NAME}!A${targetRowNumber}:I${targetRowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values],
    },
  });

  return targetRowNumber;
}

async function updateTraineeCell(rowNumber, columnLetter, value) {
  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.update({
    spreadsheetId: TRAINEE_SPREADSHEET_ID,
    range: `${TRAINEE_SHEET_NAME}!${columnLetter}${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[value]],
    },
  });
}

async function batchUpdateTraineeCells(updates) {
  if (!updates || updates.length === 0) return;

  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: TRAINEE_SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates.map(update => ({
        range: update.range,
        values: update.values,
      })),
    },
  });
}

async function deleteTraineeRow(rowNumber) {
  const sheets = await getSheetsClient();
  const spreadsheet = await getSpreadsheetMeta(TRAINEE_SPREADSHEET_ID);

  const traineeSheet = spreadsheet.sheets.find(
    s => s.properties.title === TRAINEE_SHEET_NAME
  );

  if (!traineeSheet) {
    throw new Error(`${TRAINEE_SHEET_NAME} sheet not found.`);
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: TRAINEE_SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: traineeSheet.properties.sheetId,
              dimension: 'ROWS',
              startIndex: rowNumber - 1,
              endIndex: rowNumber,
            },
          },
        },
      ],
    },
  });
}

/* ---------------- RATINGS ---------------- */

async function findRatingsRowByDiscordId(discordId) {
  const rows = await getRatingsRows();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowDiscordId = (row[18] || '').toString().trim();

    if (rowDiscordId === discordId) {
      return {
        rowNumber: i + 1,
        rowValues: row,
      };
    }
  }

  return null;
}

async function findRatingsRowsByName(name) {
  const rows = await getRatingsRows();
  const target = normalizeName(name);
  const matches = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowName = normalizeName(row[2] || '');

    if (rowName && rowName === target) {
      matches.push({
        rowNumber: i + 1,
        rowValues: row,
      });
    }
  }

  return matches;
}

async function writeRatingsRow(values) {
  const sheets = await getSheetsClient();
  const rows = await getRatingsRows();

  let targetRowNumber = null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const nameCell = (row[2] || '').toString().trim();
    const discordIdCell = (row[18] || '').toString().trim();

    if (!nameCell && !discordIdCell) {
      targetRowNumber = i + 1;
      break;
    }
  }

  if (!targetRowNumber) {
    targetRowNumber = rows.length + 1;
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
    range: `${RATINGS_SHEET_NAME}!A${targetRowNumber}:S${targetRowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values],
    },
  });

  return targetRowNumber;
}

async function updateRatingsCell(rowNumber, columnLetter, value) {
  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.update({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
    range: `${RATINGS_SHEET_NAME}!${columnLetter}${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[value]],
    },
  });
}

async function batchUpdateRatingsCells(updates) {
  if (!updates || updates.length === 0) return;

  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates.map(update => ({
        range: update.range,
        values: update.values,
      })),
    },
  });
}

module.exports = {
  normalizeName,
  getTraineeRows,
  getRatingsRows,
  findTraineeRowByDiscordId,
  findTraineeRowBySteamId64,
  findTraineeRowsByName,
  findRatingsRowByDiscordId,
  findRatingsRowsByName,
  writeTraineeRow,
  writeRatingsRow,
  updateTraineeCell,
  updateRatingsCell,
  batchUpdateTraineeCells,
  batchUpdateRatingsCells,
  deleteTraineeRow,
};