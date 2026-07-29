require('dotenv').config();
const { google } = require('googleapis');

const ID = () =>
  String(process.env.GOOGLE_SPREADSHEET_ID || '').trim();

const TABS = () => ({
  cadets: process.env.CADETS_SHEET_NAME || 'Cadets',
  roster: process.env.ROSTER_SHEET_NAME || 'Roster',
  ct: process.env.CT_NUMBERS_SHEET_NAME || 'CT Numbers',
  activity:
    process.env.GAME_ACTIVITY_SHEET_NAME ||
    'Game Activity',
});

let api;

function credentials() {
  try {
    return JSON.parse(
      String(
        process.env.GOOGLE_CREDENTIALS_JSON || ''
      ).trim()
    );
  } catch (error) {
    throw new Error(
      `Invalid GOOGLE_CREDENTIALS_JSON: ${error.message}`
    );
  }
}

async function sheets() {
  if (api) {
    return api;
  }

  const auth = new google.auth.GoogleAuth({
    credentials: credentials(),
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
    ],
  });

  api = google.sheets({
    version: 'v4',
    auth: await auth.getClient(),
  });

  return api;
}

const q = name =>
  `'${String(name).replace(/'/g, "''")}'`;

const norm = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');

const headerNorm = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

const isIgn = value =>
  ['ign', 'name'].includes(headerNorm(value));

const isNumber = value =>
  ['ct number', 'number'].includes(
    headerNorm(value)
  );

const isDiscord = value =>
  ['discord id', 'user id'].includes(
    headerNorm(value)
  );

function col(index) {
  let number = index + 1;
  let output = '';

  while (number > 0) {
    output =
      String.fromCharCode(
        65 + ((number - 1) % 26)
      ) + output;

    number = Math.floor(
      (number - 1) / 26
    );
  }

  return output;
}

function header(rows, required) {
  const wanted = required.map(headerNorm);

  for (
    let rowIndex = 0;
    rowIndex < Math.min(rows.length, 30);
    rowIndex++
  ) {
    const cells = (
      rows[rowIndex] || []
    ).map(headerNorm);

    if (
      wanted.every(item =>
        cells.includes(item)
      )
    ) {
      return {
        index: rowIndex,

        columns: Object.fromEntries(
          wanted.map(item => [
            item,
            cells.indexOf(item),
          ])
        ),
      };
    }
  }

  throw new Error(
    `Could not find headers: ${required.join(', ')}`
  );
}

async function read(
  range,
  render = 'FORMATTED_VALUE'
) {
  const client = await sheets();

  const response =
    await client.spreadsheets.values.get({
      spreadsheetId: ID(),
      range,
      valueRenderOption: render,
    });

  return response.data.values || [];
}

async function write(
  range,
  values,
  input = 'USER_ENTERED'
) {
  const client = await sheets();

  await client.spreadsheets.values.update({
    spreadsheetId: ID(),
    range,
    valueInputOption: input,

    requestBody: {
      values,
    },
  });
}

async function batchWrite(
  data,
  input = 'USER_ENTERED'
) {
  const client = await sheets();

  await client.spreadsheets.values.batchUpdate({
    spreadsheetId: ID(),

    requestBody: {
      valueInputOption: input,
      data,
    },
  });
}

async function clear(range) {
  const client = await sheets();

  await client.spreadsheets.values.clear({
    spreadsheetId: ID(),
    range,
    requestBody: {},
  });
}

async function meta(tab) {
  const client = await sheets();

  const response =
    await client.spreadsheets.get({
      spreadsheetId: ID(),

      fields:
        'sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))',
    });

  const found = (
    response.data.sheets || []
  ).find(
    sheet =>
      sheet.properties?.title === tab
  );

  if (!found) {
    throw new Error(
      `Sheet tab not found: ${tab}`
    );
  }

  return found.properties;
}

async function findPersonByDiscordId(
  tab,
  idHeader,
  discordId
) {
  const rows = await read(
    `${q(tab)}!A1:Z3000`
  );

  const foundHeader = header(rows, [
    'Name',
    idHeader,
  ]);

  const nameColumn =
    foundHeader.columns.name;

  const idColumn =
    foundHeader.columns[
      headerNorm(idHeader)
    ];

  for (
    let rowIndex =
      foundHeader.index + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    if (
      String(
        rows[rowIndex]?.[idColumn] ||
          ''
      ).trim() === String(discordId)
    ) {
      return {
        sheetName: tab,

        rowNumber:
          rowIndex + 1,

        name: String(
          rows[rowIndex]?.[
            nameColumn
          ] || ''
        ).trim(),
      };
    }
  }

  return null;
}

const findCadetByDiscordId = id =>
  findPersonByDiscordId(
    TABS().cadets,
    'User ID',
    id
  );

const findRosterByDiscordId = id =>
  findPersonByDiscordId(
    TABS().roster,
    'Discord ID',
    id
  );

async function findNameRows(
  tab,
  idHeader,
  name
) {
  const rows = await read(
    `${q(tab)}!A1:Z3000`
  );

  const foundHeader = header(rows, [
    'Name',
    idHeader,
  ]);

  const nameColumn =
    foundHeader.columns.name;

  const idColumn =
    foundHeader.columns[
      headerNorm(idHeader)
    ];

  const target = norm(name);
  const matches = [];

  for (
    let rowIndex =
      foundHeader.index + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const foundName = String(
      rows[rowIndex]?.[nameColumn] ||
        ''
    ).trim();

    if (
      foundName &&
      norm(foundName) === target
    ) {
      matches.push({
        sheetName: tab,

        rowNumber:
          rowIndex + 1,

        name: foundName,

        discordId: String(
          rows[rowIndex]?.[idColumn] ||
            ''
        ).trim(),
      });
    }
  }

  return matches;
}

const findNameOccurrencesInCadets =
  name =>
    findNameRows(
      TABS().cadets,
      'User ID',
      name
    );

const findNameOccurrencesInRoster =
  name =>
    findNameRows(
      TABS().roster,
      'Discord ID',
      name
    );

async function ctIndex() {
  const tab = TABS().ct;

  const rows = await read(
    `${q(tab)}!A1:ZZ102`
  );

  const blocks = [];

  for (
    let rowIndex = 0;
    rowIndex <
    Math.min(rows.length, 20);
    rowIndex++
  ) {
    const row = rows[rowIndex] || [];

    for (
      let columnIndex = 0;
      columnIndex < row.length - 2;
      columnIndex++
    ) {
      if (
        isIgn(row[columnIndex]) &&
        isNumber(
          row[columnIndex + 1]
        ) &&
        isDiscord(
          row[columnIndex + 2]
        )
      ) {
        blocks.push({
          headerRow: rowIndex,
          ignCol: columnIndex,

          numberCol:
            columnIndex + 1,

          discordCol:
            columnIndex + 2,
        });
      }
    }
  }

  if (!blocks.length) {
    throw new Error(
      'No IGN / CT Number / Discord ID blocks were found on CT Numbers.'
    );
  }

  const entries = [];

  for (const block of blocks) {
    for (
      let rowIndex =
        block.headerRow + 1;
      rowIndex <
      Math.min(rows.length, 102);
      rowIndex++
    ) {
      const number = String(
        rows[rowIndex]?.[
          block.numberCol
        ] ?? ''
      ).trim();

      if (!number) {
        continue;
      }

      entries.push({
        sheetName: tab,

        rowNumber:
          rowIndex + 1,

        ign: String(
          rows[rowIndex]?.[
            block.ignCol
          ] ?? ''
        ).trim(),

        name: String(
          rows[rowIndex]?.[
            block.ignCol
          ] ?? ''
        ).trim(),

        number,

        discordId: String(
          rows[rowIndex]?.[
            block.discordCol
          ] ?? ''
        ).trim(),

        ignCell:
          `${col(block.ignCol)}` +
          `${rowIndex + 1}`,

        numberCell:
          `${col(block.numberCol)}` +
          `${rowIndex + 1}`,

        discordIdCell:
          `${col(block.discordCol)}` +
          `${rowIndex + 1}`,
      });
    }
  }

  return {
    tab,
    rows,
    blocks,
    entries,
  };
}

async function findCtEntriesByName(
  name
) {
  const index = await ctIndex();

  return index.entries.filter(
    entry =>
      norm(entry.ign) === norm(name)
  );
}

async function findCtEntriesByNumber(
  number
) {
  const index = await ctIndex();

  return index.entries.filter(
    entry =>
      entry.number ===
      String(number || '').trim()
  );
}

async function findCtEntriesByDiscordId(
  discordId
) {
  const index = await ctIndex();

  return index.entries.filter(
    entry =>
      entry.discordId ===
      String(discordId || '').trim()
  );
}

async function findNextAvailableLocalCtNumber() {
  const index = await ctIndex();

  const available =
    index.entries
      .filter(entry => {
        const value =
          Number(entry.number);

        return (
          !entry.ign &&
          !entry.discordId &&
          /^\d{5}$/.test(
            entry.number
          ) &&
          value >= 53000 &&
          value <= 53999
        );
      })
      .sort(
        (first, second) =>
          Number(first.number) -
          Number(second.number)
      );

  if (!available.length) {
    throw new Error(
      'No available local CT number exists between 53000 and 53999.'
    );
  }

  return available[0];
}

async function createCustomCtEntry(
  ign,
  number,
  discordId
) {
  const tab = TABS().ct;

  const rows = await read(
    `${q(tab)}!A1:C102`
  );

  let headerRow = -1;

  for (
    let rowIndex = 0;
    rowIndex <
    Math.min(rows.length, 20);
    rowIndex++
  ) {
    if (
      isIgn(rows[rowIndex]?.[0]) &&
      isNumber(
        rows[rowIndex]?.[1]
      ) &&
      isDiscord(
        rows[rowIndex]?.[2]
      )
    ) {
      headerRow = rowIndex;
      break;
    }
  }

  if (headerRow < 0) {
    throw new Error(
      'The custom CT block was not found in columns A:C.'
    );
  }

  let targetRow = -1;

  for (
    let rowIndex = headerRow + 1;
    rowIndex < 102;
    rowIndex++
  ) {
    const hasValue = [0, 1, 2].some(
      columnIndex =>
        String(
          rows[rowIndex]?.[
            columnIndex
          ] || ''
        ).trim()
    );

    if (!hasValue) {
      targetRow = rowIndex;
      break;
    }
  }

  if (targetRow < 0) {
    throw new Error(
      'No empty row remains in the custom CT block.'
    );
  }

  const rowNumber = targetRow + 1;

  await write(
    `${q(tab)}!A${rowNumber}:C${rowNumber}`,

    [[
      ign,
      String(number),
      String(discordId),
    ]],

    'RAW'
  );

  return {
    sheetName: tab,
    rowNumber,
    ign,
    name: ign,
    number: String(number),
    discordId: String(discordId),
    ignCell: `A${rowNumber}`,
    numberCell: `B${rowNumber}`,
    discordIdCell: `C${rowNumber}`,
  };
}

async function updateCtIdentity(
  entry,
  ign,
  discordId
) {
  await batchWrite(
    [
      {
        range:
          `${q(entry.sheetName)}!` +
          `${entry.ignCell}`,

        values: [[
          String(ign || '').trim(),
        ]],
      },

      {
        range:
          `${q(entry.sheetName)}!` +
          `${entry.discordIdCell}`,

        values: [[
          String(
            discordId || ''
          ).trim(),
        ]],
      },
    ],

    'RAW'
  );
}

async function clearCtIdentity(entry) {
  const client = await sheets();

  await client.spreadsheets.values.batchClear({
    spreadsheetId: ID(),

    requestBody: {
      ranges: [
        `${q(entry.sheetName)}!${entry.ignCell}`,

        `${q(entry.sheetName)}!${entry.discordIdCell}`,
      ],
    },
  });
}

async function clearCustomCtEntry(entry) {
  await clear(
    `${q(entry.sheetName)}!` +
    `${entry.ignCell}:` +
    `${entry.discordIdCell}`
  );
}

async function addCadetRow({
  inGameName,
  timezone,
  discordUsername,
  discordId,
  joinedDate,
}) {
  const tab = TABS().cadets;

  const rows = await read(
    `${q(tab)}!A1:Z2000`
  );

  const foundHeader = header(rows, [
    'Rank',
    'Name',
    'Timezone',
    'Discord Username',
    'User ID',
    'Joined',
    'BCT',
    'Play Session 1',
    'Play Session 2',
    'Play Session 3',
    'Time Served',
  ]);

  const metadata = await meta(tab);

  let targetRow = -1;

  for (
    let rowIndex =
      foundHeader.index + 1;
    rowIndex <
    metadata.gridProperties.rowCount;
    rowIndex++
  ) {
    const nameBlank = !String(
      rows[rowIndex]?.[
        foundHeader.columns.name
      ] || ''
    ).trim();

    const idBlank = !String(
      rows[rowIndex]?.[
        foundHeader.columns[
          'user id'
        ]
      ] || ''
    ).trim();

    if (nameBlank && idBlank) {
      targetRow = rowIndex;
      break;
    }
  }

  if (targetRow < 0) {
    throw new Error(
      'No empty formatted row is available on Cadets.'
    );
  }

  const rowNumber = targetRow + 1;

  const values = {
    rank: 'CDT',
    name: inGameName,
    timezone,

    'discord username':
      discordUsername,

    'user id': discordId,
    joined: joinedDate,
    bct: false,

    'play session 1':
      false,

    'play session 2':
      false,

    'play session 3':
      false,

    'time served':
      `=TODAY()-G${rowNumber}`,
  };

  await batchWrite(
    Object.entries(values).map(
      ([key, value]) => ({
        range:
          `${q(tab)}!` +
          `${col(
            foundHeader.columns[key]
          )}` +
          `${rowNumber}`,

        values: [[value]],
      })
    )
  );

  return {
    sheetName: tab,
    rowNumber,

    clearRange:
      `A${rowNumber}:K${rowNumber}`,
  };
}

async function clearCadetRow(record) {
  await clear(
    `${q(record.sheetName)}!` +
    `${record.clearRange}`
  );
}

async function backupCadetRow(
  rowNumber
) {
  const range =
    `${q(TABS().cadets)}!` +
    `A${rowNumber}:K${rowNumber}`;

  const [display, formulas] =
    await Promise.all([
      read(range),
      read(range, 'FORMULA'),
    ]);

  const shown = display[0] || [];
  const raw = formulas[0] || [];

  return {
    range,

    values: Array.from(
      { length: 11 },

      (_, index) =>
        typeof raw[index] ===
          'string' &&
        raw[index].startsWith('=')
          ? raw[index]
          : shown[index] ?? ''
    ),
  };
}

async function restoreCadetRow(
  backup
) {
  await write(
    backup.range,
    [backup.values]
  );
}

async function findGameActivityCadetRow(
  ign
) {
  const tab = TABS().activity;

  const rows = await read(
    `${q(tab)}!A1:B3000`
  );

  let section = -1;

  for (
    let rowIndex = 0;
    rowIndex < rows.length;
    rowIndex++
  ) {
    if (
      (rows[rowIndex] || []).some(
        cell =>
          headerNorm(cell) ===
          'cadets'
      )
    ) {
      section = rowIndex;
    }
  }

  if (section < 0) {
    throw new Error(
      'The Cadets section was not found on Game Activity.'
    );
  }

  const matches = [];

  for (
    let rowIndex = section + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const rank = String(
      rows[rowIndex]?.[0] || ''
    ).trim();

    const name = String(
      rows[rowIndex]?.[1] || ''
    ).trim();

    if (
      norm(name) === norm(ign) &&
      (
        !rank ||
        rank.toUpperCase() === 'CDT'
      )
    ) {
      matches.push({
        rowIndex,

        rowNumber:
          rowIndex + 1,
      });
    }
  }

  if (matches.length > 1) {
    throw new Error(
      `Multiple Game Activity rows exist for ${ign}.`
    );
  }

  if (!matches.length) {
    return null;
  }

  const metadata = await meta(tab);

  return {
    ...matches[0],
    sheetName: tab,

    sheetId:
      metadata.sheetId,
  };
}

async function addGameActivityCadetRow(
  ign
) {
  const tab = TABS().activity;

  const rows = await read(
    `${q(tab)}!A1:B3000`
  );

  const foundHeader = header(rows, [
    'Rank',
    'Name',
  ]);

  let section = -1;

  for (
    let rowIndex =
      foundHeader.index + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    if (
      (rows[rowIndex] || []).some(
        cell =>
          headerNorm(cell) ===
          'cadets'
      )
    ) {
      section = rowIndex;
    }
  }

  if (section < 0) {
    throw new Error(
      'The Cadets section was not found on Game Activity.'
    );
  }

  let lastMemberRow = section;

  for (
    let rowIndex = section + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const rank = String(
      rows[rowIndex]?.[
        foundHeader.columns.rank
      ] || ''
    ).trim();

    const name = String(
      rows[rowIndex]?.[
        foundHeader.columns.name
      ] || ''
    ).trim();

    if (rank || name) {
      lastMemberRow = rowIndex;
    }
  }

  const insertIndex =
    lastMemberRow + 1;

  const metadata = await meta(tab);
  const client = await sheets();

  const templateIndex = Math.max(
    insertIndex - 1,
    section + 1
  );

  await client.spreadsheets.batchUpdate({
    spreadsheetId: ID(),

    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId:
                metadata.sheetId,

              dimension: 'ROWS',

              startIndex:
                insertIndex,

              endIndex:
                insertIndex + 1,
            },

            inheritFromBefore: true,
          },
        },

        {
          copyPaste: {
            source: {
              sheetId:
                metadata.sheetId,

              startRowIndex:
                templateIndex,

              endRowIndex:
                templateIndex + 1,

              startColumnIndex: 0,

              endColumnIndex:
                metadata
                  .gridProperties
                  .columnCount,
            },

            destination: {
              sheetId:
                metadata.sheetId,

              startRowIndex:
                insertIndex,

              endRowIndex:
                insertIndex + 1,

              startColumnIndex: 0,

              endColumnIndex:
                metadata
                  .gridProperties
                  .columnCount,
            },

            pasteType:
              'PASTE_FORMAT',

            pasteOrientation:
              'NORMAL',
          },
        },

        {
          copyPaste: {
            source: {
              sheetId:
                metadata.sheetId,

              startRowIndex:
                templateIndex,

              endRowIndex:
                templateIndex + 1,

              startColumnIndex: 0,

              endColumnIndex:
                metadata
                  .gridProperties
                  .columnCount,
            },

            destination: {
              sheetId:
                metadata.sheetId,

              startRowIndex:
                insertIndex,

              endRowIndex:
                insertIndex + 1,

              startColumnIndex: 0,

              endColumnIndex:
                metadata
                  .gridProperties
                  .columnCount,
            },

            pasteType:
              'PASTE_DATA_VALIDATION',

            pasteOrientation:
              'NORMAL',
          },
        },
      ],
    },
  });

  const rowNumber =
    insertIndex + 1;

  try {
    await batchWrite(
      [
        {
          range:
            `${q(tab)}!` +
            `${col(
              foundHeader.columns.rank
            )}` +
            `${rowNumber}`,

          values: [['CDT']],
        },

        {
          range:
            `${q(tab)}!` +
            `${col(
              foundHeader.columns.name
            )}` +
            `${rowNumber}`,

          values: [[ign]],
        },
      ],

      'RAW'
    );
  } catch (error) {
    await client.spreadsheets
      .batchUpdate({
        spreadsheetId: ID(),

        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId:
                    metadata.sheetId,

                  dimension:
                    'ROWS',

                  startIndex:
                    insertIndex,

                  endIndex:
                    insertIndex + 1,
                },
              },
            },
          ],
        },
      })
      .catch(rollbackError =>
        console.error(
          'Game Activity rollback failed:',
          rollbackError
        )
      );

    throw error;
  }

  return {
    sheetName: tab,

    sheetId:
      metadata.sheetId,

    rowNumber,

    rowIndex:
      insertIndex,
  };
}

async function deleteGameActivityRow(
  record
) {
  const client = await sheets();

  await client.spreadsheets.batchUpdate({
    spreadsheetId: ID(),

    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId:
                record.sheetId,

              dimension: 'ROWS',

              startIndex:
                record.rowIndex,

              endIndex:
                record.rowIndex + 1,
            },
          },
        },
      ],
    },
  });
}

module.exports = {
  normalizeName: norm,

  findCadetByDiscordId,
  findRosterByDiscordId,

  findNameOccurrencesInCadets,
  findNameOccurrencesInRoster,

  findCtEntriesByName,
  findCtEntriesByNumber,
  findCtEntriesByDiscordId,

  findNextAvailableLocalCtNumber,

  createCustomCtEntry,
  updateCtIdentity,
  clearCtIdentity,
  clearCustomCtEntry,

  addCadetRow,
  clearCadetRow,
  backupCadetRow,
  restoreCadetRow,

  findGameActivityCadetRow,
  addGameActivityCadetRow,
  deleteGameActivityRow,
};