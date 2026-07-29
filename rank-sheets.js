require('dotenv').config();

const { google } = require('googleapis');

const {
  DESTINATIONS,
  ACTIVITY_SECTIONS,
  ROSTER_BOUNDARIES,
} = require('./rank-config');

const SPREADSHEET_ID = String(
  process.env.GOOGLE_SPREADSHEET_ID || ''
).trim();

const TABS = {
  cadets:
    process.env.CADETS_SHEET_NAME ||
    'Cadets',

  roster:
    process.env.ROSTER_SHEET_NAME ||
    'Roster',

  activity:
    process.env.GAME_ACTIVITY_SHEET_NAME ||
    'Game Activity',
};

let cachedClient = null;

function normaliseHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normaliseName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normaliseSection(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[’']/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function quoteSheet(name) {
  return `'${String(name).replace(
    /'/g,
    "''"
  )}'`;
}

function columnLetter(
  indexZeroBased
) {
  let number =
    indexZeroBased + 1;

  let output = '';

  while (number > 0) {
    output =
      String.fromCharCode(
        65 +
        ((number - 1) % 26)
      ) +
      output;

    number = Math.floor(
      (number - 1) / 26
    );
  }

  return output;
}

function padRow(
  values,
  length
) {
  return Array.from(
    {
      length,
    },

    (_, index) =>
      values[index] ?? ''
  );
}

function ensureSpreadsheetId() {
  if (!SPREADSHEET_ID) {
    throw new Error(
      'GOOGLE_SPREADSHEET_ID is missing from .env.'
    );
  }
}

function parseCredentials() {
  const raw =
    String(
      process.env
        .GOOGLE_CREDENTIALS_JSON ||
        ''
    ).trim();

  if (!raw) {
    throw new Error(
      'GOOGLE_CREDENTIALS_JSON is missing from .env.'
    );
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error(
      `GOOGLE_CREDENTIALS_JSON is invalid: ${error.message}`
    );
  }
}

async function sheetsClient() {
  if (cachedClient) {
    return cachedClient;
  }

  ensureSpreadsheetId();

  const auth =
    new google.auth.GoogleAuth({
      credentials:
        parseCredentials(),

      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
      ],
    });

  cachedClient =
    google.sheets({
      version: 'v4',

      auth:
        await auth.getClient(),
    });

  return cachedClient;
}

async function readValues(
  range,
  render = 'FORMATTED_VALUE'
) {
  const client =
    await sheetsClient();

  const response =
    await client.spreadsheets.values.get({
      spreadsheetId:
        SPREADSHEET_ID,

      range,

      valueRenderOption:
        render,
    });

  return (
    response.data.values ||
    []
  );
}

async function writeValues(
  range,
  values,
  input = 'USER_ENTERED'
) {
  const client =
    await sheetsClient();

  await client.spreadsheets.values.update({
    spreadsheetId:
      SPREADSHEET_ID,

    range,

    valueInputOption:
      input,

    requestBody: {
      values,
    },
  });
}

async function clearValues(range) {
  const client =
    await sheetsClient();

  await client.spreadsheets.values.clear({
    spreadsheetId:
      SPREADSHEET_ID,

    range,

    requestBody: {},
  });
}

async function batchUpdate(
  requests
) {
  if (!requests.length) {
    return;
  }

  const client =
    await sheetsClient();

  await client.spreadsheets.batchUpdate({
    spreadsheetId:
      SPREADSHEET_ID,

    requestBody: {
      requests,
    },
  });
}

async function sheetMeta(
  sheetName
) {
  const client =
    await sheetsClient();

  const response =
    await client.spreadsheets.get({
      spreadsheetId:
        SPREADSHEET_ID,

      fields:
        'sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))',
    });

  const sheet =
    (
      response.data.sheets ||
      []
    ).find(
      item =>
        item.properties
          ?.title ===
        sheetName
    );

  if (!sheet) {
    throw new Error(
      `Sheet tab not found: ${sheetName}`
    );
  }

  return sheet.properties;
}

function findHeaderRow(
  rows,
  requiredHeaders
) {
  const required =
    requiredHeaders.map(
      normaliseHeader
    );

  for (
    let rowIndex = 0;
    rowIndex <
    Math.min(
      rows.length,
      40
    );
    rowIndex++
  ) {
    const row =
      (
        rows[rowIndex] ||
        []
      ).map(
        normaliseHeader
      );

    if (
      required.every(
        header =>
          row.includes(header)
      )
    ) {
      const columns = {};

      for (
        const header of
        required
      ) {
        columns[header] =
          row.indexOf(header);
      }

      return {
        rowIndex,

        rowNumber:
          rowIndex + 1,

        columns,
      };
    }
  }

  throw new Error(
    `Could not find these headers: ${requiredHeaders.join(', ')}`
  );
}

function rowMatchesAlias(
  row,
  alias,
  startsWith = true
) {
  const target =
    normaliseSection(alias);

  if (!target) {
    return false;
  }

  /*
   * Section headings begin in column A.
   * This prevents cells such as
   * Active service = Reserves from
   * being treated as section headings.
   */
  const value =
    normaliseSection(
      row?.[0] || ''
    );

  return startsWith
    ? value.startsWith(target)
    : value === target;
}

function findSectionRow(
  rows,
  aliases
) {
  for (
    let rowIndex = 0;
    rowIndex < rows.length;
    rowIndex++
  ) {
    if (
      aliases.some(
        alias =>
          rowMatchesAlias(
            rows[rowIndex],
            alias
          )
      )
    ) {
      return rowIndex;
    }
  }

  return -1;
}

function isRosterBoundary(row) {
  return (
    ROSTER_BOUNDARIES.some(
      alias =>
        rowMatchesAlias(
          row,
          alias
        )
    )
  );
}

function isActivityBoundary(row) {
  return (
    Object.values(
      ACTIVITY_SECTIONS
    )
      .flat()
      .some(
        alias =>
          rowMatchesAlias(
            row,
            alias
          )
      )
  );
}

function findNextBoundary(
  rows,
  startIndex,
  type
) {
  const matcher =
    type === 'roster'
      ? isRosterBoundary
      : isActivityBoundary;

  for (
    let rowIndex =
      startIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    if (
      matcher(
        rows[rowIndex]
      )
    ) {
      return rowIndex;
    }
  }

  return rows.length;
}

function destinationKeyForRosterRow(
  rows,
  rowIndex
) {
  for (
    let index =
      rowIndex - 1;
    index >= 0;
    index--
  ) {
    for (
      const destination of
      Object.values(
        DESTINATIONS
      )
    ) {
      if (
        destination
          .rosterAliases
          .some(
            alias =>
              rowMatchesAlias(
                rows[index],
                alias
              )
          )
      ) {
        return destination.key;
      }
    }
  }

  return null;
}

function activitySectionKeyForRow(
  rows,
  rowIndex
) {
  for (
    let index =
      rowIndex - 1;
    index >= 0;
    index--
  ) {
    for (
      const [
        key,
        aliases,
      ] of Object.entries(
        ACTIVITY_SECTIONS
      )
    ) {
      if (
        aliases.some(
          alias =>
            rowMatchesAlias(
              rows[index],
              alias
            )
        )
      ) {
        return key;
      }
    }
  }

  return null;
}

function findSectionPlacement(
  rows,
  sectionRowIndex,
  boundaryIndex,
  header,
  type,
  appendAfterLastMember = false
) {
  const rankColumn =
    header.columns.rank;

  const nameColumn =
    header.columns.name;

  let blankRowIndex = -1;
  let templateRowIndex = -1;
  let lastMemberRowIndex = -1;

  for (
    let rowIndex =
      sectionRowIndex + 1;
    rowIndex <
      boundaryIndex;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] ||
      [];

    const rank =
      String(
        row[rankColumn] ||
        ''
      ).trim();

    const name =
      String(
        row[nameColumn] ||
        ''
      ).trim();

    if (
      rank ||
      name
    ) {
      templateRowIndex =
        rowIndex;

      lastMemberRowIndex =
        rowIndex;
    } else if (
      blankRowIndex < 0
    ) {
      blankRowIndex =
        rowIndex;
    }
  }

  /*
   * Returned cadets on Game Activity
   * are always placed after the final
   * current cadet instead of inside an
   * old hidden or blank row.
   */
  if (
    appendAfterLastMember
  ) {
    if (
      lastMemberRowIndex < 0
    ) {
      throw new Error(
        `No formatted member row is available in the ${type} destination section.`
      );
    }

    return {
      targetRowIndex:
        lastMemberRowIndex + 1,

      inserted:
        true,

      templateRowIndex:
        lastMemberRowIndex,
    };
  }

  if (
    blankRowIndex >= 0
  ) {
    return {
      targetRowIndex:
        blankRowIndex,

      inserted:
        false,

      templateRowIndex:
        templateRowIndex >= 0
          ? templateRowIndex
          : blankRowIndex,
    };
  }

  if (
    templateRowIndex < 0
  ) {
    throw new Error(
      `No formatted member row is available in the ${type} destination section.`
    );
  }

  return {
    targetRowIndex:
      boundaryIndex,

    inserted:
      true,

    templateRowIndex,
  };
}

async function insertFormattedRow({
  sheetId,
  targetRowIndex,
  templateRowIndex,
  columnCount,
}) {
  await batchUpdate([
    {
      insertDimension: {
        range: {
          sheetId,

          dimension:
            'ROWS',

          startIndex:
            targetRowIndex,

          endIndex:
            targetRowIndex + 1,
        },

        inheritFromBefore:
          targetRowIndex > 0,
      },
    },

    {
      copyPaste: {
        source: {
          sheetId,

          startRowIndex:
            templateRowIndex,

          endRowIndex:
            templateRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            columnCount,
        },

        destination: {
          sheetId,

          startRowIndex:
            targetRowIndex,

          endRowIndex:
            targetRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            columnCount,
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
          sheetId,

          startRowIndex:
            templateRowIndex,

          endRowIndex:
            templateRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            columnCount,
        },

        destination: {
          sheetId,

          startRowIndex:
            targetRowIndex,

          endRowIndex:
            targetRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            columnCount,
        },

        pasteType:
          'PASTE_DATA_VALIDATION',

        pasteOrientation:
          'NORMAL',
      },
    },
  ]);
}

async function deleteRow(
  sheetId,
  rowIndex
) {
  await batchUpdate([
    {
      deleteDimension: {
        range: {
          sheetId,

          dimension:
            'ROWS',

          startIndex:
            rowIndex,

          endIndex:
            rowIndex + 1,
        },
      },
    },
  ]);
}

async function prepareDestinationRow({
  sheetName,
  rows,
  header,
  destinationAliases,
  type,
  appendAfterLastMember = false,
}) {
  const sectionRowIndex =
    findSectionRow(
      rows,
      destinationAliases
    );

  if (
    sectionRowIndex < 0
  ) {
    throw new Error(
      `Could not find the ${type} destination section: ${destinationAliases[0]}`
    );
  }

  const boundaryIndex =
    findNextBoundary(
      rows,
      sectionRowIndex,
      type
    );

  const placement =
    findSectionPlacement(
      rows,
      sectionRowIndex,
      boundaryIndex,
      header,
      type,
      appendAfterLastMember
    );

  const meta =
    await sheetMeta(
      sheetName
    );

  if (
    placement.inserted
  ) {
    await insertFormattedRow({
      sheetId:
        meta.sheetId,

      targetRowIndex:
        placement
          .targetRowIndex,

      templateRowIndex:
        placement
          .templateRowIndex,

      columnCount:
        meta.gridProperties
          .columnCount,
    });
  }

  return {
    ...placement,

    sheetId:
      meta.sheetId,

    columnCount:
      meta.gridProperties
        .columnCount,
  };
}

async function getCadetRecordByDiscordId(
  discordId
) {
  const rows =
    await readValues(
      `${quoteSheet(
        TABS.cadets
      )}!A1:Z3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Name',
        'Timezone',
        'Discord Username',
        'User ID',
        'Joined',
      ]
    );

  const idColumn =
    header.columns[
      'user id'
    ];

  const matches = [];

  for (
    let rowIndex =
      header.rowIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] ||
      [];

    const foundId =
      String(
        row[idColumn] ||
        ''
      ).trim();

    if (
      foundId ===
      String(discordId)
    ) {
      matches.push({
        rowIndex,
        row,
      });
    }
  }

  if (
    matches.length > 1
  ) {
    throw new Error(
      `Discord ID ${discordId} appears more than once on Cadets.`
    );
  }

  if (!matches.length) {
    return null;
  }

  const match =
    matches[0];

  return {
    sheetName:
      TABS.cadets,

    rowIndex:
      match.rowIndex,

    rowNumber:
      match.rowIndex + 1,

    rowValues:
      padRow(
        match.row,
        26
      ),

    header,

    name:
      String(
        match.row[
          header.columns.name
        ] ||
        ''
      ).trim(),

    timezone:
      String(
        match.row[
          header.columns.timezone
        ] ||
        ''
      ).trim(),

    discordUsername:
      String(
        match.row[
          header.columns[
            'discord username'
          ]
        ] ||
        ''
      ).trim(),

    discordId:
      String(
        match.row[
          idColumn
        ] ||
        ''
      ).trim(),

    joined:
      String(
        match.row[
          header.columns.joined
        ] ||
        ''
      ).trim(),
  };
}

async function getRosterRecordByDiscordId(
  discordId
) {
  const rows =
    await readValues(
      `${quoteSheet(
        TABS.roster
      )}!A1:N3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
        'Timezone',
        'Position',
        'Squad',
        'Active service',
        'Discord Username',
        'Discord ID',
        'Join Date',
      ]
    );

  const idColumn =
    header.columns[
      'discord id'
    ];

  const matches = [];

  for (
    let rowIndex =
      header.rowIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] ||
      [];

    const foundId =
      String(
        row[idColumn] ||
        ''
      ).trim();

    if (
      foundId ===
      String(discordId)
    ) {
      matches.push({
        rowIndex,
        row,
      });
    }
  }

  if (
    matches.length > 1
  ) {
    throw new Error(
      `Discord ID ${discordId} appears more than once on Roster.`
    );
  }

  if (!matches.length) {
    return null;
  }

  const match =
    matches[0];

  return {
    sheetName:
      TABS.roster,

    rowIndex:
      match.rowIndex,

    rowNumber:
      match.rowIndex + 1,

    rowValues:
      padRow(
        match.row,
        14
      ),

    header,

    name:
      String(
        match.row[
          header.columns.name
        ] ||
        ''
      ).trim(),

    rank:
      String(
        match.row[
          header.columns.rank
        ] ||
        ''
      ).trim(),

    discordId:
      String(
        match.row[
          idColumn
        ] ||
        ''
      ).trim(),

    timezone:
      String(
        match.row[
          header.columns.timezone
        ] ||
        ''
      ).trim(),

    discordUsername:
      String(
        match.row[
          header.columns[
            'discord username'
          ]
        ] ||
        ''
      ).trim(),

    joinDate:
      String(
        match.row[
          header.columns[
            'join date'
          ]
        ] ||
        ''
      ).trim(),

    sectionKey:
      destinationKeyForRosterRow(
        rows,
        match.rowIndex
      ),
  };
}

async function snapshotCadetRow(
  record
) {
  const range =
    `${quoteSheet(
      record.sheetName
    )}!A${record.rowNumber}:Z${record.rowNumber}`;

  const [
    displayRows,
    formulaRows,
  ] =
    await Promise.all([
      readValues(
        range,
        'FORMATTED_VALUE'
      ),

      readValues(
        range,
        'FORMULA'
      ),
    ]);

  const display =
    displayRows[0] ||
    [];

  const formulas =
    formulaRows[0] ||
    [];

  const values =
    Array.from(
      {
        length: 26,
      },

      (_, index) => {
        const formulaValue =
          formulas[index];

        if (
          typeof formulaValue ===
            'string' &&
          formulaValue.startsWith(
            '='
          )
        ) {
          return formulaValue;
        }

        return (
          display[index] ??
          ''
        );
      }
    );

  return {
    range,
    values,

    sheetName:
      record.sheetName,

    rowIndex:
      record.rowIndex,

    rowNumber:
      record.rowNumber,
  };
}

async function clearCadetRecord(
  record
) {
  await clearValues(
    `${quoteSheet(
      record.sheetName
    )}!A${record.rowNumber}:Z${record.rowNumber}`
  );
}

async function deleteCadetRecord(
  record
) {
  const meta =
    await sheetMeta(
      record.sheetName
    );

  await deleteRow(
    meta.sheetId,
    record.rowIndex
  );
}

async function restoreDeletedCadetSnapshot(
  snapshot
) {
  const meta =
    await sheetMeta(
      snapshot.sheetName
    );

  await batchUpdate([
    {
      insertDimension: {
        range: {
          sheetId:
            meta.sheetId,

          dimension:
            'ROWS',

          startIndex:
            snapshot.rowIndex,

          endIndex:
            snapshot.rowIndex + 1,
        },

        inheritFromBefore:
          snapshot.rowIndex > 0,
      },
    },
  ]);

  let templateRowIndex =
    null;

  if (
    snapshot.rowIndex > 0
  ) {
    templateRowIndex =
      snapshot.rowIndex - 1;
  } else if (
    snapshot.rowIndex + 1 <
    meta.gridProperties
      .rowCount +
      1
  ) {
    templateRowIndex =
      snapshot.rowIndex + 1;
  }

  if (
    templateRowIndex !==
    null
  ) {
    const copyColumnCount =
      Math.min(
        26,
        meta.gridProperties
          .columnCount
      );

    await batchUpdate([
      {
        copyPaste: {
          source: {
            sheetId:
              meta.sheetId,

            startRowIndex:
              templateRowIndex,

            endRowIndex:
              templateRowIndex + 1,

            startColumnIndex:
              0,

            endColumnIndex:
              copyColumnCount,
          },

          destination: {
            sheetId:
              meta.sheetId,

            startRowIndex:
              snapshot.rowIndex,

            endRowIndex:
              snapshot.rowIndex + 1,

            startColumnIndex:
              0,

            endColumnIndex:
              copyColumnCount,
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
              meta.sheetId,

            startRowIndex:
              templateRowIndex,

            endRowIndex:
              templateRowIndex + 1,

            startColumnIndex:
              0,

            endColumnIndex:
              copyColumnCount,
          },

          destination: {
            sheetId:
              meta.sheetId,

            startRowIndex:
              snapshot.rowIndex,

            endRowIndex:
              snapshot.rowIndex + 1,

            startColumnIndex:
              0,

            endColumnIndex:
              copyColumnCount,
          },

          pasteType:
            'PASTE_DATA_VALIDATION',

          pasteOrientation:
            'NORMAL',
        },
      },
    ]);
  }

  await writeValues(
    snapshot.range,
    [
      snapshot.values,
    ]
  );
}

async function createRosterRowFromCadet(
  cadet,
  rankAbbreviation,
  destinationKey
) {
  const destination =
    DESTINATIONS[
      destinationKey
    ];

  if (!destination) {
    throw new Error(
      `Unknown destination: ${destinationKey}`
    );
  }

  const rows =
    await readValues(
      `${quoteSheet(
        TABS.roster
      )}!A1:N3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
        'Primary Kit',
        'Secondary Kit',
        'Timezone',
        'Position',
        'Squad',
        'Active service',
        'Discord Username',
        'Discord ID',
        'Join Date',
        'Logi',
      ]
    );

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.roster,

      rows,
      header,

      destinationAliases:
        destination
          .rosterAliases,

      type:
        'roster',
    });

  const row =
    Array(14).fill('');

  row[
    header.columns.rank
  ] =
    rankAbbreviation;

  row[
    header.columns.name
  ] =
    cadet.name;

  row[
    header.columns[
      'primary kit'
    ]
  ] =
    'Rifleman';

  row[
    header.columns[
      'secondary kit'
    ]
  ] = '';

  row[
    header.columns.timezone
  ] =
    cadet.timezone;

  row[
    header.columns.position
  ] =
    'Trooper';

  row[
    header.columns.squad
  ] =
    destination.rosterSquad;

  row[
    header.columns[
      'active service'
    ]
  ] =
    destination.activeService;

  row[
    header.columns[
      'discord username'
    ]
  ] =
    cadet.discordUsername;

  row[
    header.columns[
      'discord id'
    ]
  ] =
    cadet.discordId;

  row[
    header.columns[
      'join date'
    ]
  ] =
    cadet.joined;

  /*
   * New CTs are not approved
   * to drive the Logi by default.
   */
  row[
    header.columns.logi
  ] =
    'No';

  const rowNumber =
    placement
      .targetRowIndex +
    1;

  await writeValues(
    `${quoteSheet(
      TABS.roster
    )}!A${rowNumber}:N${rowNumber}`,

    [
      row,
    ]
  );

  return {
    rowIndex:
      placement
        .targetRowIndex,

    rowNumber,

    inserted:
      placement.inserted,

    destinationKey,

    discordId:
      cadet.discordId,
  };
}

async function deleteRosterRecordByDiscordId(
  discordId
) {
  const record =
    await getRosterRecordByDiscordId(
      discordId
    );

  if (!record) {
    return;
  }

  const meta =
    await sheetMeta(
      TABS.roster
    );

  await deleteRow(
    meta.sheetId,
    record.rowIndex
  );
}

async function updateRosterRankOnly(
  record,
  rankAbbreviation
) {
  const rankColumn =
    record.header.columns.rank;

  await writeValues(
    `${quoteSheet(
      TABS.roster
    )}!${columnLetter(
      rankColumn
    )}${record.rowNumber}`,

    [
      [
        rankAbbreviation,
      ],
    ]
  );

  return {
    type:
      'rank-only',

    discordId:
      record.discordId,

    originalRowValues:
      record.rowValues,

    originalSectionKey:
      record.sectionKey,
  };
}

async function moveRosterRecord(
  record,
  rankAbbreviation,
  destinationKey
) {
  const destination =
    DESTINATIONS[
      destinationKey
    ];

  if (!destination) {
    throw new Error(
      `Unknown destination: ${destinationKey}`
    );
  }

  const rows =
    await readValues(
      `${quoteSheet(
        TABS.roster
      )}!A1:N3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
        'Timezone',
        'Position',
        'Squad',
        'Active service',
        'Discord Username',
        'Discord ID',
        'Join Date',
      ]
    );

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.roster,

      rows,
      header,

      destinationAliases:
        destination
          .rosterAliases,

      type:
        'roster',
    });

  const row =
    padRow(
      record.rowValues,
      14
    );

  row[
    header.columns.rank
  ] =
    rankAbbreviation;

  row[
    header.columns.squad
  ] =
    destination.rosterSquad;

  row[
    header.columns[
      'active service'
    ]
  ] =
    destination.activeService;

  const targetRowNumber =
    placement
      .targetRowIndex +
    1;

  await writeValues(
    `${quoteSheet(
      TABS.roster
    )}!A${targetRowNumber}:N${targetRowNumber}`,

    [
      row,
    ]
  );

  const sourceAdjustedIndex =
    record.rowIndex +
    (
      placement.inserted &&
      placement
        .targetRowIndex <=
        record.rowIndex
        ? 1
        : 0
    );

  if (
    sourceAdjustedIndex ===
    placement.targetRowIndex
  ) {
    throw new Error(
      'The Roster source and destination resolved to the same row.'
    );
  }

  await deleteRow(
    placement.sheetId,
    sourceAdjustedIndex
  );

  return {
    type:
      'move',

    discordId:
      record.discordId,

    originalRowValues:
      record.rowValues,

    originalSectionKey:
      record.sectionKey,

    destinationKey,
  };
}

async function restoreRosterChange(
  token
) {
  const current =
    await getRosterRecordByDiscordId(
      token.discordId
    );

  if (!current) {
    throw new Error(
      'Could not find the roster record during rollback.'
    );
  }

  if (
    token.type ===
    'rank-only'
  ) {
    await writeValues(
      `${quoteSheet(
        TABS.roster
      )}!A${current.rowNumber}:N${current.rowNumber}`,

      [
        padRow(
          token.originalRowValues,
          14
        ),
      ]
    );

    return;
  }

  if (
    !token.originalSectionKey
  ) {
    throw new Error(
      'Cannot automatically roll back a roster move from an unrecognised section.'
    );
  }

  const destination =
    DESTINATIONS[
      token.originalSectionKey
    ];

  if (!destination) {
    throw new Error(
      `Unknown original Roster destination: ${token.originalSectionKey}`
    );
  }

  const rows =
    await readValues(
      `${quoteSheet(
        TABS.roster
      )}!A1:N3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
        'Timezone',
        'Position',
        'Squad',
        'Active service',
        'Discord Username',
        'Discord ID',
        'Join Date',
      ]
    );

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.roster,

      rows,
      header,

      destinationAliases:
        destination
          .rosterAliases,

      type:
        'roster',
    });

  const targetRowNumber =
    placement
      .targetRowIndex +
    1;

  await writeValues(
    `${quoteSheet(
      TABS.roster
    )}!A${targetRowNumber}:N${targetRowNumber}`,

    [
      padRow(
        token.originalRowValues,
        14
      ),
    ]
  );

  const sourceAdjustedIndex =
    current.rowIndex +
    (
      placement.inserted &&
      placement
        .targetRowIndex <=
        current.rowIndex
        ? 1
        : 0
    );

  await deleteRow(
    placement.sheetId,
    sourceAdjustedIndex
  );
}

async function findActivityRecord(
  name,
  preferredSectionKey = null
) {
  const meta =
    await sheetMeta(
      TABS.activity
    );

  const endColumn =
    columnLetter(
      meta.gridProperties
        .columnCount -
      1
    );

  const rows =
    await readValues(
      `${quoteSheet(
        TABS.activity
      )}!A1:${endColumn}${meta.gridProperties.rowCount}`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
      ]
    );

  const targetName =
    normaliseName(name);

  let matches = [];

  for (
    let rowIndex =
      header.rowIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const rowName =
      normaliseName(
        rows[rowIndex]?.[
          header.columns.name
        ]
      );

    if (
      rowName ===
      targetName
    ) {
      matches.push({
        rowIndex,

        rowNumber:
          rowIndex + 1,

        rowValues:
          padRow(
            rows[rowIndex] ||
              [],

            meta.gridProperties
              .columnCount
          ),

        sectionKey:
          activitySectionKeyForRow(
            rows,
            rowIndex
          ),
      });
    }
  }

  if (
    preferredSectionKey
  ) {
    const preferred =
      matches.filter(
        match =>
          match.sectionKey ===
          preferredSectionKey
      );

    if (
      preferred.length
    ) {
      matches =
        preferred;
    }
  }

  if (
    matches.length > 1
  ) {
    throw new Error(
      `More than one Game Activity row matches ${name}.`
    );
  }

  if (!matches.length) {
    return null;
  }

  return {
    ...matches[0],

    sheetName:
      TABS.activity,

    header,

    sheetId:
      meta.sheetId,

    columnCount:
      meta.gridProperties
        .columnCount,

    rows,
  };
}

async function updateActivityRankOnly(
  record,
  rankAbbreviation
) {
  const rankColumn =
    record.header.columns.rank;

  await writeValues(
    `${quoteSheet(
      TABS.activity
    )}!${columnLetter(
      rankColumn
    )}${record.rowNumber}`,

    [
      [
        rankAbbreviation,
      ],
    ]
  );

  return {
    type:
      'rank-only',

    name:
      String(
        record.rowValues[
          record.header
            .columns.name
        ] ||
        ''
      ).trim(),

    originalRowValues:
      record.rowValues,

    originalSectionKey:
      record.sectionKey,
  };
}

async function moveActivityRecord(
  record,
  rankAbbreviation,
  destinationActivitySectionKey
) {
  const aliases =
    ACTIVITY_SECTIONS[
      destinationActivitySectionKey
    ];

  if (!aliases) {
    throw new Error(
      `Unknown Game Activity destination: ${destinationActivitySectionKey}`
    );
  }

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.activity,

      rows:
        record.rows,

      header:
        record.header,

      destinationAliases:
        aliases,

      type:
        'activity',

      appendAfterLastMember:
        destinationActivitySectionKey ===
        'cadets',
    });

  const row =
    padRow(
      record.rowValues,
      placement.columnCount
    );

  row[
    record.header.columns.rank
  ] =
    rankAbbreviation;

  const targetRowIndex =
    placement.targetRowIndex;

  const targetRowNumber =
    targetRowIndex + 1;

  const endColumn =
    columnLetter(
      placement.columnCount -
      1
    );

  const targetRange =
    `${quoteSheet(
      TABS.activity
    )}!A${targetRowNumber}:${endColumn}${targetRowNumber}`;

  let sourceDeleted =
    false;

  try {
    await writeValues(
      targetRange,
      [
        row,
      ]
    );

    const verificationRows =
      await readValues(
        targetRange
      );

    const verificationRow =
      verificationRows[0] ||
      [];

    const expectedName =
      normaliseName(
        row[
          record.header
            .columns.name
        ]
      );

    const actualName =
      normaliseName(
        verificationRow[
          record.header
            .columns.name
        ]
      );

    const expectedRank =
      String(
        rankAbbreviation
      ).trim();

    const actualRank =
      String(
        verificationRow[
          record.header
            .columns.rank
        ] ||
        ''
      ).trim();

    if (
      actualName !==
        expectedName ||
      actualRank !==
        expectedRank
    ) {
      throw new Error(
        `The new Game Activity row could not be verified at row ${targetRowNumber}. The original row has not been deleted.`
      );
    }

    let sourceRowIndex =
      record.rowIndex;

    if (
      placement.inserted &&
      targetRowIndex <=
        record.rowIndex
    ) {
      sourceRowIndex += 1;
    }

    if (
      sourceRowIndex ===
      targetRowIndex
    ) {
      throw new Error(
        'The Game Activity source and destination resolved to the same row. The original row has not been deleted.'
      );
    }

    await deleteRow(
      placement.sheetId,
      sourceRowIndex
    );

    sourceDeleted =
      true;
  } catch (error) {
    if (!sourceDeleted) {
      if (
        placement.inserted
      ) {
        await deleteRow(
          placement.sheetId,
          targetRowIndex
        ).catch(
          cleanupError =>
            console.error(
              'Failed to remove the unused Game Activity destination row:',
              cleanupError
            )
        );
      } else {
        await clearValues(
          targetRange
        ).catch(
          cleanupError =>
            console.error(
              'Failed to clear the unused Game Activity destination row:',
              cleanupError
            )
        );
      }
    }

    throw error;
  }

  return {
    type:
      'move',

    name:
      String(
        record.rowValues[
          record.header
            .columns.name
        ] ||
        ''
      ).trim(),

    originalRowValues:
      record.rowValues,

    originalSectionKey:
      record.sectionKey,

    destinationSectionKey:
      destinationActivitySectionKey,
  };
}

async function restoreActivityChange(
  token
) {
  const current =
    await findActivityRecord(
      token.name,

      token
        .destinationSectionKey ||
      token.originalSectionKey
    );

  if (!current) {
    throw new Error(
      'Could not find the Game Activity row during rollback.'
    );
  }

  if (
    token.type ===
    'rank-only'
  ) {
    const endColumn =
      columnLetter(
        current.columnCount -
        1
      );

    await writeValues(
      `${quoteSheet(
        TABS.activity
      )}!A${current.rowNumber}:${endColumn}${current.rowNumber}`,

      [
        padRow(
          token.originalRowValues,
          current.columnCount
        ),
      ]
    );

    return;
  }

  const aliases =
    ACTIVITY_SECTIONS[
      token.originalSectionKey
    ];

  if (!aliases) {
    throw new Error(
      'Cannot automatically roll back a Game Activity move from an unrecognised section.'
    );
  }

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.activity,

      rows:
        current.rows,

      header:
        current.header,

      destinationAliases:
        aliases,

      type:
        'activity',

      appendAfterLastMember:
        token
          .originalSectionKey ===
        'cadets',
    });

  const row =
    padRow(
      token.originalRowValues,
      placement.columnCount
    );

  const targetRowNumber =
    placement
      .targetRowIndex +
    1;

  const endColumn =
    columnLetter(
      placement.columnCount -
      1
    );

  await writeValues(
    `${quoteSheet(
      TABS.activity
    )}!A${targetRowNumber}:${endColumn}${targetRowNumber}`,

    [
      row,
    ]
  );

  const sourceAdjustedIndex =
    current.rowIndex +
    (
      placement.inserted &&
      placement
        .targetRowIndex <=
        current.rowIndex
        ? 1
        : 0
    );

  await deleteRow(
    placement.sheetId,
    sourceAdjustedIndex
  );
}

async function graduateCadetSheets({
  cadet,
  rankAbbreviation,
}) {
  const cadetSnapshot =
    await snapshotCadetRow(
      cadet
    );

  let rosterCreated =
    false;

  let activityToken =
    null;

  let cadetDeleted =
    false;

  try {
    await createRosterRowFromCadet(
      cadet,
      rankAbbreviation,
      'reserves'
    );

    rosterCreated =
      true;

    const activityRecord =
      await findActivityRecord(
        cadet.name,
        'cadets'
      );

    if (!activityRecord) {
      throw new Error(
        `No Game Activity row was found for cadet ${cadet.name}.`
      );
    }

    activityToken =
      await moveActivityRecord(
        activityRecord,
        rankAbbreviation,
        'reserves'
      );

    /*
     * Delete the Cadets row instead
     * of leaving a blank coloured row.
     */
    await deleteCadetRecord(
      cadet
    );

    cadetDeleted =
      true;

    return {
      type:
        'graduation',

      cadetSnapshot,

      cadet,

      rosterDiscordId:
        cadet.discordId,

      activityToken,
    };
  } catch (error) {
    if (
      cadetDeleted
    ) {
      await restoreDeletedCadetSnapshot(
        cadetSnapshot
      ).catch(
        rollbackError =>
          console.error(
            'Cadet row rollback failed:',
            rollbackError
          )
      );
    }

    if (
      activityToken
    ) {
      await restoreActivityChange(
        activityToken
      ).catch(
        rollbackError =>
          console.error(
            'Activity rollback failed:',
            rollbackError
          )
      );
    }

    if (
      rosterCreated
    ) {
      await deleteRosterRecordByDiscordId(
        cadet.discordId
      ).catch(
        rollbackError =>
          console.error(
            'Roster rollback failed:',
            rollbackError
          )
      );
    }

    throw error;
  }
}

async function rollbackCadetGraduation(
  token
) {
  await restoreDeletedCadetSnapshot(
    token.cadetSnapshot
  );

  await restoreActivityChange(
    token.activityToken
  );

  await deleteRosterRecordByDiscordId(
    token.rosterDiscordId
  );
}

async function changeExistingRosterSheets({
  roster,
  rankAbbreviation,
  destinationKey,
}) {
  let rosterToken =
    null;

  let activityToken =
    null;

  try {
    const destinationChanged =
      Boolean(
        destinationKey
      ) &&
      destinationKey !==
        roster.sectionKey;

    if (
      destinationChanged &&
      !roster.sectionKey
    ) {
      throw new Error(
        'This member is in a roster section Orion does not yet recognise. Change their rank without selecting a destination.'
      );
    }

    if (
      destinationChanged
    ) {
      rosterToken =
        await moveRosterRecord(
          roster,
          rankAbbreviation,
          destinationKey
        );
    } else {
      rosterToken =
        await updateRosterRankOnly(
          roster,
          rankAbbreviation
        );
    }

    const preferredActivitySection =
      roster.sectionKey
        ? (
            DESTINATIONS[
              roster.sectionKey
            ]?.activitySectionKey ||
            null
          )
        : null;

    const activityRecord =
      await findActivityRecord(
        roster.name,
        preferredActivitySection
      );

    if (!activityRecord) {
      throw new Error(
        `No Game Activity row was found for ${roster.name}.`
      );
    }

    const destinationActivitySection =
      destinationKey
        ? DESTINATIONS[
            destinationKey
          ].activitySectionKey
        : activityRecord
            .sectionKey;

    if (
      destinationChanged &&
      destinationActivitySection !==
        activityRecord.sectionKey
    ) {
      activityToken =
        await moveActivityRecord(
          activityRecord,
          rankAbbreviation,
          destinationActivitySection
        );
    } else {
      activityToken =
        await updateActivityRankOnly(
          activityRecord,
          rankAbbreviation
        );
    }

    return {
      type:
        'existing-roster-change',

      rosterToken,

      activityToken,
    };
  } catch (error) {
    if (
      activityToken
    ) {
      await restoreActivityChange(
        activityToken
      ).catch(
        rollbackError =>
          console.error(
            'Activity rollback failed:',
            rollbackError
          )
      );
    }

    if (
      rosterToken
    ) {
      await restoreRosterChange(
        rosterToken
      ).catch(
        rollbackError =>
          console.error(
            'Roster rollback failed:',
            rollbackError
          )
      );
    }

    throw error;
  }
}

async function rollbackExistingRosterChange(
  token
) {
  await restoreActivityChange(
    token.activityToken
  );

  await restoreRosterChange(
    token.rosterToken
  );
}

function findCadetTemplateRow(
  rows,
  header,
  targetRowIndex
) {
  const nameColumn =
    header.columns.name;

  const idColumn =
    header.columns[
      'user id'
    ];

  for (
    let distance = 1;
    distance <
      rows.length + 2;
    distance++
  ) {
    const candidates = [
      targetRowIndex -
        distance,

      targetRowIndex +
        distance,
    ];

    for (
      const candidateIndex of
      candidates
    ) {
      if (
        candidateIndex <=
          header.rowIndex ||
        candidateIndex < 0 ||
        candidateIndex >=
          rows.length
      ) {
        continue;
      }

      const candidateRow =
        rows[candidateIndex] ||
        [];

      const candidateName =
        String(
          candidateRow[
            nameColumn
          ] ||
          ''
        ).trim();

      const candidateDiscordId =
        String(
          candidateRow[
            idColumn
          ] ||
          ''
        ).trim();

      if (
        candidateName ||
        candidateDiscordId
      ) {
        return candidateIndex;
      }
    }
  }

  return -1;
}

async function createCadetRowFromRoster(
  roster
) {
  const rows =
    await readValues(
      `${quoteSheet(
        TABS.cadets
      )}!A1:Z3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
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
      ]
    );

  const duplicate =
    await getCadetRecordByDiscordId(
      roster.discordId
    );

  if (duplicate) {
    throw new Error(
      `${roster.name} already has a row on the Cadets sheet.`
    );
  }

  const nameColumn =
    header.columns.name;

  const idColumn =
    header.columns[
      'user id'
    ];

  let targetRowIndex = -1;

  /*
   * Reuse the first genuinely blank
   * Cadets row. This also fills gaps
   * left by older versions of Orion.
   */
  for (
    let rowIndex =
      header.rowIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] ||
      [];

    const name =
      String(
        row[nameColumn] ||
        ''
      ).trim();

    const discordId =
      String(
        row[idColumn] ||
        ''
      ).trim();

    if (
      !name &&
      !discordId
    ) {
      targetRowIndex =
        rowIndex;

      break;
    }
  }

  if (
    targetRowIndex < 0
  ) {
    targetRowIndex =
      Math.max(
        rows.length,
        header.rowIndex + 1
      );
  }

  const meta =
    await sheetMeta(
      TABS.cadets
    );

  if (
    targetRowIndex >=
    meta.gridProperties.rowCount
  ) {
    await batchUpdate([
      {
        appendDimension: {
          sheetId:
            meta.sheetId,

          dimension:
            'ROWS',

          length:
            targetRowIndex -
            meta.gridProperties
              .rowCount +
            1,
        },
      },
    ]);
  }

  const templateRowIndex =
    findCadetTemplateRow(
      rows,
      header,
      targetRowIndex
    );

  if (
    templateRowIndex < 0
  ) {
    throw new Error(
      'No populated Cadets row is available to copy formatting and checkbox validation from.'
    );
  }

  const copyColumnCount =
    Math.min(
      26,
      meta.gridProperties
        .columnCount
    );

  await batchUpdate([
    {
      copyPaste: {
        source: {
          sheetId:
            meta.sheetId,

          startRowIndex:
            templateRowIndex,

          endRowIndex:
            templateRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            copyColumnCount,
        },

        destination: {
          sheetId:
            meta.sheetId,

          startRowIndex:
            targetRowIndex,

          endRowIndex:
            targetRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            copyColumnCount,
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
            meta.sheetId,

          startRowIndex:
            templateRowIndex,

          endRowIndex:
            templateRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            copyColumnCount,
        },

        destination: {
          sheetId:
            meta.sheetId,

          startRowIndex:
            targetRowIndex,

          endRowIndex:
            targetRowIndex + 1,

          startColumnIndex:
            0,

          endColumnIndex:
            copyColumnCount,
        },

        pasteType:
          'PASTE_DATA_VALIDATION',

        pasteOrientation:
          'NORMAL',
      },
    },
  ]);

  const rowNumber =
    targetRowIndex + 1;

  const cadetWidth =
    Math.max(
      ...Object.values(
        header.columns
      )
    ) +
    1;

  const row =
    Array(
      cadetWidth
    ).fill('');

  row[
    header.columns.rank
  ] =
    'CDT';

  row[
    header.columns.name
  ] =
    roster.name;

  row[
    header.columns.timezone
  ] =
    roster.timezone;

  row[
    header.columns[
      'discord username'
    ]
  ] =
    roster.discordUsername;

  row[
    header.columns[
      'user id'
    ]
  ] =
    roster.discordId;

  row[
    header.columns.joined
  ] =
    roster.joinDate;

  /*
   * Boolean false displays as an
   * unchecked checkbox because its
   * validation was copied above.
   */
  row[
    header.columns.bct
  ] =
    false;

  row[
    header.columns[
      'play session 1'
    ]
  ] =
    false;

  row[
    header.columns[
      'play session 2'
    ]
  ] =
    false;

  row[
    header.columns[
      'play session 3'
    ]
  ] =
    false;

  const joinedColumn =
    columnLetter(
      header.columns.joined
    );

  row[
    header.columns[
      'time served'
    ]
  ] =
    `=TODAY()-${joinedColumn}${rowNumber}`;

  const endColumn =
    columnLetter(
      cadetWidth - 1
    );

  await writeValues(
    `${quoteSheet(
      TABS.cadets
    )}!A${rowNumber}:${endColumn}${rowNumber}`,

    [
      row,
    ]
  );

  return {
    discordId:
      roster.discordId,

    rowIndex:
      targetRowIndex,

    rowNumber,
  };
}

async function clearCreatedCadetRecord(
  discordId
) {
  const record =
    await getCadetRecordByDiscordId(
      discordId
    );

  if (record) {
    await clearCadetRecord(
      record
    );
  }
}

async function restoreDeletedRosterRecord(
  record
) {
  if (!record.sectionKey) {
    throw new Error(
      'The original Roster section was not recognised, so Orion cannot restore it automatically.'
    );
  }

  const destination =
    DESTINATIONS[
      record.sectionKey
    ];

  if (!destination) {
    throw new Error(
      `Unknown original Roster destination: ${record.sectionKey}`
    );
  }

  const rows =
    await readValues(
      `${quoteSheet(
        TABS.roster
      )}!A1:N3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Rank',
        'Name',
        'Timezone',
        'Position',
        'Squad',
        'Active service',
        'Discord Username',
        'Discord ID',
        'Join Date',
      ]
    );

  const placement =
    await prepareDestinationRow({
      sheetName:
        TABS.roster,

      rows,
      header,

      destinationAliases:
        destination
          .rosterAliases,

      type:
        'roster',
    });

  const rowNumber =
    placement
      .targetRowIndex +
    1;

  await writeValues(
    `${quoteSheet(
      TABS.roster
    )}!A${rowNumber}:N${rowNumber}`,

    [
      padRow(
        record.rowValues,
        14
      ),
    ]
  );
}

async function returnRosterMemberToCadetSheets({
  roster,
}) {
  let cadetCreated =
    false;

  let activityToken =
    null;

  let rosterDeleted =
    false;

  try {
    if (!roster.sectionKey) {
      throw new Error(
        'This member is in a Roster section Orion does not recognise.'
      );
    }

    await createCadetRowFromRoster(
      roster
    );

    cadetCreated =
      true;

    const preferredActivitySection =
      DESTINATIONS[
        roster.sectionKey
      ]?.activitySectionKey ||
      null;

    const activityRecord =
      await findActivityRecord(
        roster.name,
        preferredActivitySection
      );

    if (!activityRecord) {
      throw new Error(
        `No Game Activity row was found for ${roster.name}.`
      );
    }

    activityToken =
      await moveActivityRecord(
        activityRecord,
        'CDT',
        'cadets'
      );

    const currentRoster =
      await getRosterRecordByDiscordId(
        roster.discordId
      );

    if (!currentRoster) {
      throw new Error(
        `The Roster row for ${roster.name} disappeared before it could be removed.`
      );
    }

    const rosterMeta =
      await sheetMeta(
        TABS.roster
      );

    await deleteRow(
      rosterMeta.sheetId,
      currentRoster.rowIndex
    );

    rosterDeleted =
      true;

    return {
      type:
        'return-to-cadet',

      discordId:
        roster.discordId,

      rosterSnapshot:
        roster,

      activityToken,
    };
  } catch (error) {
    if (
      rosterDeleted
    ) {
      await restoreDeletedRosterRecord(
        roster
      ).catch(
        rollbackError =>
          console.error(
            'Roster return-to-cadet rollback failed:',
            rollbackError
          )
      );
    }

    if (
      activityToken
    ) {
      await restoreActivityChange(
        activityToken
      ).catch(
        rollbackError =>
          console.error(
            'Activity return-to-cadet rollback failed:',
            rollbackError
          )
      );
    }

    if (
      cadetCreated
    ) {
      await clearCreatedCadetRecord(
        roster.discordId
      ).catch(
        rollbackError =>
          console.error(
            'Cadet return-to-cadet rollback failed:',
            rollbackError
          )
      );
    }

    throw error;
  }
}

async function rollbackReturnToCadet(
  token
) {
  await restoreDeletedRosterRecord(
    token.rosterSnapshot
  );

  await restoreActivityChange(
    token.activityToken
  );

  await clearCreatedCadetRecord(
    token.discordId
  );
}

module.exports = {
  getCadetRecordByDiscordId,
  getRosterRecordByDiscordId,

  graduateCadetSheets,
  rollbackCadetGraduation,

  returnRosterMemberToCadetSheets,
  rollbackReturnToCadet,

  changeExistingRosterSheets,
  rollbackExistingRosterChange,
};