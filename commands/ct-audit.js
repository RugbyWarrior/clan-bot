require('dotenv').config();

const {
  SlashCommandBuilder,
} = require('discord.js');

const {
  google,
} = require('googleapis');

const {
  sendLog,
} = require('../logger');

const {
  isBotOwner,
} = require('../permissions');

const CADETS_TAB =
  process.env.CADETS_SHEET_NAME ||
  'Cadets';

const ROSTER_TAB =
  process.env.ROSTER_SHEET_NAME ||
  'Roster';

const CT_NUMBERS_TAB =
  process.env.CT_NUMBERS_SHEET_NAME ||
  'CT Numbers';

let cachedSheetsClient = null;

function normaliseHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normaliseIgn(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normaliseUsername(value) {
  return String(value || '')
    .trim()
    .replace(/^@/, '')
    .toLowerCase();
}

function quoteSheetName(name) {
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

function escapeRegex(value) {
  return String(value || '')
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );
}

function getSpreadsheetId() {
  const spreadsheetId =
    String(
      process.env
        .GOOGLE_SPREADSHEET_ID ||
        ''
    ).trim();

  if (!spreadsheetId) {
    throw new Error(
      'GOOGLE_SPREADSHEET_ID is missing from .env.'
    );
  }

  return spreadsheetId;
}

function getCredentials() {
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

async function getSheetsClient() {
  if (cachedSheetsClient) {
    return cachedSheetsClient;
  }

  const auth =
    new google.auth.GoogleAuth({
      credentials:
        getCredentials(),

      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
      ],
    });

  cachedSheetsClient =
    google.sheets({
      version: 'v4',

      auth:
        await auth.getClient(),
    });

  return cachedSheetsClient;
}

async function getValues(range) {
  const sheets =
    await getSheetsClient();

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId:
        getSpreadsheetId(),

      range,

      valueRenderOption:
        'FORMATTED_VALUE',
    });

  return response.data.values || [];
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
    Math.min(rows.length, 40);
    rowIndex++
  ) {
    const row =
      (rows[rowIndex] || []).map(
        normaliseHeader
      );

    if (
      required.every(header =>
        row.includes(header)
      )
    ) {
      const columns = {};

      for (const header of required) {
        columns[header] =
          row.indexOf(header);
      }

      return {
        rowIndex,
        columns,
      };
    }
  }

  throw new Error(
    `Could not find these headers: ${requiredHeaders.join(', ')}`
  );
}

async function getPeopleFromSheet(
  sheetName,
  idHeader
) {
  const rows =
    await getValues(
      `${quoteSheetName(
        sheetName
      )}!A1:Z3000`
    );

  const header =
    findHeaderRow(
      rows,
      [
        'Name',
        'Discord Username',
        idHeader,
      ]
    );

  const nameColumn =
    header.columns.name;

  const usernameColumn =
    header.columns[
      'discord username'
    ];

  const idColumn =
    header.columns[
      normaliseHeader(idHeader)
    ];

  const people = [];

  for (
    let rowIndex =
      header.rowIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    const name =
      String(
        row[nameColumn] || ''
      ).trim();

    const discordUsername =
      String(
        row[usernameColumn] ||
          ''
      ).trim();

    const discordId =
      String(
        row[idColumn] || ''
      ).trim();

    if (!name) {
      continue;
    }

    people.push({
      sheetName,

      rowNumber:
        rowIndex + 1,

      name,

      normalisedIgn:
        normaliseIgn(name),

      discordUsername,

      normalisedUsername:
        normaliseUsername(
          discordUsername
        ),

      discordId,
    });
  }

  return people;
}

function isIgnHeader(value) {
  return [
    'ign',
    'name',
  ].includes(
    normaliseHeader(value)
  );
}

function isCtNumberHeader(value) {
  return [
    'ct number',
    'number',
  ].includes(
    normaliseHeader(value)
  );
}

function isDiscordIdHeader(value) {
  return [
    'discord id',
    'user id',
  ].includes(
    normaliseHeader(value)
  );
}

async function getCtData() {
  const rows =
    await getValues(
      `${quoteSheetName(
        CT_NUMBERS_TAB
      )}!A1:ZZ102`
    );

  const blocks = [];

  for (
    let rowIndex = 0;
    rowIndex <
    Math.min(rows.length, 20);
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    for (
      let columnIndex = 0;
      columnIndex <
      row.length - 2;
      columnIndex++
    ) {
      if (
        isIgnHeader(
          row[columnIndex]
        ) &&
        isCtNumberHeader(
          row[columnIndex + 1]
        ) &&
        isDiscordIdHeader(
          row[columnIndex + 2]
        )
      ) {
        blocks.push({
          blockIndex:
            blocks.length,

          headerRowIndex:
            rowIndex,

          ignColumnIndex:
            columnIndex,

          numberColumnIndex:
            columnIndex + 1,

          discordColumnIndex:
            columnIndex + 2,
        });
      }
    }
  }

  if (!blocks.length) {
    throw new Error(
      'No IGN | CT Number | Discord ID blocks were found on the CT Numbers tab.'
    );
  }

  const entries = [];

  for (const block of blocks) {
    for (
      let rowIndex =
        block.headerRowIndex + 1;
      rowIndex < 102;
      rowIndex++
    ) {
      const row =
        rows[rowIndex] || [];

      const ign =
        String(
          row[
            block.ignColumnIndex
          ] || ''
        ).trim();

      const ctNumber =
        String(
          row[
            block.numberColumnIndex
          ] || ''
        ).trim();

      const discordId =
        String(
          row[
            block.discordColumnIndex
          ] || ''
        ).trim();

      if (
        !ign &&
        !ctNumber &&
        !discordId
      ) {
        continue;
      }

      entries.push({
        blockIndex:
          block.blockIndex,

        rowIndex,

        rowNumber:
          rowIndex + 1,

        ign,

        normalisedIgn:
          normaliseIgn(ign),

        ctNumber,

        discordId,

        ignCell:
          `${columnLetter(
            block.ignColumnIndex
          )}${rowIndex + 1}`,

        numberCell:
          `${columnLetter(
            block.numberColumnIndex
          )}${rowIndex + 1}`,

        discordIdCell:
          `${columnLetter(
            block.discordColumnIndex
          )}${rowIndex + 1}`,
      });
    }
  }

  return {
    rows,
    blocks,
    entries,
  };
}

function stripCtSuffix(value) {
  return String(value || '')
    .trim()
    .replace(
      /(?:\s+|-)(?:\d{2}-\d{3}|\d{5}|\d{4})\s*$/i,
      ''
    )
    .trim();
}

function stripRankPrefix(value) {
  return String(value || '')
    .trim()
    .replace(
      /^[A-Z0-9/.-]{2,10}\s+/,
      ''
    )
    .trim();
}

function makeNicknameVariants(value) {
  const raw =
    String(value || '').trim();

  if (!raw) {
    return [];
  }

  const withoutCt =
    stripCtSuffix(raw);

  const withoutRank =
    stripRankPrefix(raw);

  const withoutRankAndCt =
    stripRankPrefix(
      withoutCt
    );

  return [
    ...new Set(
      [
        normaliseIgn(raw),

        normaliseIgn(
          withoutCt
        ),

        normaliseIgn(
          withoutRank
        ),

        normaliseIgn(
          withoutRankAndCt
        ),
      ].filter(Boolean)
    ),
  ];
}

function memberSearchRecord(member) {
  const serverNickname =
    member.nickname || '';

  const displayName =
    member.displayName || '';

  const globalName =
    member.user.globalName || '';

  const username =
    member.user.username || '';

  const nameLabels = [
    ...new Set(
      [
        serverNickname,
        displayName,
        globalName,
      ].filter(Boolean)
    ),
  ];

  return {
    member,

    serverNickname,

    username,

    normalisedUsername:
      normaliseUsername(
        username
      ),

    nameLabels,

    nicknameVariants: [
      ...new Set(
        nameLabels.flatMap(
          makeNicknameVariants
        )
      ),
    ],
  };
}

function nicknameHasCtNumber(
  record,
  ctNumber
) {
  const escapedCtNumber =
    escapeRegex(
      String(
        ctNumber || ''
      ).trim()
    );

  if (
    !escapedCtNumber ||
    !record.serverNickname
  ) {
    return false;
  }

  const pattern =
    new RegExp(
      `(?:\\s+|-)${escapedCtNumber}\\s*$`,
      'i'
    );

  return pattern.test(
    record.serverNickname.trim()
  );
}

function extractCtNumbers(record) {
  const found =
    new Set();

  const nickname =
    String(
      record.serverNickname ||
        ''
    ).trim();

  if (!nickname) {
    return [];
  }

  const match =
    nickname.match(
      /(?:^|\s+|-)(\d{2}-\d{3}|\d{5}|\d{4})\s*$/i
    );

  if (match) {
    found.add(match[1]);
  }

  return [...found];
}

function levenshteinDistance(
  first,
  second
) {
  const a =
    String(first || '');

  const b =
    String(second || '');

  if (!a.length) {
    return b.length;
  }

  if (!b.length) {
    return a.length;
  }

  const previous =
    Array.from(
      {
        length:
          b.length + 1,
      },

      (_, index) => index
    );

  for (
    let aIndex = 1;
    aIndex <= a.length;
    aIndex++
  ) {
    const current =
      [aIndex];

    for (
      let bIndex = 1;
      bIndex <= b.length;
      bIndex++
    ) {
      const substitutionCost =
        a[aIndex - 1] ===
        b[bIndex - 1]
          ? 0
          : 1;

      current[bIndex] =
        Math.min(
          current[bIndex - 1] +
            1,

          previous[bIndex] + 1,

          previous[bIndex - 1] +
            substitutionCost
        );
    }

    for (
      let index = 0;
      index < current.length;
      index++
    ) {
      previous[index] =
        current[index];
    }
  }

  return previous[b.length];
}

function similarity(
  first,
  second
) {
  const a =
    normaliseIgn(first);

  const b =
    normaliseIgn(second);

  if (!a || !b) {
    return 0;
  }

  if (a === b) {
    return 1;
  }

  return (
    1 -
    levenshteinDistance(
      a,
      b
    ) /
      Math.max(
        a.length,
        b.length
      )
  );
}

function uniqueMemberRecords(records) {
  const byId =
    new Map();

  for (const record of records) {
    byId.set(
      record.member.id,
      record
    );
  }

  return [
    ...byId.values(),
  ];
}

function addToMap(
  map,
  key,
  value
) {
  if (!key) {
    return;
  }

  if (!map.has(key)) {
    map.set(key, []);
  }

  map.get(key).push(value);
}

function findDiscordMatchForCt({
  ctEntry,
  spreadsheetPeople,
  memberRecords,
  memberById,
}) {
  if (!ctEntry.normalisedIgn) {
    return {
      status: 'unmatched',

      reason:
        `CT ${ctEntry.ctNumber} has no IGN`,
    };
  }

  const spreadsheetMatches =
    spreadsheetPeople.filter(
      person =>
        person.normalisedIgn ===
        ctEntry.normalisedIgn
    );

  const spreadsheetIds = [
    ...new Set(
      spreadsheetMatches
        .map(
          person =>
            person.discordId
        )
        .filter(Boolean)
    ),
  ];

  if (
    spreadsheetIds.length === 1
  ) {
    const record =
      memberById.get(
        spreadsheetIds[0]
      );

    if (record) {
      return {
        status: 'matched',

        method:
          'Spreadsheet Discord ID',

        record,
      };
    }
  }

  if (
    spreadsheetIds.length > 1
  ) {
    return {
      status: 'ambiguous',

      reason:
        `${ctEntry.ign} has multiple Discord IDs on Cadets/Roster: ` +
        spreadsheetIds.join(', '),
    };
  }

  const spreadsheetUsernames = [
    ...new Set(
      spreadsheetMatches
        .map(
          person =>
            person.normalisedUsername
        )
        .filter(Boolean)
    ),
  ];

  if (
    spreadsheetUsernames.length
  ) {
    const usernameMatches =
      uniqueMemberRecords(
        memberRecords.filter(
          record =>
            spreadsheetUsernames.includes(
              record.normalisedUsername
            )
        )
      );

    if (
      usernameMatches.length === 1
    ) {
      return {
        status: 'matched',

        method:
          'Spreadsheet Discord username',

        record:
          usernameMatches[0],
      };
    }

    if (
      usernameMatches.length > 1
    ) {
      return {
        status: 'ambiguous',

        reason:
          `${ctEntry.ign} matches multiple Discord usernames from the spreadsheet`,
      };
    }
  }

  const ctSuffixMatches =
    uniqueMemberRecords(
      memberRecords.filter(
        record =>
          nicknameHasCtNumber(
            record,
            ctEntry.ctNumber
          )
      )
    );

  if (
    ctSuffixMatches.length === 1
  ) {
    return {
      status: 'matched',

      method:
        'CT number in nickname',

      record:
        ctSuffixMatches[0],
    };
  }

  if (
    ctSuffixMatches.length > 1
  ) {
    return {
      status: 'ambiguous',

      reason:
        `${ctEntry.ctNumber} appears in multiple Discord nicknames`,
    };
  }

  const exactNicknameMatches =
    uniqueMemberRecords(
      memberRecords.filter(
        record =>
          record.nicknameVariants.includes(
            ctEntry.normalisedIgn
          )
      )
    );

  if (
    exactNicknameMatches.length ===
    1
  ) {
    return {
      status: 'matched',

      method:
        'Exact nickname IGN',

      record:
        exactNicknameMatches[0],
    };
  }

  if (
    exactNicknameMatches.length >
    1
  ) {
    return {
      status: 'ambiguous',

      reason:
        `${ctEntry.ign} exactly matches multiple Discord nicknames`,
    };
  }

  const exactUsernameMatches =
    uniqueMemberRecords(
      memberRecords.filter(
        record =>
          record.normalisedUsername ===
          normaliseUsername(
            ctEntry.ign
          )
      )
    );

  if (
    exactUsernameMatches.length ===
    1
  ) {
    return {
      status: 'matched',

      method:
        'Exact Discord username',

      record:
        exactUsernameMatches[0],
    };
  }

  const scored =
    memberRecords
      .map(record => {
        const scores =
          record.nicknameVariants.map(
            variant =>
              similarity(
                ctEntry.normalisedIgn,
                variant
              )
          );

        return {
          record,

          score:
            scores.length
              ? Math.max(...scores)
              : 0,
        };
      })

      .filter(
        result =>
          result.score > 0
      )

      .sort(
        (first, second) =>
          second.score -
          first.score
      );

  const best =
    scored[0];

  const second =
    scored[1];

  if (!best) {
    return {
      status: 'unmatched',

      reason:
        `${ctEntry.ign} had no Discord match`,
    };
  }

  const threshold =
    ctEntry.normalisedIgn.length <=
    5
      ? 0.92
      : 0.86;

  const scoreGap =
    best.score -
    (second?.score || 0);

  if (
    best.score >= threshold &&
    scoreGap >= 0.1
  ) {
    return {
      status: 'matched',

      method:
        'Similar nickname',

      record:
        best.record,
    };
  }

  const candidates =
    scored
      .slice(0, 3)

      .map(
        result =>
          `${result.record.member.displayName} (${Math.round(
            result.score *
              100
          )}%)`
      )

      .join(', ');

  return {
    status: 'ambiguous',

    reason:
      `${ctEntry.ign} has no safe unique match. Closest: ${candidates}`,
  };
}

function resolveSpreadsheetPersonToMember({
  person,
  memberRecords,
  memberById,
}) {
  if (person.discordId) {
    const byId =
      memberById.get(
        person.discordId
      );

    if (byId) {
      return {
        status: 'matched',

        method:
          'Spreadsheet Discord ID',

        record:
          byId,
      };
    }
  }

  if (
    person.normalisedUsername
  ) {
    const usernameMatches =
      uniqueMemberRecords(
        memberRecords.filter(
          record =>
            record.normalisedUsername ===
            person.normalisedUsername
        )
      );

    if (
      usernameMatches.length === 1
    ) {
      return {
        status: 'matched',

        method:
          'Spreadsheet Discord username',

        record:
          usernameMatches[0],
      };
    }

    if (
      usernameMatches.length > 1
    ) {
      return {
        status: 'ambiguous',

        reason:
          `${person.name} has a Discord username matching multiple members`,
      };
    }
  }

  const nicknameMatches =
    uniqueMemberRecords(
      memberRecords.filter(
        record =>
          record.nicknameVariants.includes(
            person.normalisedIgn
          )
      )
    );

  if (
    nicknameMatches.length === 1
  ) {
    return {
      status: 'matched',

      method:
        'Exact nickname IGN',

      record:
        nicknameMatches[0],
    };
  }

  if (
    nicknameMatches.length > 1
  ) {
    return {
      status: 'ambiguous',

      reason:
        `${person.name} exactly matches multiple Discord nicknames`,
    };
  }

  return {
    status: 'unmatched',

    reason:
      `${person.name} could not be matched to a Discord member`,
  };
}

function findCustomBlock(ctData) {
  return ctData.blocks.find(
    block =>
      block.ignColumnIndex === 0 &&
      block.numberColumnIndex === 1 &&
      block.discordColumnIndex === 2
  );
}

function planMissingCtRecord({
  ctData,
  person,
  memberRecord,
  ctNumber,
}) {
  const cleanCtNumber =
    String(ctNumber || '').trim();

  const matchesByNumber =
    ctData.entries.filter(
      entry =>
        entry.ctNumber ===
        cleanCtNumber
    );

  /*
   * The CT number already exists
   * somewhere on the sheet.
   */
  if (
    matchesByNumber.length > 1
  ) {
    return {
      status: 'conflict',

      reason:
        `${person.name}: CT ${cleanCtNumber} appears more than once on CT Numbers`,
    };
  }

  if (
    matchesByNumber.length === 1
  ) {
    const entry =
      matchesByNumber[0];

    if (
      entry.ign &&
      normaliseIgn(
        entry.ign
      ) !==
        person.normalisedIgn
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: CT ${cleanCtNumber} is already assigned to ${entry.ign}`,
      };
    }

    if (
      entry.discordId &&
      entry.discordId !==
        memberRecord.member.id
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: CT ${cleanCtNumber} is linked to a different Discord ID`,
      };
    }

    return {
      status: 'planned',

      type:
        'fill-existing-number',

      ctNumber:
        cleanCtNumber,

      discordId:
        memberRecord.member.id,

      range:
        `${quoteSheetName(
          CT_NUMBERS_TAB
        )}!` +
        `${entry.ignCell}:` +
        `${entry.discordIdCell}`,

      values: [[
        person.name,
        cleanCtNumber,
        memberRecord.member.id,
      ]],

      description:
        `${person.name} → ${cleanCtNumber} using an existing CT-number row`,
    };
  }

  /*
   * Local Orion numbers between
   * 53000 and 53999 belong in their
   * existing numbered block.
   */
  if (
    /^53\d{3}$/.test(
      cleanCtNumber
    )
  ) {
    const numeric =
      Number(cleanCtNumber);

    const hundred =
      Math.floor(
        numeric / 100
      );

    const suffix =
      numeric % 100;

    let matchingBlock = null;

    for (
      const block of
      ctData.blocks
    ) {
      const blockNumbers =
        ctData.entries
          .filter(
            entry =>
              entry.blockIndex ===
              block.blockIndex
          )

          .map(
            entry =>
              Number(
                entry.ctNumber
              )
          )

          .filter(
            Number.isFinite
          );

      if (
        blockNumbers.some(
          number =>
            Math.floor(
              number / 100
            ) === hundred
        )
      ) {
        matchingBlock =
          block;

        break;
      }
    }

    if (!matchingBlock) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: no ${hundred}xx block was found on CT Numbers`,
      };
    }

    const targetRowIndex =
      matchingBlock
        .headerRowIndex +
      1 +
      suffix;

    if (
      targetRowIndex >= 102
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: calculated row for ${cleanCtNumber} is outside the CT-number table`,
      };
    }

    const row =
      ctData.rows[
        targetRowIndex
      ] || [];

    const currentIgn =
      String(
        row[
          matchingBlock
            .ignColumnIndex
        ] || ''
      ).trim();

    const currentNumber =
      String(
        row[
          matchingBlock
            .numberColumnIndex
        ] || ''
      ).trim();

    const currentDiscordId =
      String(
        row[
          matchingBlock
            .discordColumnIndex
        ] || ''
      ).trim();

    if (
      currentNumber &&
      currentNumber !==
        cleanCtNumber
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: expected ${cleanCtNumber}, but the target row contains ${currentNumber}`,
      };
    }

    if (
      currentIgn &&
      normaliseIgn(
        currentIgn
      ) !==
        person.normalisedIgn
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: ${cleanCtNumber} is already assigned to ${currentIgn}`,
      };
    }

    if (
      currentDiscordId &&
      currentDiscordId !==
        memberRecord.member.id
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: ${cleanCtNumber} is linked to a different Discord ID`,
      };
    }

    const rowNumber =
      targetRowIndex + 1;

    const startColumn =
      columnLetter(
        matchingBlock
          .ignColumnIndex
      );

    const endColumn =
      columnLetter(
        matchingBlock
          .discordColumnIndex
      );

    return {
      status: 'planned',

      type:
        'restore-local-number-row',

      ctNumber:
        cleanCtNumber,

      discordId:
        memberRecord.member.id,

      range:
        `${quoteSheetName(
          CT_NUMBERS_TAB
        )}!` +
        `${startColumn}${rowNumber}:` +
        `${endColumn}${rowNumber}`,

      values: [[
        person.name,
        cleanCtNumber,
        memberRecord.member.id,
      ]],

      description:
        `${person.name} → ${cleanCtNumber} restored in the ${hundred}xx block`,
    };
  }

  /*
   * All older/external CT formats go
   * into the custom block in A:C.
   *
   * Supported examples:
   *
   * 23-879
   * 9174
   * 34160
   * 5019
   * 14930
   */
  if (
    /^(?:\d{2}-\d{3}|\d{4}|\d{5})$/.test(
      cleanCtNumber
    )
  ) {
    const customBlock =
      findCustomBlock(ctData);

    if (!customBlock) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: custom CT block A:C was not found`,
      };
    }

    /*
     * Include pending proposals when
     * choosing rows so several missing
     * people are not all written into
     * the same blank custom row.
     */
    for (
      let rowIndex =
        customBlock.headerRowIndex +
        1;
      rowIndex < 102;
      rowIndex++
    ) {
      const row =
        ctData.rows[rowIndex] ||
        [];

      const ign =
        String(
          row[
            customBlock
              .ignColumnIndex
          ] || ''
        ).trim();

      const number =
        String(
          row[
            customBlock
              .numberColumnIndex
          ] || ''
        ).trim();

      const discordId =
        String(
          row[
            customBlock
              .discordColumnIndex
          ] || ''
        ).trim();

      const rowNumber =
        rowIndex + 1;

      const range =
        `${quoteSheetName(
          CT_NUMBERS_TAB
        )}!A${rowNumber}:C${rowNumber}`;

      const alreadyPlanned =
        ctData.plannedCustomRanges
          ?.has(range) || false;

      if (
        !ign &&
        !number &&
        !discordId &&
        !alreadyPlanned
      ) {
        if (
          !ctData.plannedCustomRanges
        ) {
          ctData.plannedCustomRanges =
            new Set();
        }

        ctData.plannedCustomRanges.add(
          range
        );

        return {
          status: 'planned',

          type:
            'create-custom-number',

          ctNumber:
            cleanCtNumber,

          discordId:
            memberRecord.member.id,

          range,

          values: [[
            person.name,
            cleanCtNumber,
            memberRecord.member.id,
          ]],

          description:
            `${person.name} → ${cleanCtNumber} added to the custom CT block`,
        };
      }
    }

    return {
      status: 'conflict',

      reason:
        `${person.name}: no empty row remains in the custom CT block`,
    };
  }

  return {
    status: 'conflict',

    reason:
      `${person.name}: unsupported CT-number format ${cleanCtNumber}`,
  };

  if (
    /^53\d{3}$/.test(
      ctNumber
    )
  ) {
    const numeric =
      Number(ctNumber);

    const hundred =
      Math.floor(
        numeric / 100
      );

    const suffix =
      numeric % 100;

    let matchingBlock = null;

    for (
      const block of
      ctData.blocks
    ) {
      const blockNumbers =
        ctData.entries
          .filter(
            entry =>
              entry.blockIndex ===
              block.blockIndex
          )

          .map(
            entry =>
              Number(
                entry.ctNumber
              )
          )

          .filter(
            Number.isFinite
          );

      if (
        blockNumbers.some(
          number =>
            Math.floor(
              number / 100
            ) === hundred
        )
      ) {
        matchingBlock =
          block;

        break;
      }
    }

    if (!matchingBlock) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: no ${hundred}xx block was found on CT Numbers`,
      };
    }

    const targetRowIndex =
      matchingBlock
        .headerRowIndex +
      1 +
      suffix;

    if (
      targetRowIndex >= 102
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: calculated row for ${ctNumber} is outside the CT-number table`,
      };
    }

    const row =
      ctData.rows[
        targetRowIndex
      ] || [];

    const currentIgn =
      String(
        row[
          matchingBlock
            .ignColumnIndex
        ] || ''
      ).trim();

    const currentNumber =
      String(
        row[
          matchingBlock
            .numberColumnIndex
        ] || ''
      ).trim();

    const currentDiscordId =
      String(
        row[
          matchingBlock
            .discordColumnIndex
        ] || ''
      ).trim();

    if (
      currentNumber &&
      currentNumber !==
        ctNumber
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: expected ${ctNumber}, but the target row contains ${currentNumber}`,
      };
    }

    if (
      currentIgn &&
      normaliseIgn(
        currentIgn
      ) !==
        person.normalisedIgn
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: ${ctNumber} is already assigned to ${currentIgn}`,
      };
    }

    if (
      currentDiscordId &&
      currentDiscordId !==
        memberRecord.member.id
    ) {
      return {
        status: 'conflict',

        reason:
          `${person.name}: ${ctNumber} is linked to a different Discord ID`,
      };
    }

    const rowNumber =
      targetRowIndex + 1;

    const startColumn =
      columnLetter(
        matchingBlock
          .ignColumnIndex
      );

    const endColumn =
      columnLetter(
        matchingBlock
          .discordColumnIndex
      );

    return {
      status: 'planned',

      type:
        'restore-local-number-row',

      ctNumber,

      discordId:
        memberRecord.member.id,

      range:
        `${quoteSheetName(
          CT_NUMBERS_TAB
        )}!` +
        `${startColumn}${rowNumber}:` +
        `${endColumn}${rowNumber}`,

      values: [[
        person.name,
        ctNumber,
        memberRecord.member.id,
      ]],

      description:
        `${person.name} → ${ctNumber} restored in the ${hundred}xx block`,
    };
  }

  return {
    status: 'conflict',

    reason:
      `${person.name}: unsupported CT-number format ${ctNumber}`,
  };
}

function limitedList(
  title,
  items,
  limit = 6
) {
  if (!items.length) {
    return [];
  }

  return [
    `**${title}:**`,

    ...items
      .slice(0, limit)

      .map(
        item =>
          `• ${item}`
      ),

    ...(items.length > limit
      ? [
          `• …and ${
            items.length - limit
          } more`,
        ]
      : []),
  ];
}

function fitDiscordMessage(
  lines,
  maximum = 1950
) {
  let output = '';

  for (const line of lines) {
    const next =
      output
        ? `${output}\n${line}`
        : line;

    if (
      next.length > maximum
    ) {
      return (
        `${output}\n` +
        '…additional details omitted.'
      );
    }

    output = next;
  }

  return output;
}

module.exports = {
  data:
    new SlashCommandBuilder()
      .setName('ct-audit')

      .setDescription(
        'Match CT records to Discord and repair missing CT records or Discord IDs'
      )

      .addBooleanOption(
        option =>
          option
            .setName(
              'apply_changes'
            )

            .setDescription(
              'True writes safe changes; False only previews them'
            )

            .setRequired(true)
      ),

  async execute(interaction) {
    const owner =
      isBotOwner(interaction);

    if (
      !owner &&
      interaction.channelId !==
        process.env
          .ACCEPT_CADET_CHANNEL_ID
    ) {
      await interaction.reply({
        content:
          '❌ `/ct-audit` can only be used in the configured cadet-management channel.',

        ephemeral: true,

        allowedMentions: {
          parse: [],
        },
      });

      return;
    }

    if (
      !owner &&
      !interaction.member.roles.cache.has(
        process.env
          .OFFICER_ROLE_ID
      )
    ) {
      await interaction.reply({
        content:
          '❌ Only Officers can use `/ct-audit`.',

        ephemeral: true,

        allowedMentions: {
          parse: [],
        },
      });

      return;
    }

    await interaction.deferReply({
      ephemeral: true,
    });

    const applyChanges =
      interaction.options.getBoolean(
        'apply_changes',
        true
      );

    try {
      const [
        cadets,
        roster,
        ctData,
        guildMembers,
      ] = await Promise.all([
        getPeopleFromSheet(
          CADETS_TAB,
          'User ID'
        ),

        getPeopleFromSheet(
          ROSTER_TAB,
          'Discord ID'
        ),

        getCtData(),

        interaction.guild.members.fetch(),
      ]);

      const spreadsheetPeople = [
        ...cadets,
        ...roster,
      ];

      const memberRecords =
        guildMembers
          .filter(
            member =>
              !member.user.bot
          )

          .map(
            memberSearchRecord
          );

      const memberById =
        new Map(
          memberRecords.map(
            record => [
              record.member.id,
              record,
            ]
          )
        );

      const storedByDiscordId =
        new Map();

      const storedByIgn =
        new Map();

      for (
        const entry of
        ctData.entries
      ) {
        if (entry.discordId) {
          addToMap(
            storedByDiscordId,
            entry.discordId,
            entry
          );
        }

        if (
          entry.normalisedIgn
        ) {
          addToMap(
            storedByIgn,
            entry.normalisedIgn,
            entry
          );
        }
      }

      const proposals = [];

      const proposalRanges =
        new Set();

      const proposedByDiscordId =
        new Map();

      const proposedByCtNumber =
        new Map();

      const alreadyRecorded = [];
      const unmatched = [];
      const ambiguous = [];
      const conflicts = [];

      function addProposal(
        proposal
      ) {
        const existingStored =
          storedByDiscordId.get(
            proposal.discordId
          ) || [];

        if (
          existingStored.some(
            entry =>
              entry.ctNumber !==
              proposal.ctNumber
          )
        ) {
          conflicts.push(
            `${proposal.description}, but that Discord ID is already stored against ${existingStored
              .map(
                entry =>
                  entry.ctNumber
              )
              .join(', ')}`
          );

          return;
        }

        const proposedForId =
          proposedByDiscordId.get(
            proposal.discordId
          );

        if (
          proposedForId &&
          proposedForId !==
            proposal.ctNumber
        ) {
          conflicts.push(
            `Discord ID ${proposal.discordId} matched both ${proposedForId} and ${proposal.ctNumber}`
          );

          return;
        }

        const proposedForNumber =
          proposedByCtNumber.get(
            proposal.ctNumber
          );

        if (
          proposedForNumber &&
          proposedForNumber !==
            proposal.discordId
        ) {
          conflicts.push(
            `CT ${proposal.ctNumber} matched more than one Discord account`
          );

          return;
        }

        if (
          proposalRanges.has(
            proposal.range
          )
        ) {
          return;
        }

        proposalRanges.add(
          proposal.range
        );

        proposedByDiscordId.set(
          proposal.discordId,
          proposal.ctNumber
        );

        proposedByCtNumber.set(
          proposal.ctNumber,
          proposal.discordId
        );

        proposals.push(
          proposal
        );
      }

      for (
        const entry of
        ctData.entries
      ) {
        if (
          !entry.ign ||
          !entry.ctNumber
        ) {
          continue;
        }

        if (entry.discordId) {
          const memberRecord =
            memberById.get(
              entry.discordId
            );

          if (memberRecord) {
            alreadyRecorded.push(
              `${entry.ign} → ${entry.ctNumber}`
            );
          } else {
            conflicts.push(
              `${entry.ign} (${entry.ctNumber}) stores ${entry.discordId}, but that account is not currently in the server`
            );
          }

          continue;
        }

        const match =
          findDiscordMatchForCt({
            ctEntry: entry,

            spreadsheetPeople,

            memberRecords,

            memberById,
          });

        if (
          match.status ===
          'unmatched'
        ) {
          unmatched.push(
            match.reason
          );

          continue;
        }

        if (
          match.status ===
          'ambiguous'
        ) {
          ambiguous.push(
            match.reason
          );

          continue;
        }

        addProposal({
          type:
            'fill-discord-id',

          ctNumber:
            entry.ctNumber,

          discordId:
            match.record.member.id,

          range:
            `${quoteSheetName(
              CT_NUMBERS_TAB
            )}!${entry.discordIdCell}`,

          values: [[
            match.record.member.id,
          ]],

          description:
            `${entry.ign} (${entry.ctNumber}) → ` +
            `${match.record.member.user.username} via ${match.method}`,
        });
      }

      for (
        const person of
        spreadsheetPeople
      ) {
        const matchesById =
          person.discordId
            ? storedByDiscordId.get(
                person.discordId
              ) || []
            : [];

        const matchesByIgn =
          storedByIgn.get(
            person.normalisedIgn
          ) || [];

        if (
          matchesById.length ||
          matchesByIgn.length
        ) {
          continue;
        }

        const memberMatch =
          resolveSpreadsheetPersonToMember({
            person,

            memberRecords,

            memberById,
          });

        if (
          memberMatch.status ===
          'unmatched'
        ) {
          unmatched.push(
            `${person.sheetName} row ${person.rowNumber}: ${memberMatch.reason}`
          );

          continue;
        }

        if (
          memberMatch.status ===
          'ambiguous'
        ) {
          ambiguous.push(
            `${person.sheetName} row ${person.rowNumber}: ${memberMatch.reason}`
          );

          continue;
        }

        const ctNumbers =
          extractCtNumbers(
            memberMatch.record
          );

        if (!ctNumbers.length) {
          unmatched.push(
            `${person.name}: matched ${memberMatch.record.member.displayName}, but no CT number was found at the end of their nickname`
          );

          continue;
        }

        if (
          ctNumbers.length > 1
        ) {
          ambiguous.push(
            `${person.name}: Discord nickname contains more than one possible CT number: ${ctNumbers.join(', ')}`
          );

          continue;
        }

        const plan =
          planMissingCtRecord({
            ctData,
            person,

            memberRecord:
              memberMatch.record,

            ctNumber:
              ctNumbers[0],
          });

        if (
          plan.status ===
          'conflict'
        ) {
          conflicts.push(
            plan.reason
          );

          continue;
        }

        addProposal(plan);
      }

      if (
        applyChanges &&
        proposals.length
      ) {
        const sheets =
          await getSheetsClient();

        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId:
            getSpreadsheetId(),

          requestBody: {
            valueInputOption:
              'RAW',

            data:
              proposals.map(
                proposal => ({
                  range:
                    proposal.range,

                  values:
                    proposal.values,
                })
              ),
          },
        });
      }

      const namedCtRecords =
        ctData.entries.filter(
          entry =>
            entry.ign
        ).length;

      const issueCount =
        unmatched.length +
        ambiguous.length +
        conflicts.length;

      const responseLines = [
        '✅ **CT audit complete.**',

        `**Mode:** ${
          applyChanges
            ? 'Changes applied'
            : 'Dry run only'
        }`,

        `**Discord members scanned:** ${memberRecords.length}`,

        `**Cadets/Roster records scanned:** ${spreadsheetPeople.length}`,

        `**Named CT records scanned:** ${namedCtRecords}`,

        `**Already recorded:** ${alreadyRecorded.length}`,

        `**Safe changes found:** ${proposals.length}`,

        `**Changes written:** ${
          applyChanges
            ? proposals.length
            : 0
        }`,

        `**Issues requiring attention:** ${issueCount}`,

        ...limitedList(
          'Safe changes',

          proposals.map(
            proposal =>
              proposal.description
          )
        ),

        ...limitedList(
          'Unmatched',
          unmatched
        ),

        ...limitedList(
          'Ambiguous',
          ambiguous
        ),

        ...limitedList(
          'Conflicts',
          conflicts
        ),
      ];

      if (
        !applyChanges &&
        proposals.length
      ) {
        responseLines.push(
          '',

          `Run \`/ct-audit apply_changes:True\` to write ${proposals.length} safe change(s).`
        );
      }

      await interaction.editReply({
        content:
          fitDiscordMessage(
            responseLines
          ),

        allowedMentions: {
          parse: [],
        },
      });

      await sendLog(
        interaction.guild,

        process.env.LOG_CHANNEL_ID,

        [
          '**[ORION — CT AUDIT]**',

          `**Run by:** ${interaction.user.tag} (${interaction.user.id})`,

          `**Owner bypass used:** ${
            owner ? 'Yes' : 'No'
          }`,

          `**Mode:** ${
            applyChanges
              ? 'Changes applied'
              : 'Dry run'
          }`,

          `**Cadets/Roster scanned:** ${spreadsheetPeople.length}`,

          `**Named CT records scanned:** ${namedCtRecords}`,

          `**Already recorded:** ${alreadyRecorded.length}`,

          `**Safe changes:** ${proposals.length}`,

          `**Changes written:** ${
            applyChanges
              ? proposals.length
              : 0
          }`,

          `**Issues:** ${issueCount}`,
        ].join('\n')
      ).catch(error =>
        console.error(
          'CT audit logging failed:',
          error
        )
      );
    } catch (error) {
      console.error(
        '/ct-audit failed:',
        error
      );

      await interaction.editReply({
        content:
          `❌ CT audit failed: ${
            error.message ||
            'Unknown error'
          }`,

        allowedMentions: {
          parse: [],
        },
      });
    }
  },
};