const { SlashCommandBuilder } = require('discord.js');
const { google } = require('googleapis');
const { sendLog } = require('../logger');
const {
  getRatingsRows,
} = require('../sheets');

const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const RATINGS_SHEET_NAME = (process.env.RATINGS_SHEET_NAME || 'Ratings').trim();
const RATINGS_SPREADSHEET_ID = (
  process.env.RATINGS_SPREADSHEET_ID ||
  process.env.MOS_SPREADSHEET_ID ||
  process.env.TRAINEE_SPREADSHEET_ID ||
  ''
).trim();

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

function truncateList(items, limit = 10) {
  if (items.length <= limit) return items;
  return [...items.slice(0, limit), `...and ${items.length - limit} more`];
}

async function getSheetsClient() {
  const authClient = await auth.getClient();
  return google.sheets({
    version: 'v4',
    auth: authClient,
  });
}

async function getRatingsSheetId() {
  const sheets = await getSheetsClient();
  const response = await sheets.spreadsheets.get({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
  });

  const targetSheet = response.data.sheets.find(
    s => s.properties.title === RATINGS_SHEET_NAME
  );

  if (!targetSheet) {
    throw new Error(`Could not find sheet tab named "${RATINGS_SHEET_NAME}".`);
  }

  return targetSheet.properties.sheetId;
}

async function deleteRatingsRows(rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;

  const sheets = await getSheetsClient();
  const sheetId = await getRatingsSheetId();

  const requests = [...new Set(rowNumbers)]
    .sort((a, b) => b - a)
    .map(rowNumber => ({
      deleteDimension: {
        range: {
          sheetId,
          dimension: 'ROWS',
          startIndex: rowNumber - 1,
          endIndex: rowNumber,
        },
      },
    }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
    requestBody: {
      requests,
    },
  });
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('cleanratings')
    .setDescription('Remove clearly invalid rows from the Ratings sheet.')
    .addStringOption(option =>
      option
        .setName('mode')
        .setDescription('What type of invalid rows to remove')
        .setRequired(true)
        .addChoices(
          { name: 'Trainees only', value: 'trainees' },
          { name: 'Ex Skira only', value: 'exskira' },
          { name: 'Missing Discord member only', value: 'missing_member' },
          { name: 'All clear invalid rows', value: 'all_invalid' }
        )
    ),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (interaction.channelId !== HQ_CHANNEL_ID) {
        return interaction.editReply({
          content: '❌ This command can only be used in the HQ channel.',
          allowedMentions: { users: [] },
        });
      }

      const mode = interaction.options.getString('mode');
      const allMembers = interaction.guild.members.cache.filter(member => !member.user.bot);
      const memberById = new Map();

      for (const member of allMembers.values()) {
        memberById.set(member.id, member);
      }

      const ratingsRows = await getRatingsRows();

      const traineeRows = [];
      const exSkiraRows = [];
      const missingMemberRows = [];

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        if (!rowDiscordId) {
          continue;
        }

        const member = memberById.get(rowDiscordId);

        if (!member) {
          missingMemberRows.push({ rowNumber, rowName, reason: `Missing member (${rowDiscordId})` });
          continue;
        }

        if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) {
          traineeRows.push({ rowNumber, rowName, reason: `Linked to trainee ${member.user.tag}` });
        }

        if (EX_SKIRA_ROLE_ID && member.roles.cache.has(EX_SKIRA_ROLE_ID)) {
          exSkiraRows.push({ rowNumber, rowName, reason: `Linked to Ex Skira ${member.user.tag}` });
        }
      }

      let rowsToDelete = [];

      if (mode === 'trainees') {
        rowsToDelete = traineeRows;
      } else if (mode === 'exskira') {
        rowsToDelete = exSkiraRows;
      } else if (mode === 'missing_member') {
        rowsToDelete = missingMemberRows;
      } else if (mode === 'all_invalid') {
        rowsToDelete = [...traineeRows, ...exSkiraRows, ...missingMemberRows];
      }

      if (rowsToDelete.length === 0) {
        return interaction.editReply({
          content: 'No matching invalid Ratings rows were found for that mode.',
          allowedMentions: { users: [] },
        });
      }

      await deleteRatingsRows(rowsToDelete.map(row => row.rowNumber));

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🧹 **CLEAN RATINGS**',
          `Moderator: <@${interaction.user.id}>`,
          `Mode: ${mode}`,
          `Rows Deleted: ${rowsToDelete.length}`,
          `Deleted Rows:\n- ${truncateList(rowsToDelete.map(r => `Row ${r.rowNumber} | ${r.rowName} | ${r.reason}`), 25).join('\n- ')}`,
        ].join('\n')
      );

      return interaction.editReply({
        content: [
          '**Ratings cleanup complete.**',
          `Mode: **${mode}**`,
          `Rows deleted: **${rowsToDelete.length}**`,
          '',
          `Deleted examples:\n- ${truncateList(rowsToDelete.map(r => `Row ${r.rowNumber} | ${r.rowName} | ${r.reason}`), 8).join('\n- ')}`,
        ].join('\n'),
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /cleanratings:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /cleanratings.',
        allowedMentions: { users: [] },
      });
    }
  },
};