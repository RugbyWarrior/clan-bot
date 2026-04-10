const { SlashCommandBuilder } = require('discord.js');
const { google } = require('googleapis');
const { sendLog } = require('../logger');
const { updateNickname } = require('../utils/updateNickname');
const { sortRatingsSheet } = require('../sheets');
const {
  findRatingsRowByDiscordId,
  writeRatingsRow,
  batchUpdateRatingsCells,
  findTraineeRowByDiscordId,
  findTraineeRowsByName,
  deleteTraineeRow,
} = require('../sheets');

const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const HQ_ROLE_ID = process.env.HQ_ROLE_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;

const RATINGS_SHEET_NAME = (process.env.RATINGS_SHEET_NAME || 'Ratings').trim();
const TRAINEE_SHEET_NAME = (process.env.TRAINEE_SHEET_NAME || 'Trainees').trim();

const TRAINEE_SPREADSHEET_ID = (process.env.TRAINEE_SPREADSHEET_ID || '').trim();
const RATINGS_SPREADSHEET_ID = (
  process.env.RATINGS_SPREADSHEET_ID ||
  process.env.MOS_SPREADSHEET_ID ||
  process.env.TRAINEE_SPREADSHEET_ID ||
  ''
).trim();

function canUseHqCommand(interaction) {
  if (!HQ_CHANNEL_ID || !HQ_ROLE_ID) return false;

  return (
    interaction.channelId === HQ_CHANNEL_ID &&
    interaction.member &&
    interaction.member.roles &&
    interaction.member.roles.cache &&
    interaction.member.roles.cache.has(HQ_ROLE_ID)
  );
}

function getRankOrder(rankLabel) {
  const order = {
    'Brigadier': 1,
    'Colonel': 2,
    'Major': 3,
    'Captain': 4,
    'Lieutenant': 5,
    'Squadron Leader': 6,
    'S/L': 6,
    'Second Lieutenant': 7,
    '2LT': 7,
    'Flight Lieutenant': 8,
    'Warrant Officer 1': 9,
    'WO1': 9,
    'Warrant Officer 2': 10,
    'WO2': 10,
    'Staff Sergeant': 11,
    'Sergeant': 12,
    'Corporal': 13,
    'Flying Officer': 14,
    'Flight Officer': 14,
    'F/O': 14,
    'Lance Corporal': 15,
    'Pilot Officer': 16,
    'Armour Trooper': 17,
    'Officer Cadet': 18,
    'Armour Cadet': 19,
    'Private': 20,
    'Trainee': 99,
    'Ex Skira': 999,
  };

  return order[rankLabel] ?? 999;
}

function getDisplayNameWithoutRank(member) {
  const displayName = member.nickname || member.user.username;
  if (!displayName.includes(' ')) return displayName;
  return displayName.replace(/^\S+\s+/, '').trim();
}

function getSheetSquadronName(member) {
  if (member.roles.cache.has(process.env.INFANTRY_ROLE_ID)) return '3rd Rifles';
  if (member.roles.cache.has(process.env.ARMOUR_ROLE_ID)) return '20th Hussars';
  if (member.roles.cache.has(process.env.AVIATION_ROLE_ID)) return '230th Aviation';
  return '';
}

function formatUkDate(date = new Date()) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear());
  return `${day}/${month}/${year}`;
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const authClient = await auth.getClient();

  return google.sheets({
    version: 'v4',
    auth: authClient,
  });
}

async function getSheetId(spreadsheetId, sheetName) {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  const targetSheet = response.data.sheets.find(
    s => s.properties.title === sheetName
  );

  if (!targetSheet) {
    throw new Error(`Could not find sheet tab named "${sheetName}".`);
  }

  return targetSheet.properties.sheetId;
}

async function deleteSheetRow(spreadsheetId, sheetName, rowNumber) {
  const sheets = await getSheetsClient();
  const sheetId = await getSheetId(spreadsheetId, sheetName);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId,
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

async function getSheetValues(spreadsheetId, range) {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  return response.data.values || [];
}

async function updateSheetRange(spreadsheetId, range, values) {
  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values,
    },
  });
}

async function deleteRatingsRowByDiscordId(discordId) {
  if (!discordId) return null;

  const ratingsRow = await findRatingsRowByDiscordId(discordId);
  if (!ratingsRow) return null;

  await deleteSheetRow(RATINGS_SPREADSHEET_ID, RATINGS_SHEET_NAME, ratingsRow.rowNumber);
  return ratingsRow.rowNumber;
}

async function upsertTraineeRow({ name, steamId64, discordId }) {
  const rows = await getSheetValues(
    TRAINEE_SPREADSHEET_ID,
    `${TRAINEE_SHEET_NAME}!A:I`
  );

  let targetRowNumber = null;
  let existingRow = null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowDiscordId = (row[8] || '').toString().trim();

    if (rowDiscordId && rowDiscordId === String(discordId).trim()) {
      targetRowNumber = i + 1;
      existingRow = row;
      break;
    }
  }

  if (!targetRowNumber) {
    const firstEmpty = rows.findIndex((row, index) => {
      if (index === 0) return false;
      const safeRow = row || [];
      return safeRow.every(cell => !String(cell || '').trim());
    });

    if (firstEmpty !== -1) {
      targetRowNumber = firstEmpty + 1;
    } else {
      targetRowNumber = rows.length + 1;
    }
  }

  const rowValues = new Array(9).fill('');

  if (existingRow) {
    for (let i = 0; i < Math.min(existingRow.length, 9); i++) {
      rowValues[i] = existingRow[i] || '';
    }
  }

  rowValues[0] = name || rowValues[0] || '';
  rowValues[1] = rowValues[1] || formatUkDate();
  rowValues[3] = steamId64 || rowValues[3] || '';
  rowValues[8] = discordId || rowValues[8] || '';

  await updateSheetRange(
    TRAINEE_SPREADSHEET_ID,
    `${TRAINEE_SHEET_NAME}!A${targetRowNumber}:I${targetRowNumber}`,
    [rowValues]
  );

  return targetRowNumber;
}

const rankConfig = {
  EX_SKIRA: {
    label: 'Ex Skira',
    roleEnv: 'EX_SKIRA_ROLE_ID',
    breakerEnv: null,
    allowedSquadrons: [],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: null,
    exSkiraSuffix: true,
  },
  TRAINEE: {
    label: 'Trainee',
    roleEnv: 'TRAINEE_ROLE_ID',
    breakerEnv: null,
    allowedSquadrons: [],
    traineeState: true,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'TR',
    exSkiraSuffix: false,
  },

  PRIVATE: {
    label: 'Private',
    roleEnv: 'PRIVATE_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: true,
    nicknamePrefix: 'PVT',
    exSkiraSuffix: false,
  },
  ARMOUR_CADET: {
    label: 'Armour Cadet',
    roleEnv: 'ARMOUR_CADET_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'CDT',
    exSkiraSuffix: false,
  },
  OFFICER_CADET: {
    label: 'Officer Cadet',
    roleEnv: 'OFFICER_CADET_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'O/CDT',
    exSkiraSuffix: false,
  },
  ARMOUR_TROOPER: {
    label: 'Armour Trooper',
    roleEnv: 'ARMOUR_TROOPER_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'TPR',
    exSkiraSuffix: false,
  },
  PILOT_OFFICER: {
    label: 'Pilot Officer',
    roleEnv: 'PILOT_OFFICER_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'P/O',
    exSkiraSuffix: false,
  },
  LANCE_CORPORAL: {
    label: 'Lance Corporal',
    roleEnv: 'LANCE_CORPORAL_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'LCPL',
    exSkiraSuffix: false,
  },

  FLYING_OFFICER: {
    label: 'Flying Officer',
    roleEnv: 'FLYING_OFFICER_ROLE_ID',
    breakerEnv: 'NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'F/O',
    exSkiraSuffix: false,
  },
  CORPORAL: {
    label: 'Corporal',
    roleEnv: 'CORPORAL_ROLE_ID',
    breakerEnv: 'NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'CPL',
    exSkiraSuffix: false,
  },

  SERGEANT: {
    label: 'Sergeant',
    roleEnv: 'SERGEANT_ROLE_ID',
    breakerEnv: 'SENIOR_NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID', 'AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'SGT',
    exSkiraSuffix: false,
  },
  STAFF_SERGEANT: {
    label: 'Staff Sergeant',
    roleEnv: 'STAFF_SERGEANT_ROLE_ID',
    breakerEnv: 'SENIOR_NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID', 'AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'SSGT',
    exSkiraSuffix: false,
  },
};

module.exports = {
  data: new SlashCommandBuilder()
    .setName('setrank')
    .setDescription('Set a user rank role and sync them to the system.')
    .addUserOption(option =>
      option.setName('user')
        .setDescription('User to update')
        .setRequired(true)
    )
    .addStringOption(option =>
      option.setName('rank')
        .setDescription('Rank to assign')
        .setRequired(true)
        .addChoices(
          { name: 'Ex Skira', value: 'EX_SKIRA' },
          { name: 'Trainee', value: 'TRAINEE' },
          { name: 'Private', value: 'PRIVATE' },
          { name: 'Armour Cadet', value: 'ARMOUR_CADET' },
          { name: 'Officer Cadet', value: 'OFFICER_CADET' },
          { name: 'Armour Trooper', value: 'ARMOUR_TROOPER' },
          { name: 'Pilot Officer', value: 'PILOT_OFFICER' },
          { name: 'Lance Corporal', value: 'LANCE_CORPORAL' },
          { name: 'Flying Officer', value: 'FLYING_OFFICER' },
          { name: 'Corporal', value: 'CORPORAL' },
          { name: 'Sergeant', value: 'SERGEANT' },
          { name: 'Staff Sergeant', value: 'STAFF_SERGEANT' }
        )
    )
    .addStringOption(option =>
      option
        .setName('steamid64')
        .setDescription('Optional SteamID64 (used when setting someone to Trainee)')
        .setRequired(false)
    ),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (!canUseHqCommand(interaction)) {
        await interaction.editReply({
          content: '❌ This command can only be used by the Headquarters role in the HQ channel.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const user = interaction.options.getUser('user');
      const rankKey = interaction.options.getString('rank');
      const providedSteamId64 = (interaction.options.getString('steamid64') || '').trim();

      let member = await interaction.guild.members.fetch(user.id);

      const selectedRank = rankConfig[rankKey];
      if (!selectedRank) {
        await interaction.editReply({
          content: '❌ That rank is not configured.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const targetRankRoleId = process.env[selectedRank.roleEnv];
      const breakerRoleToAdd = selectedRank.breakerEnv
        ? process.env[selectedRank.breakerEnv]
        : null;

      if (!targetRankRoleId) {
        await interaction.editReply({
          content: `❌ Missing ${selectedRank.roleEnv} in .env`,
          allowedMentions: { users: [] },
        });
        return;
      }

      if (member.roles.cache.has(targetRankRoleId)) {
        await interaction.editReply({
          content: `<@${user.id}> already has rank **${selectedRank.label}**.`,
          allowedMentions: { users: [] },
        });
        return;
      }

      const rankRoleIds = (process.env.RANK_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const breakerRoleIds = (process.env.BREAKER_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const squadronRoleIds = (process.env.SQUADRON_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const traineeRoleId = process.env.TRAINEE_ROLE_ID;
      const trainingCompanyRoleId = process.env.TRAINING_COMPANY_ROLE_ID;
      const infantryRoleId = process.env.INFANTRY_ROLE_ID;
      const skiraMemberRoleId = process.env.SKIRA_MEMBER_ROLE_ID;

      const wasCurrentlyTrainee = traineeRoleId
        ? member.roles.cache.has(traineeRoleId)
        : false;

      let traineeRow = null;
      let traineeSheetName = null;
      let traineeSteamId64 = '';

      if (wasCurrentlyTrainee) {
        traineeRow = await findTraineeRowByDiscordId(user.id);

        if (!traineeRow) {
          const nameMatches = await findTraineeRowsByName(getDisplayNameWithoutRank(member));
          if (nameMatches.length === 1) {
            traineeRow = nameMatches[0];
          }
        }

        if (traineeRow) {
          traineeSheetName = traineeRow.rowValues[0] || null;
          traineeSteamId64 = (traineeRow.rowValues[3] || '').toString().trim();
        }
      }

      const existingRatingsRow = await findRatingsRowByDiscordId(user.id);
      const ratingsSteamId64 = existingRatingsRow
        ? (existingRatingsRow.rowValues?.[19] || '').toString().trim()
        : '';

      if (selectedRank.traineeState) {
        const rolesToRemove = [];

        for (const rankRoleId of rankRoleIds) {
          if (member.roles.cache.has(rankRoleId) && rankRoleId !== targetRankRoleId) {
            rolesToRemove.push(rankRoleId);
          }
        }

        for (const breakerRoleId of breakerRoleIds) {
          if (member.roles.cache.has(breakerRoleId)) {
            rolesToRemove.push(breakerRoleId);
          }
        }

        for (const squadronRoleId of squadronRoleIds) {
          if (member.roles.cache.has(squadronRoleId)) {
            rolesToRemove.push(squadronRoleId);
          }
        }

        if (skiraMemberRoleId && member.roles.cache.has(skiraMemberRoleId)) {
          rolesToRemove.push(skiraMemberRoleId);
        }

        if (rolesToRemove.length > 0) {
          await member.roles.remove(rolesToRemove, 'Resetting member to trainee state');
        }

        if (!member.roles.cache.has(targetRankRoleId)) {
          await member.roles.add(targetRankRoleId, 'Rank set to Trainee');
        }

        if (!trainingCompanyRoleId) {
          await interaction.editReply({
            content: '❌ Missing TRAINING_COMPANY_ROLE_ID in .env',
            allowedMentions: { users: [] },
          });
          return;
        }

        if (!member.roles.cache.has(trainingCompanyRoleId)) {
          await member.roles.add(trainingCompanyRoleId, 'Training Company role added for Trainee');
        }

        await updateNickname(member, {
          prefix: selectedRank.nicknamePrefix,
          exSkira: selectedRank.exSkiraSuffix,
        });

        const ratingsRowDeleted = await deleteRatingsRowByDiscordId(user.id);

        const steamId64ToWrite =
          providedSteamId64 ||
          ratingsSteamId64 ||
          traineeSteamId64 ||
          '';

        const traineeRowNumber = await upsertTraineeRow({
          name: getDisplayNameWithoutRank(member),
          steamId64: steamId64ToWrite,
          discordId: user.id,
        });

        await interaction.editReply({
          content:
            `✅ Set <@${user.id}> to **${selectedRank.label}** and added **Training Company**.\n` +
            `Trainee row updated: **${traineeRowNumber}**\n` +
            `SteamID64 used: **${steamId64ToWrite || 'Blank'}**\n` +
            `Ratings row removed: **${ratingsRowDeleted ? `Yes (row ${ratingsRowDeleted})` : 'No existing row found'}**`,
          allowedMentions: { users: [] },
        });

        await sendLog(
          interaction.guild,
          LOG_CHANNEL_ID,
          [
            '**[SET RANK]**',
            `**User:** <@${user.id}> (${user.id})`,
            `**New Rank:** ${selectedRank.label}`,
            `**Breaker Added:** None`,
            `**Squadron Added:** None`,
            `**Skira Member Removed:** ${skiraMemberRoleId ? 'Yes if present' : 'No config'}`,
            `**Ratings Row Removed:** ${ratingsRowDeleted ? `Yes (row ${ratingsRowDeleted})` : 'No'}`,
            `**Trainee Row Upserted:** ${traineeRowNumber}`,
            `**SteamID64 Used:** ${steamId64ToWrite || 'Blank'}`,
            `**Done By:** ${interaction.user.tag}`,
            `**Channel:** <#${interaction.channelId}>`,
          ].join('\n')
        );
        return;
      }

      const currentSquadronEnv = selectedRank.allowedSquadrons.find(envName => {
        const roleId = process.env[envName];
        return roleId && member.roles.cache.has(roleId);
      });

      let squadronRoleToAdd = null;
      let autoAssignedInfantry = false;
      let addedSkiraMember = false;
      let removedSkiraMember = false;

      if (selectedRank.autoInfantryFromTrainee && wasCurrentlyTrainee) {
        if (!infantryRoleId) {
          await interaction.editReply({
            content: '❌ Missing INFANTRY_ROLE_ID in .env',
            allowedMentions: { users: [] },
          });
          return;
        }

        squadronRoleToAdd = infantryRoleId;
        autoAssignedInfantry = true;
      } else {
        if (selectedRank.allowedSquadrons.length > 0 && !currentSquadronEnv) {
          await interaction.editReply({
            content: `❌ <@${user.id}> does not have the required squadron role for **${selectedRank.label}**.`,
            allowedMentions: { users: [] },
          });
          return;
        }
      }

      const rolesToRemove = [];

      if (traineeRoleId && member.roles.cache.has(traineeRoleId)) {
        rolesToRemove.push(traineeRoleId);
      }

      if (trainingCompanyRoleId && member.roles.cache.has(trainingCompanyRoleId)) {
        rolesToRemove.push(trainingCompanyRoleId);
      }

      for (const rankRoleId of rankRoleIds) {
        if (member.roles.cache.has(rankRoleId) && rankRoleId !== targetRankRoleId) {
          rolesToRemove.push(rankRoleId);
        }
      }

      for (const breakerRoleId of breakerRoleIds) {
        if (member.roles.cache.has(breakerRoleId)) {
          rolesToRemove.push(breakerRoleId);
        }
      }

      if (squadronRoleToAdd) {
        for (const squadronRoleId of squadronRoleIds) {
          if (member.roles.cache.has(squadronRoleId) && squadronRoleId !== squadronRoleToAdd) {
            rolesToRemove.push(squadronRoleId);
          }
        }
      }

      if (rankKey === 'EX_SKIRA' && skiraMemberRoleId && member.roles.cache.has(skiraMemberRoleId)) {
        rolesToRemove.push(skiraMemberRoleId);
        removedSkiraMember = true;
      }

      if (rolesToRemove.length > 0) {
        await member.roles.remove(rolesToRemove, 'Cleaning old rank, breaker, trainee, and squadron roles');
      }

      if (!member.roles.cache.has(targetRankRoleId)) {
        await member.roles.add(targetRankRoleId, `Rank set to ${selectedRank.label}`);
      }

      if (breakerRoleToAdd && !member.roles.cache.has(breakerRoleToAdd)) {
        await member.roles.add(breakerRoleToAdd, `Breaker role added for ${selectedRank.label}`);
      }

      if (squadronRoleToAdd && !member.roles.cache.has(squadronRoleToAdd)) {
        await member.roles.add(squadronRoleToAdd, 'Auto-assigned Infantry on first promotion from trainee');
      }

      if (wasCurrentlyTrainee && skiraMemberRoleId && !member.roles.cache.has(skiraMemberRoleId) && rankKey !== 'EX_SKIRA') {
        await member.roles.add(
          skiraMemberRoleId,
          'Skira Member role added on first promotion from trainee'
        );
        addedSkiraMember = true;
      }

      await updateNickname(member, {
        prefix: selectedRank.nicknamePrefix,
        exSkira: selectedRank.exSkiraSuffix,
      });

      member = await interaction.guild.members.fetch(user.id);

      const squadronName = getSheetSquadronName(member);
      let breakerName = 'None';

      let replyMessage = `✅ Set <@${user.id}> to rank **${selectedRank.label}**.`;

      if (breakerRoleToAdd) {
        const breakerRole = interaction.guild.roles.cache.get(breakerRoleToAdd);
        if (breakerRole) {
          breakerName = breakerRole.name;
          replyMessage += ` Added breaker role **${breakerRole.name}**.`;
        }
      }

      if (squadronRoleToAdd) {
        const squadronRole = interaction.guild.roles.cache.get(squadronRoleToAdd);
        if (squadronRole) {
          replyMessage += ` Added squadron role **${squadronRole.name}**.`;
        }
      }

      if (autoAssignedInfantry) {
        replyMessage += ` Infantry was auto-assigned because they were promoted from trainee.`;
      }

      if (addedSkiraMember) {
        replyMessage += ` Added **Skira Member** role.`;
      }

      if (removedSkiraMember) {
        replyMessage += ` Removed **Skira Member** role.`;
      }

      let traineeRowRemoved = false;
      if (wasCurrentlyTrainee && traineeRow) {
        await deleteTraineeRow(traineeRow.rowNumber);
        traineeRowRemoved = true;
      }

      if (rankKey === 'EX_SKIRA') {
        const ratingsRowDeleted = await deleteRatingsRowByDiscordId(user.id);

        await interaction.editReply({
          content:
            replyMessage +
            ` Ratings row removed: **${ratingsRowDeleted ? `Yes (row ${ratingsRowDeleted})` : 'No existing row found'}**.`,
          allowedMentions: { users: [] },
        });

        await sendLog(
          interaction.guild,
          LOG_CHANNEL_ID,
          [
            '**[SET RANK]**',
            `**User:** <@${user.id}> (${user.id})`,
            `**New Rank:** ${selectedRank.label}`,
            `**Breaker Added:** ${breakerName}`,
            `**Squadron Added:** ${squadronName || 'None'}`,
            `**Auto Assigned Infantry:** ${autoAssignedInfantry ? 'Yes' : 'No'}`,
            `**Skira Member Removed:** ${removedSkiraMember ? 'Yes' : 'No'}`,
            `**Ratings Row Removed:** ${ratingsRowDeleted ? `Yes (row ${ratingsRowDeleted})` : 'No'}`,
            `**Trainee Row Removed:** ${traineeRowRemoved ? 'Yes' : 'No'}`,
            `**Done By:** ${interaction.user.tag}`,
            `**Channel:** <#${interaction.channelId}>`,
          ].join('\n')
        );
        return;
      }

      let ratingsRow = existingRatingsRow;
      let createdRatingsRow = false;

      const finalSheetName = traineeSheetName || getDisplayNameWithoutRank(member);
      const steamId64ForRatings =
        traineeSteamId64 ||
        providedSteamId64 ||
        ratingsSteamId64 ||
        '';

      const rankOrderForRatings = getRankOrder(selectedRank.label);

      if (!ratingsRow) {
        const newRow = new Array(21).fill('');
        newRow[0] = selectedRank.label;
        newRow[1] = squadronName;
        newRow[2] = finalSheetName;
        newRow[18] = user.id;
        newRow[19] = steamId64ForRatings;
        newRow[20] = rankOrderForRatings;

        const rowNumber = await writeRatingsRow(newRow);

        ratingsRow = {
          rowNumber,
          rowValues: newRow,
        };

        createdRatingsRow = true;
      }

      await batchUpdateRatingsCells([
        {
          range: `${RATINGS_SHEET_NAME}!A${ratingsRow.rowNumber}`,
          values: [[selectedRank.label]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!B${ratingsRow.rowNumber}`,
          values: [[squadronName]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!C${ratingsRow.rowNumber}`,
          values: [[finalSheetName]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!S${ratingsRow.rowNumber}`,
          values: [[user.id]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!T${ratingsRow.rowNumber}`,
          values: [[steamId64ForRatings]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!U${ratingsRow.rowNumber}`,
          values: [[rankOrderForRatings]],
        },
      ]);

      await sortRatingsSheet();

      await interaction.editReply({
        content:
          replyMessage +
          ` SteamID64 moved to Ratings: **${steamId64ForRatings || 'Blank'}**.`,
        allowedMentions: { users: [] },
      });

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '**[SET RANK]**',
          `**User:** <@${user.id}> (${user.id})`,
          `**New Rank:** ${selectedRank.label}`,
          `**Breaker Added:** ${breakerName}`,
          `**Squadron Added:** ${squadronName || 'None'}`,
          `**Auto Assigned Infantry:** ${autoAssignedInfantry ? 'Yes' : 'No'}`,
          `**Skira Member Added:** ${addedSkiraMember ? 'Yes' : 'No'}`,
          `**Ratings Sheet Name:** ${RATINGS_SHEET_NAME}`,
          `**Ratings Row:** ${ratingsRow.rowNumber}`,
          `**Ratings Row Created:** ${createdRatingsRow ? 'Yes' : 'No'}`,
          `**SteamID64 Written To Ratings:** ${steamId64ForRatings || 'Blank'}`,
          `**Rank Order Written:** ${rankOrderForRatings}`,
          `**Ratings Sheet Sorted:** Yes`,
          `**Trainee Row Found:** ${traineeRow ? 'Yes' : 'No'}`,
          `**Trainee Row Removed:** ${traineeRowRemoved ? 'Yes' : 'No'}`,
          `**Done By:** ${interaction.user.tag}`,
          `**Channel:** <#${interaction.channelId}>`,
        ].join('\n')
      );
    } catch (error) {
      console.error('Error in /setrank:', error);

      try {
        await interaction.editReply({
          content: '❌ There was an error while running /setrank. Check the console for details.',
          allowedMentions: { users: [] },
        });
      } catch {}
    }
  },
};