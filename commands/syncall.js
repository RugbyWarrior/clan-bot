const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  getTraineeRows,
  getRatingsRows,
  batchUpdateRatingsCells,
  normalizeName,
  deleteTraineeRow,
} = require('../sheets');
const { mosConfig } = require('../mosConfig');

const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;

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

function getAllowedRankRole(member, guild) {
  const rankRoleIds = (process.env.RANK_ROLE_IDS || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);

  for (const roleId of rankRoleIds) {
    if (!roleId) continue;
    if (roleId === TRAINEE_ROLE_ID || roleId === EX_SKIRA_ROLE_ID) continue;

    if (member.roles.cache.has(roleId)) {
      return guild.roles.cache.get(roleId) || null;
    }
  }

  return null;
}

function resolveSheetColumn(columnOrHeader) {
  if (!columnOrHeader) return null;

  const trimmed = String(columnOrHeader).trim();

  if (/^[A-S]$/i.test(trimmed)) {
    return trimmed.toUpperCase();
  }

  const headerToColumn = {
    Rank: 'A',
    Squadron: 'B',
    Name: 'C',
    Infantry: 'D',
    'Squad Leader': 'E',
    Shooter: 'F',
    AT: 'G',
    Engineer: 'H',
    Medic: 'I',
    Grenadier: 'J',
    Mortar: 'K',
    Pilot: 'L',
    Driver: 'M',
    'APC/IFV Gunner': 'N',
    'MBT Gunner': 'O',
    'Fire Support (R)': 'P',
    'Irregular Warfare (R)': 'Q',
    'Knife (R)': 'R',
    'Discord ID': 'S',
  };

  return headerToColumn[trimmed] || null;
}

function getMosSheetValue(member, guild, mos) {
  for (const [ratingName, roleId] of Object.entries(mos.ratings)) {
    if (!roleId) continue;

    if (member.roles.cache.has(roleId)) {
      const roleObject = guild.roles.cache.get(roleId);
      return roleObject ? roleObject.name : ratingName;
    }
  }

  return '';
}

function buildNameCandidatesFromRaw(raw) {
  const cleaned = normalizeName(raw);
  const candidates = new Set();

  if (cleaned) {
    candidates.add(cleaned);

    const spaced = String(raw || '')
      .trim()
      .toLowerCase()
      .replace(/\s+\(ex skira\)$/i, '')
      .replace(/\s+\[[^\]]+\]\s*$/g, '')
      .replace(/"[^"]*"/g, '')
      .replace(
        /^(tr|pvt|cdt|o\/cdt|tpr|p\/o|lcpl|cpl|sgt|ssgt|lt|f\/o|2lt|wo1|wo2|cpt|maj|col|brig|lieutenant|captain|major|colonel|brigadier|flight lieutenant|squadron leader)\s+/i,
        ''
      )
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const parts = spaced.split(' ').filter(Boolean);

    if (parts.length > 0) candidates.add(parts[0]);
    if (parts.length > 1) candidates.add(parts.slice(0, 2).join(''));
    if (parts.length > 1) candidates.add(parts.slice(0, 2).join(' '));
    if (parts.length > 0) candidates.add(parts.join(''));
  }

  return [...candidates];
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('syncall')
    .setDescription('Safely sync member identity and MOS data into the Ratings sheet.'),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (ALLOWED_CHANNEL_ID && interaction.channelId !== ALLOWED_CHANNEL_ID) {
        return interaction.editReply({
          content: 'This command can only be used in the allowed channel.',
          allowedMentions: { users: [] },
        });
      }

      const allMembers = interaction.guild.members.cache.filter(member => !member.user.bot);
      const traineeRows = await getTraineeRows();
      const ratingsRows = await getRatingsRows();

      const traineeByDiscordId = new Map();
      const traineeByCandidate = new Map();
      const ratingsByDiscordId = new Map();
      const ratingsByCandidate = new Map();

      for (let i = 1; i < traineeRows.length; i++) {
        const row = traineeRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[0] || '';
        const rowDiscordId = (row[8] || '').toString().trim();

        const record = {
          rowNumber,
          rowValues: row,
        };

        if (rowDiscordId) {
          traineeByDiscordId.set(rowDiscordId, record);
        }

        for (const candidate of buildNameCandidatesFromRaw(rowName)) {
          if (!candidate) continue;
          if (!traineeByCandidate.has(candidate)) traineeByCandidate.set(candidate, []);
          traineeByCandidate.get(candidate).push(record);
        }
      }

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        const record = {
          rowNumber,
          rowValues: row,
        };

        if (rowDiscordId) {
          ratingsByDiscordId.set(rowDiscordId, record);
        }

        for (const candidate of buildNameCandidatesFromRaw(rowName)) {
          if (!candidate) continue;
          if (!ratingsByCandidate.has(candidate)) ratingsByCandidate.set(candidate, []);
          ratingsByCandidate.get(candidate).push(record);
        }
      }

      const ratingsUpdates = [];
      const traineeRowsToDelete = [];
      let nextNewRatingsRow = ratingsRows.length + 1;

      let createdRows = 0;
      let updatedRows = 0;
      let removedTraineeRows = 0;
      let skippedTrainees = 0;
      let skippedExSkira = 0;
      let skippedNoAllowedRank = 0;
      let skippedNoSafeMatch = 0;
      let skippedAmbiguous = 0;

      for (const member of allMembers.values()) {
        if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) {
          skippedTrainees++;
          continue;
        }

        if (EX_SKIRA_ROLE_ID && member.roles.cache.has(EX_SKIRA_ROLE_ID)) {
          skippedExSkira++;
          continue;
        }

        const rankRole = getAllowedRankRole(member, interaction.guild);
        if (!rankRole) {
          skippedNoAllowedRank++;
          continue;
        }

        const squadronName = getSheetSquadronName(member);
        const memberCandidates = buildNameCandidatesFromRaw(member.nickname || member.user.username);

        let ratingsRow = ratingsByDiscordId.get(member.id) || null;

        if (!ratingsRow) {
          const ratingMatches = new Map();

          for (const candidate of memberCandidates) {
            for (const match of ratingsByCandidate.get(candidate) || []) {
              const existingDiscordId = (match.rowValues[18] || '').toString().trim();
              if (!existingDiscordId || existingDiscordId === member.id) {
                ratingMatches.set(match.rowNumber, match);
              }
            }
          }

          const ratingMatchesArr = [...ratingMatches.values()];
          if (ratingMatchesArr.length === 1) {
            ratingsRow = ratingMatchesArr[0];
          } else if (ratingMatchesArr.length > 1) {
            skippedAmbiguous++;
            continue;
          }
        }

        let traineeRow = traineeByDiscordId.get(member.id) || null;

        if (!traineeRow) {
          const traineeMatches = new Map();

          for (const candidate of memberCandidates) {
            for (const match of traineeByCandidate.get(candidate) || []) {
              traineeMatches.set(match.rowNumber, match);
            }
          }

          const traineeMatchesArr = [...traineeMatches.values()];
          if (traineeMatchesArr.length === 1) {
            traineeRow = traineeMatchesArr[0];
          } else if (traineeMatchesArr.length > 1) {
            skippedAmbiguous++;
            continue;
          }
        }

        const finalName =
          traineeRow?.rowValues?.[0] ||
          ratingsRow?.rowValues?.[2] ||
          getDisplayNameWithoutRank(member);

        if (!ratingsRow && !traineeRow) {
          skippedNoSafeMatch++;
          continue;
        }

        const targetRowNumber = ratingsRow ? ratingsRow.rowNumber : nextNewRatingsRow++;

        if (!ratingsRow) {
          createdRows++;
        } else {
          updatedRows++;
        }

        const rowUpdate = new Array(19).fill('');
        rowUpdate[0] = rankRole.name;
        rowUpdate[1] = squadronName;
        rowUpdate[2] = finalName;

        for (const config of Object.values(mosConfig)) {
          const column = resolveSheetColumn(config.sheetColumn);
          if (!column) continue;

          const index = column.charCodeAt(0) - 65;
          rowUpdate[index] = getMosSheetValue(member, interaction.guild, config);
        }

        rowUpdate[18] = member.id;

        ratingsUpdates.push({
          range: `${RATINGS_SHEET_NAME}!A${targetRowNumber}:S${targetRowNumber}`,
          values: [rowUpdate],
        });

        if (traineeRow) {
          traineeRowsToDelete.push(traineeRow.rowNumber);
        }
      }

      if (ratingsUpdates.length > 0) {
        await batchUpdateRatingsCells(ratingsUpdates);
      }

      const uniqueRowsToDelete = [...new Set(traineeRowsToDelete)].sort((a, b) => b - a);

      for (const rowNumber of uniqueRowsToDelete) {
        await deleteTraineeRow(rowNumber);
        removedTraineeRows++;
      }

      await sendLog(
        interaction.guild,
        process.env.LOG_CHANNEL_ID,
        [
          '**[SYNCALL]**',
          `**Created Rows:** ${createdRows}`,
          `**Updated Rows:** ${updatedRows}`,
          `**Removed Trainee Rows:** ${removedTraineeRows}`,
          `**Skipped Trainees:** ${skippedTrainees}`,
          `**Skipped Ex Skira:** ${skippedExSkira}`,
          `**Skipped No Allowed Rank:** ${skippedNoAllowedRank}`,
          `**Skipped No Safe Match:** ${skippedNoSafeMatch}`,
          `**Skipped Ambiguous:** ${skippedAmbiguous}`,
          `**Done By:** ${interaction.user.tag}`,
          `**Channel:** <#${interaction.channelId}>`,
        ].join('\n')
      );

      return interaction.editReply({
        content:
          `✅ Sync all complete.\n` +
          `Created rows: **${createdRows}**\n` +
          `Updated rows: **${updatedRows}**\n` +
          `Removed trainee rows: **${removedTraineeRows}**\n` +
          `Skipped trainees: **${skippedTrainees}**\n` +
          `Skipped Ex Skira: **${skippedExSkira}**\n` +
          `Skipped no allowed rank: **${skippedNoAllowedRank}**\n` +
          `Skipped no safe match: **${skippedNoSafeMatch}**\n` +
          `Skipped ambiguous: **${skippedAmbiguous}**`,
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /syncall:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /syncall.',
        allowedMentions: { users: [] },
      });
    }
  },
};