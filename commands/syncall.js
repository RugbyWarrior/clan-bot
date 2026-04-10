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

const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const HQ_ROLE_ID = process.env.HQ_ROLE_ID;
const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;

function canUseHqCommand(interaction) {
  return (
    interaction.channelId === HQ_CHANNEL_ID &&
    interaction.member?.roles?.cache?.has(HQ_ROLE_ID)
  );
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

  if (/^[A-T]$/i.test(trimmed)) {
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
    SteamID64: 'T',
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

function getUniqueMatchesByCandidates(candidateMap, candidates, extraFilter = null) {
  const matches = new Map();

  for (const candidate of candidates) {
    for (const match of candidateMap.get(candidate) || []) {
      if (extraFilter && !extraFilter(match)) continue;
      matches.set(match.rowNumber, match);
    }
  }

  return [...matches.values()];
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('syncall')
    .setDescription('Safely sync ranked Discord members into the Ratings sheet and remove promoted trainee rows.'),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (!canUseHqCommand(interaction)) {
        return interaction.editReply({
          content: '❌ This command can only be used by the Headquarters role in the HQ channel.',
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
      let skippedAmbiguousRatings = 0;
      let skippedAmbiguousTrainees = 0;

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
          const ratingMatches = getUniqueMatchesByCandidates(
            ratingsByCandidate,
            memberCandidates,
            match => {
              const existingDiscordId = (match.rowValues[18] || '').toString().trim();
              return !existingDiscordId || existingDiscordId === member.id;
            }
          );

          if (ratingMatches.length === 1) {
            ratingsRow = ratingMatches[0];
          } else if (ratingMatches.length > 1) {
            skippedAmbiguousRatings++;
            continue;
          }
        }

        let traineeRow = traineeByDiscordId.get(member.id) || null;

        if (!traineeRow) {
          const traineeMatches = getUniqueMatchesByCandidates(
            traineeByCandidate,
            memberCandidates
          );

          if (traineeMatches.length === 1) {
            traineeRow = traineeMatches[0];
          } else if (traineeMatches.length > 1) {
            skippedAmbiguousTrainees++;
            continue;
          }
        }

        if (!ratingsRow && !traineeRow) {
          skippedNoSafeMatch++;
          continue;
        }

        const finalName =
          traineeRow?.rowValues?.[0] ||
          ratingsRow?.rowValues?.[2] ||
          getDisplayNameWithoutRank(member);

        const existingSteamId64 =
          (ratingsRow?.rowValues?.[19] || '').toString().trim() ||
          (traineeRow?.rowValues?.[3] || '').toString().trim() ||
          '';

        const targetRowNumber = ratingsRow ? ratingsRow.rowNumber : nextNewRatingsRow++;

        if (ratingsRow) {
          updatedRows++;
        } else {
          createdRows++;
        }

        const rowUpdate = new Array(20).fill('');
        rowUpdate[0] = rankRole.name;
        rowUpdate[1] = squadronName;
        rowUpdate[2] = finalName;

        for (const config of Object.values(mosConfig)) {
          const column = resolveSheetColumn(config.sheetColumn);
          if (!column) continue;

          const index = column.charCodeAt(0) - 65;
          rowUpdate[index] = getMosSheetValue(member, interaction.guild, config);
        }

        rowUpdate[18] = member.id;          // S = Discord ID
        rowUpdate[19] = existingSteamId64;  // T = SteamID64

        ratingsUpdates.push({
          range: `${RATINGS_SHEET_NAME}!A${targetRowNumber}:T${targetRowNumber}`,
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
          `**Skipped Ambiguous Ratings:** ${skippedAmbiguousRatings}`,
          `**Skipped Ambiguous Trainees:** ${skippedAmbiguousTrainees}`,
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
          `Skipped ambiguous Ratings: **${skippedAmbiguousRatings}**\n` +
          `Skipped ambiguous Trainees: **${skippedAmbiguousTrainees}**`,
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