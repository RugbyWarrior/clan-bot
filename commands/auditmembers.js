const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  getTraineeRows,
  getRatingsRows,
  normalizeName,
} = require('../sheets');

const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;

function getDisplayNameWithoutRank(member) {
  const displayName = member.nickname || member.user.username;
  if (!displayName.includes(' ')) return displayName;
  return displayName.replace(/^\S+\s+/, '').trim();
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

function buildMemberNameCandidates(member) {
  return buildNameCandidatesFromRaw(member.nickname || member.user.username);
}

function truncateList(items, limit = 10) {
  if (items.length <= limit) return items;
  return [...items.slice(0, limit), `...and ${items.length - limit} more`];
}

function isMeaningfulTraineeRow(row) {
  if (!row || !Array.isArray(row)) return false;

  const name = (row[0] || '').toString().trim();
  const enlisted = (row[1] || '').toString().trim();
  const wl = (row[2] || '').toString().trim();
  const steamId64 = (row[3] || '').toString().trim();
  const bm = (row[4] || '').toString().trim();
  const lastSeen = (row[5] || '').toString().trim();
  const score = (row[6] || '').toString().trim();
  const notes = (row[7] || '').toString().trim();
  const discordId = (row[8] || '').toString().trim();

  const hasAnyUsefulData =
    name || enlisted || wl || steamId64 || bm || lastSeen || score || notes || discordId;

  if (!hasAnyUsefulData) return false;

  const normalizedName = name.toLowerCase();
  const normalizedEnlisted = enlisted.toLowerCase();
  const normalizedSteam = steamId64.toLowerCase();
  const normalizedNotes = notes.toLowerCase();

  const looksLikeTemplate =
    normalizedName === 'name' &&
    normalizedEnlisted === 'dd/mm/yyyy' &&
    normalizedSteam === 'steamid64' &&
    normalizedNotes === 'notes';

  if (looksLikeTemplate) return false;

  if (!name && !steamId64 && !discordId) return false;

  return true;
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('auditmembers')
    .setDescription('Audit Ratings and Trainees matching issues without changing any data.'),

  async execute(interaction) {
    try {
      await interaction.reply({
        content: 'Running auditmembers...',
        ephemeral: true,
        allowedMentions: { users: [] },
      });

      if (ALLOWED_CHANNEL_ID && interaction.channelId !== ALLOWED_CHANNEL_ID) {
        return interaction.editReply({
          content: 'This command can only be used in the allowed channel.',
          allowedMentions: { users: [] },
        });
      }

      const allMembers = interaction.guild.members.cache.filter(member => !member.user.bot);

      const traineeRows = await getTraineeRows();
      const ratingsRows = await getRatingsRows();

      const traineeByName = new Map();
      const ratingsByDiscordId = new Map();
      const ratingsByName = new Map();
      const memberByNameCandidate = new Map();

      let meaningfulTraineeRowsChecked = 0;

      for (let i = 1; i < traineeRows.length; i++) {
        const row = traineeRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[0] || '';

        if (!isMeaningfulTraineeRow(row)) {
          continue;
        }

        meaningfulTraineeRowsChecked++;

        for (const candidate of buildNameCandidatesFromRaw(rowName)) {
          if (!candidate) continue;
          if (!traineeByName.has(candidate)) traineeByName.set(candidate, []);
          traineeByName.get(candidate).push({
            rowNumber,
            rowValues: row,
          });
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
          if (!ratingsByName.has(candidate)) ratingsByName.set(candidate, []);
          ratingsByName.get(candidate).push(record);
        }
      }

      for (const member of allMembers.values()) {
        const candidates = buildMemberNameCandidates(member);
        for (const candidate of candidates) {
          if (!candidate) continue;
          if (!memberByNameCandidate.has(candidate)) memberByNameCandidate.set(candidate, []);
          memberByNameCandidate.get(candidate).push(member);
        }
      }

      const ratingsMissingId = [];
      const traineeUnmatched = [];
      const rankedMembersMissingRatings = [];

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        if (rowDiscordId) continue;

        const candidates = buildNameCandidatesFromRaw(rowName);
        const matchedMembers = new Map();

        for (const candidate of candidates) {
          const possibleMembers = memberByNameCandidate.get(candidate) || [];
          for (const member of possibleMembers) {
            matchedMembers.set(member.id, member);
          }
        }

        const matches = [...matchedMembers.values()];

        let reason = 'No match in cached Discord members';
        if (matches.length === 1) {
          const member = matches[0];
          reason = `Safe match available: ${member.user.tag} (${member.id})`;
        } else if (matches.length > 1) {
          reason = `Ambiguous: ${matches.length} Discord matches`;
        }

        ratingsMissingId.push(`Row ${rowNumber} | ${rowName} | ${reason}`);
      }

      for (let i = 1; i < traineeRows.length; i++) {
        const row = traineeRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[0] || '';
        const rowDiscordId = (row[8] || '').toString().trim();

        if (!isMeaningfulTraineeRow(row)) {
          continue;
        }

        if (rowDiscordId) continue;

        const candidates = buildNameCandidatesFromRaw(rowName);
        const matchedMembers = new Map();

        for (const candidate of candidates) {
          const possibleMembers = memberByNameCandidate.get(candidate) || [];
          for (const member of possibleMembers) {
            if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) {
              matchedMembers.set(member.id, member);
            }
          }
        }

        const matches = [...matchedMembers.values()];

        if (matches.length === 0) {
          traineeUnmatched.push(`Row ${rowNumber} | ${rowName} | No trainee Discord match`);
        } else if (matches.length > 1) {
          traineeUnmatched.push(`Row ${rowNumber} | ${rowName} | Ambiguous trainee match (${matches.length})`);
        }
      }

      for (const member of allMembers.values()) {
        if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) continue;
        if (EX_SKIRA_ROLE_ID && member.roles.cache.has(EX_SKIRA_ROLE_ID)) continue;

        const rankRole = getAllowedRankRole(member, interaction.guild);
        if (!rankRole) continue;

        const existingByDiscord = ratingsByDiscordId.get(member.id);
        if (existingByDiscord) continue;

        const candidates = buildMemberNameCandidates(member);
        const ratingMatches = new Map();
        const traineeMatches = new Map();

        for (const candidate of candidates) {
          for (const match of ratingsByName.get(candidate) || []) {
            ratingMatches.set(match.rowNumber, match);
          }
          for (const match of traineeByName.get(candidate) || []) {
            traineeMatches.set(match.rowNumber, match);
          }
        }

        const ratingMatchesArr = [...ratingMatches.values()];
        const traineeMatchesArr = [...traineeMatches.values()];

        if (ratingMatchesArr.length === 1) {
          rankedMembersMissingRatings.push(
            `${member.user.tag} | ${rankRole.name} | Existing Ratings row found by name only: row ${ratingMatchesArr[0].rowNumber}`
          );
          continue;
        }

        if (ratingMatchesArr.length > 1) {
          rankedMembersMissingRatings.push(
            `${member.user.tag} | ${rankRole.name} | Ambiguous Ratings name match (${ratingMatchesArr.length})`
          );
          continue;
        }

        if (traineeMatchesArr.length === 1) {
          rankedMembersMissingRatings.push(
            `${member.user.tag} | ${rankRole.name} | Safe trainee match exists: row ${traineeMatchesArr[0].rowNumber}`
          );
          continue;
        }

        if (traineeMatchesArr.length > 1) {
          rankedMembersMissingRatings.push(
            `${member.user.tag} | ${rankRole.name} | Ambiguous trainee name match (${traineeMatchesArr.length})`
          );
          continue;
        }

        rankedMembersMissingRatings.push(
          `${member.user.tag} | ${rankRole.name} | No safe Ratings or Trainee match`
        );
      }

      const summary = [
        `Ratings rows missing Discord ID: ${ratingsMissingId.length}`,
        `Meaningful trainee rows checked: ${meaningfulTraineeRowsChecked}`,
        `Trainee rows still unmatched/missing ID: ${traineeUnmatched.length}`,
        `Ranked Discord members missing direct Ratings link: ${rankedMembersMissingRatings.length}`,
        `Cached Discord members checked: ${allMembers.size}`,
      ];

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🧾 **AUDIT MEMBERS**',
          `Moderator: <@${interaction.user.id}>`,
          ...summary,
          ratingsMissingId.length
            ? `Ratings Missing ID:\n- ${truncateList(ratingsMissingId, 15).join('\n- ')}`
            : 'Ratings Missing ID:\n- None',
          traineeUnmatched.length
            ? `Trainee Unmatched:\n- ${truncateList(traineeUnmatched, 15).join('\n- ')}`
            : 'Trainee Unmatched:\n- None',
          rankedMembersMissingRatings.length
            ? `Ranked Missing Ratings:\n- ${truncateList(rankedMembersMissingRatings, 15).join('\n- ')}`
            : 'Ranked Missing Ratings:\n- None',
        ].join('\n')
      );

      return interaction.editReply({
        content: [
          '**Audit complete.**',
          ...summary,
          '',
          ratingsMissingId.length
            ? `Ratings missing ID examples:\n- ${truncateList(ratingsMissingId, 8).join('\n- ')}`
            : 'Ratings missing ID examples:\n- None',
          traineeUnmatched.length
            ? `Trainee unmatched examples:\n- ${truncateList(traineeUnmatched, 8).join('\n- ')}`
            : 'Trainee unmatched examples:\n- None',
          rankedMembersMissingRatings.length
            ? `Ranked missing Ratings examples:\n- ${truncateList(rankedMembersMissingRatings, 8).join('\n- ')}`
            : 'Ranked missing Ratings examples:\n- None',
        ].join('\n'),
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /auditmembers:', error);

      try {
        return interaction.editReply({
          content: '❌ There was an error while running /auditmembers.',
          allowedMentions: { users: [] },
        });
      } catch {
        return;
      }
    }
  },
};