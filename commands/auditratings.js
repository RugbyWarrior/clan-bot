const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  getRatingsRows,
  normalizeName,
} = require('../sheets');

const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;

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

function truncateList(items, limit = 10) {
  if (items.length <= limit) return items;
  return [...items.slice(0, limit), `...and ${items.length - limit} more`];
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('auditratings')
    .setDescription('Audit invalid or suspicious rows in the Ratings sheet.'),

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

      const allMembers = interaction.guild.members.cache.filter(member => !member.user.bot);
      const memberById = new Map();
      const memberByCandidate = new Map();

      for (const member of allMembers.values()) {
        memberById.set(member.id, member);

        for (const candidate of buildNameCandidatesFromRaw(member.nickname || member.user.username)) {
          if (!candidate) continue;
          if (!memberByCandidate.has(candidate)) memberByCandidate.set(candidate, []);
          memberByCandidate.get(candidate).push(member);
        }
      }

      const ratingsRows = await getRatingsRows();

      const traineeRows = [];
      const exSkiraRows = [];
      const missingMemberRows = [];
      const noDiscordIdRows = [];
      const noDiscordIdNoSafeMatchRows = [];
      const duplicateDiscordIdRows = [];

      const discordIdCounts = new Map();

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        if (rowDiscordId) {
          if (!discordIdCounts.has(rowDiscordId)) {
            discordIdCounts.set(rowDiscordId, []);
          }
          discordIdCounts.get(rowDiscordId).push({ rowNumber, rowName });
        }

        if (!rowDiscordId) {
          noDiscordIdRows.push(`Row ${rowNumber} | ${rowName}`);

          const candidates = buildNameCandidatesFromRaw(rowName);
          const matches = new Map();

          for (const candidate of candidates) {
            for (const member of memberByCandidate.get(candidate) || []) {
              matches.set(member.id, member);
            }
          }

          if (matches.size === 0) {
            noDiscordIdNoSafeMatchRows.push(`Row ${rowNumber} | ${rowName} | No safe Discord match`);
          } else if (matches.size > 1) {
            noDiscordIdNoSafeMatchRows.push(`Row ${rowNumber} | ${rowName} | Ambiguous Discord match (${matches.size})`);
          }

          continue;
        }

        const member = memberById.get(rowDiscordId);

        if (!member) {
          missingMemberRows.push(`Row ${rowNumber} | ${rowName} | Discord ID ${rowDiscordId}`);
          continue;
        }

        if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) {
          traineeRows.push(`Row ${rowNumber} | ${rowName} | ${member.user.tag} (${member.id})`);
        }

        if (EX_SKIRA_ROLE_ID && member.roles.cache.has(EX_SKIRA_ROLE_ID)) {
          exSkiraRows.push(`Row ${rowNumber} | ${rowName} | ${member.user.tag} (${member.id})`);
        }
      }

      for (const [discordId, entries] of discordIdCounts.entries()) {
        if (entries.length > 1) {
          duplicateDiscordIdRows.push(
            `Discord ID ${discordId} appears in rows: ${entries.map(e => `${e.rowNumber} (${e.rowName})`).join(', ')}`
          );
        }
      }

      const summary = [
        `Ratings rows checked: ${Math.max(ratingsRows.length - 1, 0)}`,
        `Rows linked to trainees: ${traineeRows.length}`,
        `Rows linked to Ex Skira: ${exSkiraRows.length}`,
        `Rows with Discord ID not found in cache: ${missingMemberRows.length}`,
        `Rows with blank Discord ID: ${noDiscordIdRows.length}`,
        `Blank-ID rows with no safe match: ${noDiscordIdNoSafeMatchRows.length}`,
        `Duplicate Discord IDs in Ratings: ${duplicateDiscordIdRows.length}`,
      ];

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🧾 **AUDIT RATINGS**',
          `Moderator: <@${interaction.user.id}>`,
          ...summary,
          traineeRows.length
            ? `Trainee Rows:\n- ${truncateList(traineeRows, 20).join('\n- ')}`
            : 'Trainee Rows:\n- None',
          exSkiraRows.length
            ? `Ex Skira Rows:\n- ${truncateList(exSkiraRows, 20).join('\n- ')}`
            : 'Ex Skira Rows:\n- None',
          missingMemberRows.length
            ? `Missing Member Rows:\n- ${truncateList(missingMemberRows, 20).join('\n- ')}`
            : 'Missing Member Rows:\n- None',
          noDiscordIdNoSafeMatchRows.length
            ? `Blank ID / Unsafe Rows:\n- ${truncateList(noDiscordIdNoSafeMatchRows, 20).join('\n- ')}`
            : 'Blank ID / Unsafe Rows:\n- None',
          duplicateDiscordIdRows.length
            ? `Duplicate Discord ID Rows:\n- ${truncateList(duplicateDiscordIdRows, 20).join('\n- ')}`
            : 'Duplicate Discord ID Rows:\n- None',
        ].join('\n')
      );

      return interaction.editReply({
        content: [
          '**Ratings audit complete.**',
          ...summary,
          '',
          traineeRows.length
            ? `Trainee row examples:\n- ${truncateList(traineeRows, 8).join('\n- ')}`
            : 'Trainee row examples:\n- None',
          exSkiraRows.length
            ? `Ex Skira row examples:\n- ${truncateList(exSkiraRows, 8).join('\n- ')}`
            : 'Ex Skira row examples:\n- None',
          missingMemberRows.length
            ? `Missing member examples:\n- ${truncateList(missingMemberRows, 8).join('\n- ')}`
            : 'Missing member examples:\n- None',
          noDiscordIdNoSafeMatchRows.length
            ? `Blank ID / unsafe examples:\n- ${truncateList(noDiscordIdNoSafeMatchRows, 8).join('\n- ')}`
            : 'Blank ID / unsafe examples:\n- None',
          duplicateDiscordIdRows.length
            ? `Duplicate Discord ID examples:\n- ${truncateList(duplicateDiscordIdRows, 8).join('\n- ')}`
            : 'Duplicate Discord ID examples:\n- None',
        ].join('\n'),
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /auditratings:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /auditratings.',
        allowedMentions: { users: [] },
      });
    }
  },
};