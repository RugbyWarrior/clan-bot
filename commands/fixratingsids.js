const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  getRatingsRows,
  batchUpdateRatingsCells,
  normalizeName,
} = require('../sheets');

const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';

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

module.exports = {
  data: new SlashCommandBuilder()
    .setName('fixratingsids')
    .setDescription('Safely fill missing Discord IDs in Ratings where there is exactly one clear match.'),

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
      const ratingsRows = await getRatingsRows();

      const memberByCandidate = new Map();

      for (const member of allMembers.values()) {
        const candidates = buildMemberNameCandidates(member);
        for (const candidate of candidates) {
          if (!candidate) continue;
          if (!memberByCandidate.has(candidate)) memberByCandidate.set(candidate, []);
          memberByCandidate.get(candidate).push(member);
        }
      }

      const updates = [];
      const fixed = [];
      const skippedAmbiguous = [];
      const skippedNoMatch = [];
      const skippedAlreadyHasId = [];

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        if (rowDiscordId) {
          skippedAlreadyHasId.push(`Row ${rowNumber} | ${rowName}`);
          continue;
        }

        const candidates = buildNameCandidatesFromRaw(rowName);
        const matchedMembers = new Map();

        for (const candidate of candidates) {
          const possibleMembers = memberByCandidate.get(candidate) || [];
          for (const member of possibleMembers) {
            matchedMembers.set(member.id, member);
          }
        }

        const matches = [...matchedMembers.values()];

        if (matches.length === 1) {
          const member = matches[0];

          updates.push({
            range: `${RATINGS_SHEET_NAME}!S${rowNumber}`,
            values: [[member.id]],
          });

          fixed.push(`Row ${rowNumber} | ${rowName} → ${member.user.tag} (${member.id})`);
        } else if (matches.length > 1) {
          skippedAmbiguous.push(`Row ${rowNumber} | ${rowName} | ${matches.length} possible matches`);
        } else {
          skippedNoMatch.push(`Row ${rowNumber} | ${rowName}`);
        }
      }

      if (updates.length > 0) {
        await batchUpdateRatingsCells(updates);
      }

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🛠️ **FIX RATINGS IDS**',
          `Moderator: <@${interaction.user.id}>`,
          `Fixed: ${fixed.length}`,
          `Skipped Ambiguous: ${skippedAmbiguous.length}`,
          `Skipped No Match: ${skippedNoMatch.length}`,
          fixed.length ? `Fixed Rows:\n- ${truncateList(fixed, 20).join('\n- ')}` : 'Fixed Rows:\n- None',
          skippedAmbiguous.length
            ? `Ambiguous Rows:\n- ${truncateList(skippedAmbiguous, 20).join('\n- ')}`
            : 'Ambiguous Rows:\n- None',
          skippedNoMatch.length
            ? `No Match Rows:\n- ${truncateList(skippedNoMatch, 20).join('\n- ')}`
            : 'No Match Rows:\n- None',
        ].join('\n')
      );

      return interaction.editReply({
        content: [
          '**Ratings ID fix complete.**',
          `Fixed: **${fixed.length}**`,
          `Skipped ambiguous: **${skippedAmbiguous.length}**`,
          `Skipped no match: **${skippedNoMatch.length}**`,
          '',
          fixed.length
            ? `Fixed examples:\n- ${truncateList(fixed, 8).join('\n- ')}`
            : 'Fixed examples:\n- None',
          skippedAmbiguous.length
            ? `Ambiguous examples:\n- ${truncateList(skippedAmbiguous, 8).join('\n- ')}`
            : 'Ambiguous examples:\n- None',
          skippedNoMatch.length
            ? `No match examples:\n- ${truncateList(skippedNoMatch, 8).join('\n- ')}`
            : 'No match examples:\n- None',
        ].join('\n'),
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /fixratingsids:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /fixratingsids.',
        allowedMentions: { users: [] },
      });
    }
  },
};