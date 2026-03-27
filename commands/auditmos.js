const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  getRatingsRows,
} = require('../sheets');
const { mosConfig } = require('../mosConfig');

const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';

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

function columnLetterToIndex(columnLetter) {
  return columnLetter.toUpperCase().charCodeAt(0) - 65;
}

function getMosSheetValueFromMember(member, guild, mos) {
  for (const [ratingName, roleId] of Object.entries(mos.ratings)) {
    if (!roleId) continue;

    if (member.roles.cache.has(roleId)) {
      const roleObject = guild.roles.cache.get(roleId);
      return roleObject ? roleObject.name : ratingName;
    }
  }

  return '';
}

function truncateList(items, limit = 10) {
  if (items.length <= limit) return items;
  return [...items.slice(0, limit), `...and ${items.length - limit} more`];
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('auditmos')
    .setDescription('Audit MOS values in the Ratings sheet against Discord MOS roles.'),

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
      const memberById = new Map();

      for (const member of allMembers.values()) {
        memberById.set(member.id, member);
      }

      const ratingsRows = await getRatingsRows();

      const rowsMissingDiscordId = [];
      const rowsWithMissingMember = [];
      const mismatches = [];
      const cleanMatches = [];

      for (let i = 1; i < ratingsRows.length; i++) {
        const row = ratingsRows[i] || [];
        const rowNumber = i + 1;
        const rowName = row[2] || '';
        const rowDiscordId = (row[18] || '').toString().trim();

        if (!rowDiscordId) {
          rowsMissingDiscordId.push(`Row ${rowNumber} | ${rowName}`);
          continue;
        }

        const member = memberById.get(rowDiscordId);

        if (!member) {
          rowsWithMissingMember.push(`Row ${rowNumber} | ${rowName} | Discord ID ${rowDiscordId}`);
          continue;
        }

        let rowHasMismatch = false;

        for (const mos of Object.values(mosConfig)) {
          const columnLetter = resolveSheetColumn(mos.sheetColumn);
          if (!columnLetter) continue;

          const columnIndex = columnLetterToIndex(columnLetter);
          const sheetValue = (row[columnIndex] || '').toString().trim();
          const discordValue = getMosSheetValueFromMember(member, interaction.guild, mos);

          if (sheetValue !== discordValue) {
            rowHasMismatch = true;
            mismatches.push(
              `Row ${rowNumber} | ${rowName} | ${mos.label} | Sheet: ${sheetValue || '[blank]'} | Discord: ${discordValue || '[blank]'}`
            );
          }
        }

        if (!rowHasMismatch) {
          cleanMatches.push(`Row ${rowNumber} | ${rowName}`);
        }
      }

      const summary = [
        `Ratings rows checked: ${Math.max(ratingsRows.length - 1, 0)}`,
        `Rows missing Discord ID: ${rowsMissingDiscordId.length}`,
        `Rows with Discord ID not found in cache: ${rowsWithMissingMember.length}`,
        `MOS mismatches found: ${mismatches.length}`,
        `Rows with no MOS mismatches: ${cleanMatches.length}`,
        `Cached Discord members checked: ${allMembers.size}`,
      ];

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🧾 **AUDIT MOS**',
          `Moderator: <@${interaction.user.id}>`,
          ...summary,
          rowsMissingDiscordId.length
            ? `Rows Missing Discord ID:\n- ${truncateList(rowsMissingDiscordId, 15).join('\n- ')}`
            : 'Rows Missing Discord ID:\n- None',
          rowsWithMissingMember.length
            ? `Rows With Missing Member:\n- ${truncateList(rowsWithMissingMember, 15).join('\n- ')}`
            : 'Rows With Missing Member:\n- None',
          mismatches.length
            ? `MOS Mismatches:\n- ${truncateList(mismatches, 20).join('\n- ')}`
            : 'MOS Mismatches:\n- None',
        ].join('\n')
      );

      return interaction.editReply({
        content: [
          '**MOS audit complete.**',
          ...summary,
          '',
          rowsMissingDiscordId.length
            ? `Rows missing Discord ID examples:\n- ${truncateList(rowsMissingDiscordId, 8).join('\n- ')}`
            : 'Rows missing Discord ID examples:\n- None',
          rowsWithMissingMember.length
            ? `Rows with missing member examples:\n- ${truncateList(rowsWithMissingMember, 8).join('\n- ')}`
            : 'Rows with missing member examples:\n- None',
          mismatches.length
            ? `MOS mismatch examples:\n- ${truncateList(mismatches, 8).join('\n- ')}`
            : 'MOS mismatch examples:\n- None',
        ].join('\n'),
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /auditmos:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /auditmos.',
        allowedMentions: { users: [] },
      });
    }
  },
};