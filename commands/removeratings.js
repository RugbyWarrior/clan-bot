const {
  SlashCommandBuilder,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
} = require('discord.js');
const { google } = require('googleapis');
const { sendLog } = require('../logger');
const { getRatingsRows, normalizeName } = require('../sheets');
const { mosConfig } = require('../mosConfig');
const { updateNickname } = require('../utils/updateNickname');

const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const RATINGS_SHEET_NAME = (process.env.RATINGS_SHEET_NAME || 'Ratings').trim();
const RATINGS_SPREADSHEET_ID = (
  process.env.RATINGS_SPREADSHEET_ID ||
  process.env.MOS_SPREADSHEET_ID ||
  process.env.TRAINEE_SPREADSHEET_ID ||
  ''
).trim();

const COMMUNITY_ROLE_ID = process.env.COMMUNITY_MEMBER_ID;

function isValidDiscordId(value) {
  return /^\d{17,20}$/.test(String(value || '').trim());
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

async function deleteRatingsRow(rowNumber) {
  const sheets = await getSheetsClient();
  const sheetId = await getRatingsSheetId();

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: RATINGS_SPREADSHEET_ID,
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

async function findRatingsRowByDiscordId(discordId) {
  const rows = await getRatingsRows();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowDiscordId = (row[18] || '').toString().trim();

    if (rowDiscordId === discordId) {
      return {
        rowNumber: i + 1,
        rowValues: row,
      };
    }
  }

  return null;
}

async function findRatingsRowsByName(name) {
  const rows = await getRatingsRows();
  const target = normalizeName(name);
  const matches = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowName = normalizeName(row[2] || '');

    if (rowName && rowName === target) {
      matches.push({
        rowNumber: i + 1,
        rowValues: row,
      });
    }
  }

  return matches;
}

function getAllMosRoleIds() {
  const mosRoleIds = new Set();

  for (const mos of Object.values(mosConfig)) {
    for (const roleId of Object.values(mos.ratings)) {
      if (roleId) {
        mosRoleIds.add(roleId);
      }
    }
  }

  return [...mosRoleIds];
}

function getConfiguredRoleIds(envKey) {
  return (process.env[envKey] || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('removeratings')
    .setDescription('Remove a member from the Ratings system and reset their roles.')
    .addUserOption(option =>
      option
        .setName('user')
        .setDescription('The user to remove from the Ratings sheet')
        .setRequired(false)
    )
    .addStringOption(option =>
      option
        .setName('discord_id')
        .setDescription('Discord ID to remove from the Ratings sheet')
        .setRequired(false)
    ),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (HQ_CHANNEL_ID && interaction.channelId !== HQ_CHANNEL_ID) {
        return interaction.editReply({
          content: 'This command can only be used in Headquarters.',
          allowedMentions: { users: [] },
        });
      }

      if (!COMMUNITY_ROLE_ID) {
        return interaction.editReply({
          content: 'COMMUNITY_MEMBER_ID is missing in the environment variables.',
          allowedMentions: { users: [] },
        });
      }

      const targetUser = interaction.options.getUser('user');
      const discordIdInputRaw = interaction.options.getString('discord_id');
      const discordIdInput = discordIdInputRaw ? discordIdInputRaw.trim() : null;

      if (!targetUser && !discordIdInput) {
        return interaction.editReply({
          content: 'You must provide either a user or a discord_id.',
          allowedMentions: { users: [] },
        });
      }

      if (discordIdInput && !isValidDiscordId(discordIdInput)) {
        return interaction.editReply({
          content: 'That does not look like a valid Discord ID.',
          allowedMentions: { users: [] },
        });
      }

      let foundRow = null;
      let matchedBy = null;

      if (targetUser) {
        foundRow = await findRatingsRowByDiscordId(targetUser.id);
        if (foundRow) {
          matchedBy = `Discord user (${targetUser.id})`;
        }
      }

      if (!foundRow && discordIdInput) {
        foundRow = await findRatingsRowByDiscordId(discordIdInput);
        if (foundRow) {
          matchedBy = `Discord ID (${discordIdInput})`;
        }
      }

      if (!foundRow && targetUser) {
        let member = null;

        try {
          member = await interaction.guild.members.fetch(targetUser.id);
        } catch {
          member = null;
        }

        const fallbackName = member
          ? (member.nickname || member.user.username)
          : targetUser.username;

        const nameMatches = await findRatingsRowsByName(fallbackName);

        if (nameMatches.length === 1) {
          foundRow = nameMatches[0];
          matchedBy = `Unique name match (${fallbackName})`;
        } else if (nameMatches.length > 1) {
          return interaction.editReply({
            content: 'Multiple Ratings rows matched that user name. Use discord_id instead.',
            allowedMentions: { users: [] },
          });
        }
      }

      if (!foundRow) {
        return interaction.editReply({
          content: 'No Ratings row was found for that user/Discord ID.',
          allowedMentions: { users: [] },
        });
      }

      const row = foundRow.rowValues || [];
      const ratingsRank = row[0] || 'Unknown';
      const ratingsSquadron = row[1] || 'Unknown';
      const ratingsName = row[2] || 'Unknown';
      const ratingsDiscordId = (row[18] || targetUser?.id || discordIdInput || '').toString().trim() || 'Unknown';

      const rankRoleIds = getConfiguredRoleIds('RANK_ROLE_IDS');
      const breakerRoleIds = getConfiguredRoleIds('BREAKER_ROLE_IDS');
      const squadronRoleIds = getConfiguredRoleIds('SQUADRON_ROLE_IDS');
      const mosRoleIds = getAllMosRoleIds();

      const allRemovableRoleIds = [
        ...new Set([
          ...rankRoleIds,
          ...breakerRoleIds,
          ...squadronRoleIds,
          ...mosRoleIds,
        ]),
      ].filter(roleId => roleId !== COMMUNITY_ROLE_ID);

      let member = null;
      let rolesToRemove = [];
      let stillInServer = false;

      try {
        member = await interaction.guild.members.fetch(ratingsDiscordId);
        stillInServer = true;
      } catch {
        member = null;
      }

      if (member) {
        rolesToRemove = allRemovableRoleIds.filter(roleId => member.roles.cache.has(roleId));
      }

      const preview =
        `Are you sure you want to remove **${ratingsName}** from the Ratings system?\n\n` +
        `Matched by: **${matchedBy || 'Unknown'}**\n` +
        `Row: **${foundRow.rowNumber}**\n` +
        `Rank: **${ratingsRank}**\n` +
        `Squadron: **${ratingsSquadron}**\n` +
        `Discord ID: \`${ratingsDiscordId}\`\n` +
        `Still in server: **${stillInServer ? 'Yes' : 'No'}**\n` +
        `Roles to remove: **${rolesToRemove.length}**\n` +
        `Community role to add: **Yes**\n` +
        `Nickname reset: **${stillInServer ? 'Yes (attempted)' : 'No'}**`;

      const rowButtons = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`removeratings_confirm_${interaction.id}`)
          .setLabel('Confirm')
          .setStyle(ButtonStyle.Danger),
        new ButtonBuilder()
          .setCustomId(`removeratings_cancel_${interaction.id}`)
          .setLabel('Cancel')
          .setStyle(ButtonStyle.Secondary)
      );

      await interaction.editReply({
        content: preview,
        components: [rowButtons],
        allowedMentions: { users: [] },
      });

      const message = await interaction.fetchReply();

      const buttonInteraction = await message.awaitMessageComponent({
        filter: i =>
          i.user.id === interaction.user.id &&
          (i.customId === `removeratings_confirm_${interaction.id}` ||
            i.customId === `removeratings_cancel_${interaction.id}`),
        time: 60000,
      }).catch(() => null);

      if (!buttonInteraction) {
        await interaction.editReply({
          content: `${preview}\n\n⏳ Timed out. No changes were made.`,
          components: [],
          allowedMentions: { users: [] },
        });
        return;
      }

      if (buttonInteraction.customId === `removeratings_cancel_${interaction.id}`) {
        await buttonInteraction.update({
          content: `${preview}\n\n✅ Cancelled. No changes were made.`,
          components: [],
          allowedMentions: { users: [] },
        });
        return;
      }

      let rolesRemoved = [];
      let communityAdded = false;
      let nicknameReset = false;

      if (member) {
        rolesRemoved = rolesToRemove;

        if (rolesRemoved.length > 0) {
          await member.roles.remove(
            rolesRemoved,
            `Removed via /removeratings by ${interaction.user.tag}`
          );
        }

        if (!member.roles.cache.has(COMMUNITY_ROLE_ID)) {
          await member.roles.add(
            COMMUNITY_ROLE_ID,
            `Community role added via /removeratings by ${interaction.user.tag}`
          );
          communityAdded = true;
        }

        try {
          await updateNickname(member, {
            prefix: null,
            exSkira: false,
          });
          nicknameReset = true;
        } catch {
          nicknameReset = false;
        }
      }

      await deleteRatingsRow(foundRow.rowNumber);

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🗑️ **REMOVE RATINGS ROW**',
          `Moderator: <@${interaction.user.id}>`,
          `Matched By: ${matchedBy || 'Unknown'}`,
          `Row Deleted: ${foundRow.rowNumber}`,
          `Rank: ${ratingsRank}`,
          `Squadron: ${ratingsSquadron}`,
          `Name: ${ratingsName}`,
          `Discord ID: ${ratingsDiscordId}`,
          `Still In Server: ${stillInServer ? 'Yes' : 'No'}`,
          `Roles Removed: ${rolesRemoved.length > 0 ? rolesRemoved.join(', ') : 'None'}`,
          `Community Role Added: ${communityAdded ? 'Yes' : 'No'}`,
          `Nickname Reset: ${nicknameReset ? 'Yes' : 'No'}`,
        ].join('\n')
      );

      await buttonInteraction.update({
        content:
          `Removed **${ratingsName}** from the Ratings system.\n` +
          `Matched by: **${matchedBy || 'Unknown'}**\n` +
          `Deleted row: **${foundRow.rowNumber}**\n` +
          `Rank: **${ratingsRank}**\n` +
          `Squadron: **${ratingsSquadron}**\n` +
          `Discord ID: \`${ratingsDiscordId}\`\n` +
          `Still in server: **${stillInServer ? 'Yes' : 'No'}**\n` +
          `Roles removed: **${rolesRemoved.length > 0 ? 'Yes' : 'No'}**\n` +
          `Community role added: **${communityAdded ? 'Yes' : 'No'}**\n` +
          `Nickname reset: **${nicknameReset ? 'Yes' : 'No'}**`,
        components: [],
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /removeratings:', error);

      return interaction.editReply({
        content: '❌ There was an error while running /removeratings.',
        components: [],
        allowedMentions: { users: [] },
      });
    }
  },
};