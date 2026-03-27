const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  findTraineeRowByDiscordId,
  findTraineeRowBySteamId64,
  deleteTraineeRow,
} = require('../sheets');

const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const TRAINING_COMPANY_ROLE_ID = process.env.TRAINING_COMPANY_ROLE_ID;
const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;

function isValidDiscordId(value) {
  return /^\d{17,20}$/.test(String(value || '').trim());
}

function isValidSteamId64(value) {
  return /^\d{17}$/.test(String(value || '').trim());
}

function removeTraineePrefix(displayName) {
  if (!displayName) return displayName;
  return displayName.replace(/^TR[\s._-]*/i, '').trim();
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('removetrainee')
    .setDescription('Remove a trainee from the sheet and/or remove trainee roles if they are still in the server.')
    .addUserOption(option =>
      option
        .setName('trainee')
        .setDescription('The trainee to remove')
        .setRequired(false)
    )
    .addStringOption(option =>
      option
        .setName('discord_id')
        .setDescription('Discord ID of the trainee')
        .setRequired(false)
    )
    .addStringOption(option =>
      option
        .setName('steamid64')
        .setDescription('SteamID64 of the trainee')
        .setRequired(false)
    ),

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

      const traineeUser = interaction.options.getUser('trainee');
      const discordIdInputRaw = interaction.options.getString('discord_id');
      const steamId64InputRaw = interaction.options.getString('steamid64');

      const discordIdInput = discordIdInputRaw ? discordIdInputRaw.trim() : null;
      const steamId64Input = steamId64InputRaw ? steamId64InputRaw.trim() : null;

      if (!traineeUser && !discordIdInput && !steamId64Input) {
        return interaction.editReply({
          content: 'You must provide at least one of: trainee, discord_id, or steamid64.',
          allowedMentions: { users: [] },
        });
      }

      if (discordIdInput && !isValidDiscordId(discordIdInput)) {
        return interaction.editReply({
          content: 'That does not look like a valid Discord ID.',
          allowedMentions: { users: [] },
        });
      }

      if (steamId64Input && !isValidSteamId64(steamId64Input)) {
        return interaction.editReply({
          content: 'That does not look like a valid SteamID64.',
          allowedMentions: { users: [] },
        });
      }

      let foundRecord = null;
      let matchedBy = null;

      if (traineeUser) {
        foundRecord = await findTraineeRowByDiscordId(traineeUser.id);
        matchedBy = foundRecord ? `Discord user (${traineeUser.id})` : null;
      }

      if (!foundRecord && discordIdInput) {
        foundRecord = await findTraineeRowByDiscordId(discordIdInput);
        matchedBy = foundRecord ? `Discord ID (${discordIdInput})` : null;
      }

      if (!foundRecord && steamId64Input) {
        foundRecord = await findTraineeRowBySteamId64(steamId64Input);
        matchedBy = foundRecord ? `SteamID64 (${steamId64Input})` : null;
      }

      const resolvedDiscordId =
        traineeUser?.id ||
        (foundRecord?.rowValues?.[8] || '').toString().trim() ||
        discordIdInput ||
        null;

      let traineeName = foundRecord?.rowValues?.[0] || traineeUser?.username || 'Unknown';
      let enlistedDate = foundRecord?.rowValues?.[1] || 'Unknown';
      let traineeSteamId64 = (foundRecord?.rowValues?.[3] || steamId64Input || '').toString().trim() || 'Unknown';
      let traineeDiscordId = (foundRecord?.rowValues?.[8] || resolvedDiscordId || '').toString().trim() || 'Unknown';

      let sheetRowDeleted = null;

      if (foundRecord) {
        await deleteTraineeRow(foundRecord.rowNumber);
        sheetRowDeleted = foundRecord.rowNumber;
      }

      let member = null;
      let rolesAttempted = [];
      let rolesStillPresent = [];
      let roleRemovalStatus = 'Not attempted';
      let nicknameStatus = 'No nickname change attempted';

      if (isValidDiscordId(resolvedDiscordId)) {
        try {
          member = await interaction.guild.members.fetch(resolvedDiscordId);
        } catch {
          member = null;
        }
      }

      if (member) {
        const rolesToRemove = [
          TRAINEE_ROLE_ID,
          TRAINING_COMPANY_ROLE_ID,
        ].filter(Boolean);

        rolesAttempted = rolesToRemove.filter(roleId => member.roles.cache.has(roleId));

        if (rolesAttempted.length > 0) {
          try {
            await member.roles.remove(
              rolesAttempted,
              `Removed via /removetrainee by ${interaction.user.tag}`
            );

            try {
              const currentName = member.nickname || member.user.username;
              const newName = removeTraineePrefix(currentName);

              if (newName !== currentName) {
                await member.setNickname(
                  newName,
                  'Removed trainee prefix via /removetrainee'
                );
                nicknameStatus = `Nickname changed to "${newName}"`;
              } else {
                nicknameStatus = 'No trainee prefix found in nickname';
              }
            } catch (nickError) {
              console.error('Failed to update nickname:', nickError);
              nicknameStatus = `Nickname update failed: ${nickError.message || 'Unknown error'}`;
            }

            const refreshedMember = await interaction.guild.members.fetch(member.id);

            rolesStillPresent = rolesAttempted.filter(roleId =>
              refreshedMember.roles.cache.has(roleId)
            );

            if (rolesStillPresent.length === 0) {
              roleRemovalStatus = 'Roles removed';
            } else {
              roleRemovalStatus = `Failed to remove some roles: ${rolesStillPresent.join(', ')}`;
            }
          } catch (error) {
            console.error('Error removing trainee roles:', error);
            roleRemovalStatus = `Role removal failed: ${error.message || 'Unknown error'}`;
          }
        } else {
          roleRemovalStatus = 'Member found but had no trainee roles to remove';

          try {
            const currentName = member.nickname || member.user.username;
            const newName = removeTraineePrefix(currentName);

            if (newName !== currentName) {
              await member.setNickname(
                newName,
                'Removed trainee prefix via /removetrainee'
              );
              nicknameStatus = `Nickname changed to "${newName}"`;
            } else {
              nicknameStatus = 'No trainee prefix found in nickname';
            }
          } catch (nickError) {
            console.error('Failed to update nickname:', nickError);
            nicknameStatus = `Nickname update failed: ${nickError.message || 'Unknown error'}`;
          }
        }
      } else if (resolvedDiscordId) {
        roleRemovalStatus = 'Member not in server or could not be fetched';
        nicknameStatus = 'No Discord member available to rename';
      } else {
        roleRemovalStatus = 'No Discord member available to remove roles from';
        nicknameStatus = 'No Discord member available to rename';
      }

      if (!foundRecord && !member) {
        return interaction.editReply({
          content: 'No trainee record was found, and no matching Discord member could be acted on.',
          allowedMentions: { users: [] },
        });
      }

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🗑️ **REMOVE TRAINEE**',
          `Moderator: <@${interaction.user.id}>`,
          `Matched By: ${matchedBy || 'No sheet match'}`,
          `Name: ${traineeName}`,
          `Enlisted: ${enlistedDate}`,
          `SteamID64: ${traineeSteamId64}`,
          `Discord ID: ${traineeDiscordId}`,
          `Sheet Row Deleted: ${sheetRowDeleted ?? 'None'}`,
          `Still In Server: ${member ? 'Yes' : 'No'}`,
          `Roles Attempted: ${rolesAttempted.length > 0 ? rolesAttempted.join(', ') : 'None'}`,
          `Roles Still Present: ${rolesStillPresent.length > 0 ? rolesStillPresent.join(', ') : 'None'}`,
          `Role Removal Status: ${roleRemovalStatus}`,
          `Nickname Status: ${nicknameStatus}`,
        ].join('\n')
      );

      return interaction.editReply({
        content:
          `Processed trainee removal for **${traineeName}**.\n` +
          `Matched by: **${matchedBy || 'No sheet match'}**\n` +
          `Discord ID: \`${traineeDiscordId}\`\n` +
          `SteamID64: \`${traineeSteamId64}\`\n` +
          `Sheet row deleted: **${sheetRowDeleted ?? 'No'}**\n` +
          `Still in server: **${member ? 'Yes' : 'No'}**\n` +
          `Role removal: **${roleRemovalStatus}**\n` +
          `Nickname: **${nicknameStatus}**`,
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /removetrainee:', error);

      return interaction.editReply({
        content: 'There was an error while removing the trainee.',
        allowedMentions: { users: [] },
      });
    }
  },
};