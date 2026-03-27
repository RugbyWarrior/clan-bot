const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const { updateNickname } = require('../utils/updateNickname');
const { writeTraineeRow, findTraineeRowByDiscordId } = require('../sheets');

function formatDateUK(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function isValidSteamId64(value) {
  return /^7656119\d{10}$/.test(value);
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('approve')
    .setDescription('Approve a trainee, assign roles, and add them to the trainee sheet')
    .addUserOption(option =>
      option.setName('trainee')
        .setDescription('The trainee you are approving')
        .setRequired(true)
    )
    .addStringOption(option =>
      option.setName('in_game_name')
        .setDescription('Their in-game name')
        .setRequired(true)
    )
    .addStringOption(option =>
      option.setName('steamid64')
        .setDescription('Their SteamID64 (17-digit number)')
        .setRequired(true)
    ),

  async execute(interaction) {
    if (interaction.channelId !== process.env.ALLOWED_CHANNEL_ID) {
      await interaction.reply({
        content: '❌ You can only use this command in the approvals channel.',
        ephemeral: true,
      });
      return;
    }

    await interaction.deferReply({ ephemeral: true });

    const user = interaction.options.getUser('trainee');
    const inGameName = interaction.options.getString('in_game_name').trim();
    const steamId64 = interaction.options.getString('steamid64').trim();
    const member = await interaction.guild.members.fetch(user.id);

    if (!isValidSteamId64(steamId64)) {
      await interaction.editReply({
        content: '❌ That SteamID64 does not look valid. It should be a 17-digit number starting with 7656119.',
      });
      return;
    }

    if (!process.env.TRAINEE_ROLE_ID) {
      await interaction.editReply({ content: '❌ TRAINEE_ROLE_ID missing in .env' });
      return;
    }

    if (!process.env.TRAINING_COMPANY_ROLE_ID) {
      await interaction.editReply({ content: '❌ TRAINING_COMPANY_ROLE_ID missing in .env' });
      return;
    }

    if (!process.env.TRAINEE_SPREADSHEET_ID) {
      await interaction.editReply({ content: '❌ TRAINEE_SPREADSHEET_ID missing in .env' });
      return;
    }

    if (!process.env.GOOGLE_CREDENTIALS_JSON) {
      await interaction.editReply({ content: '❌ GOOGLE_CREDENTIALS_JSON missing in .env' });
      return;
    }

    const enlistedDate = formatDateUK(new Date());

    try {
      const existingTrainee = await findTraineeRowByDiscordId(user.id);

      if (existingTrainee) {
        await interaction.editReply({
          content: `❌ <@${user.id}> is already in the **Trainees** sheet on row **${existingTrainee.rowNumber}**.`,
          allowedMentions: { users: [] },
        });
        return;
      }

      await member.roles.add(process.env.TRAINEE_ROLE_ID, 'Approved as trainee');
      await member.roles.add(process.env.TRAINING_COMPANY_ROLE_ID, 'Added to Training Company');

      await updateNickname(member, {
        exactName: `TR ${inGameName}`,
      });

      const rowNumber = await writeTraineeRow([
        inGameName,
        enlistedDate,
        '',
        steamId64,
        '',
        '',
        '',
        '',
        user.id,
      ]);

      await interaction.editReply({
        content: `✅ Approved <@${user.id}> as **${inGameName}** and added the **Trainee** and **Training Company** roles.\n📄 Filled row **${rowNumber}** in the **Trainees** sheet.`,
        allowedMentions: { users: [] },
      });

      await sendLog(
        interaction.guild,
        process.env.LOG_CHANNEL_ID,
        [
          '**[APPROVE]**',
          `**User:** <@${user.id}> (${user.id})`,
          `**Player Name:** ${inGameName}`,
          `**SteamID64:** ${steamId64}`,
          `**Enlisted Date:** ${enlistedDate}`,
          `**Roles Added:** Trainee, Training Company`,
          `**Nickname Set:** TR ${inGameName}`,
          `**Approved By:** ${interaction.user.tag}`,
          `**Sheet Updated:** Trainees`,
          `**Row Filled:** ${rowNumber}`,
          `**Channel:** <#${interaction.channelId}>`,
        ].join('\n')
      );
    } catch (error) {
      console.error('Approve command failed:', error);

      await interaction.editReply({
        content: '❌ Approval failed. Check Google credentials, spreadsheet sharing, and bot permissions.',
      });
    }
  },
};