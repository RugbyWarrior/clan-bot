const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const {
  findTraineeRowsByName,
  batchUpdateTraineeCells,
} = require('../sheets');

const ALLOWED_CHANNEL_ID = process.env.ALLOWED_CHANNEL_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const TRAINEE_SHEET_NAME = process.env.TRAINEE_SHEET_NAME || 'Trainees';

function isValidDiscordId(value) {
  return /^\d{17,20}$/.test(String(value || '').trim());
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('linktrainee')
    .setDescription('Manually link a trainee sheet row to a Discord user by sheet name.')
    .addUserOption(option =>
      option
        .setName('trainee')
        .setDescription('The Discord user to link')
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName('sheet_name')
        .setDescription('Exact trainee name as it appears on the Trainees sheet')
        .setRequired(true)
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
      const sheetName = interaction.options.getString('sheet_name')?.trim();

      if (!traineeUser || !sheetName) {
        return interaction.editReply({
          content: 'You must provide both a trainee user and a sheet_name.',
          allowedMentions: { users: [] },
        });
      }

      if (!isValidDiscordId(traineeUser.id)) {
        return interaction.editReply({
          content: 'That user does not have a valid Discord ID.',
          allowedMentions: { users: [] },
        });
      }

      const matches = await findTraineeRowsByName(sheetName);

      if (matches.length === 0) {
        return interaction.editReply({
          content: `No trainee row was found for sheet name **${sheetName}**.`,
          allowedMentions: { users: [] },
        });
      }

      if (matches.length > 1) {
        return interaction.editReply({
          content: `Multiple trainee rows matched **${sheetName}**. Please make the sheet name more specific first.`,
          allowedMentions: { users: [] },
        });
      }

      const match = matches[0];
      const rowNumber = match.rowNumber;
      const rowValues = match.rowValues || [];
      const existingDiscordId = (rowValues[8] || '').toString().trim();

      if (existingDiscordId && existingDiscordId !== traineeUser.id) {
        return interaction.editReply({
          content:
            `That trainee row already has a different Discord ID set: \`${existingDiscordId}\`.\n` +
            `Clear or correct it first before linking.`,
          allowedMentions: { users: [] },
        });
      }

      await batchUpdateTraineeCells([
        {
          range: `${TRAINEE_SHEET_NAME}!I${rowNumber}`,
          values: [[traineeUser.id]],
        },
      ]);

      await sendLog(
        interaction.guild,
        LOG_CHANNEL_ID,
        [
          '🔗 **MANUAL TRAINEE LINK**',
          `Moderator: <@${interaction.user.id}>`,
          `Discord User: <@${traineeUser.id}> (${traineeUser.id})`,
          `Sheet Name: ${sheetName}`,
          `Trainee Row: ${rowNumber}`,
          `Previous Discord ID: ${existingDiscordId || 'None'}`,
        ].join('\n')
      );

      return interaction.editReply({
        content:
          `✅ Linked <@${traineeUser.id}> to trainee row **${rowNumber}**.\n` +
          `Sheet name: **${sheetName}**\n` +
          `Discord ID written: \`${traineeUser.id}\``,
        allowedMentions: { users: [] },
      });
    } catch (error) {
      console.error('Error in /linktrainee:', error);

      return interaction.editReply({
        content: 'There was an error while manually linking the trainee.',
        allowedMentions: { users: [] },
      });
    }
  },
};