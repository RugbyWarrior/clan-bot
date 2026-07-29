const { SlashCommandBuilder } = require('discord.js');

module.exports = {
  data: new SlashCommandBuilder()
    .setName('help')
    .setDescription('Show the current Orion command guide'),

  async execute(interaction) {
    await interaction.reply({
      ephemeral: true,
      allowedMentions: { parse: [] },
      content: [
        '🌌 **ORION COMMAND GUIDE**',
        '',
        '**/accept-cadet**',
        'Accepts a new cadet and:',
        '• checks for duplicate Discord IDs and IGNs',
        '• reuses a listed mother-group CT number or allocates the next free 53000–53999 number',
        '• adds the cadet to the Cadets tab',
        '• adds the cadet to the Cadets section of Game Activity',
        '• sets the nickname to `CT-[IGN]-[CT Number]`',
        '• adds Cadet, Training Platoon and Infantry',
        '• removes Community Member',
        '',
        'Only Officers can use the command, and only in the configured acceptance channel.',
      ].join('\n'),
    });
  },
};
