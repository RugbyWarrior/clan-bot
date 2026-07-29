require('dotenv').config();

const fs = require('node:fs');
const path = require('node:path');
const {
  Client,
  Collection,
  Events,
  GatewayIntentBits,
} = require('discord.js');
const { printStartupValidationReport } = require('./utils/startupValidation');

const startupCheck = printStartupValidationReport();
if (startupCheck.status !== 'valid') {
  process.exit(1);
}

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMembers,
  ],
});

client.commands = new Collection();

const commandsPath = path.join(__dirname, 'commands');
const commandFiles = fs
  .readdirSync(commandsPath)
  .filter(file => file.endsWith('.js'));

for (const file of commandFiles) {
  const filePath = path.join(commandsPath, file);

  try {
    const command = require(filePath);

    if (!command?.data || typeof command.execute !== 'function') {
      console.warn(`Skipped invalid command file: ${file}`);
      continue;
    }

    client.commands.set(command.data.name, command);
    console.log(`Loaded command: ${command.data.name}`);
  } catch (error) {
    console.error(`Failed to load command file: ${file}`);
    console.error(error);
  }
}

client.once(Events.ClientReady, readyClient => {
  console.log(`Orion logged in as ${readyClient.user.tag}`);
});

client.on(Events.InteractionCreate, async interaction => {
  if (!interaction.isChatInputCommand()) return;

  const command = client.commands.get(interaction.commandName);
  if (!command) return;

  try {
    await command.execute(interaction);
  } catch (error) {
    console.error(`Unhandled error in /${interaction.commandName}:`, error);

    const message = '❌ Orion could not complete that command. Check the console and bot log channel.';

    try {
      if (interaction.replied || interaction.deferred) {
        await interaction.followUp({
          content: message,
          ephemeral: true,
          allowedMentions: { parse: [] },
        });
      } else {
        await interaction.reply({
          content: message,
          ephemeral: true,
          allowedMentions: { parse: [] },
        });
      }
    } catch (replyError) {
      console.error('Failed to send the interaction error message:', replyError);
    }
  }
});

client.login(process.env.DISCORD_TOKEN);
