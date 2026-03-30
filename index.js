require('dotenv').config();
const fs = require('node:fs');
const path = require('node:path');
const { Client, Collection, Events, GatewayIntentBits } = require('discord.js');
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
const commandFiles = fs.readdirSync(commandsPath).filter(file => file.endsWith('.js'));

for (const file of commandFiles) {
  const filePath = path.join(commandsPath, file);

  try {
    const command = require(filePath);

    if ('data' in command && 'execute' in command) {
      client.commands.set(command.data.name, command);
      console.log(`Loaded command: ${command.data.name} (${file})`);
    } else {
      console.warn(`Skipped invalid command file: ${file}`);
    }
  } catch (error) {
    console.error(`Failed to load command file: ${file}`);
    console.error(error);
  }
}

client.once(Events.ClientReady, async readyClient => {
  console.log(`Logged in as ${readyClient.user.tag}`);

  try {
    const guilds = await readyClient.guilds.fetch();

    for (const [, guildPreview] of guilds) {
      try {
        const guild = await readyClient.guilds.fetch(guildPreview.id);
        await guild.members.fetch();
        console.log(`Cached ${guild.members.cache.size} members for guild: ${guild.name}`);
      } catch (guildError) {
        console.error(`Failed to cache members for guild ${guildPreview.id}:`, guildError);
      }
    }
  } catch (error) {
    console.error('Failed to preload guild members:', error);
  }
});

client.on(Events.InteractionCreate, async interaction => {
  if (interaction.isAutocomplete()) {
    const command = client.commands.get(interaction.commandName);
    if (!command || !command.autocomplete) return;

    try {
      await command.autocomplete(interaction);
    } catch (error) {
      console.error('Autocomplete error:', error);
    }
    return;
  }

  if (!interaction.isChatInputCommand()) return;

  const command = client.commands.get(interaction.commandName);
  if (!command) return;

  try {
    await command.execute(interaction);
  } catch (error) {
    console.error(error);

    const message = 'Error executing command. Check permissions, role hierarchy, IDs, and config.';

    try {
      if (interaction.replied || interaction.deferred) {
        await interaction.followUp({
          content: message,
          ephemeral: true,
        });
      } else {
        await interaction.reply({
          content: message,
          ephemeral: true,
        });
      }
    } catch (replyError) {
      console.error('Failed to send error response:', replyError);
    }
  }
});

client.login(process.env.DISCORD_TOKEN);