require('dotenv').config();

const fs = require('node:fs');
const path = require('node:path');
const { REST, Routes } = require('discord.js');

const required = ['DISCORD_TOKEN', 'CLIENT_ID', 'GUILD_ID'];
const missing = required.filter(key => !String(process.env[key] || '').trim());

if (missing.length > 0) {
  console.error(`Missing required .env value(s): ${missing.join(', ')}`);
  process.exit(1);
}

const commands = [];
const commandsPath = path.join(__dirname, 'commands');
const commandFiles = fs
  .readdirSync(commandsPath)
  .filter(file => file.endsWith('.js'));

for (const file of commandFiles) {
  const command = require(path.join(commandsPath, file));

  if (!command?.data || typeof command.data.toJSON !== 'function') {
    console.warn(`Skipped invalid command file: ${file}`);
    continue;
  }

  commands.push(command.data.toJSON());
  console.log(`Prepared /${command.data.name}`);
}

const rest = new REST({ version: '10' }).setToken(process.env.DISCORD_TOKEN);

(async () => {
  try {
    console.log(`Registering ${commands.length} Orion guild command(s)...`);

    await rest.put(
      Routes.applicationGuildCommands(
        process.env.CLIENT_ID.trim(),
        process.env.GUILD_ID.trim()
      ),
      { body: commands }
    );

    console.log('Orion slash commands registered successfully.');
  } catch (error) {
    console.error('Command registration failed:', error);
    process.exitCode = 1;
  }
})();
