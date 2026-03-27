require('dotenv').config();
const fs = require('node:fs');
const path = require('node:path');
const { REST, Routes } = require('discord.js');

const commands = [];
const commandsPath = path.join(__dirname, 'commands');
const commandFiles = fs.readdirSync(commandsPath).filter(file => file.endsWith('.js'));

for (const file of commandFiles) {
  const filePath = path.join(commandsPath, file);

  try {
    const command = require(filePath);

    if (!command || !command.data || typeof command.data.toJSON !== 'function') {
      console.error(`Invalid command export in file: ${file}`);
      continue;
    }

    commands.push(command.data.toJSON());
    console.log(`Loaded command: ${file}`);
  } catch (error) {
    console.error(`Failed to load command file: ${file}`);
    console.error(error);
  }
}

const rest = new REST().setToken(process.env.DISCORD_TOKEN);

(async () => {
  try {
    console.log(`Registering ${commands.length} slash commands...`);

    await rest.put(
      Routes.applicationGuildCommands(
        process.env.CLIENT_ID,
        process.env.GUILD_ID
      ),
      { body: commands }
    );

    console.log('Successfully registered application commands.');
  } catch (error) {
    console.error(error);
  }
})();