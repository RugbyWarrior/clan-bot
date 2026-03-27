const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const { mosConfig } = require('../mosConfig');
const {
  findRatingsRowByDiscordId,
  writeRatingsRow,
  batchUpdateRatingsCells,
  findTraineeRowsByName,
} = require('../sheets');

const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';
const TRAINEE_ROLE_ID = process.env.TRAINEE_ROLE_ID;
const EX_SKIRA_ROLE_ID = process.env.EX_SKIRA_ROLE_ID;

const HEADER_TO_COLUMN = {
  'Rank': 'A',
  'Squadron': 'B',
  'Name': 'C',
  'Infantry': 'D',
  'Squad Leader': 'E',
  'Shooter': 'F',
  'AT': 'G',
  'Engineer': 'H',
  'Medic': 'I',
  'Grenadier': 'J',
  'Mortar': 'K',
  'Pilot': 'L',
  'Driver': 'M',
  'APC/IFV Gunner': 'N',
  'MBT Gunner': 'O',
  'Fire Support (R)': 'P',
  'Irregular Warfare (R)': 'Q',
  'Knife (R)': 'R',
  'Discord ID': 'S',
};

function resolveSheetColumn(columnOrHeader) {
  if (!columnOrHeader) return null;

  const trimmed = String(columnOrHeader).trim();

  if (/^[A-S]$/i.test(trimmed)) {
    return trimmed.toUpperCase();
  }

  return HEADER_TO_COLUMN[trimmed] || null;
}

function getDisplayNameWithoutRank(member) {
  const displayName = member.nickname || member.user.username;

  if (!displayName.includes(' ')) {
    return displayName;
  }

  return displayName.replace(/^\S+\s+/, '').trim();
}

function getSheetSquadronName(member) {
  if (member.roles.cache.has(process.env.INFANTRY_ROLE_ID)) return '3rd Rifles';
  if (member.roles.cache.has(process.env.ARMOUR_ROLE_ID)) return '20th Hussars';
  if (member.roles.cache.has(process.env.AVIATION_ROLE_ID)) return '230th Aviation';
  return '';
}

function getAllowedRankRole(member, guild) {
  const rankRoleIds = (process.env.RANK_ROLE_IDS || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);

  for (const roleId of rankRoleIds) {
    if (!roleId) continue;
    if (roleId === TRAINEE_ROLE_ID || roleId === EX_SKIRA_ROLE_ID) continue;

    if (member.roles.cache.has(roleId)) {
      return guild.roles.cache.get(roleId) || null;
    }
  }

  return null;
}

function getMosSheetValue(member, guild, mos) {
  for (const [ratingName, roleId] of Object.entries(mos.ratings)) {
    if (!roleId) continue;

    if (member.roles.cache.has(roleId)) {
      const roleObject = guild.roles.cache.get(roleId);
      return roleObject ? roleObject.name : ratingName;
    }
  }

  return '';
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('mos')
    .setDescription('Set a user MOS qualification')
    .addUserOption(option =>
      option.setName('user')
        .setDescription('User to update')
        .setRequired(true)
    )
    .addStringOption(option =>
      option.setName('mos')
        .setDescription('MOS field to update')
        .setRequired(true)
        .addChoices(
          { name: 'Infantry', value: 'INFANTRY' },
          { name: 'Squad Leader', value: 'SQUAD_LEADER' },
          { name: 'Shooter', value: 'SHOOTER' },
          { name: 'AT', value: 'AT' },
          { name: 'Engineer', value: 'ENGINEER' },
          { name: 'Medic', value: 'MEDIC' },
          { name: 'Grenadier', value: 'GRENADIER' },
          { name: 'Mortar', value: 'MORTAR' },
          { name: 'Pilot', value: 'PILOT' },
          { name: 'Driver', value: 'DRIVER' },
          { name: 'APC/IFV Gunner', value: 'APC_IFV_GUNNER' },
          { name: 'MBT Gunner', value: 'MBT_GUNNER' },
          { name: 'Fire Support (R)', value: 'FIRE_SUPPORT_R' },
          { name: 'Knife (R)', value: 'KNIFE_R' }
        )
    )
    .addStringOption(option =>
      option.setName('rating')
        .setDescription('Exact MOS rating to assign')
        .setRequired(true)
        .setAutocomplete(true)
    ),

  async autocomplete(interaction) {
    const mosKey = interaction.options.getString('mos');
    const focusedValue = interaction.options.getFocused().toLowerCase();

    if (!mosKey || !mosConfig[mosKey]) {
      await interaction.respond([]);
      return;
    }

    const ratingNames = Object.keys(mosConfig[mosKey].ratings);

    const filtered = ratingNames
      .filter(name => name.toLowerCase().includes(focusedValue))
      .slice(0, 25)
      .map(name => ({
        name,
        value: name,
      }));

    await interaction.respond(filtered);
  },

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (interaction.channelId !== process.env.ALLOWED_CHANNEL_ID) {
        await interaction.editReply({
          content: '❌ You can only use this command in the training/management channel.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const user = interaction.options.getUser('user');
      const mosKey = interaction.options.getString('mos');
      const ratingName = interaction.options.getString('rating');

      const mos = mosConfig[mosKey];

      if (!mos) {
        await interaction.editReply({
          content: '❌ That MOS is not configured.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const targetSheetColumn = resolveSheetColumn(mos.sheetColumn);
      if (!targetSheetColumn) {
        await interaction.editReply({
          content: `❌ MOS sheet column for ${mos.label} is invalid in mosConfig.js.`,
          allowedMentions: { users: [] },
        });
        return;
      }

      if (!(ratingName in mos.ratings)) {
        await interaction.editReply({
          content: `❌ "${ratingName}" is not a valid rating for ${mos.label}.`,
          allowedMentions: { users: [] },
        });
        return;
      }

      const member = await interaction.guild.members.fetch(user.id);

      if (TRAINEE_ROLE_ID && member.roles.cache.has(TRAINEE_ROLE_ID)) {
        await interaction.editReply({
          content: '❌ Trainees cannot be given MOS ratings. Promote them first.',
          allowedMentions: { users: [] },
        });
        return;
      }

      if (EX_SKIRA_ROLE_ID && member.roles.cache.has(EX_SKIRA_ROLE_ID)) {
        await interaction.editReply({
          content: '❌ Ex Skira members cannot be given MOS ratings.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const rankRole = getAllowedRankRole(member, interaction.guild);
      if (!rankRole) {
        await interaction.editReply({
          content: '❌ This member does not have a valid ranked role for the Ratings sheet.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const targetRoleId = mos.ratings[ratingName];
      const mosRoleIdsForCategory = Object.values(mos.ratings).filter(id => id !== null);
      const uniqueMosRoleIds = [...new Set(mosRoleIdsForCategory)];
      const rolesToRemove = uniqueMosRoleIds.filter(roleId => member.roles.cache.has(roleId));

      if (rolesToRemove.length > 0) {
        await member.roles.remove(rolesToRemove, `Cleaning ${mos.label} MOS roles`);
      }

      let roleName = 'None';
      if (targetRoleId) {
        await member.roles.add(targetRoleId, `MOS updated: ${mos.label} -> ${ratingName}`);
        const roleObject = interaction.guild.roles.cache.get(targetRoleId);
        roleName = roleObject ? roleObject.name : ratingName;
      }

      // Refresh member after role changes so sheet reflects final Discord state
      const refreshedMember = await interaction.guild.members.fetch(user.id);

      let ratingsRow = await findRatingsRowByDiscordId(user.id);
      let createdRatingsRow = false;

      const traineeNameMatches = await findTraineeRowsByName(getDisplayNameWithoutRank(refreshedMember));
      const traineeSheetName =
        traineeNameMatches.length === 1 ? (traineeNameMatches[0].rowValues[0] || '') : '';

      const finalName = traineeSheetName || getDisplayNameWithoutRank(refreshedMember);
      const finalRank = getAllowedRankRole(refreshedMember, interaction.guild)?.name || rankRole.name;
      const finalSquadron = getSheetSquadronName(refreshedMember);

      if (!ratingsRow) {
        const newRow = new Array(19).fill('');
        newRow[0] = finalRank;
        newRow[1] = finalSquadron;
        newRow[2] = finalName;
        newRow[18] = user.id;

        const rowNumber = await writeRatingsRow(newRow);

        ratingsRow = {
          rowNumber,
          rowValues: newRow,
        };

        createdRatingsRow = true;
      }

      const updates = [
        {
          range: `${RATINGS_SHEET_NAME}!A${ratingsRow.rowNumber}`,
          values: [[finalRank]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!B${ratingsRow.rowNumber}`,
          values: [[finalSquadron]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!C${ratingsRow.rowNumber}`,
          values: [[finalName]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!S${ratingsRow.rowNumber}`,
          values: [[user.id]],
        },
      ];

      // Rewrite every MOS column from actual Discord role state so sheet stays consistent
      for (const config of Object.values(mosConfig)) {
        const column = resolveSheetColumn(config.sheetColumn);
        if (!column) continue;

        updates.push({
          range: `${RATINGS_SHEET_NAME}!${column}${ratingsRow.rowNumber}`,
          values: [[getMosSheetValue(refreshedMember, interaction.guild, config)]],
        });
      }

      await batchUpdateRatingsCells(updates);

      await interaction.editReply({
        content:
          targetRoleId
            ? `✅ Set ${user.tag}'s **${mos.label}** MOS to **${ratingName}**, gave role **${roleName}**, and updated the Ratings sheet.`
            : `🧹 Set ${user.tag}'s **${mos.label}** MOS to **Unrated**, removed existing ${mos.label} MOS roles, and updated the Ratings sheet.`,
        allowedMentions: { users: [] },
      });

      await sendLog(
        interaction.guild,
        process.env.LOG_CHANNEL_ID,
        [
          '**[MOS]**',
          `**User:** ${user.tag} (${user.id})`,
          `**MOS:** ${mos.label}`,
          `**Sheet Name:** ${RATINGS_SHEET_NAME}`,
          `**Sheet Column:** ${targetSheetColumn}`,
          `**Rating Set:** ${ratingName}`,
          `**Role Added:** ${roleName}`,
          `**Ratings Row:** ${ratingsRow.rowNumber}`,
          `**Ratings Row Created:** ${createdRatingsRow ? 'Yes' : 'No'}`,
          `**Final Rank:** ${finalRank}`,
          `**Final Squadron:** ${finalSquadron || 'None'}`,
          `**Final Name:** ${finalName}`,
          `**Done By:** ${interaction.user.tag}`,
          `**Channel:** <#${interaction.channelId}>`,
        ].join('\n')
      );
    } catch (error) {
      console.error('Error in /mos:', error);

      try {
        await interaction.editReply({
          content: '❌ There was an error while running /mos. Check the console for details.',
          allowedMentions: { users: [] },
        });
      } catch {}
    }
  },
};