const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const { updateNickname } = require('../utils/updateNickname');
const {
  findRatingsRowByDiscordId,
  writeRatingsRow,
  batchUpdateRatingsCells,
  findTraineeRowByDiscordId,
  findTraineeRowsByName,
  deleteTraineeRow,
} = require('../sheets');

const RATINGS_SHEET_NAME = process.env.RATINGS_SHEET_NAME || 'Ratings';

const rankConfig = {
  EX_SKIRA: {
    label: 'Ex Skira',
    roleEnv: 'EX_SKIRA_ROLE_ID',
    breakerEnv: null,
    allowedSquadrons: [],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: null,
    exSkiraSuffix: true,
  },
  TRAINEE: {
    label: 'Trainee',
    roleEnv: 'TRAINEE_ROLE_ID',
    breakerEnv: null,
    allowedSquadrons: [],
    traineeState: true,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'TR',
    exSkiraSuffix: false,
  },

  PRIVATE: {
    label: 'Private',
    roleEnv: 'PRIVATE_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: true,
    nicknamePrefix: 'PVT',
    exSkiraSuffix: false,
  },
  ARMOUR_CADET: {
    label: 'Armour Cadet',
    roleEnv: 'ARMOUR_CADET_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'CDT',
    exSkiraSuffix: false,
  },
  OFFICER_CADET: {
    label: 'Officer Cadet',
    roleEnv: 'OFFICER_CADET_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'O/CDT',
    exSkiraSuffix: false,
  },
  ARMOUR_TROOPER: {
    label: 'Armour Trooper',
    roleEnv: 'ARMOUR_TROOPER_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'TPR',
    exSkiraSuffix: false,
  },
  PILOT_OFFICER: {
    label: 'Pilot Officer',
    roleEnv: 'PILOT_OFFICER_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'P/O',
    exSkiraSuffix: false,
  },
  LANCE_CORPORAL: {
    label: 'Lance Corporal',
    roleEnv: 'LANCE_CORPORAL_ROLE_ID',
    breakerEnv: 'ENLISTED_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'LCPL',
    exSkiraSuffix: false,
  },

  FLYING_OFFICER: {
    label: 'Flying Officer',
    roleEnv: 'FLYING_OFFICER_ROLE_ID',
    breakerEnv: 'NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'F/O',
    exSkiraSuffix: false,
  },
  CORPORAL: {
    label: 'Corporal',
    roleEnv: 'CORPORAL_ROLE_ID',
    breakerEnv: 'NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'CPL',
    exSkiraSuffix: false,
  },

  SERGEANT: {
    label: 'Sergeant',
    roleEnv: 'SERGEANT_ROLE_ID',
    breakerEnv: 'SENIOR_NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID', 'AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'SGT',
    exSkiraSuffix: false,
  },
  STAFF_SERGEANT: {
    label: 'Staff Sergeant',
    roleEnv: 'STAFF_SERGEANT_ROLE_ID',
    breakerEnv: 'SENIOR_NCO_BREAKER_ROLE_ID',
    allowedSquadrons: ['INFANTRY_ROLE_ID', 'ARMOUR_ROLE_ID', 'AVIATION_ROLE_ID'],
    traineeState: false,
    autoInfantryFromTrainee: false,
    nicknamePrefix: 'SSGT',
    exSkiraSuffix: false,
  },
};

function getDisplayNameWithoutRank(member) {
  const displayName = member.nickname || member.user.username;
  if (!displayName.includes(' ')) return displayName;
  return displayName.replace(/^\S+\s+/, '').trim();
}

function getSheetSquadronName(member) {
  if (member.roles.cache.has(process.env.INFANTRY_ROLE_ID)) return '3rd Rifles';
  if (member.roles.cache.has(process.env.ARMOUR_ROLE_ID)) return '20th Hussars';
  if (member.roles.cache.has(process.env.AVIATION_ROLE_ID)) return '230th Aviation';
  return '';
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('setrank')
    .setDescription('Set a user rank role')
    .addUserOption(option =>
      option.setName('user')
        .setDescription('User to update')
        .setRequired(true)
    )
    .addStringOption(option =>
      option.setName('rank')
        .setDescription('Rank to assign')
        .setRequired(true)
        .addChoices(
          { name: 'Ex Skira', value: 'EX_SKIRA' },
          { name: 'Trainee', value: 'TRAINEE' },
          { name: 'Private', value: 'PRIVATE' },
          { name: 'Armour Cadet', value: 'ARMOUR_CADET' },
          { name: 'Officer Cadet', value: 'OFFICER_CADET' },
          { name: 'Armour Trooper', value: 'ARMOUR_TROOPER' },
          { name: 'Pilot Officer', value: 'PILOT_OFFICER' },
          { name: 'Lance Corporal', value: 'LANCE_CORPORAL' },
          { name: 'Flying Officer', value: 'FLYING_OFFICER' },
          { name: 'Corporal', value: 'CORPORAL' },
          { name: 'Sergeant', value: 'SERGEANT' },
          { name: 'Staff Sergeant', value: 'STAFF_SERGEANT' }
        )
    ),

  async execute(interaction) {
    await interaction.deferReply({
      ephemeral: true,
      allowedMentions: { users: [] },
    });

    try {
      if (interaction.channelId !== process.env.HQ_CHANNEL_ID) {
        await interaction.editReply({
          content: '❌ You can only use this command in Headquarters.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const user = interaction.options.getUser('user');
      const rankKey = interaction.options.getString('rank');
      let member = await interaction.guild.members.fetch(user.id);

      const selectedRank = rankConfig[rankKey];
      if (!selectedRank) {
        await interaction.editReply({
          content: '❌ That rank is not configured.',
          allowedMentions: { users: [] },
        });
        return;
      }

      const targetRankRoleId = process.env[selectedRank.roleEnv];
      const breakerRoleToAdd = selectedRank.breakerEnv
        ? process.env[selectedRank.breakerEnv]
        : null;

      if (!targetRankRoleId) {
        await interaction.editReply({
          content: `❌ Missing ${selectedRank.roleEnv} in .env`,
          allowedMentions: { users: [] },
        });
        return;
      }

      const rankRoleIds = (process.env.RANK_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const breakerRoleIds = (process.env.BREAKER_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const squadronRoleIds = (process.env.SQUADRON_ROLE_IDS || '')
        .split(',')
        .map(id => id.trim())
        .filter(Boolean);

      const traineeRoleId = process.env.TRAINEE_ROLE_ID;
      const trainingCompanyRoleId = process.env.TRAINING_COMPANY_ROLE_ID;
      const infantryRoleId = process.env.INFANTRY_ROLE_ID;

      const wasCurrentlyTrainee = traineeRoleId
        ? member.roles.cache.has(traineeRoleId)
        : false;

      let traineeRow = null;
      let traineeSheetName = null;

      if (wasCurrentlyTrainee) {
        traineeRow = await findTraineeRowByDiscordId(user.id);

        if (!traineeRow) {
          const nameMatches = await findTraineeRowsByName(getDisplayNameWithoutRank(member));
          if (nameMatches.length === 1) {
            traineeRow = nameMatches[0];
          }
        }

        if (traineeRow) {
          traineeSheetName = traineeRow.rowValues[0] || null;
        }
      }

      if (selectedRank.traineeState) {
        const rolesToRemove = [];

        for (const rankRoleId of rankRoleIds) {
          if (member.roles.cache.has(rankRoleId) && rankRoleId !== targetRankRoleId) {
            rolesToRemove.push(rankRoleId);
          }
        }

        for (const breakerRoleId of breakerRoleIds) {
          if (member.roles.cache.has(breakerRoleId)) {
            rolesToRemove.push(breakerRoleId);
          }
        }

        for (const squadronRoleId of squadronRoleIds) {
          if (member.roles.cache.has(squadronRoleId)) {
            rolesToRemove.push(squadronRoleId);
          }
        }

        if (rolesToRemove.length > 0) {
          await member.roles.remove(rolesToRemove, 'Resetting member to trainee state');
        }

        if (!member.roles.cache.has(targetRankRoleId)) {
          await member.roles.add(targetRankRoleId, 'Rank set to Trainee');
        }

        if (!trainingCompanyRoleId) {
          await interaction.editReply({
            content: '❌ Missing TRAINING_COMPANY_ROLE_ID in .env',
            allowedMentions: { users: [] },
          });
          return;
        }

        if (!member.roles.cache.has(trainingCompanyRoleId)) {
          await member.roles.add(trainingCompanyRoleId, 'Training Company role added for Trainee');
        }

        await updateNickname(member, {
          prefix: selectedRank.nicknamePrefix,
          exSkira: selectedRank.exSkiraSuffix,
        });

        await interaction.editReply({
          content: `✅ Set <@${user.id}> to **${selectedRank.label}** and added **Training Company**.`,
          allowedMentions: { users: [] },
        });

        await sendLog(
          interaction.guild,
          process.env.LOG_CHANNEL_ID,
          [
            '**[SET RANK]**',
            `**User:** <@${user.id}> (${user.id})`,
            `**New Rank:** ${selectedRank.label}`,
            `**Breaker Added:** None`,
            `**Squadron Added:** None`,
            `**Ratings Sheet Updated:** No`,
            `**Trainee Row Removed:** No`,
            `**Done By:** ${interaction.user.tag}`,
            `**Channel:** <#${interaction.channelId}>`,
          ].join('\n')
        );
        return;
      }

      const currentSquadronEnv = selectedRank.allowedSquadrons.find(envName => {
        const roleId = process.env[envName];
        return roleId && member.roles.cache.has(roleId);
      });

      let squadronRoleToAdd = null;

      if (selectedRank.autoInfantryFromTrainee && wasCurrentlyTrainee) {
        if (!infantryRoleId) {
          await interaction.editReply({
            content: '❌ Missing INFANTRY_ROLE_ID in .env',
            allowedMentions: { users: [] },
          });
          return;
        }

        squadronRoleToAdd = infantryRoleId;
      } else {
        if (selectedRank.allowedSquadrons.length > 0 && !currentSquadronEnv) {
          await interaction.editReply({
            content: `❌ <@${user.id}> does not have the required squadron role for **${selectedRank.label}**.`,
            allowedMentions: { users: [] },
          });
          return;
        }
      }

      const rolesToRemove = [];

      if (traineeRoleId && member.roles.cache.has(traineeRoleId)) {
        rolesToRemove.push(traineeRoleId);
      }

      if (trainingCompanyRoleId && member.roles.cache.has(trainingCompanyRoleId)) {
        rolesToRemove.push(trainingCompanyRoleId);
      }

      for (const rankRoleId of rankRoleIds) {
        if (member.roles.cache.has(rankRoleId) && rankRoleId !== targetRankRoleId) {
          rolesToRemove.push(rankRoleId);
        }
      }

      for (const breakerRoleId of breakerRoleIds) {
        if (member.roles.cache.has(breakerRoleId)) {
          rolesToRemove.push(breakerRoleId);
        }
      }

      if (squadronRoleToAdd) {
        for (const squadronRoleId of squadronRoleIds) {
          if (member.roles.cache.has(squadronRoleId) && squadronRoleId !== squadronRoleToAdd) {
            rolesToRemove.push(squadronRoleId);
          }
        }
      }

      if (rolesToRemove.length > 0) {
        await member.roles.remove(rolesToRemove, 'Cleaning old rank, breaker, trainee, and squadron roles');
      }

      if (!member.roles.cache.has(targetRankRoleId)) {
        await member.roles.add(targetRankRoleId, `Rank set to ${selectedRank.label}`);
      }

      if (breakerRoleToAdd && !member.roles.cache.has(breakerRoleToAdd)) {
        await member.roles.add(breakerRoleToAdd, `Breaker role added for ${selectedRank.label}`);
      }

      if (squadronRoleToAdd && !member.roles.cache.has(squadronRoleToAdd)) {
        await member.roles.add(squadronRoleToAdd, 'Auto-assigned Infantry on first promotion from trainee');
      }

      await updateNickname(member, {
        prefix: selectedRank.nicknamePrefix,
        exSkira: selectedRank.exSkiraSuffix,
      });

      member = await interaction.guild.members.fetch(user.id);

      const squadronName = getSheetSquadronName(member);
      let breakerName = 'None';

      let replyMessage = `✅ Set <@${user.id}> to rank **${selectedRank.label}**.`;

      if (breakerRoleToAdd) {
        const breakerRole = interaction.guild.roles.cache.get(breakerRoleToAdd);
        if (breakerRole) {
          breakerName = breakerRole.name;
          replyMessage += ` Added breaker role **${breakerRole.name}**.`;
        }
      }

      if (squadronRoleToAdd) {
        const squadronRole = interaction.guild.roles.cache.get(squadronRoleToAdd);
        if (squadronRole) {
          replyMessage += ` Added squadron role **${squadronRole.name}**.`;
        }
      }

      let traineeRowRemoved = false;
      if (wasCurrentlyTrainee && traineeRow) {
        await deleteTraineeRow(traineeRow.rowNumber);
        traineeRowRemoved = true;
      }

      let ratingsRow = await findRatingsRowByDiscordId(user.id);
      let createdRatingsRow = false;

      const finalSheetName = traineeSheetName || getDisplayNameWithoutRank(member);

      if (!ratingsRow) {
        const newRow = new Array(19).fill('');
        newRow[0] = selectedRank.label;
        newRow[1] = squadronName;
        newRow[2] = finalSheetName;
        newRow[18] = user.id;

        const rowNumber = await writeRatingsRow(newRow);

        ratingsRow = {
          rowNumber,
          rowValues: newRow,
        };

        createdRatingsRow = true;
      }

      await batchUpdateRatingsCells([
        {
          range: `${RATINGS_SHEET_NAME}!A${ratingsRow.rowNumber}`,
          values: [[selectedRank.label]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!B${ratingsRow.rowNumber}`,
          values: [[squadronName]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!C${ratingsRow.rowNumber}`,
          values: [[finalSheetName]],
        },
        {
          range: `${RATINGS_SHEET_NAME}!S${ratingsRow.rowNumber}`,
          values: [[user.id]],
        },
      ]);

      await interaction.editReply({
        content: replyMessage,
        allowedMentions: { users: [] },
      });

      await sendLog(
        interaction.guild,
        process.env.LOG_CHANNEL_ID,
        [
          '**[SET RANK]**',
          `**User:** <@${user.id}> (${user.id})`,
          `**New Rank:** ${selectedRank.label}`,
          `**Breaker Added:** ${breakerName}`,
          `**Squadron Added:** ${squadronName || 'None'}`,
          `**Ratings Sheet Name:** ${RATINGS_SHEET_NAME}`,
          `**Ratings Row:** ${ratingsRow.rowNumber}`,
          `**Ratings Row Created:** ${createdRatingsRow ? 'Yes' : 'No'}`,
          `**Trainee Row Removed:** ${traineeRowRemoved ? 'Yes' : 'No'}`,
          `**Done By:** ${interaction.user.tag}`,
          `**Channel:** <#${interaction.channelId}>`,
        ].join('\n')
      );
    } catch (error) {
      console.error('Error in /setrank:', error);

      try {
        await interaction.editReply({
          content: '❌ There was an error while running /setrank. Check the console for details.',
          allowedMentions: { users: [] },
        });
      } catch {}
    }
  },
};