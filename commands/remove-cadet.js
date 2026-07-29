const {
  SlashCommandBuilder,
} = require('discord.js');

const {
  sendLog,
} = require('../logger');

const {
  isBotOwner,
  hasAnyRole,
} = require('../permissions');

const {
  findCadetByDiscordId,
  findRosterByDiscordId,

  findCtEntriesByName,
  findCtEntriesByDiscordId,

  updateCtIdentity,
  clearCtIdentity,

  backupCadetRow,
  restoreCadetRow,
  clearCadetRow,

  findGameActivityCadetRow,
  deleteGameActivityRow,
} = require('../sheets');

const normalise = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');

const localNumber = value =>
  /^\d{5}$/.test(
    String(value || '').trim()
  ) &&
  Number(value) >= 53000 &&
  Number(value) <= 53999;

async function checkDiscord(
  member,
  guild,
  restoreCommunity
) {
  if (!member.manageable) {
    throw new Error(
      'Orion cannot manage this member. Move Orion above their highest role.'
    );
  }

  const roleIds = [
    process.env.CADET_ROLE_ID,

    process.env
      .TRAINING_PLATOON_ROLE_ID,

    process.env.INFANTRY_ROLE_ID,

    process.env
      .UNIT_DIVIDER_ROLE_ID,
  ];

  if (restoreCommunity) {
    roleIds.push(
      process.env
        .COMMUNITY_MEMBER_ROLE_ID
    );
  }

  for (const roleId of roleIds) {
    if (!roleId) {
      throw new Error(
        'A required Discord role ID is missing from .env.'
      );
    }

    const role =
      await guild.roles.fetch(roleId);

    if (!role) {
      throw new Error(
        `Discord role not found: ${roleId}`
      );
    }

    if (!role.editable) {
      throw new Error(
        `Orion cannot manage the role “${role.name}”. Move Orion above it.`
      );
    }
  }
}

async function restoreDiscord(
  member,
  original
) {
  try {
    const roleIds = [
      process.env.CADET_ROLE_ID,

      process.env
        .TRAINING_PLATOON_ROLE_ID,

      process.env.INFANTRY_ROLE_ID,

      process.env
        .UNIT_DIVIDER_ROLE_ID,

      process.env
        .COMMUNITY_MEMBER_ROLE_ID,
    ].filter(Boolean);

    for (const roleId of roleIds) {
      const shouldHave =
        original.roles.has(roleId);

      const hasRole =
        member.roles.cache.has(roleId);

      if (
        shouldHave &&
        !hasRole
      ) {
        await member.roles.add(
          roleId,
          'Orion rollback'
        );
      }

      if (
        !shouldHave &&
        hasRole
      ) {
        await member.roles.remove(
          roleId,
          'Orion rollback'
        );
      }
    }

    if (
      member.nickname !==
      original.nickname
    ) {
      await member.setNickname(
        original.nickname,
        'Orion rollback'
      );
    }
  } catch (error) {
    console.error(
      'Discord rollback failed:',
      error
    );
  }
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('remove-cadet')

    .setDescription(
      'Remove a cadet from Discord and the cadet spreadsheets'
    )

    .addUserOption(option =>
      option
        .setName('cadet')

        .setDescription(
          'The cadet to remove'
        )

        .setRequired(true)
    )

    .addStringOption(option =>
      option
        .setName(
          'ct_number_action'
        )

        .setDescription(
          'What to do with their CT-number record'
        )

        .setRequired(true)

        .addChoices(
          {
            name:
              'Keep the CT number, IGN and Discord ID recorded',

            value: 'keep',
          },

          {
            name:
              'Release a local 53xxx number for reuse',

            value: 'release',
          }
        )
    )

    .addBooleanOption(option =>
      option
        .setName(
          'restore_community_member'
        )

        .setDescription(
          'Give Community Member back; defaults to True'
        )

        .setRequired(false)
    )

    .addStringOption(option =>
      option
        .setName('reason')

        .setDescription(
          'Optional reason'
        )

        .setMaxLength(500)
        .setRequired(false)
    ),

  async execute(interaction) {
    const owner =
      isBotOwner(interaction);

    if (
      !owner &&
      interaction.channelId !==
        process.env
          .ACCEPT_CADET_CHANNEL_ID
    ) {
      return interaction.reply({
        content:
          '❌ `/remove-cadet` can only be used in the cadet-management channel.',

        ephemeral: true,
      });
    }

    const allowed =
      owner ||
      hasAnyRole(
        interaction.member,
        [
          process.env
            .OFFICER_ROLE_ID,

          process.env
            .NCO_ROLE_ID,
        ]
      );

    if (!allowed) {
      return interaction.reply({
        content:
          '❌ Only Officers and NCOs can use `/remove-cadet`.',

        ephemeral: true,
      });
    }

    await interaction.deferReply({
      ephemeral: true,
    });

    const user =
      interaction.options.getUser(
        'cadet',
        true
      );

    const action =
      interaction.options.getString(
        'ct_number_action',
        true
      );

    const restoreCommunity =
      interaction.options.getBoolean(
        'restore_community_member'
      ) ?? true;

    const reason =
      interaction.options
        .getString('reason')
        ?.trim() ||
      'No reason supplied';

    let member;

    try {
      member =
        await interaction.guild.members.fetch(
          user.id
        );
    } catch {
      return interaction.editReply(
        '❌ That user is not currently in this server.'
      );
    }

    const original = {
      nickname: member.nickname,

      roles: new Set(
        member.roles.cache.keys()
      ),
    };

    let cadet;
    let cadetBackup;
    let ctEntry;
    let ctBackup;
    let activityRow;

    let discordChanged = false;
    let cadetCleared = false;
    let ctChanged = false;
    let finished = false;

    try {
      await checkDiscord(
        member,
        interaction.guild,
        restoreCommunity
      );

      const [
        cadetRecord,
        rosterRecord,
      ] = await Promise.all([
        findCadetByDiscordId(user.id),

        findRosterByDiscordId(
          user.id
        ),
      ]);

      if (!cadetRecord) {
        return interaction.editReply(
          rosterRecord
            ? `❌ <@${user.id}> is on the **Roster**, not **Cadets**.`
            : `❌ <@${user.id}> was not found on **Cadets**.`
        );
      }

      cadet = cadetRecord;

      cadetBackup =
        await backupCadetRow(
          cadet.rowNumber
        );

      const byDiscord =
        await findCtEntriesByDiscordId(
          user.id
        );

      if (byDiscord.length > 1) {
        return interaction.editReply(
          '❌ This Discord ID appears against more than one CT number. Correct CT Numbers first.'
        );
      }

      if (byDiscord.length === 1) {
        ctEntry = byDiscord[0];
      } else {
        const byName =
          await findCtEntriesByName(
            cadet.name
          );

        if (byName.length > 1) {
          return interaction.editReply(
            `❌ **${cadet.name}** appears against more than one CT number.`
          );
        }

        ctEntry = byName[0] || null;
      }

      if (
        ctEntry?.discordId &&
        ctEntry.discordId !== user.id
      ) {
        return interaction.editReply(
          '❌ That CT record is linked to a different Discord account. Correct CT Numbers first.'
        );
      }

      if (action === 'release') {
        if (!ctEntry) {
          return interaction.editReply(
            `❌ No CT record was found for **${cadet.name}**.`
          );
        }

        if (
          !localNumber(ctEntry.number)
        ) {
          return interaction.editReply(
            `❌ **${ctEntry.number}** is a custom/mother-group number and cannot be released.`
          );
        }
      }

      activityRow =
        await findGameActivityCadetRow(
          cadet.name
        );

      discordChanged = true;

      await member.roles.remove(
        [
          process.env.CADET_ROLE_ID,

          process.env
            .TRAINING_PLATOON_ROLE_ID,

          process.env
            .INFANTRY_ROLE_ID,

          process.env
            .UNIT_DIVIDER_ROLE_ID,
        ],

        `Removed by ${interaction.user.tag}: ${reason}`
      );

      if (restoreCommunity) {
        await member.roles.add(
          process.env
            .COMMUNITY_MEMBER_ROLE_ID,

          `Removed by ${interaction.user.tag}`
        );
      }

      await member.setNickname(
        null,

        `Removed by ${interaction.user.tag}`
      );

      await clearCadetRow({
        sheetName:
          cadet.sheetName,

        clearRange:
          `A${cadet.rowNumber}:K${cadet.rowNumber}`,
      });

      cadetCleared = true;

      if (ctEntry) {
        ctBackup = {
          ign: ctEntry.ign,

          discordId:
            ctEntry.discordId,
        };

        if (action === 'release') {
          await clearCtIdentity(
            ctEntry
          );

          ctChanged = true;
        } else if (
          ctEntry.discordId !==
            user.id ||
          normalise(ctEntry.ign) !==
            normalise(cadet.name)
        ) {
          await updateCtIdentity(
            ctEntry,
            cadet.name,
            user.id
          );

          ctChanged = true;
        }
      }

      if (activityRow) {
        await deleteGameActivityRow(
          activityRow
        );
      }

      finished = true;

      const ctResult = !ctEntry
        ? 'No CT record found; nothing changed'
        : action === 'release'
          ? `${ctEntry.number} released; IGN and Discord ID cleared`
          : `${ctEntry.number} kept with IGN and Discord ID`;

      const activityResult =
        activityRow
          ? `row ${activityRow.rowNumber} removed`
          : 'no matching row found';

      await interaction.editReply(
        [
          `✅ Removed <@${user.id}> as cadet **${cadet.name}**.`,

          `**Cadets tab:** row ${cadet.rowNumber} cleared`,

          `**Game Activity:** ${activityResult}`,

          `**CT number:** ${ctResult}`,

          `**Community Member restored:** ${
            restoreCommunity
              ? 'Yes'
              : 'No'
          }`,

          '**Nickname:** cleared',

          `**Reason:** ${reason}`,
        ].join('\n')
      );

      await sendLog(
        interaction.guild,

        process.env.LOG_CHANNEL_ID,

        [
          '**[ORION — REMOVE CADET]**',

          `**Cadet:** <@${user.id}> (${user.id})`,

          `**Discord username:** ${user.username}`,

          `**IGN:** ${cadet.name}`,

          `**CT result:** ${ctResult}`,

          `**Reason:** ${reason}`,

          `**Removed by:** ${interaction.user.tag} (${interaction.user.id})`,

          `**Owner bypass used:** ${owner ? 'Yes' : 'No'}`,
        ].join('\n')
      ).catch(error =>
        console.error(
          'Removal logging failed:',
          error
        )
      );
    } catch (error) {
      console.error(
        '/remove-cadet failed:',
        error
      );

      if (!finished) {
        if (
          ctChanged &&
          ctEntry &&
          ctBackup
        ) {
          const rollbackAction =
            !ctBackup.ign &&
            !ctBackup.discordId
              ? clearCtIdentity(
                  ctEntry
                )
              : updateCtIdentity(
                  ctEntry,
                  ctBackup.ign,

                  ctBackup.discordId
                );

          await rollbackAction.catch(
            rollbackError =>
              console.error(
                'CT rollback failed:',
                rollbackError
              )
          );
        }

        if (
          cadetCleared &&
          cadetBackup
        ) {
          await restoreCadetRow(
            cadetBackup
          ).catch(
            rollbackError =>
              console.error(
                'Cadet rollback failed:',
                rollbackError
              )
          );
        }

        if (discordChanged) {
          await restoreDiscord(
            member,
            original
          );
        }
      }

      await interaction.editReply(
        `❌ Cadet removal failed: ${
          error.message ||
          'Unknown error'
        }`
      );
    }
  },
};