const { SlashCommandBuilder } = require('discord.js');
const { sendLog } = require('../logger');
const { isBotOwner, hasAnyRole } = require('../permissions');

const {
  normalizeName,
  findCadetByDiscordId,
  findRosterByDiscordId,
  findNameOccurrencesInCadets,
  findNameOccurrencesInRoster,
  findCtEntriesByName,
  findCtEntriesByNumber,
  findCtEntriesByDiscordId,
  findNextAvailableLocalCtNumber,
  createCustomCtEntry,
  updateCtIdentity,
  clearCtIdentity,
  clearCustomCtEntry,
  addCadetRow,
  clearCadetRow,
  addGameActivityCadetRow,
  deleteGameActivityRow,
} = require('../sheets');

const TIMEZONES = [
  'GMT',
  'CET',
  'EET',
  'MSK',
  'EST',
  'CST',
  'MST',
  'PST',
  'IST',
  'AKST',
  'AEST',
  'ACST',
  'AWST',
  'NZST',
  'GST',
  'JST',
  'UTC',
  'ICT',
  'PHT',
  'HAST',
];

const dateUK = date =>
  `${String(date.getDate()).padStart(2, '0')}/` +
  `${String(date.getMonth() + 1).padStart(2, '0')}/` +
  `${date.getFullYear()}`;

const validCt = value =>
  /^(?:\d{2}-\d{3}|\d{5}|\d{4})$/.test(
    String(value || '').trim()
  );

const locations = matches =>
  matches
    .map(
      match =>
        `${match.sheetName} row ${match.rowNumber}`
    )
    .join(', ');

async function checkDiscord(member, guild) {
  if (!member.manageable) {
    throw new Error(
      'Orion cannot manage this member. Move Orion above their highest role.'
    );
  }

  const roleIds = [
    process.env.CADET_ROLE_ID,
    process.env.TRAINING_PLATOON_ROLE_ID,
    process.env.INFANTRY_ROLE_ID,
    process.env.UNIT_DIVIDER_ROLE_ID,
    process.env.COMMUNITY_MEMBER_ROLE_ID,
  ];

  for (const roleId of roleIds) {
    if (!roleId) {
      throw new Error(
        'A required Discord role ID is missing from .env.'
      );
    }

    const role = await guild.roles.fetch(roleId);

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

async function restoreDiscord(member, original) {
  try {
    const roleIds = [
      process.env.CADET_ROLE_ID,
      process.env.TRAINING_PLATOON_ROLE_ID,
      process.env.INFANTRY_ROLE_ID,
      process.env.UNIT_DIVIDER_ROLE_ID,
      process.env.COMMUNITY_MEMBER_ROLE_ID,
    ].filter(Boolean);

    for (const roleId of roleIds) {
      const shouldHave =
        original.roles.has(roleId);

      const hasRole =
        member.roles.cache.has(roleId);

      if (shouldHave && !hasRole) {
        await member.roles.add(
          roleId,
          'Orion rollback'
        );
      }

      if (!shouldHave && hasRole) {
        await member.roles.remove(
          roleId,
          'Orion rollback'
        );
      }
    }

    if (
      member.nickname !== original.nickname
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
    .setName('accept-cadet')

    .setDescription(
      'Accept a cadet and create their Discord and spreadsheet records'
    )

    .addUserOption(option =>
      option
        .setName('cadet')
        .setDescription(
          'The Discord member'
        )
        .setRequired(true)
    )

    .addStringOption(option =>
      option
        .setName('in_game_name')
        .setDescription(
          'The cadet’s in-game name'
        )
        .setMinLength(1)
        .setMaxLength(22)
        .setRequired(true)
    )

    .addStringOption(option => {
      option
        .setName('timezone')
        .setDescription(
          'The cadet’s timezone'
        )
        .setRequired(true);

      for (
        const timezone of TIMEZONES
      ) {
        option.addChoices({
          name: timezone,
          value: timezone,
        });
      }

      return option;
    })

    .addStringOption(option =>
      option
        .setName('ct_origin')
        .setDescription(
          'Reuse a mother-group number or assign a local number'
        )
        .setRequired(true)

        .addChoices(
          {
            name:
              'Mother group / existing CT number',

            value: 'existing',
          },

          {
            name:
              'New member / assign local CT number',

            value: 'new',
          }
        )
    )

    .addStringOption(option =>
      option
        .setName(
          'existing_ct_number'
        )

        .setDescription(
          'For example 23-879, 9174 or 53103'
        )

        .setRequired(false)
    )

    .addBooleanOption(option =>
      option
        .setName(
          'allow_existing_ign'
        )

        .setDescription(
          'Override an existing-IGN warning'
        )

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
          '❌ `/accept-cadet` can only be used in the configured cadet channel.',

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
          '❌ Only Officers and NCOs can use `/accept-cadet`.',

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

    const ign =
      interaction.options
        .getString(
          'in_game_name',
          true
        )
        .trim();

    const timezone =
      interaction.options.getString(
        'timezone',
        true
      );

    const origin =
      interaction.options.getString(
        'ct_origin',
        true
      );

    const supplied =
      interaction.options
        .getString(
          'existing_ct_number'
        )
        ?.trim() || null;

    const overrideIgn =
      interaction.options.getBoolean(
        'allow_existing_ign'
      ) || false;

    if (user.bot) {
      return interaction.editReply(
        '❌ A bot cannot be accepted as a cadet.'
      );
    }

    if (
      origin === 'existing' &&
      !supplied
    ) {
      return interaction.editReply(
        '❌ `existing_ct_number` is required for a mother-group cadet.'
      );
    }

    if (
      supplied &&
      !validCt(supplied)
    ) {
      return interaction.editReply(
        '❌ Use a CT number such as `23-879`, `9174` or `53103`.'
      );
    }

    const member =
      await interaction.guild.members.fetch(
        user.id
      );

    const original = {
      nickname: member.nickname,

      roles: new Set(
        member.roles.cache.keys()
      ),
    };

    let ctEntry;
    let ctRollback;
    let cadetRow;
    let activityRow;

    let discordChanged = false;
    let finished = false;

    try {
      await checkDiscord(
        member,
        interaction.guild
      );

      const [
        activeCadet,
        rosterMember,
      ] = await Promise.all([
        findCadetByDiscordId(user.id),

        findRosterByDiscordId(
          user.id
        ),
      ]);

      if (
        activeCadet ||
        rosterMember
      ) {
        const record =
          activeCadet ||
          rosterMember;

        return interaction.editReply(
          `❌ <@${user.id}> is already on **${record.sheetName}**, row **${record.rowNumber}**, as **${record.name || 'Unknown'}**.`
        );
      }

      const [
        cadetNames,
        rosterNames,
        ctNames,
        ctDiscord,
      ] = await Promise.all([
        findNameOccurrencesInCadets(
          ign
        ),

        findNameOccurrencesInRoster(
          ign
        ),

        findCtEntriesByName(ign),

        findCtEntriesByDiscordId(
          user.id
        ),
      ]);

      if (ctDiscord.length > 1) {
        return interaction.editReply(
          '❌ This Discord ID appears against more than one CT number. Correct the CT Numbers tab first.'
        );
      }

      if (ctNames.length > 1) {
        return interaction.editReply(
          `❌ **${ign}** appears against more than one CT number. Correct the duplicates first.`
        );
      }

      const duplicateNames = [
        ...cadetNames,
        ...rosterNames,

        ...ctNames.filter(
          entry =>
            entry.discordId !==
            user.id
        ),
      ];

      if (
        duplicateNames.length &&
        !overrideIgn
      ) {
        return interaction.editReply(
          [
            `⚠️ The IGN **${ign}** already exists at **${locations(duplicateNames)}**.`,

            'Use another IGN, or rerun with `allow_existing_ign: True` when this is intentional.',
          ].join('\n')
        );
      }

      let ctNumber;
      let ctSource;

      if (ctDiscord.length === 1) {
        ctEntry = ctDiscord[0];

        if (
          origin === 'existing' &&
          supplied &&
          supplied !== ctEntry.number
        ) {
          return interaction.editReply(
            `❌ This Discord account already owns **${ctEntry.number}**, not **${supplied}**.`
          );
        }

        ctRollback = {
          type: 'restore',
          entry: ctEntry,
          ign: ctEntry.ign,

          discordId:
            ctEntry.discordId,
        };

        await updateCtIdentity(
          ctEntry,
          ign,
          user.id
        );

        ctNumber = ctEntry.number;

        ctSource =
          'Existing Discord ID mapping reused';
      } else if (
        origin === 'existing'
      ) {
        const numberMatches =
          await findCtEntriesByNumber(
            supplied
          );

        if (
          numberMatches.length > 1
        ) {
          return interaction.editReply(
            `❌ **${supplied}** appears more than once on CT Numbers.`
          );
        }

        if (!numberMatches.length) {
          ctEntry =
            await createCustomCtEntry(
              ign,
              supplied,
              user.id
            );

          ctRollback = {
            type: 'delete',
            entry: ctEntry,
          };

          ctSource =
            'Mother group / existing, added to custom list';
        } else {
          ctEntry =
            numberMatches[0];

          if (
            ctEntry.ign &&
            normalizeName(
              ctEntry.ign
            ) !==
              normalizeName(ign)
          ) {
            return interaction.editReply(
              `❌ **${supplied}** is already assigned to **${ctEntry.ign}**.`
            );
          }

          if (
            ctEntry.discordId &&
            ctEntry.discordId !==
              user.id
          ) {
            return interaction.editReply(
              `❌ **${supplied}** is linked to a different Discord account.`
            );
          }

          ctRollback = {
            type: 'restore',
            entry: ctEntry,
            ign: ctEntry.ign,

            discordId:
              ctEntry.discordId,
          };

          await updateCtIdentity(
            ctEntry,
            ign,
            user.id
          );

          ctSource =
            'Mother group / existing';
        }

        ctNumber = supplied;
      } else if (
        ctNames.length === 1 &&
        overrideIgn
      ) {
        ctEntry = ctNames[0];

        if (
          ctEntry.discordId &&
          ctEntry.discordId !==
            user.id
        ) {
          return interaction.editReply(
            `❌ **${ign}** is linked to a different Discord account.`
          );
        }

        ctRollback = {
          type: 'restore',
          entry: ctEntry,
          ign: ctEntry.ign,

          discordId:
            ctEntry.discordId,
        };

        await updateCtIdentity(
          ctEntry,
          ign,
          user.id
        );

        ctNumber = ctEntry.number;

        ctSource =
          'Existing IGN mapping reused';
      } else {
        ctEntry =
          await findNextAvailableLocalCtNumber();

        ctRollback = {
          type: 'restore',
          entry: ctEntry,
          ign: ctEntry.ign,

          discordId:
            ctEntry.discordId,
        };

        await updateCtIdentity(
          ctEntry,
          ign,
          user.id
        );

        ctNumber = ctEntry.number;

        ctSource =
          'Local allocation';
      }

      const nickname =
        `CDT ${ign}-${ctNumber}`;

      if (nickname.length > 32) {
        throw new Error(
          'The resulting Discord nickname is longer than 32 characters.'
        );
      }

      cadetRow =
        await addCadetRow({
          inGameName: ign,
          timezone,

          discordUsername:
            user.username,

          discordId: user.id,

          joinedDate:
            dateUK(new Date()),
        });

      activityRow =
        await addGameActivityCadetRow(
          ign
        );

      discordChanged = true;

      await member.roles.add(
        [
          process.env.CADET_ROLE_ID,

          process.env
            .TRAINING_PLATOON_ROLE_ID,

          process.env
            .INFANTRY_ROLE_ID,

          process.env
            .UNIT_DIVIDER_ROLE_ID,
        ],

        `Accepted by ${interaction.user.tag}`
      );

      if (
        member.roles.cache.has(
          process.env
            .COMMUNITY_MEMBER_ROLE_ID
        )
      ) {
        await member.roles.remove(
          process.env
            .COMMUNITY_MEMBER_ROLE_ID,

          `Accepted by ${interaction.user.tag}`
        );
      }

      await member.setNickname(
        nickname,

        `Accepted by ${interaction.user.tag}`
      );

      finished = true;

      await interaction.editReply(
        [
          `✅ Accepted <@${user.id}> as **${ign}**.`,

          `**CT number:** ${ctNumber} (${ctSource})`,

          `**Nickname:** ${nickname}`,

          `**Cadets tab:** row ${cadetRow.rowNumber}`,

          `**Game Activity:** row ${activityRow.rowNumber}`,

          '**Roles added:** Cadet, Training Platoon, Infantry, Unit Divider',

          '**Role removed:** Community Member',
        ].join('\n')
      );

      await sendLog(
        interaction.guild,

        process.env.LOG_CHANNEL_ID,

        [
          '**[ORION — ACCEPT CADET]**',

          `**Cadet:** <@${user.id}> (${user.id})`,

          `**Discord username:** ${user.username}`,

          `**IGN:** ${ign}`,

          `**Timezone:** ${timezone}`,

          `**CT number:** ${ctNumber}`,

          `**CT source:** ${ctSource}`,

          `**Accepted by:** ${interaction.user.tag} (${interaction.user.id})`,

          `**Owner bypass used:** ${
            owner ? 'Yes' : 'No'
          }`,
        ].join('\n')
      ).catch(error =>
        console.error(
          'Acceptance logging failed:',
          error
        )
      );
    } catch (error) {
      console.error(
        '/accept-cadet failed:',
        error
      );

      if (!finished) {
        if (discordChanged) {
          await restoreDiscord(
            member,
            original
          );
        }

        if (activityRow) {
          await deleteGameActivityRow(
            activityRow
          ).catch(
            rollbackError =>
              console.error(
                'Activity rollback failed:',
                rollbackError
              )
          );
        }

        if (cadetRow) {
          await clearCadetRow(
            cadetRow
          ).catch(
            rollbackError =>
              console.error(
                'Cadet rollback failed:',
                rollbackError
              )
          );
        }

        if (
          ctRollback?.type ===
          'delete'
        ) {
          await clearCustomCtEntry(
            ctRollback.entry
          ).catch(
            rollbackError =>
              console.error(
                'CT rollback failed:',
                rollbackError
              )
          );
        } else if (
          ctRollback?.type ===
          'restore'
        ) {
          const rollbackAction =
            !ctRollback.ign &&
            !ctRollback.discordId
              ? clearCtIdentity(
                  ctRollback.entry
                )
              : updateCtIdentity(
                  ctRollback.entry,
                  ctRollback.ign,

                  ctRollback.discordId
                );

          await rollbackAction.catch(
            rollbackError =>
              console.error(
                'CT rollback failed:',
                rollbackError
              )
          );
        }
      }

      await interaction.editReply(
        `❌ Acceptance failed: ${
          error.message ||
          'Unknown error'
        }`
      );
    }
  },
};