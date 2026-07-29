const {
  SlashCommandBuilder,
} = require('discord.js');

const {
  sendLog,
} = require('../logger');

const {
  isBotOwner,
} = require('../permissions');

const {
  findCtEntriesByDiscordId,
  findCtEntriesByName,
} = require('../sheets');

const {
  ROLE_IDS,
  RANKS,
  DESTINATIONS,
  RANK_ROLE_IDS,
  UNIT_ROLE_IDS,
} = require('../rank-config');

const {
  getCadetRecordByDiscordId,
  getRosterRecordByDiscordId,

  graduateCadetSheets,
  rollbackCadetGraduation,

  returnRosterMemberToCadetSheets,
  rollbackReturnToCadet,

  changeExistingRosterSheets,
  rollbackExistingRosterChange,
} = require('../rank-sheets');

const ALL_MANAGED_ROLE_IDS = [
  ...RANK_ROLE_IDS,

  ROLE_IDS.enlisted,
  ROLE_IDS.nco,

  ROLE_IDS.cadet,
  ROLE_IDS.trainingPlatoon,

  ROLE_IDS.battalion104th,
  ROLE_IDS.zilloPlatoon,
  ROLE_IDS.unitDivider,

  ...UNIT_ROLE_IDS,
];

function unique(values) {
  return [
    ...new Set(
      values.filter(Boolean)
    ),
  ];
}

function rankChoices() {
  return Object.values(
    RANKS
  ).map(
    rank => ({
      name:
        rank.label,

      value:
        rank.key,
    })
  );
}

function destinationChoices() {
  return Object.values(
    DESTINATIONS
  ).map(
    destination => ({
      name:
        destination.label,

      value:
        destination.key,
    })
  );
}

async function resolveCtRecord(
  discordId,
  name
) {
  const byDiscordId =
    await findCtEntriesByDiscordId(
      discordId
    );

  if (
    byDiscordId.length > 1
  ) {
    throw new Error(
      'This Discord ID appears against more than one CT number. Correct CT Numbers first.'
    );
  }

  if (
    byDiscordId.length === 1
  ) {
    return byDiscordId[0];
  }

  const byName =
    await findCtEntriesByName(
      name
    );

  if (
    byName.length > 1
  ) {
    throw new Error(
      `${name} appears against more than one CT number. Correct CT Numbers first.`
    );
  }

  if (!byName.length) {
    throw new Error(
      `No CT-number record was found for ${name}. Run /ct-audit first.`
    );
  }

  const record =
    byName[0];

  if (
    record.discordId &&
    record.discordId !==
      discordId
  ) {
    throw new Error(
      `The CT-number record for ${name} is linked to a different Discord account.`
    );
  }

  return record;
}

async function validateDiscordChanges({
  guild,
  member,
  rank,
  destination,
  manageDestinationRoles,
}) {
  if (!member.manageable) {
    throw new Error(
      'Orion cannot manage this member. Move Orion above their highest role.'
    );
  }

  const guildRoles =
    await guild.roles.fetch();

  const returningToCadet =
    rank.key === 'cdt';

  const shouldManageUnits =
    returningToCadet ||
    manageDestinationRoles;

  const desiredRoleIds =
    returningToCadet
      ? unique([
          ROLE_IDS.cadet,
          ROLE_IDS.trainingPlatoon,
          ROLE_IDS.infantry,
          ROLE_IDS.unitDivider,
        ])
      : unique([
          rank.roleId,

          rank.breaker === 'nco'
            ? ROLE_IDS.nco
            : ROLE_IDS.enlisted,

          ROLE_IDS.battalion104th,
          ROLE_IDS.zilloPlatoon,
          ROLE_IDS.unitDivider,

          ...(
            manageDestinationRoles
              ? destination.roleIds
              : []
          ),
        ]);

  const rolesToRemoveFromPool =
    unique([
      ...RANK_ROLE_IDS,

      ROLE_IDS.enlisted,
      ROLE_IDS.nco,

      ROLE_IDS.cadet,
      ROLE_IDS.trainingPlatoon,

      ROLE_IDS.battalion104th,
      ROLE_IDS.zilloPlatoon,
      ROLE_IDS.unitDivider,

      ...(
        shouldManageUnits
          ? UNIT_ROLE_IDS
          : []
      ),
    ]);

  const toAdd =
    desiredRoleIds.filter(
      roleId =>
        !member.roles.cache.has(
          roleId
        )
    );

  const toRemove =
    rolesToRemoveFromPool.filter(
      roleId =>
        member.roles.cache.has(
          roleId
        ) &&
        !desiredRoleIds.includes(
          roleId
        )
    );

  for (
    const roleId of
    unique([
      ...toAdd,
      ...toRemove,
    ])
  ) {
    const role =
      guildRoles.get(
        roleId
      );

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

  return {
    toAdd,
    toRemove,
  };
}

async function applyDiscordChanges({
  member,
  toAdd,
  toRemove,
  nickname,
  reason,
}) {
  if (toRemove.length) {
    await member.roles.remove(
      toRemove,
      reason
    );
  }

  if (toAdd.length) {
    await member.roles.add(
      toAdd,
      reason
    );
  }

  await member.setNickname(
    nickname,
    reason
  );
}

async function restoreDiscord(
  member,
  snapshot
) {
  try {
    const currentManaged =
      new Set(
        ALL_MANAGED_ROLE_IDS.filter(
          roleId =>
            member.roles.cache.has(
              roleId
            )
        )
      );

    const toAdd = [
      ...snapshot.managedRoles,
    ].filter(
      roleId =>
        !currentManaged.has(
          roleId
        )
    );

    const toRemove = [
      ...currentManaged,
    ].filter(
      roleId =>
        !snapshot.managedRoles.has(
          roleId
        )
    );

    if (
      toRemove.length
    ) {
      await member.roles.remove(
        toRemove,
        'Orion setrank rollback'
      );
    }

    if (
      toAdd.length
    ) {
      await member.roles.add(
        toAdd,
        'Orion setrank rollback'
      );
    }

    if (
      member.nickname !==
      snapshot.nickname
    ) {
      await member.setNickname(
        snapshot.nickname,
        'Orion setrank rollback'
      );
    }
  } catch (error) {
    console.error(
      'Discord setrank rollback failed:',
      error
    );
  }
}

module.exports = {
  data:
    new SlashCommandBuilder()
      .setName('setrank')

      .setDescription(
        'Graduate a cadet, return someone to cadet, or change a roster member’s rank'
      )

      .addUserOption(option =>
        option
          .setName('member')

          .setDescription(
            'The member whose rank should be changed'
          )

          .setRequired(true)
      )

      .addStringOption(option =>
        option
          .setName('rank')

          .setDescription(
            'The member’s new rank'
          )

          .setRequired(true)

          .addChoices(
            ...rankChoices()
          )
      )

      .addStringOption(option =>
        option
          .setName(
            'destination'
          )

          .setDescription(
            'Optional new squad or vehicle section; leave blank to keep their current destination'
          )

          .setRequired(false)

          .addChoices(
            ...destinationChoices()
          )
      )

      .addStringOption(option =>
        option
          .setName('reason')

          .setDescription(
            'Optional promotion, demotion or movement reason'
          )

          .setMaxLength(500)

          .setRequired(false)
      ),

  async execute(interaction) {
    const owner =
      isBotOwner(
        interaction
      );

    if (
      !owner &&
      interaction.channelId !==
        process.env
          .ACCEPT_CADET_CHANNEL_ID
    ) {
      return interaction.reply({
        content:
          '❌ `/setrank` can only be used in the configured management channel.',

        ephemeral:
          true,
      });
    }

    if (
      !owner &&
      !interaction.member
        .roles.cache.has(
          process.env
            .OFFICER_ROLE_ID
        )
    ) {
      return interaction.reply({
        content:
          '❌ Only Officers can use `/setrank`.',

        ephemeral:
          true,
      });
    }

    await interaction.deferReply({
      ephemeral:
        true,
    });

    const user =
      interaction.options.getUser(
        'member',
        true
      );

    const rankKey =
      interaction.options.getString(
        'rank',
        true
      );

    const selectedDestinationKey =
      interaction.options.getString(
        'destination'
      ) ||
      null;

    const reason =
      interaction.options
        .getString('reason')
        ?.trim() ||
      'No reason supplied';

    if (user.bot) {
      return interaction.editReply(
        '❌ A bot cannot be assigned a rank.'
      );
    }

    const rank =
      RANKS[
        rankKey
      ];

    if (!rank) {
      return interaction.editReply(
        '❌ That rank is not configured.'
      );
    }

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

    const discordSnapshot = {
      nickname:
        member.nickname,

      managedRoles:
        new Set(
          ALL_MANAGED_ROLE_IDS.filter(
            roleId =>
              member.roles.cache.has(
                roleId
              )
          )
        ),
    };

    let sheetToken =
      null;

    let sheetChangeType =
      null;

    let discordChanged =
      false;

    let operationCompleted =
      false;

    try {
      const [
        cadet,
        roster,
      ] =
        await Promise.all([
          getCadetRecordByDiscordId(
            user.id
          ),

          getRosterRecordByDiscordId(
            user.id
          ),
        ]);

      if (
        cadet &&
        roster
      ) {
        throw new Error(
          'This Discord ID appears on both Cadets and Roster. Correct the spreadsheet before continuing.'
        );
      }

      if (
        !cadet &&
        !roster
      ) {
        throw new Error(
          'This member was not found on either Cadets or Roster.'
        );
      }

      const sourceName =
        cadet?.name ||
        roster.name;

      const ctRecord =
        await resolveCtRecord(
          user.id,
          sourceName
        );

      const ctNumber =
        String(
          ctRecord.number ||
          ''
        ).trim();

      if (!ctNumber) {
        throw new Error(
          `The CT-number record for ${sourceName} is blank.`
        );
      }

      let destinationKey =
        selectedDestinationKey;

      let manageDestinationRoles =
        Boolean(
          selectedDestinationKey
        );

      let sourceType;
      let oldRank;
      let oldDestination;

      if (cadet) {
        sourceType =
          'Cadet graduation';

        oldRank =
          'CDT';

        oldDestination =
          'Cadets';

        if (
          rankKey === 'cdt'
        ) {
          throw new Error(
            'This member is already a Cadet.'
          );
        }

        if (
          rankKey !== 'ct'
        ) {
          throw new Error(
            'A Cadet must first graduate to CT. They cannot be promoted directly to a higher rank.'
          );
        }

        if (
          selectedDestinationKey &&
          selectedDestinationKey !==
            'reserves'
        ) {
          throw new Error(
            'New CT graduates must be placed in Reserves. Leave destination blank or select Reserves.'
          );
        }

        destinationKey =
          'reserves';

        manageDestinationRoles =
          true;
      } else if (
        rankKey === 'cdt'
      ) {
        sourceType =
          'Returned to Cadet';

        oldRank =
          roster.rank ||
          'Unknown';

        oldDestination =
          roster.sectionKey
            ? (
                DESTINATIONS[
                  roster.sectionKey
                ]?.label ||
                roster.sectionKey
              )
            : 'Unrecognised section';

        if (
          selectedDestinationKey
        ) {
          throw new Error(
            'Do not select a destination when returning someone to Cadet.'
          );
        }

        destinationKey =
          null;

        manageDestinationRoles =
          true;
      } else {
        sourceType =
          'Roster rank change';

        oldRank =
          roster.rank ||
          'Unknown';

        oldDestination =
          roster.sectionKey
            ? (
                DESTINATIONS[
                  roster.sectionKey
                ]?.label ||
                roster.sectionKey
              )
            : 'Unchanged / unrecognised section';
      }

      const destination =
        destinationKey
          ? DESTINATIONS[
              destinationKey
            ]
          : null;

      if (
        destinationKey &&
        !destination
      ) {
        throw new Error(
          'That destination is not configured.'
        );
      }

      const nickname =
        `${rank.abbreviation} ${sourceName}-${ctNumber}`;

      if (
        nickname.length > 32
      ) {
        throw new Error(
          `The nickname “${nickname}” is longer than Discord’s 32-character limit.`
        );
      }

      const rolePlan =
        await validateDiscordChanges({
          guild:
            interaction.guild,

          member,

          rank,

          destination,

          manageDestinationRoles,
        });

      if (cadet) {
        sheetToken =
          await graduateCadetSheets({
            cadet,

            rankAbbreviation:
              rank.abbreviation,
          });

        sheetChangeType =
          'graduation';
      } else if (
        rankKey === 'cdt'
      ) {
        sheetToken =
          await returnRosterMemberToCadetSheets({
            roster,
          });

        sheetChangeType =
          'return-to-cadet';
      } else {
        sheetToken =
          await changeExistingRosterSheets({
            roster,

            rankAbbreviation:
              rank.abbreviation,

            destinationKey,
          });

        sheetChangeType =
          'existing';
      }

      discordChanged =
        true;

      await applyDiscordChanges({
        member,

        toAdd:
          rolePlan.toAdd,

        toRemove:
          rolePlan.toRemove,

        nickname,

        reason:
          `Setrank by ${interaction.user.tag}: ${reason}`,
      });

      operationCompleted =
        true;

      const destinationText =
        cadet
          ? DESTINATIONS
              .reserves.label
          : rankKey === 'cdt'
            ? 'Cadets'
            : selectedDestinationKey
              ? destination.label
              : 'Unchanged';

      await interaction.editReply(
        [
          `✅ Updated <@${user.id}> to **${rank.abbreviation}**.`,

          `**IGN:** ${sourceName}`,

          `**CT number:** ${ctNumber}`,

          `**Previous rank:** ${oldRank}`,

          `**Previous destination:** ${oldDestination}`,

          `**New destination:** ${destinationText}`,

          `**Nickname:** ${nickname}`,

          `**Reason:** ${reason}`,
        ].join('\n')
      );

      await sendLog(
        interaction.guild,

        process.env
          .LOG_CHANNEL_ID,

        [
          '**[ORION — SETRANK]**',

          `**Member:** <@${user.id}> (${user.id})`,

          `**Discord username:** ${user.username}`,

          `**IGN:** ${sourceName}`,

          `**CT number:** ${ctNumber}`,

          `**Action:** ${sourceType}`,

          `**Previous rank:** ${oldRank}`,

          `**New rank:** ${rank.abbreviation}`,

          `**Previous destination:** ${oldDestination}`,

          `**New destination:** ${destinationText}`,

          `**Reason:** ${reason}`,

          `**Changed by:** ${interaction.user.tag} (${interaction.user.id})`,

          `**Owner bypass used:** ${
            owner
              ? 'Yes'
              : 'No'
          }`,
        ].join('\n')
      ).catch(
        error =>
          console.error(
            'Setrank logging failed:',
            error
          )
      );
    } catch (error) {
      console.error(
        '/setrank failed:',
        error
      );

      if (
        operationCompleted
      ) {
        console.error(
          'Setrank completed successfully, but the response failed after completion.'
        );
      } else if (
        discordChanged
      ) {
        await restoreDiscord(
          member,
          discordSnapshot
        );
      }

      if (
        !operationCompleted &&
        sheetToken
      ) {
        if (
          sheetChangeType ===
          'graduation'
        ) {
          await rollbackCadetGraduation(
            sheetToken
          ).catch(
            rollbackError =>
              console.error(
                'Graduation rollback failed:',
                rollbackError
              )
          );
        } else if (
          sheetChangeType ===
          'return-to-cadet'
        ) {
          await rollbackReturnToCadet(
            sheetToken
          ).catch(
            rollbackError =>
              console.error(
                'Return-to-Cadet rollback failed:',
                rollbackError
              )
          );
        } else {
          await rollbackExistingRosterChange(
            sheetToken
          ).catch(
            rollbackError =>
              console.error(
                'Roster change rollback failed:',
                rollbackError
              )
          );
        }
      }

      if (
        !operationCompleted
      ) {
        await interaction.editReply(
          `❌ Setrank failed: ${
            error.message ||
            'Unknown error'
          }`
        );
      }
    }
  },
};