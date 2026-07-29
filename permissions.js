require('dotenv').config();

/**
 * Returns true when the command user is
 * the configured Orion bot owner.
 */
function isBotOwner(interactionOrUser) {
  const userId =
    interactionOrUser?.user?.id ||
    interactionOrUser?.id ||
    '';

  const ownerId = String(
    process.env.BOT_OWNER_ID || ''
  ).trim();

  return Boolean(
    ownerId &&
    String(userId) === ownerId
  );
}

/**
 * Returns true when the member has at
 * least one of the supplied role IDs.
 */
function hasAnyRole(member, roleIds) {
  if (!member?.roles?.cache) {
    return false;
  }

  return roleIds
    .filter(Boolean)
    .some(roleId =>
      member.roles.cache.has(roleId)
    );
}

module.exports = {
  isBotOwner,
  hasAnyRole,
};