const HQ_CHANNEL_ID = process.env.HQ_CHANNEL_ID;
const HQ_ROLE_ID = process.env.HQ_ROLE_ID;

// Users who can use HQ commands without needing the HQ role.
// They STILL have to use the commands inside the HQ channel.
const HQ_USER_OVERRIDES = [
  '207548782588461058', // PixelPonny
];

function canUseHqCommand(interaction) {
  if (!HQ_CHANNEL_ID) {
    return false;
  }

  // Everyone, including overrides, must be inside the HQ channel.
  if (interaction.channelId !== HQ_CHANNEL_ID) {
    return false;
  }

  const hasHqRole =
    HQ_ROLE_ID &&
    interaction.member?.roles?.cache?.has(HQ_ROLE_ID);

  const isUserOverride =
    HQ_USER_OVERRIDES.includes(interaction.user.id);

  return Boolean(hasHqRole || isUserOverride);
}

module.exports = {
  canUseHqCommand,
};