async function sendLog(guild, channelId, message) {
  if (!guild || !channelId) return;

  try {
    const channel = await guild.channels.fetch(channelId);
    if (!channel?.isTextBased()) return;

    await channel.send({
      content: message,
      allowedMentions: { parse: [] },
    });
  } catch (error) {
    console.error('Failed to send Orion log message:', error);
  }
}

module.exports = { sendLog };
