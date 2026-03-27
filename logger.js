async function sendLog(guild, channelId, message) {
  if (!channelId) return;

  try {
    const channel = await guild.channels.fetch(channelId);
    if (!channel) return;

    await channel.send({
      content: message,
      allowedMentions: { users: [] },
    });
  } catch (error) {
    console.error('Failed to send log message:', error);
  }
}

module.exports = { sendLog };