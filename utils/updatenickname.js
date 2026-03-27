const knownRankPrefixes = [
  'TR',
  'PVT',
  'LCPL',
  'CPL',
  'SGT',
  'SSGT',
  'CDT',
  'TPR',
  'O/CDT',
  'P/O',
  'F/O',
];

async function updateNickname(member, options) {
  if (!member.manageable) return false;

  const { prefix = null, exSkira = false, exactName = null } = options;

  let updatedName;

  if (exactName) {
    updatedName = exactName.trim();
  } else {
    const currentName = member.nickname || member.user.displayName || member.user.username;
    updatedName = currentName.trim();

    updatedName = updatedName.replace(/\s*\(Ex Skira\)$/i, '');

    const prefixPattern = new RegExp(
      `^(${knownRankPrefixes.map(p => p.replace('/', '\\/')).join('|')})\\s+`,
      'i'
    );

    if (prefix) {
      if (prefixPattern.test(updatedName)) {
        updatedName = updatedName.replace(prefixPattern, `${prefix} `);
      } else {
        updatedName = `${prefix} ${updatedName}`;
      }
    } else {
      updatedName = updatedName.replace(prefixPattern, '');
    }

    if (exSkira) {
      updatedName = `${updatedName} (Ex Skira)`;
    }
  }

  if (updatedName.length > 32) {
    updatedName = updatedName.slice(0, 32);
  }

  await member.setNickname(updatedName, 'Rank nickname update');
  return true;
}

module.exports = { updateNickname };