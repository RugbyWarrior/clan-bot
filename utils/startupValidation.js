const REQUIRED_DISCORD_IDS = [
  'GUILD_ID',
  'ACCEPT_CADET_CHANNEL_ID',
  'LOG_CHANNEL_ID',
  'OFFICER_ROLE_ID',
  'CADET_ROLE_ID',
  'TRAINING_PLATOON_ROLE_ID',
  'INFANTRY_ROLE_ID',
  'COMMUNITY_MEMBER_ROLE_ID',
];

const REQUIRED_TEXT = [
  'DISCORD_TOKEN',
  'CLIENT_ID',
  'GOOGLE_SPREADSHEET_ID',
  'CADETS_SHEET_NAME',
  'CT_NUMBERS_SHEET_NAME',
  'GAME_ACTIVITY_SHEET_NAME',
  'ROSTER_SHEET_NAME',
  'GOOGLE_CREDENTIALS_JSON',
];

function hasValue(key) {
  return Boolean(String(process.env[key] || '').trim());
}

function validateGoogleCredentials(errors) {
  if (!hasValue('GOOGLE_CREDENTIALS_JSON')) return;

  try {
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    const requiredKeys = ['type', 'project_id', 'private_key', 'client_email', 'token_uri'];
    const missing = requiredKeys.filter(key => !credentials[key]);

    if (missing.length > 0) {
      errors.push(`GOOGLE_CREDENTIALS_JSON is missing: ${missing.join(', ')}`);
    }
  } catch (error) {
    errors.push(`GOOGLE_CREDENTIALS_JSON is not valid JSON: ${error.message}`);
  }
}

function validateStartupConfig() {
  const errors = [];

  for (const key of REQUIRED_TEXT) {
    if (!hasValue(key)) errors.push(`Missing required .env value: ${key}`);
  }

  for (const key of REQUIRED_DISCORD_IDS) {
    const value = String(process.env[key] || '').trim();

    if (!value) {
      errors.push(`Missing required Discord ID: ${key}`);
    } else if (!/^\d{17,20}$/.test(value)) {
      errors.push(`Invalid Discord ID in ${key}: ${value}`);
    }
  }

  const clientId = String(process.env.CLIENT_ID || '').trim();
  if (clientId && !/^\d{17,20}$/.test(clientId)) {
    errors.push(`Invalid Discord application ID in CLIENT_ID: ${clientId}`);
  }

  validateGoogleCredentials(errors);

  return {
    status: errors.length > 0 ? 'invalid' : 'valid',
    errors,
  };
}

function printStartupValidationReport() {
  const result = validateStartupConfig();

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('Orion startup configuration check');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  if (result.errors.length > 0) {
    for (const error of result.errors) {
      console.error(`- ${error}`);
    }
    console.error('Startup stopped until the .env file is completed.');
  } else {
    console.log('Configuration looks valid.');
  }

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  return result;
}

module.exports = {
  validateStartupConfig,
  printStartupValidationReport,
};
