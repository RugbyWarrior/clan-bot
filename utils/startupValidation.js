function maskValue(value, visible = 4) {
  const stringValue = String(value || '');
  if (!stringValue) return '[missing]';
  if (stringValue.length <= visible) return '*'.repeat(stringValue.length);
  return `${'*'.repeat(Math.max(stringValue.length - visible, 0))}${stringValue.slice(-visible)}`;
}

function parseCsvEnv(envKey) {
  return (process.env[envKey] || '')
    .split(',')
    .map(value => value.trim())
    .filter(Boolean);
}

function validateRequiredEnv(envKey, errors) {
  const value = process.env[envKey];
  if (!value || !String(value).trim()) {
    errors.push(`Missing required env var: ${envKey}`);
    return false;
  }
  return true;
}

function validateDiscordIdEnv(envKey, errors, warnings, required = true) {
  const value = process.env[envKey];

  if (!value || !String(value).trim()) {
    if (required) {
      errors.push(`Missing required Discord ID env var: ${envKey}`);
    } else {
      warnings.push(`Optional Discord ID env var not set: ${envKey}`);
    }
    return false;
  }

  if (!/^\d{17,20}$/.test(String(value).trim())) {
    errors.push(`Invalid Discord ID format for ${envKey}: ${value}`);
    return false;
  }

  return true;
}

function validateCsvDiscordIdsEnv(envKey, errors, warnings, required = true) {
  const values = parseCsvEnv(envKey);

  if (values.length === 0) {
    if (required) {
      errors.push(`Missing required CSV Discord ID env var: ${envKey}`);
    } else {
      warnings.push(`Optional CSV Discord ID env var not set: ${envKey}`);
    }
    return false;
  }

  const invalid = values.filter(value => !/^\d{17,20}$/.test(value));
  if (invalid.length > 0) {
    errors.push(`Invalid Discord ID(s) in ${envKey}: ${invalid.join(', ')}`);
    return false;
  }

  return true;
}

function validateGoogleCredentials(errors) {
  const raw = process.env.GOOGLE_CREDENTIALS_JSON;

  if (!raw || !String(raw).trim()) {
    errors.push('Missing required env var: GOOGLE_CREDENTIALS_JSON');
    return false;
  }

  try {
    const parsed = JSON.parse(raw);

    const requiredKeys = [
      'type',
      'project_id',
      'private_key',
      'client_email',
      'client_id',
      'token_uri',
    ];

    const missingKeys = requiredKeys.filter(key => !parsed[key]);
    if (missingKeys.length > 0) {
      errors.push(`GOOGLE_CREDENTIALS_JSON is missing key(s): ${missingKeys.join(', ')}`);
      return false;
    }

    return true;
  } catch (error) {
    errors.push(`GOOGLE_CREDENTIALS_JSON is not valid JSON: ${error.message}`);
    return false;
  }
}

function validateStartupConfig() {
  const errors = [];
  const warnings = [];

  // Critical bot config
  validateRequiredEnv('DISCORD_TOKEN', errors);
  validateRequiredEnv('CLIENT_ID', errors);
  validateRequiredEnv('GUILD_ID', errors);

  // Channels
  validateDiscordIdEnv('HQ_CHANNEL_ID', errors, warnings, true);
  validateDiscordIdEnv('LOG_CHANNEL_ID', errors, warnings, true);
  validateDiscordIdEnv('ALLOWED_CHANNEL_ID', errors, warnings, false);

  // Spreadsheets / sheets
  validateRequiredEnv('TRAINEE_SPREADSHEET_ID', errors);
  validateRequiredEnv('MOS_SPREADSHEET_ID', errors);
  validateRequiredEnv('TRAINEE_SHEET_NAME', errors);
  validateRequiredEnv('RATINGS_SHEET_NAME', errors);

  // Google
  validateGoogleCredentials(errors);

  // Core role IDs
  validateDiscordIdEnv('TRAINEE_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('TRAINING_COMPANY_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('EX_SKIRA_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('PRIVATE_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('COMMUNITY_MEMBER_ID', errors, warnings, true);

  // Squadron roles
  validateDiscordIdEnv('INFANTRY_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('ARMOUR_ROLE_ID', errors, warnings, true);
  validateDiscordIdEnv('AVIATION_ROLE_ID', errors, warnings, true);

  // Rank groupings
  validateCsvDiscordIdsEnv('RANK_ROLE_IDS', errors, warnings, true);
  validateCsvDiscordIdsEnv('BREAKER_ROLE_IDS', errors, warnings, true);
  validateCsvDiscordIdsEnv('SQUADRON_ROLE_IDS', errors, warnings, true);

  // Optional but expected rank IDs used in setrank
  const rankEnvKeys = [
    'ARMOUR_CADET_ROLE_ID',
    'OFFICER_CADET_ROLE_ID',
    'ARMOUR_TROOPER_ROLE_ID',
    'PILOT_OFFICER_ROLE_ID',
    'LANCE_CORPORAL_ROLE_ID',
    'FLYING_OFFICER_ROLE_ID',
    'CORPORAL_ROLE_ID',
    'SERGEANT_ROLE_ID',
    'STAFF_SERGEANT_ROLE_ID',
    'ENLISTED_BREAKER_ROLE_ID',
    'NCO_BREAKER_ROLE_ID',
    'SENIOR_NCO_BREAKER_ROLE_ID',
  ];

  for (const envKey of rankEnvKeys) {
    validateDiscordIdEnv(envKey, errors, warnings, true);
  }

  const summary = {
    status: errors.length > 0 ? 'invalid' : 'valid',
    errors,
    warnings,
  };

  return summary;
}

function printStartupValidationReport() {
  const result = validateStartupConfig();

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('Startup Configuration Check');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  console.log(`DISCORD_TOKEN: ${process.env.DISCORD_TOKEN ? maskValue(process.env.DISCORD_TOKEN) : '[missing]'}`);
  console.log(`CLIENT_ID: ${process.env.CLIENT_ID || '[missing]'}`);
  console.log(`GUILD_ID: ${process.env.GUILD_ID || '[missing]'}`);
  console.log(`HQ_CHANNEL_ID: ${process.env.HQ_CHANNEL_ID || '[missing]'}`);
  console.log(`LOG_CHANNEL_ID: ${process.env.LOG_CHANNEL_ID || '[missing]'}`);
  console.log(`TRAINEE_SPREADSHEET_ID: ${process.env.TRAINEE_SPREADSHEET_ID || '[missing]'}`);
  console.log(`MOS_SPREADSHEET_ID: ${process.env.MOS_SPREADSHEET_ID || '[missing]'}`);
  console.log(`TRAINEE_SHEET_NAME: ${process.env.TRAINEE_SHEET_NAME || '[missing]'}`);
  console.log(`RATINGS_SHEET_NAME: ${process.env.RATINGS_SHEET_NAME || '[missing]'}`);
  console.log(`GOOGLE_CREDENTIALS_JSON: ${process.env.GOOGLE_CREDENTIALS_JSON ? '[present]' : '[missing]'}`);

  if (result.warnings.length > 0) {
    console.warn('\nWarnings:');
    for (const warning of result.warnings) {
      console.warn(`- ${warning}`);
    }
  }

  if (result.errors.length > 0) {
    console.error('\nErrors:');
    for (const error of result.errors) {
      console.error(`- ${error}`);
    }
    console.error('\nStartup aborted due to invalid configuration.');
  } else {
    console.log('\nConfiguration looks valid.');
  }

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  return result;
}

module.exports = {
  validateStartupConfig,
  printStartupValidationReport,
};