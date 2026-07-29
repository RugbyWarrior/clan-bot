const ROLE_IDS = {
  // Cadet roles
  cadet: '1210350313572143184',
  trainingPlatoon: '1210687783501430864',

  // Rank breaker roles
  enlisted: '1210051906060292168',
  nco: '1210405498101178419',

  // Permanent organisational roles
  battalion104th: '1209826414509957150',
  zilloPlatoon: '1291055161564729384',
  unitDivider: '1209826356867764274',

  // Infantry roles
  infantry: '1209823784505442305',
  reserves: '1287750051048853515',
  epsilon: '1285300578150387722',
  nova: '1290372501418938421',
  cinder: '1333106453447184495',
  mythos: '1349394669355794484',

  // Armour roles
  armour: '1209823787756027934',
  titan: '1316845610384625786',
  goliath: '1316845720170528809',
  kronos: '1339401518260031558',
  sisyphus: '1368396574220226661',
  hyperion: '1370527023617015828',

  // Aviation roles
  aviation: '1209823798615089182',
  silverSquadron: '1316451027423199295',
  hawk: '1316451134721888328',
  apollo: '1339401395459194987',
};

const RANKS = {
  cdt: {
    key: 'cdt',
    label: 'CDT - Cadet',
    abbreviation: 'CDT',
    roleId: ROLE_IDS.cadet,
    breaker: 'cadet',
  },

  ct: {
    key: 'ct',
    label: 'CT - Clone Trooper',
    abbreviation: 'CT',
    roleId: '1210035342128840724',
    breaker: 'enlisted',
  },

  spc2: {
    key: 'spc2',
    label: 'SPC2 - Specialist Second Class',
    abbreviation: 'SPC2',
    roleId: '1209664569388179506',
    breaker: 'enlisted',
  },

  spc1: {
    key: 'spc1',
    label: 'SPC1 - Specialist First Class',
    abbreviation: 'SPC1',
    roleId: '1348499044330373160',
    breaker: 'enlisted',
  },

  lcpl: {
    key: 'lcpl',
    label: 'LCPL - Lance Corporal',
    abbreviation: 'LCPL',
    roleId: '1210032580846686208',
    breaker: 'nco',
  },

  cpl: {
    key: 'cpl',
    label: 'CPL - Corporal',
    abbreviation: 'CPL',
    roleId: '1210032707698958418',
    breaker: 'nco',
  },

  sgt: {
    key: 'sgt',
    label: 'SGT - Sergeant',
    abbreviation: 'SGT',
    roleId: '1210033292506701835',
    breaker: 'nco',
  },

  ssgt: {
    key: 'ssgt',
    label: 'SSGT - Staff Sergeant',
    abbreviation: 'SSGT',
    roleId: '1378856593839624262',
    breaker: 'nco',
  },
};

const DESTINATIONS = {
  reserves: {
    key: 'reserves',
    label: 'Reserves',

    rosterAliases: [
      'Reserves',
    ],

    activitySectionKey: 'reserves',

    rosterSquad: 'Reserve Squad',
    activeService: 'Reserves',
    division: 'infantry',

    roleIds: [
      ROLE_IDS.infantry,
      ROLE_IDS.reserves,
    ],
  },

  epsilon: {
    key: 'epsilon',
    label: 'Squad 1 - Epsilon Squad',

    rosterAliases: [
      'Squad 1 - Epsilon Squad',
    ],

    activitySectionKey: 'epsilon',

    rosterSquad: 'Epsilon Squad',
    activeService: 'Active service',
    division: 'infantry',

    roleIds: [
      ROLE_IDS.infantry,
      ROLE_IDS.epsilon,
    ],
  },

  nova: {
    key: 'nova',
    label: 'Squad 2 - Nova Squad',

    rosterAliases: [
      'Squad 2 - Nova Squad',
    ],

    activitySectionKey: 'nova',

    rosterSquad: 'Nova Squad',
    activeService: 'Active service',
    division: 'infantry',

    roleIds: [
      ROLE_IDS.infantry,
      ROLE_IDS.nova,
    ],
  },

  cinder: {
    key: 'cinder',
    label: 'Squad 3 - Cinder Squad',

    rosterAliases: [
      'Squad 3 - Cinder Squad',
    ],

    activitySectionKey: 'cinder',

    rosterSquad: 'Cinder Squad',
    activeService: 'Active service',
    division: 'infantry',

    roleIds: [
      ROLE_IDS.infantry,
      ROLE_IDS.cinder,
    ],
  },

  mythos: {
    key: 'mythos',
    label: 'Squad 4 - Mythos Squad',

    rosterAliases: [
      'Squad 4 - Mythos Squad',
    ],

    activitySectionKey: 'mythos',

    rosterSquad: 'Mythos Squad',
    activeService: 'Active service',
    division: 'infantry',

    roleIds: [
      ROLE_IDS.infantry,
      ROLE_IDS.mythos,
    ],
  },

  goliath: {
    key: 'goliath',
    label: 'AT-TE 1 - Goliath',

    rosterAliases: [
      'AT-TE 1 - Goliath',
    ],

    activitySectionKey: 'titan',

    rosterSquad: 'Armor',
    activeService: 'Active service',
    division: 'armour',

    roleIds: [
      ROLE_IDS.armour,
      ROLE_IDS.titan,
      ROLE_IDS.goliath,
    ],
  },

  kronos: {
    key: 'kronos',
    label: 'AT-TE 2 - Kronos',

    rosterAliases: [
      'AT-TE 2 - Kronos',
    ],

    activitySectionKey: 'titan',

    rosterSquad: 'Armor',
    activeService: 'Active service',
    division: 'armour',

    roleIds: [
      ROLE_IDS.armour,
      ROLE_IDS.titan,
      ROLE_IDS.kronos,
    ],
  },

  sisyphus: {
    key: 'sisyphus',
    label: 'TX-130 1 - Sisyphus / Buff',

    rosterAliases: [
      'TX-130 1 - Sisyphus',
    ],

    activitySectionKey: 'titan',

    rosterSquad: 'Armor',
    activeService: 'Active service',
    division: 'armour',

    roleIds: [
      ROLE_IDS.armour,
      ROLE_IDS.titan,
      ROLE_IDS.sisyphus,
    ],
  },

  hyperion: {
    key: 'hyperion',
    label: 'TX-130 2 - Hyperion',

    rosterAliases: [
      'TX-130 2 - Hyperion',
    ],

    activitySectionKey: 'titan',

    rosterSquad: 'Armor',
    activeService: 'Active service',
    division: 'armour',

    roleIds: [
      ROLE_IDS.armour,
      ROLE_IDS.titan,
      ROLE_IDS.hyperion,
    ],
  },

  hawk: {
    key: 'hawk',
    label: 'LAAT 1 - Hawk',

    rosterAliases: [
      'LAAT 1 - Hawk',
    ],

    activitySectionKey: 'silver',

    rosterSquad: 'Aviation',
    activeService: 'Active service',
    division: 'aviation',

    roleIds: [
      ROLE_IDS.aviation,
      ROLE_IDS.silverSquadron,
      ROLE_IDS.hawk,
    ],
  },

  apollo: {
    key: 'apollo',
    label: 'LAAT 2 - Apollo',

    rosterAliases: [
      'LAAT 2 - Apollo',
    ],

    activitySectionKey: 'silver',

    rosterSquad: 'Aviation',
    activeService: 'Active service',
    division: 'aviation',

    roleIds: [
      ROLE_IDS.aviation,
      ROLE_IDS.silverSquadron,
      ROLE_IDS.apollo,
    ],
  },
};

const ACTIVITY_SECTIONS = {
  headquarters: [
    'Headquarters',
  ],

  epsilon: [
    'Squad 1 - Epsilon Squad',
  ],

  nova: [
    'Squad 2 - Nova Squad',
  ],

  cinder: [
    'Squad 3 - Cinder Squad',
  ],

  mythos: [
    'Squad 4 - Mythos Squad',
  ],

  titan: [
    'Squad 5 - Titan Squad',
  ],

  silver: [
    'Squad 6 - Silver Squadron',
  ],

  reserves: [
    'Reserves',
  ],

  cadets: [
    'Cadets',
  ],
};

const ROSTER_BOUNDARIES = [
  'Platoon 1 - ZilloPlatoon',
  'Platoon 1 - Zillo Platoon',

  'Section 1 - Infantry',

  'Squad 1 - Epsilon Squad',
  'Squad 2 - Nova Squad',
  'Squad 3 - Cinder Squad',
  'Squad 4 - Mythos Squad',

  'Reserves',

  'Section 2 - Motorized',
  'Armour Division - Titan Squad',

  'AT-TE 1 - Goliath',
  'AT-TE 2 - Kronos',
  'TX-130 1 - Sisyphus',
  'TX-130 2 - Hyperion',

  'Aviation Division - Silver Squadron',

  'LAAT 1 - Hawk',
  'LAAT 2 - Apollo',
  'LAAT 3',
  'LAAT 4',
];

const RANK_ROLE_IDS = Object.values(
  RANKS
).map(rank => rank.roleId);

const UNIT_ROLE_IDS = [
  // Infantry
  ROLE_IDS.infantry,
  ROLE_IDS.reserves,
  ROLE_IDS.epsilon,
  ROLE_IDS.nova,
  ROLE_IDS.cinder,
  ROLE_IDS.mythos,

  // Armour
  ROLE_IDS.armour,
  ROLE_IDS.titan,
  ROLE_IDS.goliath,
  ROLE_IDS.kronos,
  ROLE_IDS.sisyphus,
  ROLE_IDS.hyperion,

  // Aviation
  ROLE_IDS.aviation,
  ROLE_IDS.silverSquadron,
  ROLE_IDS.hawk,
  ROLE_IDS.apollo,
];

module.exports = {
  ROLE_IDS,
  RANKS,
  DESTINATIONS,
  ACTIVITY_SECTIONS,
  ROSTER_BOUNDARIES,
  RANK_ROLE_IDS,
  UNIT_ROLE_IDS,
};