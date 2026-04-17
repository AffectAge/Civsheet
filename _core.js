/*******************************************
 * Civ-подобный движок для Google Sheets
 * Хранение:
 * - Named ranges
 * - JSON Lines (1 строка = 1 JSON-объект)
 * - Батч чтение/запись
 * - Автосоздание листов и диапазонов
 * - Карта хранится как JSON-объекты гексов
 * - Приказы только в JSON
 * - В приказах используются названия гексов
 * - Отдельная политическая карта
 *******************************************/

const CONFIG = {
  SHEETS: {
    CORE: 'Core',
    MAP: 'Карта',
    POLITICAL_MAP: 'Политическая карта',
    ORDERS: 'Приказы',
    REPORTS: 'Отчёты',
    SWITCHBOARD: 'Переключатели',
  },

  RANGES: {
    GAME_META: 'NR_GAME_META',
    PLAYERS: 'NR_PLAYERS',
    CITIES: 'NR_CITIES',
    UNITS: 'NR_UNITS',
    ORDERS: 'NR_ORDERS',
    LOG: 'NR_LOG',
    HEXES: 'NR_HEXES',
    UNIT_TYPES: 'NR_UNIT_TYPES',
    BUILDING_TYPES: 'NR_BUILDING_TYPES',
    TECH_TYPES: 'NR_TECH_TYPES',
    RULES: 'NR_RULES',
    DISTRICTS: 'NR_DISTRICTS',
    WONDERS: 'NR_WONDERS',
    HEX_IMPROVEMENTS: 'NR_HEX_IMPROVEMENTS',
    DISTRICT_TYPES: 'NR_DISTRICT_TYPES',
    WONDER_TYPES: 'NR_WONDER_TYPES',
    HEX_IMPROVEMENT_TYPES: 'NR_HEX_IMPROVEMENT_TYPES',
  },

  DEFAULT_CAPACITY: {
    NR_GAME_META: 10,
    NR_PLAYERS: 50,
    NR_CITIES: 500,
    NR_UNITS: 5000,
    NR_ORDERS: 5000,
    NR_LOG: 10000,
    NR_HEXES: 10000,
    NR_UNIT_TYPES: 500,
    NR_BUILDING_TYPES: 500,
    NR_TECH_TYPES: 500,
    NR_RULES: 200,
    NR_DISTRICTS: 2000,
    NR_WONDERS: 200,
    NR_HEX_IMPROVEMENTS: 5000,
    NR_DISTRICT_TYPES: 100,
    NR_WONDER_TYPES: 100,
    NR_HEX_IMPROVEMENT_TYPES: 100,
  },

  CORE_LAYOUT: {
    NR_GAME_META: { col: 1 },
    NR_PLAYERS: { col: 2 },
    NR_CITIES: { col: 3 },
    NR_UNITS: { col: 4 },
    NR_ORDERS: { col: 5 },
    NR_LOG: { col: 6 },
    NR_HEXES: { col: 7 },
    NR_UNIT_TYPES: { col: 8 },
    NR_BUILDING_TYPES: { col: 9 },
    NR_TECH_TYPES: { col: 10 },
    NR_RULES: { col: 11 },
    NR_DISTRICTS: { col: 12 },
    NR_WONDERS: { col: 13 },
    NR_HEX_IMPROVEMENTS: { col: 14 },
    NR_DISTRICT_TYPES: { col: 15 },
    NR_WONDER_TYPES: { col: 16 },
    NR_HEX_IMPROVEMENT_TYPES: { col: 17 },
  },

  CORE_START_ROW: 2,

  MAP_RENDER: {
    START_ROW: 2,
    START_COL: 2,
  },

  HEX_DIRECTIONS_AXIAL: [
    { dq: 1, dr: 0 },
    { dq: 1, dr: -1 },
    { dq: 0, dr: -1 },
    { dq: -1, dr: 0 },
    { dq: -1, dr: 1 },
    { dq: 0, dr: 1 },
  ],
};

/*******************************************
 * МЕНЮ
 *******************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Цивилизация')
    .addItem('Подготовить хранилище', 'bootstrapGameStorage')
    .addItem('Новая игра', 'initNewGame')
    .addSeparator()
    .addItem('Импорт приказов', 'importOrdersFromSheet')
    .addItem('Выполнить ход', 'resolveTurnSafe')
    .addSeparator()
    .addItem('Отрисовать карту', 'renderMap')
    .addItem('Отрисовать политическую карту', 'renderPoliticalMap')
    .addItem('Отрисовать отчёты', 'renderReports')
    .addItem('Отрисовать всё', 'renderAll')
    .addSeparator()
    .addItem('Выполнить по переключателям', 'runActionsFromSwitchboard')
    .addSeparator()
    .addItem('Очистить лист приказов', 'clearOrdersSheet')
    .addToUi();
}

/*******************************************
 * ПОДГОТОВКА ХРАНИЛИЩА
 *******************************************/

function bootstrapGameStorage() {
  const coreSheet = getOrCreateSheet_(CONFIG.SHEETS.CORE);
  const mapSheet = getOrCreateSheet_(CONFIG.SHEETS.MAP);
  const politicalMapSheet = getOrCreateSheet_(CONFIG.SHEETS.POLITICAL_MAP);
  const ordersSheet = getOrCreateSheet_(CONFIG.SHEETS.ORDERS);
  const reportsSheet = getOrCreateSheet_(CONFIG.SHEETS.REPORTS);
  const switchboardSheet = getOrCreateSheet_(CONFIG.SHEETS.SWITCHBOARD);

  setupCoreSheet_(coreSheet);
  setupMapSheet_(mapSheet);
  setupPoliticalMapSheet_(politicalMapSheet);
  setupOrdersSheet_(ordersSheet);
  setupReportsSheet_(reportsSheet);
  setupSwitchboardSheet_(switchboardSheet);

  ensureAllNamedRanges_();
}

function setupSwitchboardSheet_(sheet) {
  const headers = ['action', 'run', 'description'];
  const actions = getSwitchboardActions_();
  const existing = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
  const existingRunByAction = {};
  existing.forEach((row) => {
    const id = String(row[0] || '').trim();
    if (!id) return;
    existingRunByAction[id] = row[1] === true;
  });

  const rows = actions.map((item) => [item.id, Boolean(existingRunByAction[item.id]), item.description]);
  const targetRows = Math.max(rows.length + 1, 30);
  ensureSheetSize_(sheet, targetRows, headers.length + 1);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const checkboxRange = sheet.getRange(2, 2, rows.length, 1);
  checkboxRange.insertCheckboxes();
  checkboxRange.setValues(rows.map((row) => [row[1]]));

  if (sheet.getLastRow() > rows.length + 1) {
    sheet.getRange(rows.length + 2, 1, sheet.getLastRow() - rows.length - 1, headers.length).clearContent();
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 560);
}

function getSwitchboardActions_() {
  return [
    { id: 'bootstrapGameStorage', description: 'Подготовить хранилище и named ranges' },
    { id: 'initNewGame', description: 'Создать новую игру (перезапишет текущее состояние)' },
    { id: 'importOrdersFromSheet', description: 'Импортировать JSON-приказы с листа "Приказы"' },
    { id: 'resolveTurnSafe', description: 'Выполнить ход' },
    { id: 'renderMap', description: 'Отрисовать карту' },
    { id: 'renderPoliticalMap', description: 'Отрисовать политическую карту' },
    { id: 'renderReports', description: 'Отрисовать отчёты' },
    { id: 'renderAll', description: 'Отрисовать всё' },
    { id: 'clearOrdersSheet', description: 'Очистить лист приказов' },
  ];
}

function runActionsFromSwitchboard() {
  bootstrapGameStorage();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.SWITCHBOARD);
  const actions = getSwitchboardActions_();
  const values = sheet.getRange(2, 1, actions.length, 2).getValues();
  const messages = [];

  values.forEach((row, i) => {
    const actionId = String(row[0] || '').trim();
    const shouldRun = row[1] === true;
    if (!shouldRun) return;

    const actionDef = actions.find((a) => a.id === actionId);
    if (!actionDef) {
      messages.push(`⚠ Неизвестное действие: ${actionId}`);
      return;
    }
    const fn = this[actionDef.id];
    if (typeof fn !== 'function') {
      messages.push(`⚠ Функция не найдена: ${actionDef.id}`);
      return;
    }
    fn();
    messages.push(`✅ Выполнено: ${actionDef.id}`);
    sheet.getRange(2 + i, 2).setValue(false);
  });

  SpreadsheetApp.getActive().toast(messages.length ? messages.join('\n') : 'Нет включённых переключателей');
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setupCoreSheet_(sheet) {
  const headers = [];
  Object.keys(CONFIG.CORE_LAYOUT).forEach((rangeName) => {
    const col = CONFIG.CORE_LAYOUT[rangeName].col;
    headers[col - 1] = rangeName;
    sheet.setColumnWidth(col, 300);
  });

  const currentHeaderValues = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needWriteHeaders = headers.some((h, i) => String(currentHeaderValues[i] || '') !== h);

  if (needWriteHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  ensureSheetSize_(sheet, CONFIG.CORE_START_ROW + 1000, headers.length + 2);
}

function setupMapSheet_(sheet) {
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1).setValue('Карта');
    sheet.getRange(1, 1).setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function setupPoliticalMapSheet_(sheet) {
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1).setValue('Политическая карта');
    sheet.getRange(1, 1).setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function setupOrdersSheet_(sheet) {
  const headers = ['json'];

  const currentHeaderValues = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needWriteHeaders = headers.some((h, i) => String(currentHeaderValues[i] || '') !== h);

  if (needWriteHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 1200);
}

function setupReportsSheet_(sheet) {
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1).setValue('Отчёты');
    sheet.getRange(1, 1).setFontWeight('bold');
  }
}

function ensureAllNamedRanges_() {
  Object.keys(CONFIG.RANGES).forEach((key) => {
    ensureNamedRange_(CONFIG.RANGES[key]);
  });
}

function ensureNamedRange_(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getRangeByName(rangeName);
  if (existing) return existing;

  const coreSheet = getOrCreateSheet_(CONFIG.SHEETS.CORE);
  const col = CONFIG.CORE_LAYOUT[rangeName].col;
  const capacity = CONFIG.DEFAULT_CAPACITY[rangeName] || 100;
  const startRow = CONFIG.CORE_START_ROW;

  ensureSheetSize_(coreSheet, startRow + capacity + 5, col + 2);

  const range = coreSheet.getRange(startRow, col, capacity, 1);
  ss.setNamedRange(rangeName, range);
  return range;
}

function ensureSheetSize_(sheet, minRows, minCols) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < minCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minCols - sheet.getMaxColumns());
  }
}

/*******************************************
 * НОВАЯ ИГРА
 *******************************************/

function initNewGame() {
  bootstrapGameStorage();

  const state = createInitialState_();
  saveState(state);
  renderAll();

  SpreadsheetApp.getActive().toast('Новая игра создана');
}

function getInitialDistrictTypes_() {
  return [
    { id: 'campus', nameRu: 'Кампус', category: 'science', populationRequired: 3, cost: 90, yields: { science: 4 }, allowedBuildings: ['library', 'university'], maxPerCity: 1, icon: '🎓' },
    { id: 'commercial_hub', nameRu: 'Торговый центр', category: 'gold', populationRequired: 2, cost: 80, yields: { gold: 4 }, allowedBuildings: ['market', 'bank'], maxPerCity: 1, icon: '💰' },
    { id: 'harbor', nameRu: 'Гавань', category: 'gold', populationRequired: 2, cost: 75, yields: { gold: 2, food: 1 }, allowedBuildings: ['lighthouse', 'shipyard'], maxPerCity: 1, requiresCoast: true, icon: '⚓' },
    { id: 'holy_site', nameRu: 'Священное место', category: 'faith', populationRequired: 1, cost: 70, yields: { faith: 4 }, allowedBuildings: ['shrine', 'temple'], maxPerCity: 1, icon: '⛪' },
    { id: 'encampment', nameRu: 'Военный лагерь', category: 'military', populationRequired: 3, cost: 90, yields: { production: 1 }, allowedBuildings: ['barracks', 'stable'], maxPerCity: 1, icon: '🏕' },
    { id: 'theater_square', nameRu: 'Театральная площадь', category: 'culture', populationRequired: 3, cost: 90, yields: { culture: 4 }, allowedBuildings: ['amphitheater', 'art_museum'], maxPerCity: 1, icon: '🎭' },
    { id: 'industrial_zone', nameRu: 'Промышленная зона', category: 'production', populationRequired: 4, cost: 100, yields: { production: 4 }, allowedBuildings: ['workshop', 'factory'], maxPerCity: 1, icon: '🏭' },
  ];
}

function getInitialWonderTypes_() {
  return [
    { id: 'pyramids', nameRu: 'Пирамиды', cost: 220, era: 'ancient', yields: { culture: 2, faith: 2 }, cityBonus: { housing: 2 }, unique: true, icon: '🔺', description: 'Все будущие поселенцы строятся с бонусом +1 ход.' },
    { id: 'colosseum', nameRu: 'Колизей', cost: 400, era: 'classical', yields: { culture: 2, gold: 3 }, cityBonus: { amenities: 3 }, unique: true, icon: '🏟', description: '+3 удобства во всех городах в радиусе 6 тайлов.' },
    { id: 'great_library', nameRu: 'Великая Библиотека', cost: 260, era: 'classical', yields: { science: 4, culture: 2 }, cityBonus: {}, unique: true, icon: '📚', description: '+2 великих учёных и +2 слота реликвий.' },
    { id: 'stonehenge', nameRu: 'Стоунхендж', cost: 180, era: 'ancient', yields: { faith: 6 }, cityBonus: {}, unique: true, icon: '🗿', description: 'Бесплатная религия для построившего.' },
    { id: 'colossus', nameRu: 'Колосс', cost: 400, era: 'classical', yields: { gold: 6 }, cityBonus: {}, unique: true, requiresCoast: true, icon: '🗽', description: 'Торговый маршрут в портовых городах приносит +1 золото.' },
    { id: 'oracle', nameRu: 'Оракул', cost: 290, era: 'classical', yields: { culture: 2, faith: 2 }, cityBonus: {}, unique: true, icon: '🔮', description: 'Все технологии и гражданские науки исследуются на 20% быстрее.' },
  ];
}

function getInitialHexImprovementTypes_() {
  return [
    { id: 'farm', nameRu: 'Ферма', icon: '🌾', cost: 0, yields: { food: 1 }, allowedTerrains: ['grassland', 'plains', 'desert', 'tundra'], requiredTech: null, removesFeature: false },
    { id: 'mine', nameRu: 'Шахта', icon: '⛏', cost: 0, yields: { production: 2 }, allowedTerrains: ['hill', 'plains', 'desert', 'tundra'], requiredTech: 'mining', removesFeature: false },
    { id: 'lumber_mill', nameRu: 'Лесопилка', icon: '🪚', cost: 0, yields: { production: 2 }, allowedTerrains: ['grassland', 'plains'], requiredFeature: 'forest', requiredTech: 'mining', removesFeature: false },
    { id: 'pasture', nameRu: 'Пастбище', icon: '🐄', cost: 0, yields: { food: 1, production: 1 }, allowedTerrains: ['grassland', 'plains', 'tundra'], requiredTech: null, removesFeature: false },
    { id: 'quarry', nameRu: 'Каменоломня', icon: '🪨', cost: 0, yields: { production: 1, gold: 1 }, allowedTerrains: ['grassland', 'plains', 'hill'], requiredTech: 'mining', removesFeature: false },
    { id: 'road', nameRu: 'Дорога', icon: '🛤', cost: 0, yields: {}, movementBonus: -1, allowedTerrains: ['grassland', 'plains', 'desert', 'tundra', 'hill', 'coast'], requiredTech: null, removesFeature: false },
  ];
}

function getDistrictBuildingTypes_() {
  return [
    { id: 'library', nameRu: 'Библиотека', cost: 90, yields: { science: 2 }, housing: 0, modifiers: [], districtId: 'campus' },
    { id: 'university', nameRu: 'Университет', cost: 250, yields: { science: 4 }, housing: 0, modifiers: [], districtId: 'campus' },
    { id: 'market', nameRu: 'Рынок', cost: 80, yields: { gold: 3 }, housing: 0, modifiers: [], districtId: 'commercial_hub' },
    { id: 'bank', nameRu: 'Банк', cost: 260, yields: { gold: 5 }, housing: 0, modifiers: [], districtId: 'commercial_hub' },
    { id: 'lighthouse', nameRu: 'Маяк', cost: 80, yields: { food: 1, gold: 1 }, housing: 1, modifiers: [], districtId: 'harbor' },
    { id: 'shipyard', nameRu: 'Верфь', cost: 240, yields: { production: 2 }, housing: 0, modifiers: [], districtId: 'harbor' },
    { id: 'shrine', nameRu: 'Святилище', cost: 70, yields: { faith: 2 }, housing: 0, modifiers: [], districtId: 'holy_site' },
    { id: 'temple', nameRu: 'Храм', cost: 200, yields: { faith: 4 }, housing: 1, modifiers: [], districtId: 'holy_site' },
    { id: 'barracks', nameRu: 'Казармы', cost: 90, yields: { production: 1 }, housing: 1, modifiers: [], districtId: 'encampment' },
    { id: 'stable', nameRu: 'Конюшня', cost: 160, yields: { production: 1, gold: 1 }, housing: 1, modifiers: [], districtId: 'encampment' },
    { id: 'amphitheater', nameRu: 'Амфитеатр', cost: 90, yields: { culture: 2 }, housing: 0, modifiers: [], districtId: 'theater_square' },
    { id: 'art_museum', nameRu: 'Художественный музей', cost: 240, yields: { culture: 3 }, housing: 0, modifiers: [], districtId: 'theater_square' },
    { id: 'workshop', nameRu: 'Мастерская', cost: 100, yields: { production: 2 }, housing: 0, modifiers: [], districtId: 'industrial_zone' },
    { id: 'factory', nameRu: 'Завод', cost: 300, yields: { production: 4 }, housing: 0, modifiers: [], districtId: 'industrial_zone' },
  ];
}

function createInitialState_() {
  const unitTypes = [
    {
      id: 'warrior',
      nameRu: 'Воин',
      category: 'melee',
      domain: 'land',
      cost: 40,
      combat: { melee: 20, ranged: 0, range: 0 },
      movement: { max: 2 },
      tags: ['military'],
    },
    {
      id: 'settler',
      nameRu: 'Поселенец',
      category: 'civilian',
      domain: 'land',
      cost: 80,
      combat: { melee: 0, ranged: 0, range: 0 },
      movement: { max: 2 },
      tags: ['civilian', 'found_city'],
    },
    {
      id: 'slinger',
      nameRu: 'Пращник',
      category: 'ranged',
      domain: 'land',
      cost: 35,
      combat: { melee: 5, ranged: 15, range: 2 },
      movement: { max: 2 },
      tags: ['military'],
    },
    {
      id: 'builder',
      nameRu: 'Строитель',
      category: 'civilian',
      domain: 'land',
      cost: 50,
      combat: { melee: 0, ranged: 0, range: 0 },
      movement: { max: 2 },
      builderCharges: 3,
      tags: ['civilian', 'builder'],
    },
  ];

  const buildingTypes = [
    {
      id: 'palace',
      nameRu: 'Дворец',
      cost: 0,
      yields: { gold: 5, science: 2, culture: 1 },
      housing: 1,
      modifiers: [],
    },
    {
      id: 'monument',
      nameRu: 'Монумент',
      cost: 60,
      yields: { culture: 2 },
      housing: 0,
      modifiers: [],
    },
    {
      id: 'granary',
      nameRu: 'Амбар',
      cost: 65,
      yields: { food: 1 },
      housing: 2,
      modifiers: [],
    },
  ].concat(getDistrictBuildingTypes_());

  const techTypes = [
    {
      id: 'pottery',
      nameRu: 'Гончарное дело',
      cost: 20,
      requires: [],
      unlocks: { buildings: ['granary'], units: [] },
    },
    {
      id: 'mining',
      nameRu: 'Добыча',
      cost: 25,
      requires: [],
      unlocks: { buildings: [], units: [] },
    },
    {
      id: 'archery',
      nameRu: 'Стрельба',
      cost: 35,
      requires: [],
      unlocks: { buildings: [], units: ['slinger'] },
    },
    {
      id: 'bronze_working',
      nameRu: 'Обработка бронзы',
      cost: 40,
      requires: ['mining'],
      unlocks: { buildings: [], units: [] },
    },
  ];

  const rules = [
    {
      id: 'economy',
      baseFoodToGrow: 15,
      growthFactor: 1.15,
      baseCityHealth: 200,
      baseCityFood: 2,
      foodPerPop: 1,
      baseCityProduction: 1,
      productionPerTwoPop: 1,
      baseCityScience: 1,
      sciencePerPop: 1,
      baseCityGold: 1,
    },
    {
      id: 'combat',
      baseDamage: 24,
      minDamage: 6,
      maxDamage: 40,
      cityBaseStrength: 12,
    },
    {
      id: 'movement',
      terrainCost: {
        grassland: 1,
        plains: 1,
        desert: 1,
        tundra: 1,
        coast: 1,
        hill: 2,
        mountain: 999,
        water: 999,
      },
      featureCost: {
        forest: 1,
        jungle: 1,
        marsh: 1,
      },
    },
    {
      id: 'turn',
      refreshUnitMovesEachTurn: true,
      allowMultipleOrdersPerEntity: false,
    },
  ];
  const districtTypes = getInitialDistrictTypes_();
  const wonderTypes = getInitialWonderTypes_();
  const hexImprovementTypes = getInitialHexImprovementTypes_();

  const players = [
    {
      id: 'P1',
      name: 'Рим',
      leader: 'Траян',
      isAI: false,
      color: '#c84b4b',
      textColor: '#ffffff',
      gold: 120,
      scienceStock: 0,
      cultureStock: 0,
      faithStock: 0,
      goldPerTurn: 0,
      sciencePerTurn: 0,
      culturePerTurn: 0,
      faithPerTurn: 0,
      knownTechs: ['pottery', 'mining'],
      knownCivics: ['code_of_laws'],
      currentResearch: 'bronze_working',
      researchProgress: 0,
      government: 'chiefdom',
      policies: ['discipline'],
      visibility: {},
    },
    {
      id: 'P2',
      name: 'Египет',
      leader: 'Клеопатра',
      isAI: true,
      color: '#d9b43b',
      textColor: '#000000',
      gold: 100,
      scienceStock: 0,
      cultureStock: 0,
      faithStock: 0,
      goldPerTurn: 0,
      sciencePerTurn: 0,
      culturePerTurn: 0,
      faithPerTurn: 0,
      knownTechs: ['pottery'],
      knownCivics: ['code_of_laws'],
      currentResearch: 'mining',
      researchProgress: 0,
      government: 'chiefdom',
      policies: [],
      visibility: {},
    },
  ];

  const hexes = generateDemoHexMap_(10, 8);

  const cities = [
    {
      id: 'C1',
      playerId: 'P1',
      name: 'Рим',
      hexId: 'H_2_3',
      population: 3,
      foodStored: 0,
      productionStored: 0,
      housing: 5,
      amenities: 1,
      health: 200,
      buildings: ['palace'],
      districts: [],
      queue: [{ kind: 'building', typeId: 'monument' }],
      specialists: [],
      modifiers: [],
      workedHexIds: getNearbyHexIds_('H_2_3', hexes, 1).slice(0, 6),
    },
    {
      id: 'C2',
      playerId: 'P2',
      name: 'Мемфис',
      hexId: 'H_7_4',
      population: 2,
      foodStored: 0,
      productionStored: 0,
      housing: 4,
      amenities: 0,
      health: 200,
      buildings: ['palace'],
      districts: [],
      queue: [{ kind: 'unit', typeId: 'warrior' }],
      specialists: [],
      modifiers: [],
      workedHexIds: getNearbyHexIds_('H_7_4', hexes, 1).slice(0, 6),
    },
  ];

  const units = [
    {
      id: 'U1',
      playerId: 'P1',
      type: 'warrior',
      hexId: 'H_3_3',
      hp: 100,
      movesLeft: 2,
      status: 'idle',
      xp: 0,
      promotions: [],
      task: null,
    },
    {
      id: 'U2',
      playerId: 'P1',
      type: 'settler',
      hexId: 'H_2_2',
      hp: 100,
      movesLeft: 2,
      status: 'idle',
      xp: 0,
      promotions: [],
      task: null,
    },
    {
      id: 'U3',
      playerId: 'P2',
      type: 'warrior',
      hexId: 'H_6_4',
      hp: 100,
      movesLeft: 2,
      status: 'idle',
      xp: 0,
      promotions: [],
      task: null,
    },
  ];

  assignHexOwnershipFromCities_(hexes, cities);
  assignHexOccupants_(hexes, cities, units);

  const log = [
    {
      turn: 1,
      type: 'info',
      playerId: null,
      text: 'Игра создана',
      payload: {},
    },
  ];

  const meta = {
    id: 'meta',
    version: 1,
    turn: 1,
    phase: 'player',
    activePlayerId: 'P1',
    map: {
      width: 10,
      height: 8,
      axial: true,
    },
    victory: {
      science: true,
      culture: true,
      domination: true,
      score: true,
    },
  };

  return {
    meta,
    players,
    cities,
    units,
    orders: [],
    log,
    hexes,
    unitTypes,
    buildingTypes,
    techTypes,
    rules,
    districts: [],
    wonders: [],
    hexImprovements: [],
    districtTypes,
    wonderTypes,
    hexImprovementTypes,
  };
}

/*******************************************
 * JSON LINES
 *******************************************/

function getNamedRangeOrCreate_(rangeName) {
  return ensureNamedRange_(rangeName);
}

function parseJsonSafe_(text, fallback = null) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return fallback;
  }
}

function readJsonLines_(rangeName) {
  const range = getNamedRangeOrCreate_(rangeName);
  const values = range.getValues();
  const out = [];

  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    const text = String(cell || '').trim();
    if (!text) continue;

    const obj = parseJsonSafe_(text, null);
    if (!obj) {
      throw new Error(`Некорректный JSON в диапазоне ${rangeName}, строка ${i + 1}`);
    }
    out.push(obj);
  }

  return out;
}

function writeJsonLines_(rangeName, items) {
  let range = getNamedRangeOrCreate_(rangeName);
  const needed = Math.max(1, items.length);
  const capacity = range.getNumRows();

  if (needed > capacity) {
    range = expandNamedRange_(rangeName, needed);
  }

  const values = items.map((item) => [JSON.stringify(item)]);
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const clearRows = range.getNumRows();

  sheet.getRange(startRow, startCol, clearRows, 1).clearContent();

  if (values.length > 0) {
    sheet.getRange(startRow, startCol, values.length, 1).setValues(values);
  }
}

function expandNamedRange_(rangeName, minNeededRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldRange = getNamedRangeOrCreate_(rangeName);
  const sheet = oldRange.getSheet();
  const startRow = oldRange.getRow();
  const startCol = oldRange.getColumn();
  const newSize = Math.max(oldRange.getNumRows() * 2, minNeededRows + 20);

  ensureSheetSize_(sheet, startRow + newSize + 5, startCol + 2);

  const newRange = sheet.getRange(startRow, startCol, newSize, 1);
  ss.setNamedRange(rangeName, newRange);
  return newRange;
}

function readSingleJsonObject_(rangeName) {
  const items = readJsonLines_(rangeName);
  if (items.length === 0) return null;
  if (items.length > 1) {
    throw new Error(`В ${rangeName} должен быть максимум один JSON-объект`);
  }
  return items[0];
}

function writeSingleJsonObject_(rangeName, obj) {
  writeJsonLines_(rangeName, obj ? [obj] : []);
}

/*******************************************
 * ЗАГРУЗКА / СОХРАНЕНИЕ
 *******************************************/

function loadState() {
  bootstrapGameStorage();

  const state = {
    meta: readSingleJsonObject_(CONFIG.RANGES.GAME_META),
    players: readJsonLines_(CONFIG.RANGES.PLAYERS),
    cities: readJsonLines_(CONFIG.RANGES.CITIES),
    units: readJsonLines_(CONFIG.RANGES.UNITS),
    orders: readJsonLines_(CONFIG.RANGES.ORDERS),
    log: readJsonLines_(CONFIG.RANGES.LOG),
    hexes: readJsonLines_(CONFIG.RANGES.HEXES),
    unitTypes: readJsonLines_(CONFIG.RANGES.UNIT_TYPES),
    buildingTypes: readJsonLines_(CONFIG.RANGES.BUILDING_TYPES),
    techTypes: readJsonLines_(CONFIG.RANGES.TECH_TYPES),
    rules: readJsonLines_(CONFIG.RANGES.RULES),
    districts: readJsonLines_(CONFIG.RANGES.DISTRICTS),
    wonders: readJsonLines_(CONFIG.RANGES.WONDERS),
    hexImprovements: readJsonLines_(CONFIG.RANGES.HEX_IMPROVEMENTS),
    districtTypes: readJsonLines_(CONFIG.RANGES.DISTRICT_TYPES),
    wonderTypes: readJsonLines_(CONFIG.RANGES.WONDER_TYPES),
    hexImprovementTypes: readJsonLines_(CONFIG.RANGES.HEX_IMPROVEMENT_TYPES),
  };
  if (!state.districtTypes.length) state.districtTypes = getInitialDistrictTypes_();
  if (!state.wonderTypes.length) state.wonderTypes = getInitialWonderTypes_();
  if (!state.hexImprovementTypes.length) state.hexImprovementTypes = getInitialHexImprovementTypes_();

  if (!state.meta) {
    throw new Error('Игра не инициализирована. Запусти "Новая игра".');
  }

  buildIndexes_(state);
  return state;
}

function saveState(state) {
  rebuildStateDerivedFields_(state);

  writeSingleJsonObject_(CONFIG.RANGES.GAME_META, state.meta);
  writeJsonLines_(CONFIG.RANGES.PLAYERS, state.players);
  writeJsonLines_(CONFIG.RANGES.CITIES, state.cities);
  writeJsonLines_(CONFIG.RANGES.UNITS, state.units);
  writeJsonLines_(CONFIG.RANGES.ORDERS, state.orders);
  writeJsonLines_(CONFIG.RANGES.LOG, state.log);
  writeJsonLines_(CONFIG.RANGES.HEXES, state.hexes);
  writeJsonLines_(CONFIG.RANGES.UNIT_TYPES, state.unitTypes);
  writeJsonLines_(CONFIG.RANGES.BUILDING_TYPES, state.buildingTypes);
  writeJsonLines_(CONFIG.RANGES.TECH_TYPES, state.techTypes);
  writeJsonLines_(CONFIG.RANGES.RULES, state.rules);
  writeJsonLines_(CONFIG.RANGES.DISTRICTS, state.districts || []);
  writeJsonLines_(CONFIG.RANGES.WONDERS, state.wonders || []);
  writeJsonLines_(CONFIG.RANGES.HEX_IMPROVEMENTS, state.hexImprovements || []);
  writeJsonLines_(CONFIG.RANGES.DISTRICT_TYPES, state.districtTypes || []);
  writeJsonLines_(CONFIG.RANGES.WONDER_TYPES, state.wonderTypes || []);
  writeJsonLines_(CONFIG.RANGES.HEX_IMPROVEMENT_TYPES, state.hexImprovementTypes || []);
}

function buildIndexes_(state) {
  state.index = {
    playerById: indexById_(state.players),
    cityById: indexById_(state.cities),
    unitById: indexById_(state.units),
    hexById: indexById_(state.hexes),
    hexByName: {},
    unitTypeById: indexById_(state.unitTypes),
    buildingTypeById: indexById_(state.buildingTypes),
    techTypeById: indexById_(state.techTypes),
    ruleById: indexById_(state.rules),
    districtById: indexById_(state.districts || []),
    wonderById: indexById_(state.wonders || []),
    hexImprovementById: indexById_(state.hexImprovements || []),
    districtTypeById: indexById_(state.districtTypes || []),
    wonderTypeById: indexById_(state.wonderTypes || []),
    hexImprovementTypeById: indexById_(state.hexImprovementTypes || []),
    cityByHexId: {},
    unitsByHexId: {},
    districtsByHexId: {},
    wonderByHexId: {},
    improvementByHexId: {},
    districtsByCityId: {},
    wondersByCityId: {},
    wonderByTypeId: {},
  };

  state.hexes.forEach((hex) => {
    state.index.hexByName[normalizeHexName_(hex.name)] = hex;
  });

  state.cities.forEach((city) => {
    state.index.cityByHexId[city.hexId] = city;
  });

  state.units.forEach((unit) => {
    if (!state.index.unitsByHexId[unit.hexId]) {
      state.index.unitsByHexId[unit.hexId] = [];
    }
    state.index.unitsByHexId[unit.hexId].push(unit);
  });

  (state.districts || []).forEach((district) => {
    state.index.districtsByHexId[district.hexId] = district;
    if (!state.index.districtsByCityId[district.cityId]) state.index.districtsByCityId[district.cityId] = [];
    state.index.districtsByCityId[district.cityId].push(district);
  });

  (state.wonders || []).forEach((wonder) => {
    state.index.wonderByHexId[wonder.hexId] = wonder;
    state.index.wonderByTypeId[wonder.typeId] = wonder;
    if (!state.index.wondersByCityId[wonder.cityId]) state.index.wondersByCityId[wonder.cityId] = [];
    state.index.wondersByCityId[wonder.cityId].push(wonder);
  });

  (state.hexImprovements || []).forEach((imp) => {
    state.index.improvementByHexId[imp.hexId] = imp;
  });
}

function indexById_(items) {
  const out = {};
  items.forEach((item) => {
    if (item && item.id !== undefined && item.id !== null) {
      out[item.id] = item;
    }
  });
  return out;
}

function rebuildStateDerivedFields_(state) {
  assignHexOwnershipFromCities_(state.hexes, state.cities);
  assignHexOccupants_(state.hexes, state.cities, state.units);
  buildIndexes_(state);
}

/*******************************************
 * ХОД ИГРЫ
 *******************************************/

function resolveTurnSafe() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    resolveTurn();
  } finally {
    lock.releaseLock();
  }
}

function resolveTurn() {
  const state = loadState();

  validateState_(state);
  validateOrders_(state);

  resolveOrders_(state);
  resolveCombatPhase_(state);
  resolveEconomy_(state);
  resolveCityGrowth_(state);
  resolveProduction_(state);
  resolveResearch_(state);
  refreshUnits_(state);
  finalizeTurn_(state);

  saveState(state);
  renderAll();

  SpreadsheetApp.getActive().toast(`Ход ${state.meta.turn} начался`);
}

/*******************************************
 * ВАЛИДАЦИЯ
 *******************************************/

function validateState_(state) {
  if (!state.meta) {
    throw new Error('Отсутствует мета-состояние игры');
  }

  state.cities.forEach((city) => {
    if (!state.index.playerById[city.playerId]) {
      throw new Error(`Город ${city.id} содержит неверный playerId ${city.playerId}`);
    }
    if (!state.index.hexById[city.hexId]) {
      throw new Error(`Город ${city.id} содержит неверный hexId ${city.hexId}`);
    }
    (city.buildings || []).forEach((buildingId) => {
      if (!state.index.buildingTypeById[buildingId]) {
        throw new Error(`Город ${city.id} содержит неверное здание ${buildingId}`);
      }
    });
  });

  state.units.forEach((unit) => {
    if (!state.index.playerById[unit.playerId]) {
      throw new Error(`Юнит ${unit.id} содержит неверный playerId ${unit.playerId}`);
    }
    if (!state.index.unitTypeById[unit.type]) {
      throw new Error(`Юнит ${unit.id} содержит неверный тип ${unit.type}`);
    }
    if (!state.index.hexById[unit.hexId]) {
      throw new Error(`Юнит ${unit.id} содержит неверный hexId ${unit.hexId}`);
    }
  });
}

function validateOrders_(state) {
  const currentTurn = state.meta.turn;
  const turnRule = state.index.ruleById['turn'] || {};
  const seenEntityOrders = new Set();

  state.orders.forEach((order) => {
    if (Number(order.turn) !== Number(currentTurn)) return;
    if ((order.status || 'pending') !== 'pending') return;

    if (!state.index.playerById[order.playerId]) {
      throw new Error(`Приказ ${order.id}: неизвестный игрок ${order.playerId}`);
    }

    if (order.entityType === 'city') {
      if (!state.index.cityById[order.entityId]) {
        throw new Error(`Приказ ${order.id}: город ${order.entityId} не найден`);
      }
    } else if (order.entityType === 'unit') {
      if (!state.index.unitById[order.entityId]) {
        throw new Error(`Приказ ${order.id}: юнит ${order.entityId} не найден`);
      }
    } else {
      throw new Error(`Приказ ${order.id}: неподдерживаемый entityType ${order.entityType}`);
    }

    if (turnRule.allowMultipleOrdersPerEntity === false) {
      const key = `${order.entityType}:${order.entityId}`;
      if (seenEntityOrders.has(key)) {
        throw new Error(`Несколько приказов для ${key} запрещены в одном ходу`);
      }
      seenEntityOrders.add(key);
    }
  });
}

/*******************************************
 * ПРИКАЗЫ
 *******************************************/

function importOrdersFromSheet() {
  bootstrapGameStorage();
  const state = loadState();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.ORDERS);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    SpreadsheetApp.getActive().toast('Нет приказов для импорта');
    return;
  }

  const orders = [];

  for (let i = 1; i < values.length; i++) {
    const raw = String(values[i][0] || '').trim();
    if (!raw) continue;

    const parsed = parseJsonSafe_(raw, null);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error(`Некорректный JSON приказа в строке ${i + 1}`);
    }

    if (!parsed.playerId) throw new Error(`В строке ${i + 1} отсутствует playerId`);
    if (!parsed.entityType) throw new Error(`В строке ${i + 1} отсутствует entityType`);
    if (!parsed.entityId) throw new Error(`В строке ${i + 1} отсутствует entityId`);
    if (!parsed.action) throw new Error(`В строке ${i + 1} отсутствует action`);

    const payload = parsed.payload && typeof parsed.payload === 'object' ? parsed.payload : {};

    if (payload.targetHexName && !payload.targetHexId) {
      const targetHex = getHexByName_(state, payload.targetHexName);
      if (!targetHex) {
        throw new Error(`В строке ${i + 1} не найден гекс с названием "${payload.targetHexName}"`);
      }
      payload.targetHexId = targetHex.id;
    }

    if (payload.hexName && !payload.hexId) {
      const hex = getHexByName_(state, payload.hexName);
      if (!hex) {
        throw new Error(`В строке ${i + 1} не найден гекс с названием "${payload.hexName}"`);
      }
      payload.hexId = hex.id;
    }

    orders.push({
      id: nextId_('O', state.orders.concat(orders)),
      turn: state.meta.turn,
      playerId: String(parsed.playerId),
      entityType: String(parsed.entityType),
      entityId: String(parsed.entityId),
      action: String(parsed.action),
      payload,
      status: 'pending',
      source: { sheetRow: i + 1 },
    });
  }

  state.orders = state.orders
    .filter((o) => Number(o.turn) !== Number(state.meta.turn) || (o.status || 'pending') !== 'pending')
    .concat(orders);

  saveState(state);
  SpreadsheetApp.getActive().toast(`Импортировано приказов: ${orders.length}`);
}

function clearOrdersSheet() {
  bootstrapGameStorage();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.ORDERS);
  const maxRows = Math.max(2, sheet.getMaxRows());
  sheet.getRange(2, 1, maxRows - 1, 1).clearContent();
  SpreadsheetApp.getActive().toast('Лист приказов очищен');
}

function resolveOrders_(state) {
  const currentTurn = state.meta.turn;

  state.orders.forEach((order) => {
    if (Number(order.turn) !== Number(currentTurn)) return;
    if ((order.status || 'pending') !== 'pending') return;

    try {
      if (order.entityType === 'city') {
        applyCityOrder_(state, order);
      } else if (order.entityType === 'unit') {
        applyUnitOrder_(state, order);
      } else {
        throw new Error(`Неподдерживаемый entityType ${order.entityType}`);
      }

      order.status = 'done';
    } catch (e) {
      order.status = 'failed';
      order.error = String(e && e.message ? e.message : e);

      appendLog_(state, {
        turn: currentTurn,
        type: 'order_failed',
        playerId: order.playerId,
        text: `Приказ ${order.id} не выполнен: ${order.error}`,
        payload: { orderId: order.id },
      });
    }
  });

  rebuildStateDerivedFields_(state);
}

function calcHexPurchaseCost_(dist, city) {
  const baseByDist = { 1: 20, 2: 35, 3: 70 };
  const base = baseByDist[dist] || 70;
  const owned = (city.ownedHexIds || []).length;
  return Math.round(base * (1 + owned * 0.15));
}

function applyCityOrder_buyHex_(state, city, payload) {
  const hex = payload.hexId ? state.index.hexById[payload.hexId] : getHexByName_(state, payload.hexName);
  if (!hex) throw new Error(`Гекс не найден: ${payload.hexName || payload.hexId}`);
  if (hex.ownerPlayerId) throw new Error(`Гекс ${hex.name} уже принадлежит игроку ${hex.ownerPlayerId}`);

  const cityHex = state.index.hexById[city.hexId];
  const dist = axialDistance_(cityHex.q, cityHex.r, hex.q, hex.r);
  if (dist > 3) throw new Error(`Гекс ${hex.name} слишком далеко (максимум 3 клетки)`);

  const cost = calcHexPurchaseCost_(dist, city);
  const player = state.index.playerById[city.playerId];
  if (Number(player.gold || 0) < cost) {
    throw new Error(`Недостаточно золота для покупки гекса ${hex.name}: нужно ${cost}, есть ${player.gold}`);
  }

  player.gold -= cost;
  hex.ownerPlayerId = city.playerId;
  hex.ownerCityId = city.id;
  city.ownedHexIds = city.ownedHexIds || [];
  if (!city.ownedHexIds.includes(hex.id)) city.ownedHexIds.push(hex.id);

  appendLog_(state, {
    turn: state.meta.turn,
    type: 'hex_purchased',
    playerId: city.playerId,
    text: `Город ${city.name} купил гекс ${hex.name} за ${cost} золота`,
    payload: { cityId: city.id, hexId: hex.id, cost },
  });
}

function applyCityOrder_buildDistrict_(state, city, payload) {
  const typeId = payload.districtTypeId;
  const def = state.index.districtTypeById[typeId];
  if (!def) throw new Error(`Неизвестный тип района: ${typeId}`);
  if (Number(city.population || 1) < Number(def.populationRequired || 0)) {
    throw new Error(`Для района "${def.nameRu}" нужно население ${def.populationRequired}, в городе ${city.population}`);
  }

  const existing = (state.index.districtsByCityId[city.id] || []).find((d) => d.typeId === typeId);
  if (existing) throw new Error(`Район "${def.nameRu}" уже построен в городе ${city.name}`);

  const hex = payload.hexId ? state.index.hexById[payload.hexId] : getHexByName_(state, payload.hexName);
  if (!hex) throw new Error(`Гекс не найден: ${payload.hexName || payload.hexId}`);
  if (hex.ownerCityId && hex.ownerCityId !== city.id) throw new Error(`Гекс ${hex.name} принадлежит другому городу`);
  if (state.index.districtsByHexId[hex.id]) throw new Error(`На гексе ${hex.name} уже есть район`);
  if (state.index.wonderByHexId[hex.id]) throw new Error(`На гексе ${hex.name} стоит чудо света`);
  if (def.requiresCoast) {
    const hasCoastNeighbor = getHexNeighbors_(state, hex.id).some((n) => n.terrain === 'coast' || n.terrain === 'water');
    if (!hasCoastNeighbor) throw new Error(`Район "${def.nameRu}" требует соседства с водой`);
  }

  state.districts.push({
    id: nextId_('D', state.districts),
    typeId,
    cityId: city.id,
    playerId: city.playerId,
    hexId: hex.id,
    buildings: [],
    turnsBuilt: state.meta.turn,
  });

  appendLog_(state, {
    turn: state.meta.turn,
    type: 'district_built',
    playerId: city.playerId,
    text: `Город ${city.name} построил район "${def.nameRu}" на гексе ${hex.name}`,
    payload: { cityId: city.id, typeId, hexId: hex.id },
  });

  rebuildStateDerivedFields_(state);
}

function applyCityOrder_(state, order) {
  const city = state.index.cityById[order.entityId];
  const payload = order.payload || {};

  switch (order.action) {
    case 'setProduction': {
      city.queue = Array.isArray(payload.queue) ? payload.queue : [];
      appendLog_(state, {
        turn: state.meta.turn,
        type: 'production',
        playerId: city.playerId,
        text: `Город ${city.name} изменил очередь производства`,
        payload: { cityId: city.id },
      });
      return;
    }

    case 'buyBuilding': {
      const player = state.index.playerById[city.playerId];
      const typeId = payload.typeId;
      const def = state.index.buildingTypeById[typeId];
      if (!def) throw new Error(`Неизвестный тип здания ${typeId}`);

      const cost = Number(def.cost || 0) * 2;
      if (Number(player.gold || 0) < cost) {
        throw new Error(`Недостаточно золота для покупки здания ${getBuildingNameRu_(state, typeId)}`);
      }
      if ((city.buildings || []).includes(typeId)) {
        throw new Error(`В городе ${city.name} уже есть ${getBuildingNameRu_(state, typeId)}`);
      }

      player.gold -= cost;
      city.buildings = city.buildings || [];
      city.buildings.push(typeId);

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'purchase',
        playerId: city.playerId,
        text: `Город ${city.name} купил здание ${getBuildingNameRu_(state, typeId)} за ${cost} золота`,
        payload: { cityId: city.id, typeId, cost },
      });
      return;
    }
    case 'buyHex':
      applyCityOrder_buyHex_(state, city, payload);
      return;
    case 'buildDistrict':
      applyCityOrder_buildDistrict_(state, city, payload);
      return;

    default:
      throw new Error(`Неподдерживаемое действие города: ${order.action}`);
  }
}

function applyUnitOrder_(state, order) {
  const unit = state.index.unitById[order.entityId];
  const payload = order.payload || {};

  switch (order.action) {
    case 'buildImprovement': {
      const typeId = payload.improvementTypeId;
      const def = state.index.hexImprovementTypeById[typeId];
      if (!def) throw new Error(`Неизвестный тип улучшения: ${typeId}`);

      const unitDef = state.index.unitTypeById[unit.type];
      if (!(unitDef.tags || []).includes('builder')) {
        throw new Error(`Юнит ${unit.id} не может строить улучшения (нужен тег 'builder')`);
      }

      const hex = state.index.hexById[unit.hexId];
      if (!hex) throw new Error('Гекс не найден');
      if (def.allowedTerrains && !def.allowedTerrains.includes(hex.terrain)) {
        throw new Error(`Улучшение "${def.nameRu}" нельзя строить на "${getTerrainNameRu_(hex.terrain)}"`);
      }
      if (def.requiredFeature && hex.feature !== def.requiredFeature) {
        throw new Error(`Улучшение "${def.nameRu}" требует особенности "${def.requiredFeature}"`);
      }
      if (def.requiredTech) {
        const player = state.index.playerById[unit.playerId];
        if (!(player.knownTechs || []).includes(def.requiredTech)) {
          throw new Error(`Для "${def.nameRu}" нужна технология "${def.requiredTech}"`);
        }
      }
      if (state.index.improvementByHexId[hex.id]) {
        throw new Error(`На гексе ${hex.name} уже есть улучшение`);
      }

      state.hexImprovements.push({
        id: nextId_('I', state.hexImprovements),
        typeId,
        hexId: hex.id,
        playerId: unit.playerId,
        turnsBuilt: state.meta.turn,
      });

      if (typeof unit.builderCharges === 'number') {
        unit.builderCharges -= 1;
        if (unit.builderCharges <= 0) {
          state.units = state.units.filter((u) => u.id !== unit.id);
        }
      }

      unit.movesLeft = 0;
      unit.status = 'building';
      appendLog_(state, {
        turn: state.meta.turn,
        type: 'improvement_built',
        playerId: unit.playerId,
        text: `Построено улучшение "${def.nameRu}" на гексе ${hex.name}`,
        payload: { unitId: unit.id, hexId: hex.id, typeId },
      });
      rebuildStateDerivedFields_(state);
      return;
    }
    case 'move': {
      const targetHexId = payload.targetHexId;
      if (!targetHexId) throw new Error('Не указан targetHexName или targetHexId');

      const path = findShortestHexPath_(state, unit.hexId, targetHexId, unit);
      if (!path || path.length < 2) {
        throw new Error(`Нет пути от ${getHexNameById_(state, unit.hexId)} до ${getHexNameById_(state, targetHexId)}`);
      }

      const nextHexId = path[1];
      const moveCost = getUnitMoveCostToHex_(state, unit, nextHexId);
      if (unit.movesLeft < moveCost) {
        throw new Error(`У юнита ${getUnitNameRu_(state, unit.type)} недостаточно очков перемещения`);
      }

      const occupyingEnemy = getEnemyUnitsOnHex_(state, nextHexId, unit.playerId);
      if (occupyingEnemy.length > 0) {
        throw new Error(`Целевой гекс ${getHexNameById_(state, nextHexId)} занят вражескими юнитами`);
      }

      unit.hexId = nextHexId;
      unit.movesLeft -= moveCost;
      unit.status = 'moved';

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'movement',
        playerId: unit.playerId,
        text: `${getUnitNameRu_(state, unit.type)} ${unit.id} переместился в ${getHexNameById_(state, nextHexId)}`,
        payload: { unitId: unit.id, hexId: nextHexId },
      });
      return;
    }

    case 'moveTo': {
      const targetHexId = payload.targetHexId;
      if (!targetHexId) throw new Error('Не указан targetHexName или targetHexId');

      let moved = false;
      while (unit.movesLeft > 0 && unit.hexId !== targetHexId) {
        const path = findShortestHexPath_(state, unit.hexId, targetHexId, unit);
        if (!path || path.length < 2) break;

        const nextHexId = path[1];
        const moveCost = getUnitMoveCostToHex_(state, unit, nextHexId);
        if (unit.movesLeft < moveCost) break;

        const occupyingEnemy = getEnemyUnitsOnHex_(state, nextHexId, unit.playerId);
        if (occupyingEnemy.length > 0) break;

        unit.hexId = nextHexId;
        unit.movesLeft -= moveCost;
        moved = true;
      }

      if (!moved) {
        throw new Error(
          `${getUnitNameRu_(state, unit.type)} ${unit.id} не смог продвинуться к ${getHexNameById_(state, targetHexId)}`
        );
      }

      unit.status = 'moved';

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'movement',
        playerId: unit.playerId,
        text:
          `${getUnitNameRu_(state, unit.type)} ${unit.id} продвинулся к ${getHexNameById_(state, targetHexId)} ` +
          `и теперь находится в ${getHexNameById_(state, unit.hexId)}`,
        payload: { unitId: unit.id, targetHexId, hexId: unit.hexId },
      });
      return;
    }

    case 'attackUnit': {
      const targetUnitId = payload.targetUnitId;
      const target = state.index.unitById[targetUnitId];
      if (!target) throw new Error(`Целевой юнит ${targetUnitId} не найден`);
      if (target.playerId === unit.playerId) throw new Error('Нельзя атаковать собственный юнит');

      const attackerHex = state.index.hexById[unit.hexId];
      const defenderHex = state.index.hexById[target.hexId];
      const dist = axialDistance_(attackerHex.q, attackerHex.r, defenderHex.q, defenderHex.r);
      const unitDef = state.index.unitTypeById[unit.type];
      const range = Number((unitDef.combat && unitDef.combat.range) || 1);

      if (dist > range) throw new Error('Цель вне радиуса атаки');
      if (unit.movesLeft <= 0) throw new Error('У юнита нет очков хода');

      resolveUnitVsUnitCombat_(state, unit, target);
      unit.movesLeft = 0;
      unit.status = 'attacked';
      return;
    }

    case 'attackCity': {
      const targetCityId = payload.targetCityId;
      const targetCity = state.index.cityById[targetCityId];
      if (!targetCity) throw new Error(`Целевой город ${targetCityId} не найден`);
      if (targetCity.playerId === unit.playerId) throw new Error('Нельзя атаковать собственный город');

      const attackerHex = state.index.hexById[unit.hexId];
      const defenderHex = state.index.hexById[targetCity.hexId];
      const dist = axialDistance_(attackerHex.q, attackerHex.r, defenderHex.q, defenderHex.r);
      const unitDef = state.index.unitTypeById[unit.type];
      const range = Number((unitDef.combat && unitDef.combat.range) || 1);

      if (dist > range) throw new Error('Город вне радиуса атаки');
      if (unit.movesLeft <= 0) throw new Error('У юнита нет очков хода');

      resolveUnitVsCityCombat_(state, unit, targetCity);
      unit.movesLeft = 0;
      unit.status = 'attacked';
      return;
    }

    case 'foundCity': {
      const unitDef = state.index.unitTypeById[unit.type];
      const canFound = (unitDef.tags || []).includes('found_city');
      if (!canFound) throw new Error(`Юнит ${unit.id} не может основывать город`);

      if (state.index.cityByHexId[unit.hexId]) {
        throw new Error(`На гексе ${getHexNameById_(state, unit.hexId)} уже есть город`);
      }

      const hex = state.index.hexById[unit.hexId];
      if (!hex.passable) throw new Error('Нельзя основать город на непроходимом гексе');
      if (hex.terrain === 'mountain' || hex.terrain === 'water') {
        throw new Error(`Нельзя основать город на местности "${getTerrainNameRu_(hex.terrain)}"`);
      }

      const cityName = payload.name || `Город ${nextNumericId_('C', state.cities)}`;
      const newCity = {
        id: nextId_('C', state.cities),
        playerId: unit.playerId,
        name: cityName,
        hexId: unit.hexId,
        population: 1,
        foodStored: 0,
        productionStored: 0,
        housing: 3,
        amenities: 0,
        health: 200,
        buildings: ['palace'],
        districts: [],
        queue: [],
        specialists: [],
        modifiers: [],
        workedHexIds: getNearbyHexIds_(unit.hexId, state.hexes, 1).slice(0, 6),
      };

      state.cities.push(newCity);
      state.units = state.units.filter((u) => u.id !== unit.id);

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'city_founded',
        playerId: newCity.playerId,
        text: `Основан город ${cityName} на гексе ${getHexNameById_(state, newCity.hexId)}`,
        payload: { cityId: newCity.id, hexId: newCity.hexId },
      });

      rebuildStateDerivedFields_(state);
      return;
    }

    default:
      throw new Error(`Неподдерживаемое действие юнита: ${order.action}`);
  }
}

/*******************************************
 * БОЙ
 *******************************************/

function resolveCombatPhase_(state) {
  state.units = state.units.filter((u) => Number(u.hp || 0) > 0);
  state.cities = state.cities.filter((c) => Number(c.health || 0) > 0);
  const aliveCityIds = new Set(state.cities.map((c) => c.id));
  state.districts = (state.districts || []).filter((d) => aliveCityIds.has(d.cityId));
  state.hexImprovements = (state.hexImprovements || []).filter((i) => !i.cityId || aliveCityIds.has(i.cityId));
  // Чудеса не уничтожаются автоматически при падении города.
  rebuildStateDerivedFields_(state);
}

function resolveUnitVsUnitCombat_(state, attacker, defender) {
  const combatRule = state.index.ruleById['combat'] || {
    baseDamage: 24,
    minDamage: 6,
    maxDamage: 40,
  };

  const atkDef = state.index.unitTypeById[attacker.type];
  const defDef = state.index.unitTypeById[defender.type];

  const attackerStrength = Number(
    (atkDef.combat && (atkDef.combat.ranged > 0 ? atkDef.combat.ranged : atkDef.combat.melee)) || 0
  );
  const defenderStrength = Number((defDef.combat && defDef.combat.melee) || 0);

  const diff = attackerStrength - defenderStrength;
  let damageToDef = Number(combatRule.baseDamage || 24) + diff * 1.5;
  damageToDef = clamp_(Math.round(damageToDef), Number(combatRule.minDamage || 6), Number(combatRule.maxDamage || 40));

  const isRanged =
    Number((atkDef.combat && atkDef.combat.range) || 0) > 1 ||
    Number((atkDef.combat && atkDef.combat.ranged) || 0) > 0;

  let damageToAtk = 0;

  if (!isRanged) {
    damageToAtk = clamp_(
      Math.round(Number(combatRule.baseDamage || 24) - diff),
      Number(combatRule.minDamage || 6),
      Number(combatRule.maxDamage || 40)
    );
  }

  defender.hp = Number(defender.hp || 0) - damageToDef;
  attacker.hp = Number(attacker.hp || 0) - damageToAtk;

  appendLog_(state, {
    turn: state.meta.turn,
    type: 'combat',
    playerId: attacker.playerId,
    text:
      `${getUnitNameRu_(state, attacker.type)} ${attacker.id} атаковал ` +
      `${getUnitNameRu_(state, defender.type)} ${defender.id}: нанесено ${damageToDef}, получено ${damageToAtk}`,
    payload: {
      attackerId: attacker.id,
      defenderId: defender.id,
      damageToDef,
      damageToAtk,
    },
  });
}

function resolveUnitVsCityCombat_(state, attacker, city) {
  const combatRule = state.index.ruleById['combat'] || {
    baseDamage: 24,
    minDamage: 6,
    maxDamage: 40,
    cityBaseStrength: 12,
  };

  const atkDef = state.index.unitTypeById[attacker.type];
  const attackerStrength = Number(
    (atkDef.combat && (atkDef.combat.ranged > 0 ? atkDef.combat.ranged : atkDef.combat.melee)) || 0
  );
  const cityStrength =
    Number(combatRule.cityBaseStrength || 12) +
    Math.floor(Number(city.population || 1) * 1.5) +
    Math.floor((city.buildings || []).length / 2);

  const diff = attackerStrength - cityStrength;
  const damageToCity = clamp_(
    Math.round(Number(combatRule.baseDamage || 24) + diff),
    Number(combatRule.minDamage || 6),
    Number(combatRule.maxDamage || 40)
  );

  const isRanged =
    Number((atkDef.combat && atkDef.combat.range) || 0) > 1 ||
    Number((atkDef.combat && atkDef.combat.ranged) || 0) > 0;

  const damageToAtk = isRanged
    ? 0
    : clamp_(
        Math.round(Number(combatRule.baseDamage || 24) - diff),
        Number(combatRule.minDamage || 6),
        Number(combatRule.maxDamage || 40)
      );

  city.health = Number(city.health || 0) - damageToCity;
  attacker.hp = Number(attacker.hp || 0) - damageToAtk;

  appendLog_(state, {
    turn: state.meta.turn,
    type: 'city_combat',
    playerId: attacker.playerId,
    text:
      `${getUnitNameRu_(state, attacker.type)} ${attacker.id} атаковал город ${city.name}: ` +
      `нанесено городу ${damageToCity}, получено ${damageToAtk}`,
    payload: {
      attackerId: attacker.id,
      cityId: city.id,
      damageToCity,
      damageToAtk,
    },
  });
}

/*******************************************
 * ЭКОНОМИКА
 *******************************************/

function resolveEconomy_(state) {
  state.players.forEach((player) => {
    player.goldPerTurn = 0;
    player.sciencePerTurn = 0;
    player.culturePerTurn = 0;
    player.faithPerTurn = 0;
  });

  state.cities.forEach((city) => {
    const player = state.index.playerById[city.playerId];
    const yields = getCityYield_(state, city);
    city._turnYields = yields;

    player.goldPerTurn += Number(yields.gold || 0);
    player.sciencePerTurn += Number(yields.science || 0);
    player.culturePerTurn += Number(yields.culture || 0);
    player.faithPerTurn += Number(yields.faith || 0);

    city.foodStored = Number(city.foodStored || 0) + Number(yields.food || 0);
    city.productionStored = Number(city.productionStored || 0) + Number(yields.production || 0);
  });

  state.players.forEach((player) => {
    player.gold = Number(player.gold || 0) + Number(player.goldPerTurn || 0);
    player.scienceStock = Number(player.scienceStock || 0) + Number(player.sciencePerTurn || 0);
    player.cultureStock = Number(player.cultureStock || 0) + Number(player.culturePerTurn || 0);
    player.faithStock = Number(player.faithStock || 0) + Number(player.faithPerTurn || 0);
  });
}

function getCityYield_(state, city) {
  const econRule = state.index.ruleById['economy'] || {};
  const yields = {
    food: Number(econRule.baseCityFood || 2) + Number(city.population || 1) * Number(econRule.foodPerPop || 1),
    production:
      Number(econRule.baseCityProduction || 1) +
      Math.floor(Number(city.population || 1) / 2) * Number(econRule.productionPerTwoPop || 1),
    gold: Number(econRule.baseCityGold || 1),
    science: Number(econRule.baseCityScience || 1) + Number(city.population || 1) * Number(econRule.sciencePerPop || 1),
    culture: 0,
    faith: 0,
  };

  const workedHexIds = Array.isArray(city.workedHexIds) ? city.workedHexIds : [];
  workedHexIds.forEach((hexId) => {
    const hex = state.index.hexById[hexId];
    if (!hex) return;
    addYield_(yields, getHexYield_(hex));
    const imp = state.index.improvementByHexId[hexId];
    if (imp) {
      const impType = state.index.hexImprovementTypeById[imp.typeId];
      if (impType && impType.yields) addYield_(yields, impType.yields);
    }
  });

  (city.buildings || []).forEach((buildingId) => {
    const def = state.index.buildingTypeById[buildingId];
    if (def && def.yields) {
      addYield_(yields, def.yields);
    }
  });

  const districts = state.index.districtsByCityId[city.id] || [];
  districts.forEach((district) => {
    const distDef = state.index.districtTypeById[district.typeId];
    if (distDef && distDef.yields) addYield_(yields, distDef.yields);
    (district.buildings || []).forEach((buildingId) => {
      const def = state.index.buildingTypeById[buildingId];
      if (def && def.yields) addYield_(yields, def.yields);
    });
  });

  const wonders = state.index.wondersByCityId[city.id] || [];
  wonders.forEach((wonder) => {
    const wDef = state.index.wonderTypeById[wonder.typeId];
    if (wDef && wDef.yields) addYield_(yields, wDef.yields);
  });

  return yields;
}

function getHexYield_(hex) {
  const y = { food: 0, production: 0, gold: 0, science: 0, culture: 0, faith: 0 };

  switch (hex.terrain) {
    case 'grassland':
      y.food += 2;
      break;
    case 'plains':
      y.food += 1;
      y.production += 1;
      break;
    case 'desert':
      y.gold += 1;
      break;
    case 'tundra':
      y.food += 1;
      break;
    case 'coast':
      y.food += 1;
      y.gold += 1;
      break;
    case 'hill':
      y.production += 2;
      break;
  }

  switch (hex.feature) {
    case 'forest':
      y.production += 1;
      break;
    case 'jungle':
      y.food += 1;
      break;
    case 'marsh':
      y.food += 1;
      break;
  }

  if (hex.resource === 'wheat') y.food += 1;
  if (hex.resource === 'horses') y.production += 1;
  if (hex.resource === 'gold') y.gold += 2;

  return y;
}

function addYield_(base, extra) {
  Object.keys(extra || {}).forEach((k) => {
    base[k] = Number(base[k] || 0) + Number(extra[k] || 0);
  });
}

/*******************************************
 * РОСТ / ПРОИЗВОДСТВО / ИССЛЕДОВАНИЯ
 *******************************************/

function resolveCityGrowth_(state) {
  const econRule = state.index.ruleById['economy'] || {
    baseFoodToGrow: 15,
    growthFactor: 1.15,
  };

  state.cities.forEach((city) => {
    const pop = Number(city.population || 1);
    const needed = Math.floor(
      Number(econRule.baseFoodToGrow || 15) * Math.pow(Number(econRule.growthFactor || 1.15), pop - 1)
    );

    if (Number(city.foodStored || 0) >= needed) {
      city.foodStored -= needed;
      city.population = pop + 1;

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'growth',
        playerId: city.playerId,
        text: `Город ${city.name} вырос до населения ${city.population}`,
        payload: { cityId: city.id, population: city.population },
      });
    }
  });
}

function resolveProduction_(state) {
  state.cities.forEach((city) => {
    const queue = city.queue || [];
    if (!queue.length) return;

    const current = queue[0];
    const cost = getProductionCost_(state, current);

    if (Number(city.productionStored || 0) >= cost) {
      city.productionStored -= cost;
      finishProductionItem_(state, city, current);
      city.queue.shift();
    }
  });
}

function getProductionCost_(state, item) {
  if (item.kind === 'unit') {
    const def = state.index.unitTypeById[item.typeId];
    if (!def) throw new Error(`Неизвестный тип юнита ${item.typeId}`);
    return Number(def.cost || 999999);
  }
  if (item.kind === 'building') {
    const def = state.index.buildingTypeById[item.typeId];
    if (!def) throw new Error(`Неизвестный тип здания ${item.typeId}`);
    return Number(def.cost || 999999);
  }
  if (item.kind === 'wonder') {
    const def = state.index.wonderTypeById[item.typeId];
    if (!def) throw new Error(`Неизвестное чудо ${item.typeId}`);
    return Number(def.cost || 999999);
  }
  if (item.kind === 'district') {
    const def = state.index.districtTypeById[item.typeId];
    if (!def) throw new Error(`Неизвестный тип района ${item.typeId}`);
    return Number(def.cost || 999999);
  }
  throw new Error(`Неизвестный вид производства ${item.kind}`);
}

function finishProductionItem_(state, city, item) {
  if (item.kind === 'wonder') {
    const def = state.index.wonderTypeById[item.typeId];
    if (!def) throw new Error(`Неизвестный тип чуда: ${item.typeId}`);
    if (state.index.wonderByTypeId[item.typeId]) throw new Error(`Чудо "${def.nameRu}" уже построено в мире`);
    const hex = item.hexId ? state.index.hexById[item.hexId] : getHexByName_(state, item.hexName);
    if (!hex) throw new Error(`Для чуда "${def.nameRu}" не указан гекс`);
    if (state.index.districtsByHexId[hex.id]) throw new Error(`На гексе ${hex.name} уже стоит район`);
    if (state.index.wonderByHexId[hex.id]) throw new Error(`На гексе ${hex.name} уже стоит другое чудо`);
    if (def.requiresCoast) {
      const hasCoast = getHexNeighbors_(state, hex.id).some((n) => n.terrain === 'coast' || n.terrain === 'water');
      if (!hasCoast) throw new Error(`Чудо "${def.nameRu}" требует гекса рядом с водой`);
    }
    state.wonders.push({
      id: nextId_('W', state.wonders),
      typeId: item.typeId,
      cityId: city.id,
      playerId: city.playerId,
      hexId: hex.id,
      turnsBuilt: state.meta.turn,
    });
    Object.keys(def.cityBonus || {}).forEach((k) => {
      city[k] = Number(city[k] || 0) + Number(def.cityBonus[k] || 0);
    });
    appendLog_(state, {
      turn: state.meta.turn,
      type: 'wonder_built',
      playerId: city.playerId,
      text: `${city.name} завершил строительство чуда "${def.nameRu}"! ${def.description || ''}`,
      payload: { cityId: city.id, typeId: item.typeId, hexId: hex.id },
    });
    rebuildStateDerivedFields_(state);
    return;
  }

  if (item.kind === 'district') {
    applyCityOrder_buildDistrict_(state, city, { districtTypeId: item.typeId, hexId: item.hexId, hexName: item.hexName });
    return;
  }

  if (item.kind === 'building') {
    city.buildings = city.buildings || [];
    if (!city.buildings.includes(item.typeId)) {
      city.buildings.push(item.typeId);
    }

    appendLog_(state, {
      turn: state.meta.turn,
      type: 'production_complete',
      playerId: city.playerId,
      text: `Город ${city.name} завершил строительство: ${getBuildingNameRu_(state, item.typeId)}`,
      payload: { cityId: city.id, typeId: item.typeId },
    });
    return;
  }

  if (item.kind === 'unit') {
    const def = state.index.unitTypeById[item.typeId];
    const newUnit = {
      id: nextId_('U', state.units),
      playerId: city.playerId,
      type: item.typeId,
      hexId: city.hexId,
      hp: 100,
      movesLeft: Number((def.movement && def.movement.max) || 2),
      status: 'idle',
      xp: 0,
      promotions: [],
      task: null,
      builderCharges: typeof def.builderCharges === 'number' ? def.builderCharges : undefined,
    };

    state.units.push(newUnit);

    appendLog_(state, {
      turn: state.meta.turn,
      type: 'production_complete',
      playerId: city.playerId,
      text: `Город ${city.name} создал юнит: ${getUnitNameRu_(state, item.typeId)}`,
      payload: { cityId: city.id, unitId: newUnit.id, typeId: item.typeId },
    });
    return;
  }

  throw new Error(`Неподдерживаемый вид производства ${item.kind}`);
}

function resolveResearch_(state) {
  state.players.forEach((player) => {
    if (!player.currentResearch) return;

    const tech = state.index.techTypeById[player.currentResearch];
    if (!tech) return;

    player.researchProgress = Number(player.researchProgress || 0) + Number(player.sciencePerTurn || 0);

    if (Number(player.researchProgress || 0) >= Number(tech.cost || 0)) {
      player.knownTechs = player.knownTechs || [];
      if (!player.knownTechs.includes(tech.id)) {
        player.knownTechs.push(tech.id);
      }

      appendLog_(state, {
        turn: state.meta.turn,
        type: 'research_complete',
        playerId: player.id,
        text: `${player.name} завершил исследование: ${getTechNameRu_(state, tech.id)}`,
        payload: { techId: tech.id },
      });

      player.currentResearch = null;
      player.researchProgress = 0;
    }
  });
}

/*******************************************
 * ЗАВЕРШЕНИЕ ХОДА
 *******************************************/

function refreshUnits_(state) {
  const turnRule = state.index.ruleById['turn'] || {};
  if (!turnRule.refreshUnitMovesEachTurn) return;

  state.units.forEach((unit) => {
    const def = state.index.unitTypeById[unit.type];
    unit.movesLeft = Number((def.movement && def.movement.max) || 2);
    if (unit.status === 'moved' || unit.status === 'attacked') {
      unit.status = 'idle';
    }
  });
}

function finalizeTurn_(state) {
  state.meta.turn = Number(state.meta.turn || 1) + 1;
  state.meta.phase = 'player';

  state.orders = state.orders.filter((order) => {
    return Number(order.turn) >= Number(state.meta.turn) - 1;
  });

  appendLog_(state, {
    turn: state.meta.turn,
    type: 'turn',
    playerId: null,
    text: `Начался ход ${state.meta.turn}`,
    payload: {},
  });
}

function appendLog_(state, entry) {
  state.log.push(entry);
  const maxLog = 8000;
  if (state.log.length > maxLog) {
    state.log = state.log.slice(state.log.length - maxLog);
  }
}

/*******************************************
 * ГЕКСЫ
 *******************************************/

function generateDemoHexMap_(width, height) {
  const hexes = [];

  for (let r = 0; r < height; r++) {
    for (let q = 0; q < width; q++) {
      const id = `H_${q}_${r}`;
      const terrain = getDemoTerrain_(q, r, width, height);
      const feature = getDemoFeature_(q, r);
      const resource = getDemoResource_(q, r);
      const name = generateHexName_(q, r);

      hexes.push({
        id,
        name,
        q,
        r,
        terrain,
        feature,
        resource,
        passable: terrain !== 'mountain' && terrain !== 'water',
        ownerPlayerId: null,
        ownerCityId: null,
        unitIds: [],
        cityId: null,
      });
    }
  }

  return hexes;
}

function generateHexName_(q, r) {
  return `Гекс ${q}:${r}`;
}

function normalizeHexName_(name) {
  return String(name || '').trim().toLowerCase();
}

function getHexByName_(state, name) {
  return state.index.hexByName[normalizeHexName_(name)] || null;
}

function getHexNameById_(state, hexId) {
  const hex = state.index.hexById[hexId];
  return hex ? hex.name : hexId;
}

function getDemoTerrain_(q, r, width, height) {
  if (q === 0 || r === 0 || q === width - 1 || r === height - 1) return 'coast';
  if ((q + r) % 11 === 0) return 'mountain';
  if ((q * 2 + r) % 9 === 0) return 'hill';
  if ((q + r * 2) % 7 === 0) return 'plains';
  if ((q + r) % 5 === 0) return 'desert';
  return 'grassland';
}

function getDemoFeature_(q, r) {
  if ((q + r) % 6 === 0) return 'forest';
  if ((q * 3 + r) % 13 === 0) return 'jungle';
  return null;
}

function getDemoResource_(q, r) {
  if ((q + r) % 10 === 0) return 'wheat';
  if ((q * 2 + r) % 12 === 0) return 'horses';
  if ((q + r * 4) % 17 === 0) return 'gold';
  return null;
}

function assignHexOwnershipFromCities_(hexes, cities) {
  const hexById = indexById_(hexes);

  hexes.forEach((hex) => {
    hex.ownerCityId = null;
    hex.ownerPlayerId = null;
    hex.cityId = null;
  });

  cities.forEach((city) => {
    const cityHex = hexById[city.hexId];
    if (cityHex) {
      cityHex.cityId = city.id;
      cityHex.ownerCityId = city.id;
      cityHex.ownerPlayerId = city.playerId;
    }

    const radius1 = getNearbyHexIds_(city.hexId, hexes, 1);
    radius1.concat(city.ownedHexIds || []).forEach((hexId) => {
      const hex = hexById[hexId];
      if (!hex) return;
      if (!hex.ownerCityId) {
        hex.ownerCityId = city.id;
        hex.ownerPlayerId = city.playerId;
      }
    });
  });
}

function assignHexOccupants_(hexes, cities, units) {
  const hexById = indexById_(hexes);

  hexes.forEach((hex) => {
    hex.unitIds = [];
    hex.cityId = null;
  });

  cities.forEach((city) => {
    const hex = hexById[city.hexId];
    if (hex) hex.cityId = city.id;
  });

  units.forEach((unit) => {
    const hex = hexById[unit.hexId];
    if (hex) {
      hex.unitIds = hex.unitIds || [];
      hex.unitIds.push(unit.id);
    }
  });
}

function getHexNeighbors_(state, hexId) {
  const hex = state.index.hexById[hexId];
  if (!hex) return [];

  const out = [];
  CONFIG.HEX_DIRECTIONS_AXIAL.forEach((dir) => {
    const q = Number(hex.q) + dir.dq;
    const r = Number(hex.r) + dir.dr;
    const neighborId = `H_${q}_${r}`;
    const neighbor = state.index.hexById[neighborId];
    if (neighbor) out.push(neighbor);
  });
  return out;
}

function getNearbyHexIds_(hexId, hexes, radius) {
  const hexById = indexById_(hexes);
  const origin = hexById[hexId];
  if (!origin) return [];

  const out = [];
  hexes.forEach((hex) => {
    const dist = axialDistance_(origin.q, origin.r, hex.q, hex.r);
    if (dist <= radius) {
      out.push(hex.id);
    }
  });
  return out;
}

function axialDistance_(q1, r1, q2, r2) {
  return (Math.abs(q1 - q2) + Math.abs(q1 + r1 - q2 - r2) + Math.abs(r1 - r2)) / 2;
}

function findShortestHexPath_(state, startHexId, targetHexId, unit) {
  if (startHexId === targetHexId) return [startHexId];

  const queue = [startHexId];
  const visited = new Set([startHexId]);
  const prev = {};

  while (queue.length) {
    const current = queue.shift();
    const neighbors = getHexNeighbors_(state, current);

    for (const neighbor of neighbors) {
      if (visited.has(neighbor.id)) continue;
      if (!canUnitEnterHex_(state, unit, neighbor.id)) continue;

      visited.add(neighbor.id);
      prev[neighbor.id] = current;

      if (neighbor.id === targetHexId) {
        return reconstructPath_(prev, startHexId, targetHexId);
      }

      queue.push(neighbor.id);
    }
  }

  return null;
}

function reconstructPath_(prev, start, target) {
  const path = [target];
  let current = target;

  while (current !== start) {
    current = prev[current];
    if (!current) return null;
    path.push(current);
  }

  path.reverse();
  return path;
}

function canUnitEnterHex_(state, unit, hexId) {
  const hex = state.index.hexById[hexId];
  if (!hex) return false;
  if (!hex.passable) return false;

  const enemyUnits = getEnemyUnitsOnHex_(state, hexId, unit.playerId);
  if (enemyUnits.length > 0) return false;

  return true;
}

function getEnemyUnitsOnHex_(state, hexId, playerId) {
  const units = state.index.unitsByHexId[hexId] || [];
  return units.filter((u) => u.playerId !== playerId);
}

function getUnitMoveCostToHex_(state, unit, hexId) {
  const hex = state.index.hexById[hexId];
  if (!hex) return 999;

  const moveRule = state.index.ruleById['movement'] || {};
  const terrainCostTable = moveRule.terrainCost || {};
  const featureCostTable = moveRule.featureCost || {};

  let cost = Number(terrainCostTable[hex.terrain] || 1);
  if (hex.feature) {
    cost += Number(featureCostTable[hex.feature] || 0);
  }
  const impOnHex = state.index.improvementByHexId ? state.index.improvementByHexId[hexId] : null;
  if (impOnHex) {
    const impType = state.index.hexImprovementTypeById[impOnHex.typeId];
    if (impType && impType.movementBonus) {
      cost = Math.max(1, cost + Number(impType.movementBonus));
    }
  }

  return Math.max(1, cost);
}

/*******************************************
 * ОТРИСОВКА ОСНОВНОЙ КАРТЫ
 *******************************************/

function renderMap() {
  bootstrapGameStorage();
  const state = loadState();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.MAP);
  sheet.clear();

  const width = Number(state.meta.map.width || 0);
  const height = Number(state.meta.map.height || 0);
  const startRow = CONFIG.MAP_RENDER.START_ROW;
  const startCol = CONFIG.MAP_RENDER.START_COL;

  ensureSheetSize_(sheet, startRow + height + 10, startCol + width + 10);

  sheet.getRange(1, 1).setValue(`Карта — ход ${state.meta.turn}`);
  sheet.getRange(1, 1).setFontWeight('bold');

  for (let q = 0; q < width; q++) {
    sheet.getRange(startRow - 1, startCol + q).setValue(q);
  }

  for (let r = 0; r < height; r++) {
    sheet.getRange(startRow + r, startCol - 1).setValue(r);
  }

  const values = [];
  const backgrounds = [];
  const notes = [];

  for (let r = 0; r < height; r++) {
    const rowValues = [];
    const rowBackgrounds = [];
    const rowNotes = [];

    for (let q = 0; q < width; q++) {
      const hexId = `H_${q}_${r}`;
      const hex = state.index.hexById[hexId];

      rowValues.push(getHexRenderText_(state, hex));
      rowBackgrounds.push(getHexBackgroundColor_(hex));
      rowNotes.push(getHexNote_(state, hex));
    }

    values.push(rowValues);
    backgrounds.push(rowBackgrounds);
    notes.push(rowNotes);
  }

  const range = sheet.getRange(startRow, startCol, height, width);
  range.setValues(values);
  range.setBackgrounds(backgrounds);
  range.setNotes(notes);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontWeight('bold');
  range.setWrap(true);

  for (let c = 0; c < width; c++) {
    sheet.setColumnWidth(startCol + c, 90);
  }
  for (let r = 0; r < height; r++) {
    sheet.setRowHeight(startRow + r, 52);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function getHexRenderText_(state, hex) {
  const city = state.index.cityByHexId[hex.id];
  const units = state.index.unitsByHexId[hex.id] || [];
  const district = state.index.districtsByHexId ? state.index.districtsByHexId[hex.id] : null;
  const wonder = state.index.wonderByHexId ? state.index.wonderByHexId[hex.id] : null;
  const improvement = state.index.improvementByHexId ? state.index.improvementByHexId[hex.id] : null;

  if (city) return `🏛 ${city.name}`;
  if (wonder) {
    const wDef = state.index.wonderTypeById ? state.index.wonderTypeById[wonder.typeId] : null;
    return `${(wDef && wDef.icon) || '✨'} ${wDef ? wDef.nameRu : wonder.typeId}`;
  }
  if (district) {
    const dDef = state.index.districtTypeById ? state.index.districtTypeById[district.typeId] : null;
    return `${(dDef && dDef.icon) || '🏗'} ${dDef ? dDef.nameRu : district.typeId}`;
  }
  if (units.length > 0) {
    if (units.length === 1) return `⚔ ${getUnitNameRu_(state, units[0].type)}`;
    return `⚔ x${units.length}`;
  }
  if (improvement) {
    const iDef = state.index.hexImprovementTypeById ? state.index.hexImprovementTypeById[improvement.typeId] : null;
    return [terrainShortRu_(hex.terrain), (iDef && iDef.icon) || '🔧'].join(' ');
  }

  const t = terrainShortRu_(hex.terrain);
  const f = hex.feature ? featureShortRu_(hex.feature) : '';
  const r = hex.resource ? resourceShortRu_(hex.resource) : '';
  return [t, f, r].filter(Boolean).join(' ');
}

function getHexBackgroundColor_(hex) {
  switch (hex.terrain) {
    case 'grassland':
      return '#b7e1a1';
    case 'plains':
      return '#f6e3a1';
    case 'desert':
      return '#f4d68f';
    case 'tundra':
      return '#d9e2f3';
    case 'coast':
      return '#9fc5e8';
    case 'hill':
      return '#d5c4a1';
    case 'mountain':
      return '#999999';
    case 'water':
      return '#6fa8dc';
    default:
      return '#ffffff';
  }
}

function getHexNote_(state, hex) {
  const city = state.index.cityByHexId[hex.id];
  const units = state.index.unitsByHexId[hex.id] || [];
  const district = state.index.districtsByHexId ? state.index.districtsByHexId[hex.id] : null;
  const wonder = state.index.wonderByHexId ? state.index.wonderByHexId[hex.id] : null;
  const improvement = state.index.improvementByHexId ? state.index.improvementByHexId[hex.id] : null;
  const parts = [
    `Гекс: ${hex.name}`,
    `ID: ${hex.id}`,
    `q=${hex.q}, r=${hex.r}`,
    `Местность: ${getTerrainNameRu_(hex.terrain)}`,
    `Особенность: ${getFeatureNameRu_(hex.feature)}`,
    `Ресурс: ${getResourceNameRu_(hex.resource)}`,
    `Владелец-игрок: ${hex.ownerPlayerId || '-'}`,
    `Владелец-город: ${hex.ownerCityId || '-'}`,
  ];

  if (city) parts.push(`Город: ${city.name} (${city.id})`);
  if (district) {
    const dDef = state.index.districtTypeById ? state.index.districtTypeById[district.typeId] : null;
    parts.push(`Район: ${dDef ? dDef.nameRu : district.typeId} (${district.id})`);
    if (district.buildings && district.buildings.length) {
      parts.push(`  Здания: ${district.buildings.map((b) => getBuildingNameRu_(state, b)).join(', ')}`);
    }
  }
  if (wonder) {
    const wDef = state.index.wonderTypeById ? state.index.wonderTypeById[wonder.typeId] : null;
    parts.push(`Чудо: ${wDef ? wDef.nameRu : wonder.typeId} (${wonder.id})`);
  }
  if (improvement) {
    const iDef = state.index.hexImprovementTypeById ? state.index.hexImprovementTypeById[improvement.typeId] : null;
    parts.push(`Улучшение: ${iDef ? iDef.nameRu : improvement.typeId}`);
  }
  if (units.length) {
    parts.push(`Юниты: ${units.map((u) => `${u.id}:${getUnitNameRu_(state, u.type)}:${u.playerId}`).join(', ')}`);
  }

  return parts.join('\n');
}

/*******************************************
 * ОТРИСОВКА ПОЛИТИЧЕСКОЙ КАРТЫ
 *******************************************/

function renderPoliticalMap() {
  bootstrapGameStorage();
  const state = loadState();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.POLITICAL_MAP);
  sheet.clear();

  const width = Number(state.meta.map.width || 0);
  const height = Number(state.meta.map.height || 0);
  const startRow = CONFIG.MAP_RENDER.START_ROW;
  const startCol = CONFIG.MAP_RENDER.START_COL;

  ensureSheetSize_(sheet, startRow + height + 10, startCol + width + 10);

  sheet.getRange(1, 1).setValue(`Политическая карта — ход ${state.meta.turn}`);
  sheet.getRange(1, 1).setFontWeight('bold');

  for (let q = 0; q < width; q++) {
    sheet.getRange(startRow - 1, startCol + q).setValue(q);
  }

  for (let r = 0; r < height; r++) {
    sheet.getRange(startRow + r, startCol - 1).setValue(r);
  }

  const values = [];
  const backgrounds = [];
  const fontColors = [];
  const notes = [];

  for (let r = 0; r < height; r++) {
    const rowValues = [];
    const rowBackgrounds = [];
    const rowFontColors = [];
    const rowNotes = [];

    for (let q = 0; q < width; q++) {
      const hexId = `H_${q}_${r}`;
      const hex = state.index.hexById[hexId];

      rowValues.push(getPoliticalHexText_(state, hex));
      rowBackgrounds.push(getPoliticalHexBackground_(state, hex));
      rowFontColors.push(getPoliticalHexFontColor_(state, hex));
      rowNotes.push(getPoliticalHexNote_(state, hex));
    }

    values.push(rowValues);
    backgrounds.push(rowBackgrounds);
    fontColors.push(rowFontColors);
    notes.push(rowNotes);
  }

  const range = sheet.getRange(startRow, startCol, height, width);
  range.setValues(values);
  range.setBackgrounds(backgrounds);
  range.setFontColors(fontColors);
  range.setNotes(notes);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontWeight('bold');
  range.setWrap(true);

  for (let c = 0; c < width; c++) {
    sheet.setColumnWidth(startCol + c, 90);
  }
  for (let r = 0; r < height; r++) {
    sheet.setRowHeight(startRow + r, 52);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function getPoliticalHexText_(state, hex) {
  const city = state.index.cityByHexId[hex.id];
  const units = state.index.unitsByHexId[hex.id] || [];

  if (city) {
    return `🏛 ${city.name}`;
  }

  if (units.length > 0) {
    return `⚔ ${getUnitNameRu_(state, units[0].type)}`;
  }

  if (!hex.ownerPlayerId) {
    return '—';
  }

  return '';
}

function getPoliticalHexBackground_(state, hex) {
  if (!hex.ownerPlayerId) {
    return '#d9d9d9';
  }

  const player = state.index.playerById[hex.ownerPlayerId];
  if (!player) {
    return '#d9d9d9';
  }

  return player.color || '#d9d9d9';
}

function getPoliticalHexFontColor_(state, hex) {
  if (!hex.ownerPlayerId) {
    return '#000000';
  }

  const player = state.index.playerById[hex.ownerPlayerId];
  if (!player) {
    return '#000000';
  }

  return player.textColor || '#000000';
}

function getPoliticalHexNote_(state, hex) {
  const city = state.index.cityByHexId[hex.id];
  const units = state.index.unitsByHexId[hex.id] || [];

  const ownerPlayer = hex.ownerPlayerId
    ? state.index.playerById[hex.ownerPlayerId]
    : null;

  const ownerCity = hex.ownerCityId
    ? state.index.cityById[hex.ownerCityId]
    : null;

  const parts = [
    `Гекс: ${hex.name}`,
    `ID: ${hex.id}`,
    `Владелец: ${ownerPlayer ? ownerPlayer.name : 'Нейтральный'}`,
    `Город-владелец: ${ownerCity ? ownerCity.name : '-'}`,
    `Город на гексе: ${city ? city.name : '-'}`,
    `Юниты: ${units.length ? units.map((u) => `${u.id}:${getUnitNameRu_(state, u.type)}`).join(', ') : '-'}`,
  ];

  return parts.join('\n');
}

/*******************************************
 * ОТЧЁТЫ
 *******************************************/

function renderReports() {
  bootstrapGameStorage();
  const state = loadState();
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.REPORTS);
  sheet.clear();

  let row = 1;

  sheet.getRange(row, 1).setValue(`Ход ${state.meta.turn}`);
  sheet.getRange(row, 1).setFontWeight('bold');
  row += 2;

  sheet.getRange(row, 1).setValue('Игроки');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;

  const playerHeaders = [
    'id',
    'название',
    'лидер',
    'золото',
    'золото/ход',
    'наука/ход',
    'культура/ход',
    'вера/ход',
    'текущее исследование',
    'прогресс исследования',
  ];
  sheet.getRange(row, 1, 1, playerHeaders.length).setValues([playerHeaders]).setFontWeight('bold');
  row++;

  const playerValues = state.players.map((p) => [
    p.id,
    p.name,
    p.leader || '',
    p.gold,
    p.goldPerTurn,
    p.sciencePerTurn,
    p.culturePerTurn,
    p.faithPerTurn,
    p.currentResearch ? getTechNameRu_(state, p.currentResearch) : '',
    p.researchProgress || 0,
  ]);
  if (playerValues.length) {
    sheet.getRange(row, 1, playerValues.length, playerHeaders.length).setValues(playerValues);
    row += playerValues.length + 2;
  }

  sheet.getRange(row, 1).setValue('Города');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;

  const cityHeaders = [
    'id',
    'игрок',
    'название',
    'гекс',
    'население',
    'еда',
    'производство',
    'жильё',
    'здоровье',
    'здания',
    'очередь',
  ];
  sheet.getRange(row, 1, 1, cityHeaders.length).setValues([cityHeaders]).setFontWeight('bold');
  row++;

  const cityValues = state.cities.map((c) => [
    c.id,
    c.playerId,
    c.name,
    getHexNameById_(state, c.hexId),
    c.population,
    c.foodStored,
    c.productionStored,
    c.housing,
    c.health,
    (c.buildings || []).map((b) => getBuildingNameRu_(state, b)).join(', '),
    formatQueueRu_(state, c.queue || []),
  ]);
  if (cityValues.length) {
    sheet.getRange(row, 1, cityValues.length, cityHeaders.length).setValues(cityValues);
    row += cityValues.length + 2;
  }

  sheet.getRange(row, 1).setValue('Юниты');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;

  const unitHeaders = [
    'id',
    'игрок',
    'тип',
    'гекс',
    'здоровье',
    'очки хода',
    'статус',
  ];
  sheet.getRange(row, 1, 1, unitHeaders.length).setValues([unitHeaders]).setFontWeight('bold');
  row++;

  const unitValues = state.units.map((u) => [
    u.id,
    u.playerId,
    getUnitNameRu_(state, u.type),
    getHexNameById_(state, u.hexId),
    u.hp,
    u.movesLeft,
    getUnitStatusRu_(u.status),
  ]);
  if (unitValues.length) {
    sheet.getRange(row, 1, unitValues.length, unitHeaders.length).setValues(unitValues);
    row += unitValues.length + 2;
  }

  row = renderReportsExtraTables_(state, sheet, row);

  sheet.getRange(row, 1).setValue('Последние события');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;

  const logHeaders = ['ход', 'тип', 'игрок', 'текст'];
  sheet.getRange(row, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight('bold');
  row++;

  const recentLog = state.log.slice(-40).map((e) => [
    e.turn,
    e.type,
    e.playerId || '',
    e.text,
  ]);
  if (recentLog.length) {
    sheet.getRange(row, 1, recentLog.length, logHeaders.length).setValues(recentLog);
  }

  sheet.autoResizeColumns(1, 17);
}

function renderReportsExtraTables_(state, sheet, startRow) {
  let row = startRow;

  sheet.getRange(row, 1).setValue('Районы').setFontWeight('bold');
  row++;
  const districtHeaders = ['id', 'город', 'тип', 'гекс', 'здания', 'ход постройки'];
  sheet.getRange(row, 1, 1, districtHeaders.length).setValues([districtHeaders]).setFontWeight('bold');
  row++;
  const districtValues = (state.districts || []).map((d) => {
    const dDef = state.index.districtTypeById[d.typeId];
    return [d.id, d.cityId, dDef ? dDef.nameRu : d.typeId, getHexNameById_(state, d.hexId), (d.buildings || []).map((b) => getBuildingNameRu_(state, b)).join(', '), d.turnsBuilt || ''];
  });
  if (districtValues.length) {
    sheet.getRange(row, 1, districtValues.length, districtHeaders.length).setValues(districtValues);
    row += districtValues.length;
  }
  row += 2;

  sheet.getRange(row, 1).setValue('Чудеса света').setFontWeight('bold');
  row++;
  const wonderHeaders = ['id', 'тип', 'город', 'гекс', 'игрок', 'ход постройки'];
  sheet.getRange(row, 1, 1, wonderHeaders.length).setValues([wonderHeaders]).setFontWeight('bold');
  row++;
  const wonderValues = (state.wonders || []).map((w) => {
    const wDef = state.index.wonderTypeById[w.typeId];
    return [w.id, wDef ? `${wDef.icon || ''} ${wDef.nameRu}` : w.typeId, w.cityId, getHexNameById_(state, w.hexId), w.playerId, w.turnsBuilt || ''];
  });
  if (wonderValues.length) {
    sheet.getRange(row, 1, wonderValues.length, wonderHeaders.length).setValues(wonderValues);
    row += wonderValues.length;
  }
  row += 2;

  sheet.getRange(row, 1).setValue('Улучшения гексов').setFontWeight('bold');
  row++;
  const impHeaders = ['id', 'тип', 'гекс', 'игрок', 'ход постройки'];
  sheet.getRange(row, 1, 1, impHeaders.length).setValues([impHeaders]).setFontWeight('bold');
  row++;
  const impValues = (state.hexImprovements || []).map((imp) => {
    const iDef = state.index.hexImprovementTypeById[imp.typeId];
    return [imp.id, iDef ? `${iDef.icon || ''} ${iDef.nameRu}` : imp.typeId, getHexNameById_(state, imp.hexId), imp.playerId, imp.turnsBuilt || ''];
  });
  if (impValues.length) {
    sheet.getRange(row, 1, impValues.length, impHeaders.length).setValues(impValues);
    row += impValues.length;
  }

  return row + 2;
}

/*******************************************
 * ОБЩАЯ ОТРИСОВКА
 *******************************************/

function renderAll() {
  renderMap();
  renderPoliticalMap();
  renderReports();
}

/*******************************************
 * РУССКИЕ НАЗВАНИЯ
 *******************************************/

function getUnitNameRu_(state, typeId) {
  const def = state.index.unitTypeById[typeId];
  return def ? (def.nameRu || typeId) : typeId;
}

function getBuildingNameRu_(state, typeId) {
  const def = state.index.buildingTypeById[typeId];
  return def ? (def.nameRu || typeId) : typeId;
}

function getTechNameRu_(state, techId) {
  const def = state.index.techTypeById[techId];
  return def ? (def.nameRu || techId) : techId;
}

function getDistrictNameRu_(state, typeId) {
  const def = state.index.districtTypeById[typeId];
  return def ? (def.nameRu || typeId) : typeId;
}

function getWonderNameRu_(state, typeId) {
  const def = state.index.wonderTypeById[typeId];
  return def ? (def.nameRu || typeId) : typeId;
}

function getImprovementNameRu_(state, typeId) {
  const def = state.index.hexImprovementTypeById[typeId];
  return def ? (def.nameRu || typeId) : typeId;
}

function getTerrainNameRu_(terrain) {
  switch (terrain) {
    case 'grassland': return 'Луга';
    case 'plains': return 'Равнины';
    case 'desert': return 'Пустыня';
    case 'tundra': return 'Тундра';
    case 'coast': return 'Берег';
    case 'hill': return 'Холмы';
    case 'mountain': return 'Горы';
    case 'water': return 'Вода';
    default: return '-';
  }
}

function getFeatureNameRu_(feature) {
  switch (feature) {
    case 'forest': return 'Лес';
    case 'jungle': return 'Джунгли';
    case 'marsh': return 'Болото';
    case null:
    case undefined:
    case '':
      return '-';
    default:
      return feature;
  }
}

function getResourceNameRu_(resource) {
  switch (resource) {
    case 'wheat': return 'Пшеница';
    case 'horses': return 'Лошади';
    case 'gold': return 'Золото';
    case null:
    case undefined:
    case '':
      return '-';
    default:
      return resource;
  }
}

function getUnitStatusRu_(status) {
  switch (status) {
    case 'idle': return 'ожидает';
    case 'moved': return 'переместился';
    case 'attacked': return 'атаковал';
    case 'building': return 'строит улучшение';
    default: return status || '-';
  }
}

function formatQueueRu_(state, queue) {
  if (!queue || !queue.length) return '';
  return queue.map((item) => {
    if (item.kind === 'unit') return `Юнит: ${getUnitNameRu_(state, item.typeId)}`;
    if (item.kind === 'building') return `Здание: ${getBuildingNameRu_(state, item.typeId)}`;
    if (item.kind === 'district') return `Район: ${getDistrictNameRu_(state, item.typeId)}`;
    if (item.kind === 'wonder') return `Чудо: ${getWonderNameRu_(state, item.typeId)}`;
    return JSON.stringify(item);
  }).join(' → ');
}

function terrainShortRu_(terrain) {
  switch (terrain) {
    case 'grassland': return 'Луг';
    case 'plains': return 'Рав';
    case 'desert': return 'Пус';
    case 'tundra': return 'Тун';
    case 'coast': return 'Бер';
    case 'hill': return 'Хол';
    case 'mountain': return 'Гор';
    case 'water': return 'Вод';
    default: return '?';
  }
}

function featureShortRu_(feature) {
  switch (feature) {
    case 'forest': return '🌲';
    case 'jungle': return '🌿';
    case 'marsh': return '🟫';
    default: return '';
  }
}

function resourceShortRu_(resource) {
  switch (resource) {
    case 'wheat': return '🌾';
    case 'horses': return '🐎';
    case 'gold': return '🪙';
    default: return 'Рес';
  }
}

/*******************************************
 * ВСПОМОГАТЕЛЬНОЕ
 *******************************************/

function clamp_(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function nextId_(prefix, items) {
  return `${prefix}${nextNumericId_(prefix, items)}`;
}

function nextNumericId_(prefix, items) {
  let max = 0;
  items.forEach((item) => {
    const id = String(item.id || '');
    const m = id.match(new RegExp(`^${prefix}(\\d+)$`));
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max + 1;
}
