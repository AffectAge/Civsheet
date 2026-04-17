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

  setupCoreSheet_(coreSheet);
  setupMapSheet_(mapSheet);
  setupPoliticalMapSheet_(politicalMapSheet);
  setupOrdersSheet_(ordersSheet);
  setupReportsSheet_(reportsSheet);

  ensureAllNamedRanges_();
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
  ];

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
  };

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
    cityByHexId: {},
    unitsByHexId: {},
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

    default:
      throw new Error(`Неподдерживаемое действие города: ${order.action}`);
  }
}

function applyUnitOrder_(state, order) {
  const unit = state.index.unitById[order.entityId];
  const payload = order.payload || {};

  switch (order.action) {
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
  });

  (city.buildings || []).forEach((buildingId) => {
    const def = state.index.buildingTypeById[buildingId];
    if (def && def.yields) {
      addYield_(yields, def.yields);
    }
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
  throw new Error(`Неизвестный вид производства ${item.kind}`);
}

function finishProductionItem_(state, city, item) {
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
    radius1.forEach((hexId) => {
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

  if (city) return `🏛 ${city.name}`;
  if (units.length > 0) {
    if (units.length === 1) return `⚔ ${getUnitNameRu_(state, units[0].type)}`;
    return `⚔ x${units.length}`;
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

  sheet.autoResizeColumns(1, 14);
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
    default: return status || '-';
  }
}

function formatQueueRu_(state, queue) {
  if (!queue || !queue.length) return '';
  return queue.map((item) => {
    if (item.kind === 'unit') return `Юнит: ${getUnitNameRu_(state, item.typeId)}`;
    if (item.kind === 'building') return `Здание: ${getBuildingNameRu_(state, item.typeId)}`;
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
