const STORAGE_KEY = "fs42-scheduler-state-v4";
const LEGACY_KEYS = ["fs42-scheduler-state-v3", "fs42-scheduler-state-v2", "fs42-scheduler-state-v1"];
const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const CATEGORY_OPTIONS = ["Entertainment", "Events", "Lifestyle", "Movies", "Drama", "Commercials", "Continuity"];
const ITEM_TYPE_META = {
  Programme: { color: "#50bfff", defaultCategory: "Entertainment", commercial: false },
  "Ad Break": { color: "#ff8c4d", defaultCategory: "Commercials", commercial: true },
  "Ad Spot": { color: "#ffc15e", defaultCategory: "Commercials", commercial: true },
  Bumper: { color: "#9e7dff", defaultCategory: "Continuity", commercial: true },
  Promo: { color: "#59d2b1", defaultCategory: "Continuity", commercial: true },
  Continuity: { color: "#8898ba", defaultCategory: "Continuity", commercial: true },
};
const SLOT_COLORS = {
  Breakfast: "#f9844a",
  Daytime: "#4cc9f0",
  "Prime Time": "#ff4d8d",
  Late: "#7b6dff",
};
const IMPORTANCE_COLORS = {
  Critical: "#ff5d73",
  High: "#ffab5c",
  Medium: "#57d6b9",
  Low: "#7d91b9",
};
const NETWORK_TYPES = ["standard", "web", "guide", "loop", "streaming"];
const BREAK_STRATEGIES = ["standard", "end", "center"];
const SCHEDULE_INCREMENTS = [5, 10, 15, 30, 60];
const WORKSPACE_META = {
  schedule: {
    eyebrow: "Planning Board",
    title: "Linear Channel Scheduling",
    description: "Plan on the timeline first, then use Review and Export when the schedule is ready.",
  },
  channels: {
    eyebrow: "Channel Setup",
    title: "FS42 Station Config",
    description: "Edit the FS42 station_conf values that define each channel before export.",
  },
  review: {
    eyebrow: "Review Board",
    title: "Audit and Validation",
    description: "Inspect the table, conflicts, and readiness notes without crowding the planner.",
  },
  export: {
    eyebrow: "Export Handoff",
    title: "Readiness and Output",
    description: "Choose the export profile, check blockers, and generate JSON or CSV handoff files.",
  },
  help: {
    eyebrow: "Help / Settings",
    title: "Quick Workflow Guide",
    description: "Short workflow notes, theme settings, and local-first guidance.",
  },
};
const DEFAULT_SECTION_OPEN_STATE = {
  "planning-controls": false,
  "colour-legend": false,
  "channel-identity": true,
  "channel-schedule": true,
  "channel-content": false,
  "channel-runtime": false,
  "channel-advanced": false,
  "review-notes": false,
  "export-notes": false,
  "help-workflow": false,
  "help-appearance": false,
};
const DAY_START = 6 * 60;
const DAY_END = 24 * 60;
const SNAP_MINUTES = 5;
const MAGNETIC_THRESHOLD = 7;
const HOUR_WIDTH = getTimelineHourWidth();
const DEBUG_TIMELINE_DND = false;

const DEFAULT_CHANNELS = [
  createChannel("FS42 Main", "Entertainment", "#ff8a5b", "Flagship entertainment feed"),
  createChannel("FS42 Stories", "Factual", "#67d4ff", "Docs and factual blocks"),
  createChannel("FS42 Action", "Sport", "#9bff9f", "Sports and event windows"),
  createChannel("FS42 Extra", "Extensions", "#ffd86b", "Catch-up and themed nights"),
];

const elements = {
  workspaceButtons: Array.from(document.querySelectorAll("[data-workspace-button]")),
  heroWorkspaceLabel: document.getElementById("heroWorkspaceLabel"),
  heroWorkspaceTitle: document.getElementById("heroWorkspaceTitle"),
  heroWorkspaceDescription: document.getElementById("heroWorkspaceDescription"),
  viewMode: document.getElementById("viewMode"),
  timelineScale: document.getElementById("timelineScale"),
  colorMode: document.getElementById("colorMode"),
  exportProfile: document.getElementById("exportProfile"),
  dayFilter: document.getElementById("dayFilter"),
  channelFilter: document.getElementById("channelFilter"),
  channel: document.getElementById("channel"),
  day: document.getElementById("day"),
  showForm: document.getElementById("showForm"),
  showId: document.getElementById("showId"),
  title: document.getElementById("title"),
  itemType: document.getElementById("itemType"),
  category: document.getElementById("category"),
  start: document.getElementById("start"),
  duration: document.getElementById("duration"),
  importance: document.getElementById("importance"),
  watershedRestricted: document.getElementById("watershedRestricted"),
  primeTime: document.getElementById("primeTime"),
  mustRun: document.getElementById("mustRun"),
  slot: document.getElementById("slot"),
  blockGroup: document.getElementById("blockGroup"),
  assetCode: document.getElementById("assetCode"),
  regenerateAssetCode: document.getElementById("regenerateAssetCode"),
  notes: document.getElementById("notes"),
  itemValidation: document.getElementById("itemValidation"),
  formTitle: document.getElementById("formTitle"),
  formHint: document.getElementById("formHint"),
  resetForm: document.getElementById("resetForm"),
  duplicateItem: document.getElementById("duplicateItem"),
  channelForm: document.getElementById("channelForm"),
  channelId: document.getElementById("channelId"),
  channelName: document.getElementById("channelName"),
  channelGroup: document.getElementById("channelGroup"),
  channelTagline: document.getElementById("channelTagline"),
  channelNumber: document.getElementById("channelNumber"),
  networkType: document.getElementById("networkType"),
  scheduleIncrement: document.getElementById("scheduleIncrement"),
  breakStrategy: document.getElementById("breakStrategy"),
  commercialFree: document.getElementById("commercialFree"),
  breakDuration: document.getElementById("breakDuration"),
  contentDir: document.getElementById("contentDir"),
  commercialDir: document.getElementById("commercialDir"),
  bumpDir: document.getElementById("bumpDir"),
  clipShows: document.getElementById("clipShows"),
  signOffVideo: document.getElementById("signOffVideo"),
  offAirVideo: document.getElementById("offAirVideo"),
  standbyImage: document.getElementById("standbyImage"),
  beRightBackMedia: document.getElementById("beRightBackMedia"),
  logoDir: document.getElementById("logoDir"),
  showLogo: document.getElementById("showLogo"),
  defaultLogo: document.getElementById("defaultLogo"),
  logoPermanent: document.getElementById("logoPermanent"),
  channelMultiLogoMode: document.getElementById("channelMultiLogoMode"),
  channelMultiLogoPanel: document.getElementById("channelMultiLogoPanel"),
  channelMultiLogoProfile: document.getElementById("channelMultiLogoProfile"),
  channelColor: document.getElementById("channelColor"),
  resetChannelForm: document.getElementById("resetChannelForm"),
  channelList: document.getElementById("channelList"),
  legend: document.getElementById("legend"),
  legendModeLabel: document.getElementById("legendModeLabel"),
  exportReadiness: document.getElementById("exportReadiness"),
  timelineHeading: document.getElementById("timelineHeading"),
  timelineSubheading: document.getElementById("timelineSubheading"),
  selectedItemPanel: document.getElementById("selectedItemPanel"),
  timelineShell: document.getElementById("timelineShell"),
  timelineView: document.getElementById("timelineView"),
  tableView: document.getElementById("tableView"),
  tableBody: document.getElementById("tableBody"),
  showList: document.getElementById("showList"),
  insights: document.getElementById("insights"),
  statBlocks: document.getElementById("statBlocks"),
  statRuntime: document.getElementById("statRuntime"),
  statPrime: document.getElementById("statPrime"),
  seedButton: document.getElementById("seedButton"),
  seedButtonSchedule: document.getElementById("seedButtonSchedule"),
  exportCsv: document.getElementById("exportCsv"),
  exportJson: document.getElementById("exportJson"),
  collapsiblePanels: Array.from(document.querySelectorAll("[data-collapsible]")),
  themeMode: document.getElementById("themeMode"),
};

const dragState = {
  itemId: null,
  pointerOffsetX: 0,
  lastClientX: null,
  lastTimelineX: null,
  draggedItemSnapshot: null,
  lastEventType: null,
  lastTargetLabel: null,
  lastCurrentTargetLabel: null,
  lastClosestLaneLabel: null,
  dropAllowed: "unknown",
  dropBlockedReason: "-",
  lastDropCommitted: false,
  commitSnapshot: null,
  timingSnapshot: null,
};

const pointerDragState = {
  active: false,
  itemId: null,
  pointerId: null,
  startClientX: 0,
  startClientY: 0,
  originLaneEl: null,
  originChannelId: null,
  originCategory: null,
  originLaneLabel: "-",
  originStartMinutes: null,
  originStartTime: "-",
  originItemSnapshot: null,
  pointerOffsetX: 0,
  pointerOffsetY: 0,
  lastClientX: null,
  lastClientY: null,
  lastLaneEl: null,
  lastChannelId: null,
  lastCategory: null,
  lastVisibleLaneX: null,
  lastTimelineContentX: null,
  lastAdjustedX: null,
  lastPreSnapMinutes: null,
  lastSnappedMinutes: null,
  lastClampedMinutes: null,
  previewItem: null,
  pendingFrame: false,
  rafId: 0,
};

const strategicDragState = {
  armed: false,
  dragging: false,
  itemId: null,
  pointerId: null,
  mode: null,
  originItemSnapshot: null,
  originChannelId: null,
  originCategory: null,
  originDay: null,
  originPlanningWeek: null,
  lastTargetCellEl: null,
  lastTargetKey: "-",
  lastTargetChannelId: null,
  lastTargetCategory: null,
  lastTargetColumnValue: null,
  previewItem: null,
  startClientX: 0,
  startClientY: 0,
  pendingFrame: false,
  rafId: 0,
  recentClickSuppressionId: null,
};

let timelineDebugPanel = null;

const resizeState = {
  itemId: null,
  edge: null,
  startX: 0,
  startMinutes: 0,
  duration: 0,
};

let state = loadState();
let itemFormDirty = false;

init();

function init() {
  populateStaticSelects();
  bindEvents();
  if (DEBUG_TIMELINE_DND) initTimelineDebugPanel();
  syncControls();
  syncTheme();
  syncCollapsibleSections();
  resetItemForm();
  resetChannelForm();
  render();
}

function createChannel(nameOrConfig, group, color, tagline, multiLogoMode = false, multiLogoProfile = "") {
  const config =
    typeof nameOrConfig === "object" && nameOrConfig !== null
      ? nameOrConfig
      : {
          name: nameOrConfig,
          group,
          color,
          tagline,
          multiLogoMode,
          multiLogoProfile,
        };

  const name = String(config.name || "").trim() || "Untitled channel";
  const channelNumber = parsePositiveInteger(config.channelNumber);
  const pathBase = getChannelConfigBase(name, channelNumber || 1);

  return {
    id: crypto.randomUUID(),
    name,
    group: String(config.group || "").trim(),
    color: config.color || "#7d91b9",
    tagline: String(config.tagline || "").trim(),
    channelNumber: channelNumber || null,
    networkType: NETWORK_TYPES.includes(config.networkType) ? config.networkType : "standard",
    scheduleIncrement: parsePositiveInteger(config.scheduleIncrement) || 30,
    breakStrategy: BREAK_STRATEGIES.includes(config.breakStrategy) ? config.breakStrategy : "standard",
    commercialFree: typeof config.commercialFree === "boolean" ? config.commercialFree : false,
    breakDuration: parsePositiveInteger(config.breakDuration) || 120,
    contentDir: String(config.contentDir || "").trim() || `catalog/${pathBase}`,
    commercialDir: String(config.commercialDir || "").trim() || `commercial/${pathBase}`,
    bumpDir: String(config.bumpDir || "").trim() || `bump/${pathBase}`,
    clipShows: normalizeClipShowList(config.clipShows),
    signOffVideo: String(config.signOffVideo || "").trim() || "runtime/signoff.mp4",
    offAirVideo: String(config.offAirVideo || "").trim() || "runtime/off_air_pattern.mp4",
    standbyImage: String(config.standbyImage || "").trim() || "runtime/standby.png",
    beRightBackMedia: String(config.beRightBackMedia || "").trim() || "runtime/brb.png",
    logoDir: String(config.logoDir || "").trim() || `logos/${pathBase}`,
    showLogo: typeof config.showLogo === "boolean" ? config.showLogo : true,
    defaultLogo: String(config.defaultLogo || "").trim() || `${pathBase}.png`,
    logoPermanent: typeof config.logoPermanent === "boolean" ? config.logoPermanent : true,
    multiLogoMode: Boolean(config.multiLogoMode),
    multiLogoProfile: String(config.multiLogoProfile || "").trim(),
  };
}

function normalizeFlags(flags = {}) {
  return {
    watershedRestricted: Boolean(flags.watershedRestricted),
    primeTime: Boolean(flags.primeTime),
    mustRun: Boolean(flags.mustRun),
  };
}

function createItem(
  title,
  itemType,
  category,
  channelId,
  day,
  start,
  duration,
  importance,
  blockGroup,
  assetCode,
  flags = {},
  notes,
) {
  return {
    id: crypto.randomUUID(),
    title,
    itemType,
    category,
    channelId,
    day,
    start,
    duration,
    importance,
    slot: getSlotFromMinutes(timeToMinutes(start)),
    blockGroup,
    assetCode,
    weekOrder: flags.weekOrder ?? null,
    monthOrder: flags.monthOrder ?? null,
    planningWeek: flags.planningWeek ?? null,
    ...normalizeFlags(flags),
    notes,
  };
}

function createDefaultState() {
  const channels = normalizeChannels(structuredClone(DEFAULT_CHANNELS));
  return {
    workspace: "schedule",
    viewMode: "timeline",
    timelineScale: "day",
    colorMode: "channel",
    exportProfile: "fs42-native",
    theme: "dark",
    selectedDay: "Monday",
    selectedChannelId: "all",
    sectionOpenState: getDefaultSectionOpenState(),
    channels,
    items: seedStrategicOrders([
      createItem("FS42 Breakfast Live", "Programme", "Entertainment", channels[0].id, "Monday", "06:00", 180, "High", "Morning studio", "FS42-MAIN-001", {}, "Live breakfast block."),
      createItem("Breakfast Break A", "Ad Break", "Commercials", channels[0].id, "Monday", "07:28", 6, "Medium", "Breakfast ads", "FS42-ADS-101", {}, "Contains 4 spots and bumper."),
      createItem("Coffee Sponsor Bumper", "Bumper", "Continuity", channels[0].id, "Monday", "07:34", 1, "Medium", "Breakfast ads", "FS42-BUMP-007", {}, "Lead back into programme."),
      createItem("The Ad Vault", "Programme", "Lifestyle", channels[1].id, "Monday", "09:00", 60, "Medium", "Archive docs", "FS42-STO-011", {}, "Advert history strand."),
      createItem("Midday Promo Wheel", "Promo", "Continuity", channels[1].id, "Monday", "12:00", 3, "Low", "Promo junction", "FS42-PRO-031", {}, "Cross-channel promotion."),
      createItem("Matchday Reload", "Programme", "Events", channels[2].id, "Monday", "18:00", 120, "Critical", "Sport", "FS42-ACT-021", { primeTime: true, mustRun: true }, "Highest sponsorship priority."),
      createItem("Prime Time Break", "Ad Break", "Commercials", channels[2].id, "Monday", "18:55", 5, "High", "Prime ads", "FS42-ADS-202", { primeTime: true }, "Mid-match sponsored break."),
      createItem("Late Signal", "Programme", "Drama", channels[3].id, "Monday", "22:00", 90, "Low", "Overnight", "FS42-EXT-003", { watershedRestricted: true }, "Experimental and repeat-friendly slot."),
      createItem("Spotlight UK", "Programme", "Entertainment", channels[0].id, "Tuesday", "20:00", 60, "Critical", "Feature", "FS42-MAIN-010", { primeTime: true, mustRun: true }, "Signature weekly tentpole."),
    ]),
  };
}

function normalizeChannels(channels = []) {
  return channels.map((channel, index) => normalizeChannel(channel, index));
}

function normalizeChannel(channel = {}, index = 0) {
  const name = String(channel.name || "").trim() || "Untitled channel";
  const providedNumber = parsePositiveInteger(channel.channelNumber ?? channel.channel_number);
  const channelNumber = providedNumber || index + 1;
  const pathBase = getChannelConfigBase(name, channelNumber);
  return {
    id: channel.id || crypto.randomUUID(),
    name,
    group: String(channel.group || "").trim(),
    color: channel.color || "#7d91b9",
    tagline: String(channel.tagline || "").trim(),
    channelNumber,
    networkType: NETWORK_TYPES.includes(channel.networkType || channel.network_type) ? channel.networkType || channel.network_type : "standard",
    scheduleIncrement: parsePositiveInteger(channel.scheduleIncrement ?? channel.schedule_increment) || 30,
    breakStrategy: BREAK_STRATEGIES.includes(channel.breakStrategy || channel.break_strategy)
      ? channel.breakStrategy || channel.break_strategy
      : "standard",
    commercialFree:
      typeof channel.commercialFree === "boolean"
        ? channel.commercialFree
        : typeof channel.commercial_free === "boolean"
          ? channel.commercial_free
          : false,
    breakDuration: parsePositiveInteger(channel.breakDuration ?? channel.break_duration) || 120,
    contentDir: String(channel.contentDir || channel.content_dir || "").trim() || `catalog/${pathBase}`,
    commercialDir: String(channel.commercialDir || channel.commercial_dir || "").trim() || `commercial/${pathBase}`,
    bumpDir: String(channel.bumpDir || channel.bump_dir || "").trim() || `bump/${pathBase}`,
    clipShows: normalizeClipShowList(channel.clipShows ?? channel.clip_shows),
    signOffVideo: String(channel.signOffVideo || channel.sign_off_video || "").trim() || "runtime/signoff.mp4",
    offAirVideo: String(channel.offAirVideo || channel.off_air_video || "").trim() || "runtime/off_air_pattern.mp4",
    standbyImage: String(channel.standbyImage || channel.standby_image || "").trim() || "runtime/standby.png",
    beRightBackMedia:
      String(channel.beRightBackMedia || channel.be_right_back_media || "").trim() || "runtime/brb.png",
    logoDir: String(channel.logoDir || channel.logo_dir || "").trim() || `logos/${pathBase}`,
    showLogo:
      typeof channel.showLogo === "boolean"
        ? channel.showLogo
        : typeof channel.show_logo === "boolean"
          ? channel.show_logo
          : true,
    defaultLogo: String(channel.defaultLogo || channel.default_logo || "").trim() || `${pathBase}.png`,
    logoPermanent:
      typeof channel.logoPermanent === "boolean"
        ? channel.logoPermanent
        : typeof channel.logo_permanent === "boolean"
          ? channel.logo_permanent
          : true,
    multiLogoMode:
      typeof channel.multiLogoMode === "boolean"
        ? channel.multiLogoMode
        : typeof channel.multi_logo === "string"
          ? Boolean(channel.multi_logo.trim())
          : false,
    multiLogoProfile: String(channel.multiLogoProfile || channel.multi_logo || "").trim(),
  };
}

function getNextAvailableChannelNumber(channels = state.channels, excludeId = null) {
  const used = new Set(
    channels
      .filter((channel) => channel.id !== excludeId)
      .map((channel) => parsePositiveInteger(channel.channelNumber))
      .filter((value) => Number.isInteger(value)),
  );
  let candidate = 1;
  while (used.has(candidate)) candidate += 1;
  return candidate;
}

function parsePositiveInteger(value) {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : null;
}

function getChannelConfigBase(name, channelNumber) {
  const safeNumber = parsePositiveInteger(channelNumber) || 1;
  const safeName = slugify(name || "channel") || "channel";
  return `ch${String(safeNumber).padStart(2, "0")}_${safeName}`;
}

function normalizeClipShowList(value) {
  if (Array.isArray(value)) {
    return uniqueValues(value.map((entry) => String(entry || "").trim()).filter(Boolean));
  }
  if (typeof value === "string") {
    return uniqueValues(
      value
        .split(/[\n,]/)
        .map((entry) => String(entry || "").trim())
        .filter(Boolean),
    );
  }
  return [];
}

function formatClipShowList(value) {
  return normalizeClipShowList(value).join(", ");
}

function populateStaticSelects() {
  DAY_NAMES.forEach((day) => {
    elements.dayFilter.appendChild(createOption(day, day));
    elements.day.appendChild(createOption(day, day));
  });

  CATEGORY_OPTIONS.forEach((category) => {
    elements.category.appendChild(createOption(category, category));
  });

  NETWORK_TYPES.forEach((networkType) => {
    if (elements.networkType) {
      elements.networkType.appendChild(createOption(networkType, networkType));
    }
  });

  BREAK_STRATEGIES.forEach((strategy) => {
    if (elements.breakStrategy) {
      elements.breakStrategy.appendChild(createOption(strategy, strategy));
    }
  });

  SCHEDULE_INCREMENTS.forEach((increment) => {
    if (elements.scheduleIncrement) {
      elements.scheduleIncrement.appendChild(createOption(String(increment), `${increment} min`));
    }
  });
}

function bindEvents() {
  elements.workspaceButtons.forEach((button) => {
    button.addEventListener("click", () => handleWorkspaceChange(button.dataset.workspaceButton));
  });
  elements.viewMode.addEventListener("change", handleControlChange);
  elements.timelineScale.addEventListener("change", handleControlChange);
  elements.colorMode.addEventListener("change", handleControlChange);
  elements.exportProfile.addEventListener("change", handleControlChange);
  elements.dayFilter.addEventListener("change", handleControlChange);
  elements.channelFilter.addEventListener("change", handleControlChange);
  elements.itemType.addEventListener("change", handleItemTypeChange);
  elements.title.addEventListener("blur", syncNewItemExportCodeFromTitle);
  const markItemFormDirty = () => {
    itemFormDirty = true;
    updateItemValidationSummary();
  };
  elements.showForm.addEventListener("input", markItemFormDirty);
  elements.showForm.addEventListener("change", markItemFormDirty);

  elements.showForm.addEventListener("submit", handleItemSubmit);
  elements.resetForm.addEventListener("click", resetItemForm);
  elements.duplicateItem.addEventListener("click", () => duplicateActiveItem());
  elements.regenerateAssetCode?.addEventListener("click", regenerateActiveItemExportCode);
  elements.channelForm.addEventListener("submit", handleChannelSubmit);
  elements.resetChannelForm.addEventListener("click", resetChannelForm);
  elements.commercialFree.addEventListener("change", syncChannelCommercialFreeControls);
  elements.channelMultiLogoMode.addEventListener("change", syncChannelMultiLogoControls);
  const restoreDemoData = () => {
    state = createDefaultState();
    syncControls();
    resetItemForm();
    resetChannelForm();
    persistAndRender();
  };
  elements.seedButton?.addEventListener("click", restoreDemoData);
  elements.seedButtonSchedule?.addEventListener("click", restoreDemoData);
  elements.exportCsv.addEventListener("click", () => exportSchedule("csv"));
  elements.exportJson.addEventListener("click", () => exportSchedule("json"));
  elements.themeMode?.addEventListener("change", handleThemeChange);
  elements.collapsiblePanels.forEach((panel) => {
    panel.addEventListener("toggle", handleCollapsibleToggle);
  });

  document.addEventListener("pointermove", handleBlockPointerMove, true);
  document.addEventListener("pointerup", handleBlockPointerUp, true);
  document.addEventListener("pointercancel", handleBlockPointerCancel, true);
  document.addEventListener("pointermove", handleStrategicCardPointerMove, true);
  document.addEventListener("pointerup", handleStrategicCardPointerUp, true);
  document.addEventListener("pointercancel", handleStrategicCardPointerCancel, true);

  window.addEventListener("mousemove", handleResizeMove);
  window.addEventListener("mouseup", stopResize);
  window.addEventListener("keydown", handleKeyboardShortcuts);
}

function handleControlChange(event) {
  if (event.target === elements.viewMode) state.viewMode = event.target.value;
  if (event.target === elements.timelineScale) state.timelineScale = event.target.value;
  if (event.target === elements.colorMode) state.colorMode = event.target.value;
  if (event.target === elements.exportProfile) state.exportProfile = normalizeExportProfile(event.target.value);
  if (event.target === elements.dayFilter) state.selectedDay = event.target.value;
  if (event.target === elements.channelFilter) state.selectedChannelId = event.target.value;
  syncControls();
  persistAndRender();
}

function handleWorkspaceChange(workspace) {
  state.workspace = normalizeWorkspace(workspace);
  if (state.workspace === "review") {
    state.viewMode = "table";
  } else if (state.workspace === "schedule") {
    state.viewMode = "timeline";
  }
  syncControls();
  persistAndRender();
}

function handleThemeChange(event) {
  state.theme = normalizeTheme(event.target.value);
  syncTheme();
  persistState();
}

function handleCollapsibleToggle(event) {
  const panel = event.currentTarget;
  const id = panel?.dataset?.collapsibleId;
  if (!id) return;

  state.sectionOpenState = {
    ...getDefaultSectionOpenState(),
    ...normalizeSectionOpenState(state.sectionOpenState),
    [id]: Boolean(panel.open),
  };
  persistState();
}

function handleItemTypeChange() {
  const meta = ITEM_TYPE_META[elements.itemType.value];
  if (!meta) return;
  elements.category.value = meta.defaultCategory;
  if (meta.commercial) {
    elements.importance.value = "Medium";
    elements.duration.value = elements.itemType.value === "Bumper" ? "5" : elements.duration.value || "5";
  }
}

function syncNewItemExportCodeFromTitle() {
  if (elements.showId.value) return;
  const title = elements.title.value.trim();
  if (!title || elements.assetCode.value.trim()) return;
  elements.assetCode.value = getUniqueExportCode(generateDefaultExportCode(title), state.items);
  updateItemValidationSummary();
}

function handleItemSubmit(event) {
  event.preventDefault();
  const itemId = elements.showId.value || crypto.randomUUID();
  const title = elements.title.value.trim();
  const generatedAssetCode = generateDefaultExportCode(title);
  const manualAssetCode = elements.assetCode.value.trim();
  const assetCode = manualAssetCode || getUniqueExportCode(generatedAssetCode, state.items, itemId);
  const item = {
    id: itemId,
    title,
    itemType: elements.itemType.value,
    category: elements.category.value,
    channelId: elements.channel.value,
    day: elements.day.value,
    start: elements.start.value,
    duration: getSafeDuration(elements.duration.value, elements.itemType.value),
    importance: elements.importance.value,
    watershedRestricted: elements.watershedRestricted.checked,
    primeTime: elements.primeTime.checked,
    mustRun: elements.mustRun.checked,
    slot: elements.slot.value,
    blockGroup: elements.blockGroup.value.trim(),
    assetCode,
    notes: elements.notes.value.trim(),
  };

  const validation = validateItem(item, state.items, state.exportProfile);
  const blockers = validation.filter((entry) => entry.severity === "blocker");
  if (blockers.length > 0) {
    updateItemValidationSummary(item, validation);
    return;
  }

    if (!state.items.some((entry) => entry.id === item.id)) {
      item.weekOrder = getNextStrategicOrderValue(item, state.items, "week");
      item.monthOrder = getNextStrategicOrderValue(item, state.items, "month");
    } else {
      const existing = state.items.find((entry) => entry.id === item.id);
      item.weekOrder = Number.isFinite(existing?.weekOrder) ? existing.weekOrder : null;
      item.monthOrder = Number.isFinite(existing?.monthOrder) ? existing.monthOrder : null;
      item.planningWeek = Number.isInteger(existing?.planningWeek) ? existing.planningWeek : null;
    }

  const index = state.items.findIndex((entry) => entry.id === item.id);
  if (index >= 0) {
    state.items.splice(index, 1, item);
  } else {
    state.items.push(item);
  }

  state.selectedDay = item.day;
  state.selectedChannelId = "all";
  syncControls();
  resetItemForm();
  persistAndRender();
}

function handleChannelSubmit(event) {
  event.preventDefault();
  const channel = collectChannelFormState();

  const index = state.channels.findIndex((entry) => entry.id === channel.id);
  if (index >= 0) {
    state.channels.splice(index, 1, channel);
  } else {
    state.channels.push(channel);
  }

  resetChannelForm();
  if (!state.channels.some((entry) => entry.id === state.selectedChannelId)) {
    state.selectedChannelId = "all";
  }
  persistAndRender();
}

function refreshChannelSelects() {
  const currentFormValue = elements.channel.value;
  const currentFilterValue = state.selectedChannelId;
  elements.channel.innerHTML = "";
  elements.channelFilter.innerHTML = "";
  elements.channelFilter.appendChild(createOption("all", "All channels"));

  state.channels.forEach((channel) => {
    elements.channel.appendChild(createOption(channel.id, channel.name));
    elements.channelFilter.appendChild(createOption(channel.id, channel.name));
  });

  elements.channel.value = state.channels.some((entry) => entry.id === currentFormValue)
    ? currentFormValue
    : state.channels[0]?.id || "";
  elements.channelFilter.value = state.channels.some((entry) => entry.id === currentFilterValue)
    ? currentFilterValue
    : "all";
}

function resetItemForm() {
  itemFormDirty = false;
  elements.showForm.reset();
  refreshChannelSelects();
  elements.showId.value = "";
  elements.channel.value = state.channels[0]?.id || "";
  elements.day.value = state.selectedDay;
  elements.itemType.value = "Programme";
  elements.category.value = "Entertainment";
  elements.importance.value = "Critical";
  elements.watershedRestricted.checked = false;
  elements.primeTime.checked = false;
  elements.mustRun.checked = false;
  elements.start.value = "20:00";
  elements.duration.value = "60";
  elements.slot.value = "Prime Time";
  elements.formTitle.textContent = "Add Item";
  elements.formHint.textContent = "Programming, commercials, and bumpers";
  updateItemValidationSummary();
}

function resetChannelForm() {
  elements.channelForm.reset();
  elements.channelId.value = "";
  const draft = buildDefaultChannelDraft();
  applyChannelFormState(draft, true);
  syncChannelMultiLogoControls();
}

function syncChannelMultiLogoControls() {
  const enabled = Boolean(elements.channelMultiLogoMode.checked);
  if (elements.channelMultiLogoPanel) elements.channelMultiLogoPanel.hidden = !enabled;
  if (elements.channelMultiLogoProfile) elements.channelMultiLogoProfile.disabled = !enabled;
}

function syncChannelCommercialFreeControls() {
  const locked = Boolean(elements.commercialFree.checked);
  if (elements.breakStrategy) elements.breakStrategy.disabled = locked;
  if (elements.breakDuration) {
    elements.breakDuration.disabled = locked;
    if (parsePositiveInteger(elements.breakDuration.value) === null || parsePositiveInteger(elements.breakDuration.value) <= 0) {
      elements.breakDuration.value = "120";
    }
  }
}

function render() {
  renderWorkspaceChrome();
  syncCollapsibleSections();
  syncTheme();
  refreshChannelSelects();
  renderLegend();
  renderPlanner();
  renderSelectedItemPanel();
  renderTable();
  renderShowList();
  renderChannelList();
  renderInsights();
  renderStats();
  syncExportControls();
  updateItemValidationSummary();
}

function renderWorkspaceChrome() {
  const workspace = normalizeWorkspace(state.workspace);
  const meta = WORKSPACE_META[workspace] || WORKSPACE_META.schedule;
  state.workspace = workspace;
  document.body.dataset.workspace = workspace;

  if (elements.heroWorkspaceLabel) elements.heroWorkspaceLabel.textContent = meta.eyebrow;
  if (elements.heroWorkspaceTitle) elements.heroWorkspaceTitle.textContent = meta.title;
  if (elements.heroWorkspaceDescription) elements.heroWorkspaceDescription.textContent = meta.description;

  elements.workspaceButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.workspaceButton === workspace);
    button.setAttribute("aria-pressed", button.dataset.workspaceButton === workspace ? "true" : "false");
  });
}

function getDefaultSectionOpenState() {
  return { ...DEFAULT_SECTION_OPEN_STATE };
}

function normalizeSectionOpenState(value = {}) {
  const defaults = getDefaultSectionOpenState();
  const normalized = { ...defaults };

  Object.entries(value || {}).forEach(([key, entry]) => {
    if (Object.prototype.hasOwnProperty.call(defaults, key)) {
      normalized[key] = Boolean(entry);
    }
  });

  return normalized;
}

function syncCollapsibleSections() {
  const sections = elements.collapsiblePanels.length > 0 ? elements.collapsiblePanels : Array.from(document.querySelectorAll("[data-collapsible]"));
  const sectionOpenState = normalizeSectionOpenState(state.sectionOpenState);

  sections.forEach((panel) => {
    const id = panel.dataset.collapsibleId;
    if (!id) return;
    panel.open = Boolean(sectionOpenState[id]);
  });
}

function syncTheme() {
  const theme = normalizeTheme(state.theme);
  state.theme = theme;
  document.body.dataset.theme = theme;
  if (elements.themeMode && elements.themeMode.value !== theme) {
    elements.themeMode.value = theme;
  }
}

function renderLegend() {
  elements.legend.innerHTML = "";
  const items =
    state.colorMode === "channel"
      ? state.channels.map((channel) => ({ label: channel.name, color: channel.color }))
      : state.colorMode === "slot"
        ? Object.entries(SLOT_COLORS).map(([label, color]) => ({ label, color }))
        : state.colorMode === "importance"
          ? Object.entries(IMPORTANCE_COLORS).map(([label, color]) => ({ label, color }))
          : Object.entries(ITEM_TYPE_META).map(([label, meta]) => ({ label, color: meta.color }));

  elements.legendModeLabel.textContent =
    state.colorMode === "channel"
      ? "Channel"
      : state.colorMode === "slot"
        ? "Time Slot"
        : state.colorMode === "importance"
          ? "Importance"
          : "Item Type";

  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "legend-item";
    const swatch = document.createElement("span");
    swatch.className = "legend-swatch";
    swatch.style.background = item.color;
    const label = document.createElement("span");
    label.textContent = item.label;
    row.append(swatch, label);
    elements.legend.appendChild(row);
  });
}

function renderPlanner() {
  const timelineVisible = state.viewMode === "timeline";
  elements.timelineView.classList.toggle("is-hidden", !timelineVisible);
  elements.tableView.classList.toggle("is-hidden", timelineVisible);

  if (!timelineVisible) return;

  if (state.timelineScale === "day") {
    renderDayTimeline();
  } else {
    renderStrategicBoard(state.timelineScale);
  }
}

function renderSelectedItemPanel() {
  if (!elements.selectedItemPanel) return;

  const isSchedule = normalizeWorkspace(state.workspace) === "schedule";
  const item = isSchedule ? state.items.find((entry) => entry.id === elements.showId.value) || null : null;
  elements.selectedItemPanel.classList.toggle("is-hidden", !item);
  if (!item) {
    elements.selectedItemPanel.innerHTML = "";
    return;
  }

  const channel = state.channels.find((entry) => entry.id === item.channelId) || null;
  const issues = validateItem(item, state.items, state.exportProfile);
  const status = issues.length ? `<span class="status-pill ${getStatusClass(issues)}">${getStatusLabel(issues)}</span>` : "";
  const flags = getFlagList(item)
    .map((flag) => `<span class="flag-pill">${flag.label}</span>`)
    .join("");
  elements.selectedItemPanel.innerHTML = `
    <div class="selected-item-header">
      <div>
        <p class="selected-item-eyebrow">Selected item</p>
        <h4>${item.title}</h4>
        <p>${channel?.name || "Unknown channel"} | ${item.day} | ${item.category}</p>
      </div>
      <div class="selected-item-actions">
        ${status}
        <span class="flag-pill">${item.itemType}</span>
      </div>
    </div>
    <div class="selected-item-meta">
      <span><strong>Start:</strong> ${item.start}</span>
      <span><strong>Duration:</strong> ${item.duration} mins</span>
      <span><strong>Channel:</strong> ${channel?.name || "-"}</span>
      <span><strong>Day:</strong> ${item.day}</span>
      <span><strong>Category:</strong> ${item.category}</span>
    </div>
    <div class="selected-item-flags">${flags || "<span class=\"field-help\">No flags set.</span>"}</div>
    <div class="button-row selected-item-buttons">
      <button type="button" class="button tertiary edit-button">Edit</button>
      <button type="button" class="button tertiary duplicate-button">Duplicate</button>
      <button type="button" class="button tertiary danger delete-button">Delete</button>
    </div>
  `;
  elements.selectedItemPanel.querySelector(".edit-button")?.addEventListener("click", () => hydrateItemForm(item.id));
  elements.selectedItemPanel
    .querySelector(".duplicate-button")
    ?.addEventListener("click", () => duplicateItem(item.id));
  elements.selectedItemPanel.querySelector(".delete-button")?.addEventListener("click", () => deleteItem(item.id));
}

function renderDayTimeline() {
  const visibleChannels = getVisibleChannels();
  const visibleItems = getVisibleItems();
  const renderItems = getTimelineRenderItems(visibleItems);
  const conflictIds = new Set(getConflicts(renderItems));
  const validation = validateSchedule(state.exportProfile);

  elements.timelineHeading.textContent = `${state.selectedDay} Operational Schedule`;
  elements.timelineSubheading.textContent =
    "Blocks snap into lanes by channel and category. Drag to move, drag edges to change duration.";

  const shell = document.createElement("div");
  const header = document.createElement("div");
  header.className = "timeline-header";

  const headerLabel = document.createElement("div");
  headerLabel.className = "header-label";
  headerLabel.innerHTML = `<strong>Lane</strong><span>Channel / category</span>`;

  const scale = document.createElement("div");
  scale.className = "time-scale";
  for (let hour = 6; hour < 24; hour += 1) {
    const cell = document.createElement("div");
    cell.className = "hour-cell";
    cell.textContent = `${String(hour).padStart(2, "0")}:00`;
    scale.appendChild(cell);
  }
  header.append(headerLabel, scale);
  shell.appendChild(header);

  visibleChannels.forEach((channel) => {
    getLaneCategories(channel.id).forEach((category) => {
      const laneItems = renderItems
        .filter((item) => item.channelId === channel.id && item.category === category)
        .sort(compareItems);
      const layout = computeLaneLayout(laneItems);

      const row = document.createElement("div");
      row.className = "day-row";

      const label = document.createElement("div");
      label.className = "lane-label";
      label.innerHTML = `<strong>${channel.name}</strong><small>${channel.group || "Custom"} | ${category}</small>`;

      const lane = document.createElement("div");
      lane.className = "day-lane";
      lane.dataset.channelId = channel.id;
      lane.dataset.category = category;
      lane.style.minHeight = `${Math.max(54, layout.levels * 40 + 8)}px`;
      renderSlotBands(lane);

      laneItems.forEach((item) => {
        const block = createDayBlock(
          item,
          layout.positions.get(item.id) || 0,
          conflictIds.has(item.id),
          validation.itemIssues.get(item.id) || [],
        );
        lane.appendChild(block);
      });

      row.append(label, lane);
      shell.appendChild(row);
    });
  });

  elements.timelineShell.innerHTML = "";
  elements.timelineShell.appendChild(shell);
}

function renderStrategicBoard(scaleMode) {
  const visibleChannels = getVisibleChannels();
  const columns = scaleMode === "week" ? DAY_NAMES : ["Week 1", "Week 2", "Week 3", "Week 4"];
  const renderItems = getStrategicRenderItems(getVisibleItems());

  elements.timelineHeading.textContent =
    scaleMode === "week" ? "Weekly Planning Board" : "Monthly Planning Board";
  elements.timelineSubheading.textContent =
    scaleMode === "week"
      ? "Use this for recurring weekly shape and lane balance."
      : "Use this for recurring month-wide templates and campaign planning.";

  const shell = document.createElement("div");
  const header = document.createElement("div");
  header.className = "timeline-header";

  const headerLabel = document.createElement("div");
  headerLabel.className = "header-label";
  headerLabel.innerHTML = `<strong>Lane</strong><span>Strategic board</span>`;

  const boardHeader = document.createElement("div");
  boardHeader.className = `board-grid ${scaleMode}`;
  columns.forEach((column) => {
    const cell = document.createElement("div");
    cell.className = "board-cell-header";
    cell.dataset.columnValue = column;
    cell.dataset.columnIndex = String(columns.indexOf(column));
    cell.textContent = column;
    boardHeader.appendChild(cell);
  });
  header.append(headerLabel, boardHeader);
  shell.appendChild(header);

  visibleChannels.forEach((channel) => {
    getLaneCategories(channel.id).forEach((category) => {
      const row = document.createElement("div");
      row.className = "board-row";

      const label = document.createElement("div");
      label.className = "lane-label";
      label.innerHTML = `<strong>${channel.name}</strong><small>${channel.group || "Custom"} | ${category}</small>`;

      const board = document.createElement("div");
      board.className = `board-grid ${scaleMode}`;

      columns.forEach((_, index) => {
        const cell = document.createElement("div");
        cell.className = "board-cell";
        const columnValue = scaleMode === "week" ? columns[index] : String(index + 1);
        cell.dataset.channelId = channel.id;
        cell.dataset.category = category;
        cell.dataset.columnIndex = String(index);
        cell.dataset.columnValue = columnValue;
        const cellKey = getStrategicCellKey(channel.id, category, columnValue);
        cell.classList.toggle(
          "is-strategic-drop-target",
          strategicDragState.dragging && strategicDragState.lastTargetKey === cellKey,
        );
        const cellItems = getStrategicItems(renderItems, channel.id, category, scaleMode, index);

        if (cellItems.length === 0) {
          const tag = document.createElement("span");
          tag.className = "lane-tag";
          tag.textContent = scaleMode === "week" ? "No items" : "Template clear";
          cell.appendChild(tag);
        } else {
          cellItems.forEach((item) => {
            const card = document.createElement("button");
            card.type = "button";
            card.className = "mini-card";
            card.dataset.itemId = item.id;
            card.style.background = getItemColor(item);
            card.innerHTML = `<strong>${item.title}</strong><small>${item.day} | ${item.start} | ${item.itemType}</small>`;
            card.classList.toggle("is-strategic-dragging", strategicDragState.dragging && strategicDragState.itemId === item.id);
            card.addEventListener("pointerdown", handleStrategicCardPointerDown);
            card.addEventListener("click", handleStrategicCardClick);
            cell.appendChild(card);
          });
        }

        board.appendChild(cell);
      });

      row.append(label, board);
      shell.appendChild(row);
    });
  });

  elements.timelineShell.innerHTML = "";
  elements.timelineShell.appendChild(shell);
}

function getStrategicRenderItems(items) {
  if (!strategicDragState.dragging || !strategicDragState.previewItem) return items;
  return items.map((item) => (item.id === strategicDragState.itemId ? strategicDragState.previewItem : item));
}

function getStrategicCellKey(channelId, category, columnValue) {
  return `${channelId}|${category}|${columnValue}`;
}

function getStrategicMonthBucket(item) {
  const stored = Number(item.planningWeek);
  if (Number.isInteger(stored) && stored >= 1 && stored <= 4) return stored;
  const dayIndex = DAY_NAMES.indexOf(item.day);
  if (dayIndex < 0) return 1;
  return Math.min(4, Math.floor((dayIndex * 4) / DAY_NAMES.length) + 1);
}

function renderSlotBands(lane) {
  const bands = [
    { start: 6, end: 10, color: SLOT_COLORS.Breakfast },
    { start: 10, end: 18, color: SLOT_COLORS.Daytime },
    { start: 18, end: 22, color: SLOT_COLORS["Prime Time"] },
    { start: 22, end: 24, color: SLOT_COLORS.Late },
  ];

  bands.forEach((band) => {
    const bandElement = document.createElement("div");
    bandElement.className = "time-slot-band";
    bandElement.style.left = `${timelineMinutesToX(band.start * 60)}px`;
    bandElement.style.width = `${timelineMinutesToX(band.end * 60) - timelineMinutesToX(band.start * 60)}px`;
    bandElement.style.background = band.color;
    lane.appendChild(bandElement);
  });
}

function createDayBlock(item, level, hasConflict, issues = []) {
  const block = document.createElement("div");
  const startMinutes = clampPlanningMinutes(getSafeStartMinutes(item.start));
  const duration = getSafeDuration(item.duration, item.itemType);
  const left = timelineMinutesToX(startMinutes);
  const naturalWidth = (duration / 60) * HOUR_WIDTH;
  const width = Math.max(naturalWidth, duration <= 30 ? 24 : ITEM_TYPE_META[item.itemType]?.commercial ? 42 : 48);
  block.className = "schedule-block";
  block.dataset.itemId = item.id;
  block.dataset.channelId = item.channelId;
  block.dataset.category = item.category;
  block.title = `${item.title} | ${minutesToTime(startMinutes)} | ${duration} mins`;
  if (elements.showId.value === item.id) block.classList.add("is-active");
  if (pointerDragState.active && pointerDragState.itemId === item.id) block.classList.add("is-pointer-dragging");
  block.style.left = `${left}px`;
  block.style.top = `${6 + level * 38}px`;
  block.style.width = `${width}px`;
  block.style.background = getItemColor(item);
  if (hasConflict) block.classList.add("conflict");
  if (issues.some((issue) => issue.severity === "blocker")) block.classList.add("validation-blocked");
  else if (issues.some((issue) => issue.severity === "warning")) block.classList.add("validation-warning");
  if (ITEM_TYPE_META[item.itemType]?.commercial) block.classList.add("is-commercial");

  const content = document.createElement("div");
  content.className = "schedule-block-content";
  const flags = getFlagList(item)
    .map((flag) => `<span class="flag-pill">${flag.short}</span>`)
    .join("");
  const status = issues.length ? `<span class="status-pill ${getStatusClass(issues)}">${getStatusLabel(issues)}</span>` : "";
  content.innerHTML = `
    <h4>${item.title}</h4>
    <p>${minutesToTime(startMinutes)} | ${duration} mins | ${item.itemType}</p>
    <div class="schedule-block-flags">${status}${flags}</div>
  `;
  content.addEventListener("click", () => hydrateItemForm(item.id));

  const leftHandle = document.createElement("span");
  leftHandle.className = "resize-handle left";
  leftHandle.addEventListener("mousedown", (event) => startResize(event, item.id, "left"));

  const rightHandle = document.createElement("span");
  rightHandle.className = "resize-handle right";
  rightHandle.addEventListener("mousedown", (event) => startResize(event, item.id, "right"));

  block.append(leftHandle, content, rightHandle);
  block.addEventListener("pointerdown", handleBlockPointerDown);
  return block;
}

function getTimelineRenderItems(items) {
  if (!pointerDragState.active || !pointerDragState.previewItem) return items;
  return items.map((item) => (item.id === pointerDragState.itemId ? pointerDragState.previewItem : item));
}

function handleBlockPointerDown(event) {
  if (state.timelineScale !== "day") return;
  if (event.pointerType === "mouse" && event.button !== 0) return;
  if (event.target?.closest?.(".resize-handle")) return;

  const item = state.items.find((entry) => entry.id === event.currentTarget.dataset.itemId);
  if (!item) return;

  event.preventDefault();
  event.stopPropagation();

  const originLane = event.currentTarget.closest(".day-lane");
  const pointerContext = getPointerTimelineContext(event, originLane);
  const blockLeftX = getItemTimelineLeftX(item);

  pointerDragState.active = true;
  pointerDragState.itemId = item.id;
  pointerDragState.pointerId = event.pointerId ?? null;
  pointerDragState.startClientX = event.clientX;
  pointerDragState.startClientY = event.clientY;
  pointerDragState.originLaneEl = originLane;
  pointerDragState.originChannelId = item.channelId;
  pointerDragState.originCategory = item.category;
  pointerDragState.originLaneLabel = originLane ? `${getChannelName(item.channelId)} / ${item.category}` : "-";
  pointerDragState.originStartMinutes = getSafeStartMinutes(item.start);
  pointerDragState.originStartTime = item.start;
  pointerDragState.originItemSnapshot = {
    id: item.id,
    title: item.title,
    itemType: item.itemType,
    category: item.category,
    channelId: item.channelId,
    start: item.start,
    duration: item.duration,
    slot: item.slot,
    importance: item.importance,
    watershedRestricted: item.watershedRestricted,
    primeTime: item.primeTime,
    mustRun: item.mustRun,
    blockGroup: item.blockGroup,
    assetCode: item.assetCode,
    notes: item.notes,
  };
  pointerDragState.pointerOffsetX = Number.isFinite(pointerContext.timelineContentX)
    ? pointerContext.timelineContentX - blockLeftX
    : 0;
  pointerDragState.pointerOffsetY = Number.isFinite(pointerContext.clientY) && originLane
    ? pointerContext.clientY - originLane.getBoundingClientRect().top
    : 0;
  pointerDragState.lastClientX = pointerContext.clientX;
  pointerDragState.lastClientY = pointerContext.clientY;
  pointerDragState.lastLaneEl = originLane;
  pointerDragState.lastChannelId = item.channelId;
  pointerDragState.lastCategory = item.category;
  pointerDragState.lastVisibleLaneX = null;
  pointerDragState.lastTimelineContentX = null;
  pointerDragState.lastAdjustedX = null;
  pointerDragState.lastPreSnapMinutes = null;
  pointerDragState.lastSnappedMinutes = null;
  pointerDragState.lastClampedMinutes = null;
  pointerDragState.previewItem = { ...pointerDragState.originItemSnapshot };
  pointerDragState.pendingFrame = false;

  dragState.itemId = item.id;
  dragState.pointerOffsetX = pointerDragState.pointerOffsetX;
  dragState.lastClientX = pointerContext.clientX;
  dragState.lastTimelineX = null;
  dragState.draggedItemSnapshot = { ...pointerDragState.originItemSnapshot };
  dragState.commitSnapshot = null;
  dragState.timingSnapshot = {
    status: "dragging",
    hover: null,
    commit: null,
    final: null,
    postRender: null,
    matches: {
      hoverToCommit: "-",
      commitToSaved: "-",
      savedToRecheck: "-",
    },
    itemId: item.id,
    itemTitle: item.title,
    originalStartMinutes: pointerDragState.originStartMinutes,
    originalStartTime: item.start,
    revertedAfterCommit: "no",
  };

  document.body.classList.add("is-pointer-dragging");

  if (DEBUG_TIMELINE_DND) {
    const snapshot = buildTimelineDebugSnapshot("pointerdown", event, originLane, item, {
      dropStatus: "POINTER DOWN",
      dropFired: "no",
      dropCommitted: "no",
      dropAllowed: "unknown",
      dropBlockedReason: "-",
      targetLabel: describeTimelineNode(event.target),
      currentTargetLabel: describeTimelineNode(event.currentTarget),
      closestLaneLabel: describeTimelineNode(originLane),
      renderNote: "pointer drag started",
      timingSnapshot: dragState.timingSnapshot,
    });
    updateTimelineDebugPanel(snapshot);
      console.log("[timeline-dnd] pointerdown", {
        itemId: item.id,
        originalStart: item.start,
        originLane: pointerDragState.originLaneLabel,
        pointerOffsetX: pointerDragState.pointerOffsetX,
        blockLeftX,
        pointerTimelineX: pointerContext.timelineContentX,
      });
    }

  schedulePointerDragRender();
}

function handleBlockPointerMove(event) {
  if (!pointerDragState.active) return;
  if (pointerDragState.pointerId !== null && event.pointerId !== pointerDragState.pointerId) return;

  event.preventDefault();
  updatePointerDragPreviewFromEvent(event, "move");
  schedulePointerDragRender();
}

function handleBlockPointerUp(event) {
  if (!pointerDragState.active) return;
  if (pointerDragState.pointerId !== null && event.pointerId !== pointerDragState.pointerId) return;

  const movedX = Math.abs(event.clientX - pointerDragState.startClientX);
  const movedY = Math.abs(event.clientY - pointerDragState.startClientY);
  const clickThreshold = 6;

  event.preventDefault();

  // Treat a near-stationary pointer interaction as a click/select, not a drag.
  if (movedX < clickThreshold && movedY < clickThreshold) {
    const selectedItemId = pointerDragState.itemId;
    clearPointerDragState();
    hydrateItemForm(selectedItemId);
    return;
  }

  updatePointerDragPreviewFromEvent(event, "up");
  commitPointerDrag(event);
}

function handleBlockPointerCancel(event) {
  if (!pointerDragState.active) return;
  if (pointerDragState.pointerId !== null && event.pointerId !== pointerDragState.pointerId) return;
  cancelPointerDrag("pointercancel", event);
}

function handleStrategicCardPointerDown(event) {
  if (state.timelineScale === "day") return;
  if (event.pointerType === "mouse" && event.button !== 0) return;

  const item = state.items.find((entry) => entry.id === event.currentTarget.dataset.itemId);
  if (!item) return;

  strategicDragState.armed = true;
  strategicDragState.dragging = false;
  strategicDragState.itemId = item.id;
  strategicDragState.pointerId = event.pointerId ?? null;
  strategicDragState.mode = state.timelineScale;
  strategicDragState.originItemSnapshot = { ...item };
  strategicDragState.originChannelId = item.channelId;
  strategicDragState.originCategory = item.category;
  strategicDragState.originDay = item.day;
  strategicDragState.originPlanningWeek = getStrategicMonthBucket(item);
  strategicDragState.startClientX = event.clientX;
  strategicDragState.startClientY = event.clientY;
  strategicDragState.lastTargetChannelId = item.channelId;
  strategicDragState.lastTargetCategory = item.category;
  strategicDragState.lastTargetColumnValue = state.timelineScale === "week" ? item.day : String(strategicDragState.originPlanningWeek);
  strategicDragState.lastTargetKey = getStrategicCellKey(
    item.channelId,
    item.category,
    strategicDragState.lastTargetColumnValue,
  );
  strategicDragState.lastTargetCellEl = event.currentTarget.closest(".board-cell");
  strategicDragState.previewItem = { ...item };
  strategicDragState.pendingFrame = false;

  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {
    // Some browsers may not allow pointer capture on buttons.
  }

  if (DEBUG_TIMELINE_DND) {
    console.log("[strategic-dnd] pointerdown", {
      itemId: item.id,
      title: item.title,
      mode: state.timelineScale,
      day: item.day,
      planningWeek: item.planningWeek || "-",
    });
  }
}

function handleStrategicCardPointerMove(event) {
  if (!strategicDragState.armed) return;
  if (strategicDragState.pointerId !== null && event.pointerId !== strategicDragState.pointerId) return;

  const movedX = Math.abs(event.clientX - strategicDragState.startClientX);
  const movedY = Math.abs(event.clientY - strategicDragState.startClientY);
  if (!strategicDragState.dragging && movedX < 6 && movedY < 6) return;

  event.preventDefault();

  if (!strategicDragState.dragging) {
    strategicDragState.dragging = true;
    document.body.classList.add("is-strategic-dragging");
  }

  updateStrategicDragPreviewFromEvent(event);
  scheduleStrategicDragRender();
}

function handleStrategicCardPointerUp(event) {
  if (!strategicDragState.armed) return;
  if (strategicDragState.pointerId !== null && event.pointerId !== strategicDragState.pointerId) return;

  if (!strategicDragState.dragging) {
    clearStrategicDragState();
    return;
  }

  event.preventDefault();
  updateStrategicDragPreviewFromEvent(event);
  commitStrategicDrag(event);
}

function handleStrategicCardPointerCancel(event) {
  if (!strategicDragState.armed) return;
  if (strategicDragState.pointerId !== null && event.pointerId !== strategicDragState.pointerId) return;
  cancelStrategicDrag("pointercancel", event);
}

function handleStrategicCardClick(event) {
  if (strategicDragState.recentClickSuppressionId === event.currentTarget.dataset.itemId) {
    event.preventDefault();
    event.stopPropagation();
    return;
  }

  const itemId = event.currentTarget.dataset.itemId;
  hydrateItemForm(itemId);
}

function scheduleStrategicDragRender() {
  if (!strategicDragState.dragging || strategicDragState.pendingFrame) return;
  strategicDragState.pendingFrame = true;
  strategicDragState.rafId = window.requestAnimationFrame(() => {
    strategicDragState.pendingFrame = false;
    if (!strategicDragState.dragging) return;
    renderPlanner();
  });
}

function getStrategicPointerCellFromEvent(event) {
  const pointElement = document.elementFromPoint(event.clientX, event.clientY);
  return pointElement?.closest?.(".board-cell") || strategicDragState.lastTargetCellEl || null;
}

function updateStrategicDragPreviewFromEvent(event) {
  const sourceItem = strategicDragState.originItemSnapshot || state.items.find((entry) => entry.id === strategicDragState.itemId) || null;
  if (!sourceItem) return null;

  const cell = getStrategicPointerCellFromEvent(event);
  const mode = strategicDragState.mode || state.timelineScale;
  const channelId = cell?.dataset.channelId || strategicDragState.lastTargetChannelId || sourceItem.channelId;
  const category = cell?.dataset.category || strategicDragState.lastTargetCategory || sourceItem.category;
  const columnValue = cell?.dataset.columnValue || strategicDragState.lastTargetColumnValue;
  const field = getStrategicOrderField(mode);
  const previewItem = { ...sourceItem, channelId, category };

  if (mode === "week") {
    previewItem.day = columnValue || sourceItem.day;
  } else {
    const bucket = Number(cell?.dataset.columnIndex);
    previewItem.planningWeek = Number.isInteger(bucket) ? bucket + 1 : strategicDragState.originPlanningWeek;
  }

  previewItem[field] = getStrategicCellOrder(cell, event, mode, sourceItem.id, sourceItem);

  strategicDragState.previewItem = previewItem;
  strategicDragState.lastTargetCellEl = cell;
  strategicDragState.lastTargetChannelId = channelId;
  strategicDragState.lastTargetCategory = category;
  strategicDragState.lastTargetColumnValue = columnValue || strategicDragState.lastTargetColumnValue;
  strategicDragState.lastTargetKey = getStrategicCellKey(channelId, category, strategicDragState.lastTargetColumnValue);

  if (DEBUG_TIMELINE_DND) {
    console.log("[strategic-dnd] hover", {
      itemId: sourceItem.id,
      title: sourceItem.title,
      mode,
      targetChannelId: channelId,
      targetCategory: category,
      targetColumn: strategicDragState.lastTargetColumnValue,
      previewDay: previewItem.day,
      previewPlanningWeek: previewItem.planningWeek ?? "-",
      previewOrder: previewItem[field],
    });
  }

  return previewItem;
}

function commitStrategicDrag(event) {
  const item = state.items.find((entry) => entry.id === strategicDragState.itemId);
  if (!item) {
    cancelStrategicDrag("missing item", event);
    return;
  }

  const preview = strategicDragState.previewItem || strategicDragState.originItemSnapshot || item;
  const before = { day: item.day, planningWeek: item.planningWeek || null, channelId: item.channelId, category: item.category };
  item.channelId = preview.channelId;
  item.category = preview.category;
  if ((strategicDragState.mode || state.timelineScale) === "week") {
    item.day = preview.day || item.day;
  } else {
    item.planningWeek = Number.isInteger(preview.planningWeek) ? preview.planningWeek : strategicDragState.originPlanningWeek;
  }
  const orderField = getStrategicOrderField(strategicDragState.mode || state.timelineScale);
  item[orderField] = Number.isFinite(preview[orderField]) ? preview[orderField] : item[orderField];

  if (DEBUG_TIMELINE_DND) {
    console.log("[strategic-dnd] commit", {
      itemId: item.id,
      title: item.title,
      before,
      after: { day: item.day, planningWeek: item.planningWeek || null, channelId: item.channelId, category: item.category, order: item[orderField] },
    });
  }

  clearStrategicDragState();
  persistAndRender();
  strategicDragState.recentClickSuppressionId = item.id;
  window.setTimeout(() => {
    if (strategicDragState.recentClickSuppressionId === item.id) strategicDragState.recentClickSuppressionId = null;
  }, 0);
}

function cancelStrategicDrag(reason, event) {
  if (DEBUG_TIMELINE_DND) {
    console.log("[strategic-dnd] cancel", {
      itemId: strategicDragState.itemId,
      reason,
    });
  }
  clearStrategicDragState();
  if (event) event.preventDefault();
  renderPlanner();
}

function clearStrategicDragState() {
  strategicDragState.armed = false;
  strategicDragState.dragging = false;
  strategicDragState.itemId = null;
  strategicDragState.pointerId = null;
  strategicDragState.mode = null;
  strategicDragState.originItemSnapshot = null;
  strategicDragState.originChannelId = null;
  strategicDragState.originCategory = null;
  strategicDragState.originDay = null;
  strategicDragState.originPlanningWeek = null;
  strategicDragState.lastTargetCellEl = null;
  strategicDragState.lastTargetKey = "-";
  strategicDragState.lastTargetChannelId = null;
  strategicDragState.lastTargetCategory = null;
  strategicDragState.lastTargetColumnValue = null;
  strategicDragState.previewItem = null;
  strategicDragState.startClientX = 0;
  strategicDragState.startClientY = 0;
  strategicDragState.pendingFrame = false;
  if (strategicDragState.rafId) {
    window.cancelAnimationFrame(strategicDragState.rafId);
    strategicDragState.rafId = 0;
  }
  document.body.classList.remove("is-strategic-dragging");
}

function schedulePointerDragRender() {
  if (!pointerDragState.active || pointerDragState.pendingFrame) return;
  pointerDragState.pendingFrame = true;
  pointerDragState.rafId = window.requestAnimationFrame(() => {
    pointerDragState.pendingFrame = false;
    if (!pointerDragState.active) return;
    renderPlanner();
  });
}

function getPointerLaneFromEvent(event, fallbackLane = null) {
  const pointElement = document.elementFromPoint(event.clientX, event.clientY);
  return pointElement?.closest?.(".day-lane") || fallbackLane || pointerDragState.lastLaneEl || pointerDragState.originLaneEl || null;
}

function getPointerTimelineContext(event, fallbackLane = null) {
  const lane = getPointerLaneFromEvent(event, fallbackLane);
  const shell = elements.timelineShell;
  const laneRect = lane?.getBoundingClientRect?.() || null;
  const clientX = Number.isFinite(event?.clientX) ? event.clientX : pointerDragState.lastClientX;
  const clientY = Number.isFinite(event?.clientY) ? event.clientY : pointerDragState.lastClientY;
  const shellScrollLeft = shell?.scrollLeft ?? 0;
  const shellClientWidth = shell?.clientWidth ?? null;
  const shellScrollWidth = shell?.scrollWidth ?? null;
  const visibleLaneX = Number.isFinite(clientX) && laneRect ? clientX - laneRect.left : null;
  const timelineContentX = Number.isFinite(visibleLaneX) ? visibleLaneX + shellScrollLeft : null;

  return {
    lane,
    clientX,
    clientY,
    laneLeft: laneRect ? laneRect.left : null,
    laneWidth: laneRect ? laneRect.width : null,
    shellScrollLeft,
    shellClientWidth,
    shellScrollWidth,
    visibleLaneX,
    timelineContentX,
  };
}

function updatePointerDragPreviewFromEvent(event, stage) {
  const sourceItem = pointerDragState.originItemSnapshot || state.items.find((entry) => entry.id === pointerDragState.itemId) || null;
  if (!sourceItem) return null;

  const context = getPointerTimelineContext(event, pointerDragState.lastLaneEl || pointerDragState.originLaneEl);
  const lane = context.lane || pointerDragState.lastLaneEl || pointerDragState.originLaneEl || null;
  const adjustedX = Number.isFinite(context.timelineContentX) ? context.timelineContentX - pointerDragState.pointerOffsetX : null;
  const preSnapMinutes = Number.isFinite(adjustedX) ? timelineXToMinutes(adjustedX) : null;
  const snappedMinutes = Number.isFinite(preSnapMinutes) ? snapStartMinutes(sourceItem, preSnapMinutes) : null;
  const duration = getSafeDuration(sourceItem.duration, sourceItem.itemType);
  const clampedMinutes = Number.isFinite(snappedMinutes)
    ? Math.max(DAY_START, Math.min(DAY_END - duration, snappedMinutes))
    : null;
  const finalMinutes = Number.isFinite(clampedMinutes) ? clampedMinutes : pointerDragState.originStartMinutes;
  const targetChannelId = lane?.dataset?.channelId || pointerDragState.lastChannelId || pointerDragState.originChannelId;
  const targetLaneCategory = lane?.dataset?.category || pointerDragState.lastCategory || pointerDragState.originCategory;
  const targetLaneLabel = lane ? `${getChannelName(targetChannelId)} / ${targetLaneCategory}` : pointerDragState.originLaneLabel;
  const finalStart = minutesToTime(clampPlanningMinutes(finalMinutes));

  pointerDragState.lastClientX = context.clientX;
  pointerDragState.lastClientY = context.clientY;
  pointerDragState.lastLaneEl = lane;
  pointerDragState.lastChannelId = targetChannelId;
  pointerDragState.lastCategory = targetLaneCategory;
  pointerDragState.lastVisibleLaneX = context.visibleLaneX;
  pointerDragState.lastTimelineContentX = context.timelineContentX;
  pointerDragState.lastAdjustedX = adjustedX;
  pointerDragState.lastPreSnapMinutes = preSnapMinutes;
  pointerDragState.lastSnappedMinutes = snappedMinutes;
  pointerDragState.lastClampedMinutes = clampedMinutes;
  pointerDragState.previewItem = {
    ...pointerDragState.originItemSnapshot,
    channelId: targetChannelId,
    category: targetLaneCategory,
    start: finalStart,
    duration,
    slot: getSlotFromMinutes(finalMinutes),
  };

  dragState.lastClientX = context.clientX;
  dragState.lastTimelineX = context.timelineContentX;
  dragState.lastEventType = `pointer${stage}`;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(lane);
  dragState.dropAllowed = lane ? "yes" : "unknown";
  dragState.dropBlockedReason = lane ? "-" : "no lane under pointer";

  if (!dragState.timingSnapshot) {
    dragState.timingSnapshot = {
      status: "dragging",
      hover: null,
      commit: null,
      final: null,
      postRender: null,
      matches: { hoverToCommit: "-", commitToSaved: "-", savedToRecheck: "-" },
      itemId: sourceItem.id,
      itemTitle: sourceItem.title,
      originalStartMinutes: pointerDragState.originStartMinutes,
      originalStartTime: pointerDragState.originStartTime,
      revertedAfterCommit: "no",
    };
  }

  dragState.timingSnapshot.status = "hovering";
  dragState.timingSnapshot.itemId = sourceItem.id;
  dragState.timingSnapshot.itemTitle = sourceItem.title;
  dragState.timingSnapshot.originalStartMinutes = pointerDragState.originStartMinutes;
  dragState.timingSnapshot.originalStartTime = pointerDragState.originStartTime;
  dragState.timingSnapshot.hover = {
    itemId: sourceItem.id,
    itemTitle: sourceItem.title,
    originalStartMinutes: pointerDragState.originStartMinutes,
    originalStartTime: pointerDragState.originStartTime,
    visibleLaneX: context.visibleLaneX,
    timelineContentX: context.timelineContentX,
    adjustedX,
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    targetChannelId,
    targetLaneCategory,
    targetLaneLabel,
  };

  if (DEBUG_TIMELINE_DND) {
    const snapshot = buildTimelineDebugSnapshot(`pointer${stage}`, event, lane, sourceItem, {
      timelineContentX: context.timelineContentX,
      visibleLaneX: context.visibleLaneX,
      adjustedX,
      preSnapMinutes,
      snappedMinutes,
      clampedMinutes,
      dropAllowed: dragState.dropAllowed,
      dropBlockedReason: dragState.dropBlockedReason,
      dropStatus: stage === "move" ? "POINTER HOVER" : "POINTER SETTLED",
      dropFired: "no",
      dropCommitted: "no",
      targetLabel: describeTimelineNode(event.target),
      currentTargetLabel: describeTimelineNode(event.currentTarget),
      closestLaneLabel: describeTimelineNode(lane),
      renderNote: stage === "move" ? "live pointer preview" : "pointer release preview",
      timingSnapshot: dragState.timingSnapshot,
    });
    updateTimelineDebugPanel(snapshot);
    console.log("[timeline-dnd] hover snapshot", {
      itemId: sourceItem.id,
      originalStart: pointerDragState.originStartTime,
      hoverSnapped: snappedMinutes,
      hoverClamped: clampedMinutes,
      targetChannelId,
      targetLaneCategory,
      visibleLaneX: context.visibleLaneX,
      timelineContentX: context.timelineContentX,
      adjustedX,
    });
  }

  return {
    lane,
    context,
    adjustedX,
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    finalMinutes,
    finalStart,
    targetChannelId,
    targetLaneCategory,
    targetLaneLabel,
  };
}

function commitPointerDrag(event) {
  if (!pointerDragState.active) return;

  const item = state.items.find((entry) => entry.id === pointerDragState.itemId);
  if (!item) {
    cancelPointerDrag("missing item", event);
    return;
  }

  const preview = pointerDragState.previewItem || item;
  const sameLane =
    pointerDragState.originChannelId === preview.channelId &&
    pointerDragState.originCategory === preview.category;
  const commitStartMinutes = getSafeStartMinutes(preview.start);
  const commitTargetLaneLabel = pointerDragState.lastLaneEl
    ? `${getChannelName(preview.channelId)} / ${preview.category}`
    : pointerDragState.originLaneLabel;
  const commitSnapshot = {
    itemId: item.id,
    itemTitle: item.title,
    rawX: pointerDragState.lastTimelineContentX,
    adjustedX: pointerDragState.lastAdjustedX,
    preSnapMinutes: pointerDragState.lastPreSnapMinutes,
    snappedMinutes: pointerDragState.lastSnappedMinutes,
    clampedMinutes: pointerDragState.lastClampedMinutes,
    targetChannelId: preview.channelId,
    targetLaneCategory: preview.category,
    targetLaneLabel: commitTargetLaneLabel,
    sameLane,
    newStartTime: preview.start,
    newChannelId: preview.channelId,
    newLaneCategory: preview.category,
    dropFired: "yes",
    dropCommitted: "yes",
  };

  dragState.commitSnapshot = {
    source: "pointer",
    dropFired: true,
    dropCommitted: true,
    itemId: item.id,
    itemTitle: item.title,
    originalStartMinutes: pointerDragState.originStartMinutes,
    originalStartTime: pointerDragState.originStartTime,
    timelineContentX: pointerDragState.lastTimelineContentX,
    adjustedX: pointerDragState.lastAdjustedX,
    preSnapMinutes: pointerDragState.lastPreSnapMinutes,
    snappedMinutes: pointerDragState.lastSnappedMinutes,
    clampedMinutes: pointerDragState.lastClampedMinutes,
    targetChannelId: preview.channelId,
    targetLaneCategory: preview.category,
    targetLaneLabel: commitTargetLaneLabel,
    sameLane,
    targetDetail: describeTimelineNode(event?.target),
    targetLabel: describeTimelineNode(event?.target),
    currentTargetLabel: describeTimelineNode(event?.currentTarget),
    closestLaneLabel: describeTimelineNode(pointerDragState.lastLaneEl),
    oldStart: item.start,
    oldChannelId: item.channelId,
    oldLaneCategory: item.category,
    newStartTime: preview.start,
    newStartMinutes: commitStartMinutes,
    newChannelId: preview.channelId,
    newLaneCategory: preview.category,
    stateUpdated: false,
    renderReflected: false,
    revertedAfterCommit: false,
    renderedLeft: null,
    renderVisibleTimeLabel: "-",
    status: "commit-pending",
  };

  dragState.timingSnapshot.commit = {
    itemId: item.id,
    itemTitle: item.title,
    rawX: pointerDragState.lastTimelineContentX,
    adjustedX: pointerDragState.lastAdjustedX,
    preSnapMinutes: pointerDragState.lastPreSnapMinutes,
    snappedMinutes: pointerDragState.lastSnappedMinutes,
    clampedMinutes: pointerDragState.lastClampedMinutes,
    targetChannelId: preview.channelId,
    targetLaneCategory: preview.category,
    targetLaneLabel: commitTargetLaneLabel,
    sameLane,
    newStartTime: preview.start,
    newChannelId: preview.channelId,
    newLaneCategory: preview.category,
  };
  dragState.timingSnapshot.status = "committing";

  console.log("[timeline-dnd] commit snapshot", {
    itemId: item.id,
    originalStart: pointerDragState.originStartTime,
    hoverSnapped: dragState.timingSnapshot.hover?.snappedMinutes,
    hoverClamped: dragState.timingSnapshot.hover?.clampedMinutes,
    commitSnapped: commitSnapshot.snappedMinutes,
    commitClamped: commitSnapshot.clampedMinutes,
    finalSavedStart: preview.start,
    revertedAfterCommit: "no",
  });

  console.log("[timeline-dnd] commit before update", {
    itemId: item.id,
    oldStart: item.start,
    oldChannelId: item.channelId,
    oldLaneCategory: item.category,
    newStart: preview.start,
    newChannelId: preview.channelId,
    newLaneCategory: preview.category,
  });

  item.start = preview.start;
  item.channelId = preview.channelId;
  item.category = preview.category;
  item.duration = preview.duration;
  item.slot = preview.slot;

  console.log("[timeline-dnd] commit after update", {
    itemId: item.id,
    newStart: item.start,
    newStartMinutes: getSafeStartMinutes(item.start),
    newChannelId: item.channelId,
    newLaneCategory: item.category,
  });

  dragState.commitSnapshot.stateUpdated = true;
  dragState.commitSnapshot.newStartMinutes = getSafeStartMinutes(item.start);
  dragState.commitSnapshot.newStartTime = item.start;
  dragState.commitSnapshot.newChannelId = item.channelId;
  dragState.commitSnapshot.newLaneCategory = item.category;
  dragState.commitSnapshot.status = "DROP FIRED AND COMMITTED";

  persistAndRender();

  const finalSavedItem = state.items.find((entry) => entry.id === item.id) || null;
  const finalSavedStart = finalSavedItem?.start || item.start;
  const finalSavedMinutes = finalSavedItem ? getSafeStartMinutes(finalSavedItem.start) : getSafeStartMinutes(item.start);
  const finalRenderedLeft = finalSavedItem ? timelineMinutesToX(finalSavedMinutes) : timelineMinutesToX(getSafeStartMinutes(item.start));
  const finalSavedMatchesCommit = finalSavedStart === preview.start;
  const hoverMatchesCommit =
    dragState.timingSnapshot.hover?.snappedMinutes === dragState.timingSnapshot.commit?.snappedMinutes &&
    dragState.timingSnapshot.hover?.clampedMinutes === dragState.timingSnapshot.commit?.clampedMinutes;
  const commitMatchesSaved = finalSavedStart === commitSnapshot.newStartTime;

  dragState.timingSnapshot.final = {
    itemId: item.id,
    savedStart: finalSavedStart,
    savedMinutes: finalSavedMinutes,
    savedTimeLabel: finalSavedStart,
    savedChannelId: finalSavedItem?.channelId || item.channelId,
    savedLaneCategory: finalSavedItem?.category || item.category,
    renderedLeft: finalRenderedLeft,
    visibleTimeLabel: finalSavedStart,
  };
  dragState.timingSnapshot.matches = {
    hoverToCommit: hoverMatchesCommit ? "yes" : "no",
    commitToSaved: commitMatchesSaved ? "yes" : "no",
    savedToRecheck: finalSavedMatchesCommit ? "yes" : "no",
  };
  dragState.timingSnapshot.postRender = {
    itemId: item.id,
    expectedCommittedStart: preview.start,
    postCheckStart: finalSavedStart,
    revertedAfterCommit: finalSavedMatchesCommit ? "no" : "yes",
  };
  dragState.timingSnapshot.revertedAfterCommit = finalSavedMatchesCommit ? "no" : "yes";
  dragState.timingSnapshot.status = finalSavedMatchesCommit ? "saved" : "reverted";

  console.log("[timeline-dnd] final saved snapshot", {
    itemId: item.id,
    originalStart: pointerDragState.originStartTime,
    hoverSnapped: dragState.timingSnapshot.hover?.snappedMinutes,
    hoverClamped: dragState.timingSnapshot.hover?.clampedMinutes,
    commitSnapped: dragState.timingSnapshot.commit?.snappedMinutes,
    commitClamped: dragState.timingSnapshot.commit?.clampedMinutes,
    finalSavedStart,
    finalSavedMinutes,
    revertedAfterCommit: finalSavedMatchesCommit ? "no" : "yes",
  });

  console.log("[timeline-dnd] post-render recheck", {
    itemId: item.id,
    expectedCommittedStart: preview.start,
    currentStart: finalSavedStart,
    currentStartMinutes: finalSavedMinutes,
    currentChannelId: finalSavedItem?.channelId || item.channelId,
    currentLaneCategory: finalSavedItem?.category || item.category,
    renderedLeft: finalRenderedLeft,
    visibleTimeLabel: finalSavedStart,
    revertedAfterCommit: finalSavedMatchesCommit ? "no" : "yes",
  });

  dragState.commitSnapshot.renderReflected = finalSavedMatchesCommit;
  dragState.commitSnapshot.renderedLeft = finalRenderedLeft;
  dragState.commitSnapshot.renderVisibleTimeLabel = finalSavedStart;
  dragState.commitSnapshot.revertedAfterCommit = finalSavedMatchesCommit ? false : true;
  dragState.commitSnapshot.postCheckStart = finalSavedStart;
  dragState.commitSnapshot.postCheckMinutes = finalSavedMinutes;
  dragState.commitSnapshot.postCheckTargetChannelId = finalSavedItem?.channelId || item.channelId;
  dragState.commitSnapshot.postCheckTargetLaneCategory = finalSavedItem?.category || item.category;
  dragState.commitSnapshot.postCheckMatched = finalSavedMatchesCommit;

  dragState.lastDropCommitted = true;

  if (DEBUG_TIMELINE_DND) {
    const snapshot = buildTimelineDebugSnapshot("pointerup", event, pointerDragState.lastLaneEl || event.currentTarget, item, {
      timelineContentX: pointerDragState.lastTimelineContentX,
      adjustedX: pointerDragState.lastAdjustedX,
      preSnapMinutes: pointerDragState.lastPreSnapMinutes,
      snappedMinutes: pointerDragState.lastSnappedMinutes,
      clampedMinutes: pointerDragState.lastClampedMinutes,
      finalMinutes: finalSavedMinutes,
      finalTime: finalSavedStart,
      renderedLeft: finalRenderedLeft,
      dropStatus: "POINTER DROP COMMITTED",
      dropFired: "yes",
      dropCommitted: "yes",
      dropAllowed: "yes",
      dropBlockedReason: "-",
      commitSnapshot: dragState.commitSnapshot,
      timingSnapshot: dragState.timingSnapshot,
      targetLabel: describeTimelineNode(event.target),
      currentTargetLabel: describeTimelineNode(event.currentTarget),
      closestLaneLabel: describeTimelineNode(pointerDragState.lastLaneEl),
      renderNote: "pointer commit complete",
    });
    updateTimelineDebugPanel(snapshot);
  }

  clearPointerDragState();
  renderPlanner();
}

function cancelPointerDrag(reason, event) {
  if (DEBUG_TIMELINE_DND) {
    console.log("[timeline-dnd] pointer drag canceled", {
      itemId: pointerDragState.itemId,
      reason,
      originalStart: pointerDragState.originStartTime,
    });
  }
  clearPointerDragState();
  dragState.timingSnapshot = dragState.timingSnapshot || {};
  dragState.timingSnapshot.status = "canceled";
  dragState.timingSnapshot.revertedAfterCommit = "no";
  if (event) {
    event.preventDefault();
  }
  renderPlanner();
}

function clearPointerDragState() {
  pointerDragState.active = false;
  pointerDragState.itemId = null;
  pointerDragState.pointerId = null;
  pointerDragState.startClientX = 0;
  pointerDragState.startClientY = 0;
  pointerDragState.originLaneEl = null;
  pointerDragState.originChannelId = null;
  pointerDragState.originCategory = null;
  pointerDragState.originLaneLabel = "-";
  pointerDragState.originStartMinutes = null;
  pointerDragState.originStartTime = "-";
  pointerDragState.originItemSnapshot = null;
  pointerDragState.pointerOffsetX = 0;
  pointerDragState.pointerOffsetY = 0;
  pointerDragState.lastClientX = null;
  pointerDragState.lastClientY = null;
  pointerDragState.lastLaneEl = null;
  pointerDragState.lastChannelId = null;
  pointerDragState.lastCategory = null;
  pointerDragState.lastVisibleLaneX = null;
  pointerDragState.lastTimelineContentX = null;
  pointerDragState.lastAdjustedX = null;
  pointerDragState.lastPreSnapMinutes = null;
  pointerDragState.lastSnappedMinutes = null;
  pointerDragState.lastClampedMinutes = null;
  pointerDragState.previewItem = null;
  pointerDragState.pendingFrame = false;
  if (pointerDragState.rafId) {
    window.cancelAnimationFrame(pointerDragState.rafId);
    pointerDragState.rafId = 0;
  }
  document.body.classList.remove("is-pointer-dragging");
  dragState.itemId = null;
  dragState.pointerOffsetX = 0;
  dragState.lastClientX = null;
  dragState.lastTimelineX = null;
  dragState.draggedItemSnapshot = null;
  dragState.dropAllowed = "unknown";
  dragState.dropBlockedReason = "-";
  dragState.lastEventType = null;
  dragState.lastTargetLabel = null;
  dragState.lastCurrentTargetLabel = null;
  dragState.lastClosestLaneLabel = null;
}

function renderTable() {
  elements.tableBody.innerHTML = "";
  if (state.viewMode !== "table") return;

  const items = getVisibleItems();
  if (items.length === 0) {
    const row = document.createElement("tr");
    row.innerHTML = `<td colspan="12" class="empty-state">No items in the current view.</td>`;
    elements.tableBody.appendChild(row);
    return;
  }

  items.forEach((item) => {
    const issues = validateItem(item, state.items, state.exportProfile);
    const statusLabel = getStatusLabel(issues);
    const flagList = getFlagList(item)
      .map((flag) => `<span class="flag-pill">${flag.short}</span>`)
      .join("") || "<span class='flag-pill'>-</span>";
    const row = document.createElement("tr");
    if (elements.showId.value === item.id) row.classList.add("is-active");
    row.innerHTML = `
      <td>${item.title}</td>
      <td>${item.itemType}</td>
      <td>${item.category}</td>
      <td>${getChannelName(item.channelId)}</td>
      <td>${item.day}</td>
      <td>${item.start}</td>
      <td>${getSafeDuration(item.duration, item.itemType)} mins</td>
      <td>${item.slot}</td>
      <td><span class="status-pill ${getStatusClass(issues)}">${statusLabel}</span></td>
      <td>${flagList}</td>
      <td>${item.assetCode || "-"}</td>
      <td><div class="table-actions"></div></td>
    `;

    const actions = row.querySelector(".table-actions");
    const editButton = document.createElement("button");
    editButton.type = "button";
    editButton.className = "button tertiary";
    editButton.textContent = "Edit";
    editButton.addEventListener("click", () => hydrateItemForm(item.id));

    const duplicateButton = document.createElement("button");
    duplicateButton.type = "button";
    duplicateButton.className = "button tertiary";
    duplicateButton.textContent = "Duplicate";
    duplicateButton.addEventListener("click", () => duplicateItem(item.id));

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "button tertiary danger";
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => deleteItem(item.id));

    actions.append(editButton, duplicateButton, deleteButton);
    elements.tableBody.appendChild(row);
  });
}

function renderShowList() {
  elements.showList.innerHTML = "";
  const items = getVisibleItems();

  if (items.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No scheduled items match this view yet.";
    elements.showList.appendChild(empty);
    return;
  }

  const template = document.getElementById("showCardTemplate");
  items.forEach((item) => {
    const node = template.content.firstElementChild.cloneNode(true);
    const issues = validateItem(item, state.items, state.exportProfile);
    node.querySelector(".show-card-title").textContent = item.title;
    node.querySelector(".show-card-meta").textContent =
      `${item.day} | ${getChannelName(item.channelId)} | ${item.start} | ${getSafeDuration(item.duration, item.itemType)} mins | ${item.category}`;
    node.querySelector(".show-chip").textContent = `${item.itemType} / ${item.importance}`;
    node.querySelector(".show-chip").style.border = `1px solid ${getItemColor(item)}`;
    node.classList.toggle("is-active", elements.showId.value === item.id);
    node.querySelector(".show-card-flags").innerHTML =
      `${issues.length ? `<span class="status-pill ${getStatusClass(issues)}">${getStatusLabel(issues)}</span>` : ""}${getFlagList(item)
        .map((flag) => `<span class="flag-pill">${flag.short}</span>`)
        .join("")}`;
    node.querySelector(".show-card-notes").textContent = item.notes || "No notes added.";
    node.querySelector(".edit-button").addEventListener("click", () => hydrateItemForm(item.id));
    node.querySelector(".duplicate-button").addEventListener("click", () => duplicateItem(item.id));
    node.querySelector(".delete-button").addEventListener("click", () => deleteItem(item.id));
    elements.showList.appendChild(node);
  });
}

function renderChannelList() {
  elements.channelList.innerHTML = "";
  state.channels.forEach((channel) => {
    const row = document.createElement("article");
    row.className = "channel-card";
    const channelIssues = validateChannelConfig(channel, state.items, state.exportProfile, state.channels);
    const configSummary = [
      `#${String(parsePositiveInteger(channel.channelNumber) || 1).padStart(2, "0")}`,
      channel.networkType || "standard",
      `inc ${parsePositiveInteger(channel.scheduleIncrement) || 30}m`,
      `break ${channel.breakStrategy || "standard"} / ${parsePositiveInteger(channel.breakDuration) || 120}s`,
      channel.commercialFree ? "commercial-free" : "commercials on",
      channel.showLogo ? "logo on" : "logo off",
      channel.multiLogoMode ? `multi-logo ${channel.multiLogoProfile || "profile required"}` : "multi-logo off",
    ].join(" | ");
    const pathSummary = [
      channel.contentDir || "-",
      channel.commercialDir || "-",
      channel.bumpDir || "-",
    ].join(" | ");
    const assetSummary = [
      channel.signOffVideo || "sign-off default",
      channel.offAirVideo || "off-air default",
      channel.standbyImage || "standby default",
    ].join(" | ");
    const clipSummary = normalizeClipShowList(channel.clipShows).length
      ? normalizeClipShowList(channel.clipShows).join(", ")
      : "No clip shows";
    row.innerHTML = `
      <div class="channel-card-main">
        <span class="channel-swatch" style="background:${channel.color}"></span>
        <div class="channel-card-copy">
          <strong>${channel.name}</strong>
          <p>${channel.group || "Ungrouped"} | ${channel.tagline || "No tagline set"}</p>
          <div class="channel-card-meta">${configSummary}</div>
          <div class="channel-card-paths">${pathSummary}</div>
          <div class="channel-card-assets">${assetSummary}</div>
          <div class="channel-card-clips">Clip shows: ${clipSummary}</div>
        </div>
      </div>
      <div class="channel-card-status"></div>
    `;

    const status = row.querySelector(".channel-card-status");
    const statusPill = document.createElement("span");
    statusPill.className = `status-pill ${channelIssues.some((issue) => issue.severity === "blocker") ? "blocked" : channelIssues.some((issue) => issue.severity === "warning") ? "warning" : "ok"}`;
    statusPill.textContent = channelIssues.some((issue) => issue.severity === "blocker")
      ? "Config blocked"
      : channelIssues.some((issue) => issue.severity === "warning")
        ? "Config warning"
        : "Config ready";
    status.appendChild(statusPill);

    if (channelIssues.length > 0) {
      const issueList = document.createElement("div");
      issueList.className = "channel-card-issues";
      issueList.innerHTML = channelIssues.slice(0, 3).map((issue) => `<div class="validation-line">${issue.message}</div>`).join("");
      status.appendChild(issueList);
    }

    const actions = document.createElement("div");
    actions.className = "card-actions";

    const editButton = document.createElement("button");
    editButton.type = "button";
    editButton.className = "button tertiary";
    editButton.textContent = "Edit";
    editButton.addEventListener("click", () => hydrateChannelForm(channel.id));

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "button tertiary danger";
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => deleteChannel(channel.id));

    actions.append(editButton, deleteButton);
    row.appendChild(actions);
    elements.channelList.appendChild(row);
  });
}

function renderInsights() {
  const items = getVisibleItems();
  const insights = [];
  const validation = validateSchedule(state.exportProfile);

  if (items.length === 0) {
    insights.push("This view is empty, so it is a good place to sketch the next template.");
  } else {
    const commercials = items.filter((item) => ITEM_TYPE_META[item.itemType]?.commercial).length;
    const conflicts = getConflicts(items).length;
    const assetCodes = items.filter((item) => item.assetCode).length;
    const latest = items.reduce((max, item) => (timeToMinutes(item.start) > timeToMinutes(max.start) ? item : max), items[0]);

    insights.push(
      `This view carries ${Math.round(items.reduce((sum, item) => sum + getSafeDuration(item.duration, item.itemType), 0) / 60)} hours of planned runtime.`,
    );
    insights.push(`${commercials} commercial or continuity items are planned, including ad breaks, spots, bumpers, or promos.`);
    insights.push(conflicts > 0 ? `${conflicts} overlap flags are active and should be resolved before export.` : "No overlaps are detected in the current view.");
    insights.push(`${assetCodes} items already have export codes, which will help downstream ingest into FS42.`);
    insights.push(
      validation.blockers.length > 0
        ? `Export readiness is blocked by ${validation.blockers.length} issue${validation.blockers.length === 1 ? "" : "s"} in the current profile.`
        : validation.warnings.length > 0
          ? `Export readiness has ${validation.warnings.length} advisory warning${validation.warnings.length === 1 ? "" : "s"}.`
          : "Export readiness is clear for the selected profile.",
    );
    insights.push(`The latest start in the current view is ${latest.start} with "${latest.title}".`);
  }

  elements.insights.innerHTML = insights.map((entry) => `<li>${entry}</li>`).join("");
}

function renderStats() {
  const items = getVisibleItems();
  const runtimeHours = items.reduce((sum, item) => sum + getSafeDuration(item.duration, item.itemType), 0) / 60;
  const commercials = items.filter((item) => ITEM_TYPE_META[item.itemType]?.commercial).length;
  elements.statBlocks.textContent = String(items.length);
  elements.statRuntime.textContent = `${runtimeHours.toFixed(runtimeHours % 1 === 0 ? 0 : 1)}h`;
  elements.statPrime.textContent = String(commercials);
}

function syncExportControls() {
  elements.exportProfile.value = state.exportProfile;
  elements.exportJson.textContent =
    state.exportProfile === "internal" ? "Export Internal JSON" : "Export FS42 Native JSON";
  renderExportReadiness();
}

function renderExportReadiness() {
  const validation = state.exportProfile === "internal" ? validateSchedule("internal") : validateFs42NativeExport();
  const readiness = elements.exportReadiness;
  const label = state.exportProfile === "internal" ? "Internal scheduler JSON" : "FS42 station config";
  const blockerCount = validation.blockers.length;
  const warningCount = validation.warnings.length;
  const items = blockerCount > 0 ? validation.blockers : validation.warnings;
  const uniqueIssues = items.filter(
    (issue, index, collection) =>
      index ===
      collection.findIndex(
        (entry) => entry.itemId === issue.itemId && entry.code === issue.code && entry.message === issue.message,
      ),
  );
  const className = blockerCount > 0 ? "is-blocked" : warningCount > 0 ? "is-warning" : "is-ready";
  const headline =
    blockerCount > 0
      ? `${label} blocked by ${blockerCount} issue${blockerCount === 1 ? "" : "s"}.`
      : warningCount > 0
        ? `${label} is ready with ${warningCount} warning${warningCount === 1 ? "" : "s"}.`
        : `${label} is ready for export.`;
  const lines = uniqueIssues.slice(0, 3).map((issue) => `<div class="validation-line">- ${renderIssueLine(issue)}</div>`).join("");

  readiness.className = `export-readiness ${className}`;
  readiness.innerHTML = `<strong>${headline}</strong>${lines || "<div class='validation-line'>No blockers are active.</div>"}`;
}

function updateItemValidationSummary(draftItem = null, draftIssues = null) {
  const currentItem = draftItem || collectItemFormState();
  const issues = draftIssues || validateItem(currentItem, state.items, state.exportProfile);
  const summary = elements.itemValidation;

  if (!elements.showId.value && !itemFormDirty && !draftItem) {
    summary.innerHTML = "<div class='validation-line'>Fill the form to see export readiness for this item.</div>";
    return;
  }

  if (issues.length === 0) {
    summary.innerHTML = "<div class='validation-line'><span class='status-pill ok'>Ready</span><span>Item passes current validation rules.</span></div>";
    return;
  }

  summary.innerHTML = issues
    .slice(0, 4)
    .map((issue) => `<div class="validation-line"><span class="status-pill ${issue.severity === "blocker" ? "blocked" : "warning"}">${issue.severity === "blocker" ? "Blocker" : "Warning"}</span><span>${issue.message}</span></div>`)
    .join("");
}

function collectItemFormState() {
  return {
    id: elements.showId.value || "",
    title: elements.title.value,
    itemType: elements.itemType.value,
    category: elements.category.value,
    channelId: elements.channel.value,
    day: elements.day.value,
    start: elements.start.value,
    duration: getSafeDuration(elements.duration.value, elements.itemType.value),
    importance: elements.importance.value,
    watershedRestricted: elements.watershedRestricted.checked,
    primeTime: elements.primeTime.checked,
    mustRun: elements.mustRun.checked,
    slot: elements.slot.value,
    blockGroup: elements.blockGroup.value,
    assetCode: elements.assetCode.value,
    notes: elements.notes.value,
  };
}

function duplicateItem(itemId) {
  const source = state.items.find((entry) => entry.id === itemId);
  if (!source) return;

  const sourceStart = getSafeStartMinutes(source.start);
  const sourceDuration = getSafeDuration(source.duration, source.itemType);
  const startMinutes = Math.min(sourceStart + 30, DAY_END - sourceDuration);
  const duplicateTitle = source.title.endsWith(" Copy") ? `${source.title} 2` : `${source.title} Copy`;
  const duplicateAssetCode = getUniqueExportCode(generateDefaultExportCode(duplicateTitle), state.items);
  const duplicate = {
    ...structuredClone(source),
    id: crypto.randomUUID(),
    title: duplicateTitle,
    duration: sourceDuration,
    start: minutesToTime(Math.max(DAY_START, startMinutes)),
    assetCode: duplicateAssetCode,
  };
  state.items.push(duplicate);
  hydrateItemForm(duplicate.id);
  state.selectedDay = duplicate.day;
  syncControls();
  persistAndRender();
}

function duplicateActiveItem() {
  if (!elements.showId.value) return;
  duplicateItem(elements.showId.value);
}

function handleKeyboardShortcuts(event) {
  if (event.altKey && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
    if (["INPUT", "TEXTAREA", "SELECT"].includes(document.activeElement?.tagName || "")) return;
    if (!elements.showId.value) return;
    event.preventDefault();
    const delta = event.shiftKey ? 15 : 5;
    adjustActiveItemStart(event.key === "ArrowLeft" ? -delta : delta);
  }
}

function adjustActiveItemStart(deltaMinutes) {
  const item = state.items.find((entry) => entry.id === elements.showId.value);
  if (!item) return;
  const duration = getSafeDuration(item.duration, item.itemType);
  const startMinutes = getSafeStartMinutes(item.start);
  const nextStart = Math.max(DAY_START, Math.min(DAY_END - duration, startMinutes + deltaMinutes));
  item.duration = duration;
  item.start = minutesToTime(snapStartMinutes(item, nextStart));
  item.slot = getSlotFromMinutes(timeToMinutes(item.start));
  hydrateItemForm(item.id);
  persistAndRender();
}

function hydrateItemForm(itemId) {
  const item = state.items.find((entry) => entry.id === itemId);
  if (!item) return;
  itemFormDirty = false;
  state.workspace = "schedule";

  elements.showId.value = item.id;
  elements.title.value = item.title;
  elements.itemType.value = item.itemType;
  elements.category.value = item.category;
  elements.channel.value = item.channelId;
  elements.day.value = item.day;
  elements.start.value = item.start;
  elements.duration.value = String(item.duration);
  elements.importance.value = item.importance;
  elements.watershedRestricted.checked = Boolean(item.watershedRestricted);
  elements.primeTime.checked = Boolean(item.primeTime);
  elements.mustRun.checked = Boolean(item.mustRun);
  elements.slot.value = item.slot;
  elements.blockGroup.value = item.blockGroup;
  elements.assetCode.value = item.assetCode || "";
  elements.notes.value = item.notes;
  elements.formTitle.textContent = "Edit Item";
  elements.formHint.textContent = `Updating ${item.title}`;
  updateItemValidationSummary(item);
  render();
}

function regenerateActiveItemExportCode() {
  const title = elements.title.value.trim();
  const generated = generateDefaultExportCode(title);
  if (!generated) {
    elements.assetCode.value = "";
    itemFormDirty = true;
    updateItemValidationSummary();
    return;
  }

  elements.assetCode.value = getUniqueExportCode(generated, state.items, elements.showId.value || null);
  itemFormDirty = true;
  updateItemValidationSummary();
}

function hydrateChannelForm(channelId) {
  const channel = state.channels.find((entry) => entry.id === channelId);
  if (!channel) return;

  applyChannelFormState(channel, false);
  syncChannelMultiLogoControls();
}

function buildDefaultChannelDraft() {
  const channelNumber = getNextAvailableChannelNumber();
  const base = getChannelConfigBase("Channel", channelNumber);
  return {
    id: "",
    name: "",
    group: "",
    color: state.channels[0]?.color || "#ff8a5b",
    tagline: "",
    channelNumber,
    networkType: "standard",
    scheduleIncrement: 30,
    breakStrategy: "standard",
    commercialFree: false,
    breakDuration: 120,
    contentDir: `catalog/${base}`,
    commercialDir: `commercial/${base}`,
    bumpDir: `bump/${base}`,
    clipShows: [],
    signOffVideo: "runtime/signoff.mp4",
    offAirVideo: "runtime/off_air_pattern.mp4",
    standbyImage: "runtime/standby.png",
    beRightBackMedia: "runtime/brb.png",
    logoDir: `logos/${base}`,
    showLogo: true,
    defaultLogo: `${base}.png`,
    logoPermanent: true,
    multiLogoMode: false,
    multiLogoProfile: "",
  };
}

function applyChannelFormState(channel, preserveBlankName = false) {
  const normalized = normalizeChannel(channel, 0);
  elements.channelId.value = normalized.id || "";
  elements.channelName.value = preserveBlankName ? String(channel?.name || "") : normalized.name || "";
  elements.channelGroup.value = normalized.group || "";
  elements.channelTagline.value = normalized.tagline || "";
  elements.channelNumber.value = String(normalized.channelNumber || "");
  elements.channelColor.value = normalized.color || "#ff8a5b";
  elements.networkType.value = normalized.networkType || "standard";
  elements.scheduleIncrement.value = String(normalized.scheduleIncrement || 30);
  elements.breakStrategy.value = normalized.breakStrategy || "standard";
  elements.commercialFree.checked = Boolean(normalized.commercialFree);
  elements.breakDuration.value = String(normalized.breakDuration || 120);
  elements.contentDir.value = normalized.contentDir || "";
  elements.commercialDir.value = normalized.commercialDir || "";
  elements.bumpDir.value = normalized.bumpDir || "";
  elements.clipShows.value = formatClipShowList(normalized.clipShows);
  elements.signOffVideo.value = normalized.signOffVideo || "";
  elements.offAirVideo.value = normalized.offAirVideo || "";
  elements.standbyImage.value = normalized.standbyImage || "";
  elements.beRightBackMedia.value = normalized.beRightBackMedia || "";
  elements.logoDir.value = normalized.logoDir || "";
  elements.showLogo.checked = Boolean(normalized.showLogo);
  elements.defaultLogo.value = normalized.defaultLogo || "";
  elements.logoPermanent.checked = Boolean(normalized.logoPermanent);
  elements.channelMultiLogoMode.checked = Boolean(normalized.multiLogoMode);
  elements.channelMultiLogoProfile.value = normalized.multiLogoProfile || "";
  syncChannelCommercialFreeControls();
  syncChannelMultiLogoControls();
}

function collectChannelFormState() {
  const multiLogoMode = Boolean(elements.channelMultiLogoMode.checked);
  const logoDir = String(elements.logoDir?.value || "").trim();
  return normalizeChannel({
    id: elements.channelId.value || crypto.randomUUID(),
    name: elements.channelName.value.trim(),
    group: elements.channelGroup.value.trim(),
    color: elements.channelColor.value,
    tagline: elements.channelTagline.value.trim(),
    channelNumber: parsePositiveInteger(elements.channelNumber.value),
    networkType: elements.networkType.value,
    scheduleIncrement: parsePositiveInteger(elements.scheduleIncrement.value),
    breakStrategy: elements.breakStrategy.value,
    commercialFree: Boolean(elements.commercialFree.checked),
    breakDuration: parsePositiveInteger(elements.breakDuration.value) || 120,
    contentDir: elements.contentDir.value.trim(),
    commercialDir: elements.commercialDir.value.trim(),
    bumpDir: elements.bumpDir.value.trim(),
    clipShows: normalizeClipShowList(elements.clipShows.value),
    signOffVideo: elements.signOffVideo.value.trim(),
    offAirVideo: elements.offAirVideo.value.trim(),
    standbyImage: elements.standbyImage.value.trim(),
    beRightBackMedia: elements.beRightBackMedia.value.trim(),
    logoDir,
    showLogo: Boolean(elements.showLogo.checked),
    defaultLogo: elements.defaultLogo.value.trim(),
    logoPermanent: Boolean(elements.logoPermanent.checked),
    multiLogoMode,
    multiLogoProfile: multiLogoMode ? elements.channelMultiLogoProfile.value.trim() : "",
  });
}

function deleteItem(itemId) {
  state.items = state.items.filter((entry) => entry.id !== itemId);
  if (elements.showId.value === itemId) resetItemForm();
  persistAndRender();
}

function deleteChannel(channelId) {
  if (state.channels.length === 1) {
    window.alert("At least one channel is required.");
    return;
  }

  const fallback = state.channels.find((channel) => channel.id !== channelId);
  const affected = state.items.filter((item) => item.channelId === channelId).length;
  const confirmMessage =
    affected > 0
      ? `Delete this channel and move ${affected} items to ${fallback?.name}?`
      : "Delete this channel?";

  if (!window.confirm(confirmMessage)) return;

  state.items = state.items.map((item) =>
    item.channelId === channelId ? { ...item, channelId: fallback.id } : item,
  );
  state.channels = state.channels.filter((channel) => channel.id !== channelId);

  if (state.selectedChannelId === channelId) state.selectedChannelId = "all";
  if (elements.channelId.value === channelId) resetChannelForm();
  persistAndRender();
}

function getVisibleChannels() {
  return state.channels.filter(
    (channel) => state.selectedChannelId === "all" || state.selectedChannelId === channel.id,
  );
}

function getVisibleItems() {
  return state.items
    .filter((item) => state.selectedChannelId === "all" || item.channelId === state.selectedChannelId)
    .filter((item) => {
      if (state.timelineScale === "day" || state.viewMode === "table") {
        return item.day === state.selectedDay;
      }
      return true;
    })
    .sort(compareItems);
}

function getLaneCategories(channelId) {
  const categories = new Set(
    state.items
      .filter((item) => item.channelId === channelId)
      .filter((item) => state.timelineScale === "day" ? item.day === state.selectedDay : true)
      .map((item) => item.category),
  );
  return categories.size > 0 ? Array.from(categories) : ["Entertainment"];
}

function getStrategicItems(items, channelId, category, scaleMode, index) {
  const laneItems = items
    .filter((item) => item.channelId === channelId && item.category === category)
    .sort((a, b) => compareStrategicItems(a, b, scaleMode));

  if (scaleMode === "week") {
    return laneItems.filter((item) => DAY_NAMES.indexOf(item.day) === index);
  }

  return laneItems.filter((item) => getStrategicMonthBucket(item) === index + 1);
}

function compareStrategicItems(a, b, mode) {
  const field = getStrategicOrderField(mode);
  const orderA = Number.isFinite(a[field]) ? Number(a[field]) : Number.MAX_SAFE_INTEGER;
  const orderB = Number.isFinite(b[field]) ? Number(b[field]) : Number.MAX_SAFE_INTEGER;
  if (orderA !== orderB) return orderA - orderB;
  return compareItems(a, b);
}

function getStrategicCellOrder(cell, event, mode, draggedItemId, fallbackItem = null) {
  const draggedSelector = `[data-item-id="${draggedItemId}"]`;
  const cards = Array.from(cell?.querySelectorAll?.(".mini-card") || []).filter((card) => card.dataset.itemId !== draggedItemId);
  if (cards.length === 0) {
    const base = fallbackItem ? Number(fallbackItem[getStrategicOrderField(mode)]) : NaN;
    return Number.isFinite(base) ? base + 10 : 10;
  }

  const pointerY = event.clientY;
  const cardInfo = cards
    .map((card) => {
      const item = state.items.find((entry) => entry.id === card.dataset.itemId);
      const rect = card.getBoundingClientRect();
      return {
        item,
        order: item ? Number(item[getStrategicOrderField(mode)]) : NaN,
        centerY: rect.top + rect.height / 2,
      };
    })
    .filter((entry) => entry.item && Number.isFinite(entry.order))
    .sort((left, right) => left.centerY - right.centerY);

  if (cardInfo.length === 0) return 10;

  let insertIndex = cardInfo.findIndex((entry) => pointerY < entry.centerY);
  if (insertIndex === -1) insertIndex = cardInfo.length;

  if (insertIndex <= 0) return cardInfo[0].order - 10;
  if (insertIndex >= cardInfo.length) return cardInfo[cardInfo.length - 1].order + 10;
  return (cardInfo[insertIndex - 1].order + cardInfo[insertIndex].order) / 2;
}

function computeLaneLayout(items) {
  const levelsEnd = [];
  const positions = new Map();

  items.forEach((item) => {
    const start = timeToMinutes(item.start);
    const end = start + getSafeDuration(item.duration, item.itemType);
    let level = levelsEnd.findIndex((existingEnd) => existingEnd <= start);
    if (level === -1) level = levelsEnd.length;
    levelsEnd[level] = end;
    positions.set(item.id, level);
  });

  return { positions, levels: Math.max(levelsEnd.length, 1) };
}

function getConflicts(items) {
  const conflicts = new Set();
  const lanes = new Map();

  items.forEach((item) => {
    const key = `${item.channelId}|${item.category}|${item.day}`;
    if (!lanes.has(key)) lanes.set(key, []);
    lanes.get(key).push(item);
  });

  lanes.forEach((laneItems) => {
    const sorted = [...laneItems].sort(compareItems);
    for (let index = 0; index < sorted.length - 1; index += 1) {
      const current = sorted[index];
      const next = sorted[index + 1];
      const currentEnd = timeToMinutes(current.start) + current.duration;
      if (currentEnd > timeToMinutes(next.start)) {
        conflicts.add(current.id);
        conflicts.add(next.id);
      }
    }
  });

  return Array.from(conflicts);
}

function getItemColor(item) {
  if (state.colorMode === "slot") return SLOT_COLORS[item.slot];
  if (state.colorMode === "importance") return IMPORTANCE_COLORS[item.importance];
  if (state.colorMode === "type") return ITEM_TYPE_META[item.itemType]?.color || "#7d91b9";
  return state.channels.find((channel) => channel.id === item.channelId)?.color || "#7d91b9";
}

function getChannelName(channelId) {
  return state.channels.find((channel) => channel.id === channelId)?.name || "Unknown channel";
}

function compareItems(a, b) {
  return timeToMinutes(a.start) - timeToMinutes(b.start);
}

function parseTimeMinutes(value) {
  const match = /^(\d{2}):(\d{2})$/.exec(String(value || ""));
  if (!match) return NaN;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours > 23 || minutes > 59) return NaN;
  return hours * 60 + minutes;
}

function getTimelineHourWidth() {
  const raw = getComputedStyle(document.documentElement).getPropertyValue("--hour-width").trim();
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : 72;
}

function timelineMinutesToX(minutes) {
  return ((minutes - DAY_START) / 60) * HOUR_WIDTH;
}

function timelineXToMinutes(x) {
  return DAY_START + (x / HOUR_WIDTH) * 60;
}

function timelinePixelsToMinutes(px) {
  return (px / HOUR_WIDTH) * 60;
}

function getItemTimelineLeftX(item) {
  return timelineMinutesToX(clampPlanningMinutes(getSafeStartMinutes(item.start)));
}

function getSafeStartMinutes(value, fallback = DAY_START) {
  const parsed = parseTimeMinutes(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function getSafeDuration(value, itemType = "Programme") {
  const parsed = Number(value);
  const fallback = ITEM_TYPE_META[itemType]?.commercial ? 30 : 60;
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.max(5, parsed);
}

function normalizeExportProfile(value) {
  return value === "internal" ? "internal" : "fs42-native";
}

function seedStrategicOrders(items) {
  const seeded = items.map((item) => ({ ...item }));
  seedStrategicOrderField(seeded, "weekOrder", (item) => `${item.channelId}|${item.category}|${item.day}`);
  seedStrategicOrderField(seeded, "monthOrder", (item) => `${item.channelId}|${item.category}|${getStrategicMonthBucket(item)}`);
  return seeded;
}

function seedStrategicOrderField(items, field, groupKeyFn) {
  const groups = new Map();
  items.forEach((item) => {
    const key = groupKeyFn(item);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  });

  groups.forEach((groupItems) => {
    const withOrder = groupItems
      .filter((item) => Number.isFinite(item[field]))
      .sort((a, b) => Number(a[field]) - Number(b[field]));
    const withoutOrder = groupItems.filter((item) => !Number.isFinite(item[field])).sort(compareItems);
    let nextOrder = withOrder.length > 0 ? Number(withOrder[withOrder.length - 1][field]) + 10 : 10;

    withoutOrder.forEach((item) => {
      item[field] = nextOrder;
      nextOrder += 10;
    });
  });
}

function getStrategicOrderField(mode) {
  return mode === "week" ? "weekOrder" : "monthOrder";
}

function getStrategicOrderGroupKey(item, mode) {
  return mode === "week"
    ? `${item.channelId}|${item.category}|${item.day}`
    : `${item.channelId}|${item.category}|${getStrategicMonthBucket(item)}`;
}

function getNextStrategicOrderValue(item, items, mode) {
  const field = getStrategicOrderField(mode);
  const groupKey = getStrategicOrderGroupKey(item, mode);
  const peers = items.filter((entry) => entry.id !== item.id && getStrategicOrderGroupKey(entry, mode) === groupKey);
  const orders = peers.map((entry) => Number(entry[field])).filter((value) => Number.isFinite(value)).sort((a, b) => a - b);
  if (orders.length === 0) return 10;
  return orders[orders.length - 1] + 10;
}

function clampPlanningMinutes(minutes, fallback = DAY_START) {
  const parsed = Number(minutes);
  const safe = Number.isFinite(parsed) ? parsed : fallback;
  return Math.max(DAY_START, Math.min(DAY_END, safe));
}

function getFlagList(item) {
  const flags = [];
  if (item.watershedRestricted) flags.push({ label: "Watershed restricted", short: "WS" });
  if (item.primeTime) flags.push({ label: "Prime time", short: "PT" });
  if (item.mustRun) flags.push({ label: "Must-run", short: "MR" });
  return flags;
}

function getStatusLabel(issues) {
  if (issues.some((issue) => issue.severity === "blocker")) return "Blocked";
  if (issues.some((issue) => issue.severity === "warning")) return "Warning";
  return "Ready";
}

function getStatusClass(issues) {
  if (issues.some((issue) => issue.severity === "blocker")) return "blocked";
  if (issues.some((issue) => issue.severity === "warning")) return "warning";
  return "ok";
}

function validateItem(item, allItems = state.items, profile = state.exportProfile) {
  const issues = [];
  const strict = profile === "fs42-native";
  const exportedItem = !ITEM_TYPE_META[item.itemType]?.commercial;
  const title = String(item.title || "").trim();
  const channelValid = state.channels.some((channel) => channel.id === item.channelId);
  const startMinutes = parseTimeMinutes(item.start);
  const duration = Number(item.duration);
  const assetCode = String(item.assetCode || "").trim();

  if (!title) {
    issues.push({ severity: "blocker", code: "title", message: "Title is required." });
  }
  if (!channelValid) {
    issues.push({ severity: "blocker", code: "channel", message: "Channel is missing or invalid." });
  }
  if (!DAY_NAMES.includes(item.day)) {
    issues.push({ severity: "blocker", code: "day", message: "Day selection is invalid." });
  }
  if (!Number.isFinite(startMinutes)) {
    issues.push({ severity: "blocker", code: "start", message: "Start time is invalid." });
  }
  if (!Number.isFinite(duration) || duration <= 0) {
    issues.push({ severity: "blocker", code: "duration", message: "Duration must be greater than zero." });
  } else if (duration < 5) {
    issues.push({ severity: "blocker", code: "duration-min", message: "Duration must be at least 5 minutes." });
  }
  if (Number.isFinite(startMinutes) && Number.isFinite(duration)) {
    const end = startMinutes + duration;
    if (startMinutes < DAY_START) {
      issues.push({
        severity: strict && exportedItem ? "blocker" : "warning",
        code: "day-start",
        message: `Start time is before the planning day opens at ${minutesToTime(DAY_START)}.`,
      });
    }
    if (end > DAY_END) {
      issues.push({
        severity: strict && exportedItem ? "blocker" : "warning",
        code: "day-end",
        message: `Block runs past the planning day end at ${minutesToTime(DAY_END)}.`,
      });
    }
  }
  if (!assetCode) {
    issues.push({ severity: "warning", code: "assetCode", message: "Export code is missing." });
  }

  const laneItems = allItems.filter(
    (entry) =>
      entry.id !== item.id &&
      entry.channelId === item.channelId &&
      entry.category === item.category &&
      entry.day === item.day &&
      (!strict || !exportedItem || !ITEM_TYPE_META[entry.itemType]?.commercial),
  );
  if (Number.isFinite(startMinutes) && Number.isFinite(duration)) {
    const end = startMinutes + duration;
    laneItems.forEach((entry) => {
      const entryStart = parseTimeMinutes(entry.start);
      const entryEnd = Number.isFinite(entryStart) ? entryStart + getSafeDuration(entry.duration, entry.itemType) : NaN;
      if (Number.isFinite(entryStart) && Number.isFinite(entryEnd) && end > entryStart && startMinutes < entryEnd) {
        issues.push({
          severity: strict && exportedItem ? "blocker" : "warning",
          code: "overlap",
          message: `Overlaps with ${entry.title} in the same lane.`,
        });
      }
    });
  }

  if (item.primeTime && item.slot !== "Prime Time") {
    issues.push({
      severity: "warning",
      code: "prime-time",
      message: "Prime time flag is set, but the item is not in the Prime Time daypart.",
    });
  }

  return issues;
}

function validateSchedule(profile = state.exportProfile) {
  const itemIssues = new Map();
  const blockers = [];
  const warnings = [];

  const validationItems =
    profile === "fs42-native" ? state.items.filter((item) => !ITEM_TYPE_META[item.itemType]?.commercial) : state.items;

  validationItems.forEach((item) => {
    const issues = validateItem(item, validationItems, profile);
    itemIssues.set(item.id, issues);
    issues.forEach((issue) => {
      const payload = { ...issue, itemId: item.id, title: item.title, channelId: item.channelId };
      if (issue.severity === "blocker") blockers.push(payload);
      else warnings.push(payload);
    });
  });

  return {
    profile,
    itemIssues,
    blockers,
    warnings,
    ready: blockers.length === 0,
  };
}

function validateChannelConfig(channel, scheduleItems = state.items, profile = state.exportProfile, channels = state.channels) {
  const issues = [];
  const title = channel.name || "Untitled channel";
  const channelNumber = parsePositiveInteger(channel.channelNumber);
  const networkType = String(channel.networkType || "").trim();
  const scheduleIncrement = parsePositiveInteger(channel.scheduleIncrement);
  const breakStrategy = String(channel.breakStrategy || "").trim();
  const breakDuration = Number(channel.breakDuration);
  const commercialFree = typeof channel.commercialFree === "boolean" ? channel.commercialFree : null;
  const clipShows = normalizeClipShowList(channel.clipShows);
  const commercialCount = scheduleItems.filter(
    (item) => item.channelId === channel.id && ITEM_TYPE_META[item.itemType]?.commercial,
  ).length;
  const duplicateNumber = channelNumber
    ? channels.some((entry) => entry.id !== channel.id && parsePositiveInteger(entry.channelNumber) === channelNumber)
    : false;

  if (!channelNumber) {
    issues.push({
      severity: "blocker",
      code: "channel-number",
      message: "Channel number must be a positive integer.",
    });
  } else if (duplicateNumber) {
    issues.push({
      severity: "blocker",
      code: "channel-number-duplicate",
      message: `Channel number ${channelNumber} is already used by another channel.`,
    });
  }

  if (!NETWORK_TYPES.includes(networkType)) {
    issues.push({
      severity: "blocker",
      code: "network-type",
      message: "Network type must be standard, web, guide, loop, or streaming.",
    });
  }

  if (!scheduleIncrement || !Number.isInteger(scheduleIncrement)) {
    issues.push({
      severity: "blocker",
      code: "schedule-increment",
      message: "Schedule increment must be a positive whole number.",
    });
  }

  if (!BREAK_STRATEGIES.includes(breakStrategy)) {
    issues.push({
      severity: "blocker",
      code: "break-strategy",
      message: "Break strategy must be standard, end, or center.",
    });
  }

  if (commercialFree === null) {
    issues.push({
      severity: "blocker",
      code: "commercial-free",
      message: "Commercial free must be set to yes or no.",
    });
  }

  if (!Number.isInteger(breakDuration) || breakDuration <= 0) {
    issues.push({
      severity: "blocker",
      code: "break-duration",
      message: "Break duration must be a positive whole number.",
    });
  }

  if (clipShows.some((entry) => !entry)) {
    issues.push({
      severity: "warning",
      code: "clip-shows",
      message: "Clip shows should not contain blank entries.",
    });
  }

  if (commercialCount > 0 && commercialFree) {
    issues.push({
      severity: "warning",
      code: "commercial-free-mismatch",
      message: "Commercial items exist on this channel, but commercial_free is set to yes.",
    });
  }

  if (commercialCount > 0 && breakStrategy === "end") {
    issues.push({
      severity: "warning",
      code: "break-strategy-note",
      message: "This channel uses end breaks. Check that this matches the channel's real commercial pattern.",
    });
  }

  if (channel.multiLogoMode && !String(channel.multiLogoProfile || "").trim()) {
    issues.push({
      severity: "blocker",
      code: "multi-logo-profile",
      message: "Multi-logo mode is on, but no profile name is set.",
    });
  }

  return issues;
}

function validateFs42NativeExport() {
  const scheduleValidation = validateSchedule("fs42-native");
  const nativeBlockers = [];
  const nativeWarnings = [];

  state.channels.forEach((channel, index) => {
    const channelNumber = parsePositiveInteger(channel.channelNumber) || index + 1;
    const channelItems = state.items.filter((item) => item.channelId === channel.id);
    const channelIssues = validateChannelConfig(channel, channelItems, "fs42-native", state.channels);
    channelIssues.forEach((issue) => {
      const payload = {
        ...issue,
        title: channel.name || `Channel ${channelNumber}`,
        channelId: channel.id,
      };
      if (issue.severity === "warning") nativeWarnings.push(payload);
      else nativeBlockers.push(payload);
    });

    const stationConf = buildFs42NativeStationConfig(channel, channelItems, channelNumber);
    const issues = validateFs42NativePayload({ station_conf: stationConf }, channel, channelNumber);
    issues.forEach((issue) => {
      nativeBlockers.push({
        ...issue,
        itemId: issue.itemId || channel.id,
        title: issue.title || channel.name || `Channel ${channelNumber}`,
        channelId: channel.id,
      });
    });
  });

  const blockers = [...scheduleValidation.blockers, ...nativeBlockers];
  const warnings = [...scheduleValidation.warnings, ...nativeWarnings];

  return {
    profile: "fs42-native",
    blockers,
    warnings,
    ready: blockers.length === 0,
  };
}

function renderIssueLine(issue) {
  return `${issue.title}: ${issue.message}`;
}

function initTimelineDebugPanel() {
  if (timelineDebugPanel) return;

  const panel = document.createElement("aside");
  panel.className = "timeline-debug-panel";
  panel.innerHTML = `
    <div class="timeline-debug-title">Timeline DnD Debug</div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Drag event</div>
      <div class="timeline-debug-row"><span>Drop status</span><code data-debug-field="dropStatus">-</code></div>
      <div class="timeline-debug-row"><span>Drop fired</span><code data-debug-field="dropFired">-</code></div>
      <div class="timeline-debug-row"><span>Drop committed</span><code data-debug-field="dropCommitted">-</code></div>
      <div class="timeline-debug-row"><span>Drop allowed</span><code data-debug-field="dropAllowed">-</code></div>
      <div class="timeline-debug-row"><span>Blocked reason</span><code data-debug-field="dropBlockedReason">-</code></div>
      <div class="timeline-debug-row"><span>Stage</span><code data-debug-field="stage">idle</code></div>
      <div class="timeline-debug-row"><span>Last event</span><code data-debug-field="lastEventType">-</code></div>
      <div class="timeline-debug-row"><span>Item</span><code data-debug-field="item">-</code></div>
      <div class="timeline-debug-row"><span>Item id</span><code data-debug-field="itemId">-</code></div>
      <div class="timeline-debug-row"><span>Channel / lane</span><code data-debug-field="lane">-</code></div>
      <div class="timeline-debug-row"><span>Event</span><code data-debug-field="event">-</code></div>
      <div class="timeline-debug-row"><span>Target</span><code data-debug-field="target">-</code></div>
      <div class="timeline-debug-row"><span>Current target</span><code data-debug-field="currentTarget">-</code></div>
      <div class="timeline-debug-row"><span>Closest lane</span><code data-debug-field="closestLane">-</code></div>
      <div class="timeline-debug-row"><span>Pointer offset</span><code data-debug-field="pointerOffset">-</code></div>
      <div class="timeline-debug-row"><span>Original mins</span><code data-debug-field="originalMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Original time</span><code data-debug-field="originalTime">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Geometry</div>
      <div class="timeline-debug-row"><span>Lane left</span><code data-debug-field="laneLeft">-</code></div>
      <div class="timeline-debug-row"><span>Lane width</span><code data-debug-field="laneWidth">-</code></div>
      <div class="timeline-debug-row"><span>Shell scrollLeft</span><code data-debug-field="scrollLeft">-</code></div>
      <div class="timeline-debug-row"><span>Shell width</span><code data-debug-field="shellWidth">-</code></div>
      <div class="timeline-debug-row"><span>Shell scrollWidth</span><code data-debug-field="shellScrollWidth">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Derived X</div>
      <div class="timeline-debug-row"><span>Visible lane X</span><code data-debug-field="visibleLaneX">-</code></div>
      <div class="timeline-debug-row"><span>Timeline content X</span><code data-debug-field="timelineContentX">-</code></div>
      <div class="timeline-debug-row"><span>Adjusted X</span><code data-debug-field="adjustedX">-</code></div>
      <div class="timeline-debug-row"><span>Fallback X</span><code data-debug-field="fallbackX">-</code></div>
      <div class="timeline-debug-row"><span>X source</span><code data-debug-field="xSource">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Minutes</div>
      <div class="timeline-debug-row"><span>Pre-snap</span><code data-debug-field="preSnapMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Snapped</span><code data-debug-field="snappedMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Clamped</span><code data-debug-field="clampedMinutes">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Final result</div>
      <div class="timeline-debug-row"><span>Saved minutes</span><code data-debug-field="finalMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Saved time</span><code data-debug-field="finalTime">-</code></div>
      <div class="timeline-debug-row"><span>Rendered left</span><code data-debug-field="renderedLeft">-</code></div>
      <div class="timeline-debug-row"><span>Render note</span><code data-debug-field="renderNote">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Commit snapshot</div>
      <div class="timeline-debug-row"><span>Commit status</span><code data-debug-field="commitStatus">-</code></div>
      <div class="timeline-debug-row"><span>Drop fired</span><code data-debug-field="commitDropFired">-</code></div>
      <div class="timeline-debug-row"><span>Drop committed</span><code data-debug-field="commitDropCommitted">-</code></div>
      <div class="timeline-debug-row"><span>Committed item</span><code data-debug-field="commitItem">-</code></div>
      <div class="timeline-debug-row"><span>Committed item id</span><code data-debug-field="commitItemId">-</code></div>
      <div class="timeline-debug-row"><span>Original start</span><code data-debug-field="commitOriginalTime">-</code></div>
      <div class="timeline-debug-row"><span>Original mins</span><code data-debug-field="commitOriginalMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Target lane</span><code data-debug-field="commitTargetLane">-</code></div>
      <div class="timeline-debug-row"><span>Target channel</span><code data-debug-field="commitTargetChannelId">-</code></div>
      <div class="timeline-debug-row"><span>Target category</span><code data-debug-field="commitTargetCategory">-</code></div>
      <div class="timeline-debug-row"><span>Same lane</span><code data-debug-field="commitSameLane">-</code></div>
      <div class="timeline-debug-row"><span>Content X</span><code data-debug-field="commitTimelineContentX">-</code></div>
      <div class="timeline-debug-row"><span>Adjusted X</span><code data-debug-field="commitAdjustedX">-</code></div>
      <div class="timeline-debug-row"><span>Pre-snap</span><code data-debug-field="commitPreSnapMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Snapped</span><code data-debug-field="commitSnappedMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Clamped</span><code data-debug-field="commitClampedMinutes">-</code></div>
      <div class="timeline-debug-row"><span>New start</span><code data-debug-field="commitNewStart">-</code></div>
      <div class="timeline-debug-row"><span>State updated</span><code data-debug-field="commitStateUpdated">-</code></div>
      <div class="timeline-debug-row"><span>Render reflected</span><code data-debug-field="commitRenderReflected">-</code></div>
      <div class="timeline-debug-row"><span>Rendered left</span><code data-debug-field="commitRenderedLeft">-</code></div>
      <div class="timeline-debug-row"><span>Rendered label</span><code data-debug-field="commitRenderedLabel">-</code></div>
      <div class="timeline-debug-row"><span>Post-check start</span><code data-debug-field="commitPostCheckStart">-</code></div>
      <div class="timeline-debug-row"><span>Reverted</span><code data-debug-field="commitReverted">-</code></div>
      <div class="timeline-debug-row"><span>Target detail</span><code data-debug-field="commitTargetDetail">-</code></div>
    </div>
    <div class="timeline-debug-section">
      <div class="timeline-debug-section-title">Last timing snapshot</div>
      <div class="timeline-debug-row"><span>Snapshot status</span><code data-debug-field="timingStatus">-</code></div>
      <div class="timeline-debug-row"><span>Item</span><code data-debug-field="timingItem">-</code></div>
      <div class="timeline-debug-row"><span>Item id</span><code data-debug-field="timingItemId">-</code></div>
      <div class="timeline-debug-row"><span>Original start</span><code data-debug-field="timingOriginalTime">-</code></div>
      <div class="timeline-debug-row"><span>Original mins</span><code data-debug-field="timingOriginalMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Hover pre-snap</span><code data-debug-field="hoverPreSnapMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Hover snapped</span><code data-debug-field="hoverSnappedMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Hover clamped</span><code data-debug-field="hoverClampedMinutes">-</code></div>
      <div class="timeline-debug-row"><span>Hover lane</span><code data-debug-field="hoverTargetLane">-</code></div>
      <div class="timeline-debug-row"><span>Hover channel</span><code data-debug-field="hoverTargetChannelId">-</code></div>
      <div class="timeline-debug-row"><span>Commit pre-snap</span><code data-debug-field="commitPreSnapMinutesTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit snapped</span><code data-debug-field="commitSnappedMinutesTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit clamped</span><code data-debug-field="commitClampedMinutesTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit lane</span><code data-debug-field="commitTargetLaneTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit channel</span><code data-debug-field="commitTargetChannelIdTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit category</span><code data-debug-field="commitTargetCategoryTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit raw X</span><code data-debug-field="commitRawXTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit adjusted X</span><code data-debug-field="commitAdjustedXTiming">-</code></div>
      <div class="timeline-debug-row"><span>Commit new start</span><code data-debug-field="commitNewStartTiming">-</code></div>
      <div class="timeline-debug-row"><span>Final saved start</span><code data-debug-field="finalSavedStartTiming">-</code></div>
      <div class="timeline-debug-row"><span>Final saved time</span><code data-debug-field="finalSavedTimeTiming">-</code></div>
      <div class="timeline-debug-row"><span>Final rendered left</span><code data-debug-field="finalRenderedLeftTiming">-</code></div>
      <div class="timeline-debug-row"><span>Hover -> commit</span><code data-debug-field="matchHoverCommit">-</code></div>
      <div class="timeline-debug-row"><span>Commit -> saved</span><code data-debug-field="matchCommitSaved">-</code></div>
      <div class="timeline-debug-row"><span>Saved -> recheck</span><code data-debug-field="matchSavedRecheck">-</code></div>
      <div class="timeline-debug-row"><span>Reverted after commit</span><code data-debug-field="timingReverted">-</code></div>
      <div class="timeline-debug-row"><span>Post-check start</span><code data-debug-field="timingPostCheckStart">-</code></div>
    </div>
  `;

  document.body.appendChild(panel);
  timelineDebugPanel = {
    root: panel,
    fields: Object.fromEntries(
      Array.from(panel.querySelectorAll("[data-debug-field]"), (node) => [node.dataset.debugField, node]),
    ),
  };
}

function setTimelineDebugField(name, value) {
  if (!DEBUG_TIMELINE_DND || !timelineDebugPanel) return;
  const node = timelineDebugPanel.fields[name];
  if (!node) return;
  node.textContent = formatTimelineDebugValue(value);
}

function formatTimelineDebugValue(value) {
  if (value === null || value === undefined) return "-";
  if (typeof value === "number" && Number.isFinite(value)) return Number.isInteger(value) ? String(value) : value.toFixed(2);
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "string" && value.length === 0) return "-";
  return String(value);
}

function logTimelineDebug(stage, snapshot) {
  if (!DEBUG_TIMELINE_DND) return;
  console.log(`[timeline-dnd] ${stage}`, snapshot);
}

function describeTimelineNode(node) {
  if (!node) return "-";
  const tagName = node.tagName || node.nodeName || "node";
  const className =
    typeof node.className === "string"
      ? node.className.trim().replace(/\s+/g, ".")
      : node.classList
        ? Array.from(node.classList).join(".")
        : "";
  const datasetParts = [];
  if (node.dataset) {
    if (node.dataset.channelId) datasetParts.push(`channel=${node.dataset.channelId}`);
    if (node.dataset.category) datasetParts.push(`category=${node.dataset.category}`);
    if (node.dataset.itemId) datasetParts.push(`item=${node.dataset.itemId}`);
  }
  return [tagName, className ? `.${className}` : "", datasetParts.length > 0 ? ` [${datasetParts.join(" ")}]` : ""].join("");
}

function resolveTimelineLaneFromEvent(event) {
  const path = typeof event?.composedPath === "function" ? event.composedPath() : [];
  for (const node of path) {
    if (node?.classList?.contains?.("day-lane")) return node;
  }
  if (event?.target?.closest) return event.target.closest(".day-lane");
  if (event?.currentTarget?.classList?.contains?.("day-lane")) return event.currentTarget;
  return null;
}

function updateTimelineDebugPanel(snapshot) {
  if (!DEBUG_TIMELINE_DND || !timelineDebugPanel || !snapshot) return;

  setTimelineDebugField("stage", snapshot.stage);
  setTimelineDebugField("dropStatus", snapshot.dropStatus);
  setTimelineDebugField("dropFired", snapshot.dropFired);
  setTimelineDebugField("dropCommitted", snapshot.dropCommitted);
  setTimelineDebugField("dropAllowed", snapshot.dropAllowed);
  setTimelineDebugField("dropBlockedReason", snapshot.dropBlockedReason);
  setTimelineDebugField("lastEventType", snapshot.lastEventType);
  setTimelineDebugField("item", snapshot.itemLabel);
  setTimelineDebugField("itemId", snapshot.itemId);
  setTimelineDebugField("lane", snapshot.laneLabel);
  setTimelineDebugField("event", snapshot.eventLabel);
  setTimelineDebugField("target", snapshot.targetLabel);
  setTimelineDebugField("currentTarget", snapshot.currentTargetLabel);
  setTimelineDebugField("closestLane", snapshot.closestLaneLabel);
  setTimelineDebugField("pointerOffset", snapshot.pointerOffsetX);
  setTimelineDebugField("originalMinutes", snapshot.originalMinutes);
  setTimelineDebugField("originalTime", snapshot.originalTime);
  setTimelineDebugField("laneLeft", snapshot.laneLeft);
  setTimelineDebugField("laneWidth", snapshot.laneWidth);
  setTimelineDebugField("scrollLeft", snapshot.scrollLeft);
  setTimelineDebugField("shellWidth", snapshot.shellClientWidth);
  setTimelineDebugField("shellScrollWidth", snapshot.shellScrollWidth);
  setTimelineDebugField("visibleLaneX", snapshot.visibleLaneX);
  setTimelineDebugField("timelineContentX", snapshot.timelineContentX);
  setTimelineDebugField("adjustedX", snapshot.adjustedX);
  setTimelineDebugField("fallbackX", snapshot.fallbackX);
  setTimelineDebugField("xSource", snapshot.xSource);
  setTimelineDebugField("preSnapMinutes", snapshot.preSnapMinutes);
  setTimelineDebugField("snappedMinutes", snapshot.snappedMinutes);
  setTimelineDebugField("clampedMinutes", snapshot.clampedMinutes);
  setTimelineDebugField("finalMinutes", snapshot.finalMinutes);
  setTimelineDebugField("finalTime", snapshot.finalTime);
  setTimelineDebugField("renderedLeft", snapshot.renderedLeft);
  setTimelineDebugField("renderNote", snapshot.renderNote);
  const commit = snapshot.commitSnapshot || dragState.commitSnapshot;
  setTimelineDebugField("commitItem", commit?.itemTitle);
  setTimelineDebugField("commitStatus", commit?.status);
  setTimelineDebugField("commitDropFired", commit?.dropFired);
  setTimelineDebugField("commitDropCommitted", commit?.dropCommitted);
  setTimelineDebugField("commitItemId", commit?.itemId);
  setTimelineDebugField("commitOriginalTime", commit?.originalStartTime);
  setTimelineDebugField("commitOriginalMinutes", commit?.originalStartMinutes);
  setTimelineDebugField("commitTargetLane", commit?.targetLaneLabel);
  setTimelineDebugField("commitTargetChannelId", commit?.targetChannelId);
  setTimelineDebugField("commitTargetCategory", commit?.targetLaneCategory);
  setTimelineDebugField("commitSameLane", commit?.sameLane);
  setTimelineDebugField("commitTimelineContentX", commit?.timelineContentX);
  setTimelineDebugField("commitAdjustedX", commit?.adjustedX);
  setTimelineDebugField("commitPreSnapMinutes", commit?.preSnapMinutes);
  setTimelineDebugField("commitSnappedMinutes", commit?.snappedMinutes);
  setTimelineDebugField("commitClampedMinutes", commit?.clampedMinutes);
  setTimelineDebugField("commitNewStart", commit?.newStartTime);
  setTimelineDebugField("commitStateUpdated", commit?.stateUpdated);
  setTimelineDebugField("commitRenderReflected", commit?.renderReflected);
  setTimelineDebugField("commitRenderedLeft", commit?.renderedLeft);
  setTimelineDebugField("commitRenderedLabel", commit?.renderVisibleTimeLabel);
  setTimelineDebugField("commitPostCheckStart", commit?.postCheckStart);
  setTimelineDebugField("commitReverted", commit?.revertedAfterCommit);
  setTimelineDebugField("commitTargetDetail", commit?.targetDetail);

  const timing = snapshot.timingSnapshot || dragState.timingSnapshot;
  setTimelineDebugField("timingStatus", timing?.status);
  setTimelineDebugField("timingItem", timing?.itemTitle);
  setTimelineDebugField("timingItemId", timing?.itemId);
  setTimelineDebugField("timingOriginalTime", timing?.originalStartTime);
  setTimelineDebugField("timingOriginalMinutes", timing?.originalStartMinutes);
  setTimelineDebugField("hoverPreSnapMinutes", timing?.hover?.preSnapMinutes);
  setTimelineDebugField("hoverSnappedMinutes", timing?.hover?.snappedMinutes);
  setTimelineDebugField("hoverClampedMinutes", timing?.hover?.clampedMinutes);
  setTimelineDebugField("hoverTargetLane", timing?.hover?.targetLaneLabel);
  setTimelineDebugField("hoverTargetChannelId", timing?.hover?.targetChannelId);
  setTimelineDebugField("commitPreSnapMinutesTiming", timing?.commit?.preSnapMinutes);
  setTimelineDebugField("commitSnappedMinutesTiming", timing?.commit?.snappedMinutes);
  setTimelineDebugField("commitClampedMinutesTiming", timing?.commit?.clampedMinutes);
  setTimelineDebugField("commitTargetLaneTiming", timing?.commit?.targetLaneLabel);
  setTimelineDebugField("commitTargetChannelIdTiming", timing?.commit?.targetChannelId);
  setTimelineDebugField("commitTargetCategoryTiming", timing?.commit?.targetLaneCategory);
  setTimelineDebugField("commitRawXTiming", timing?.commit?.rawX);
  setTimelineDebugField("commitAdjustedXTiming", timing?.commit?.adjustedX);
  setTimelineDebugField("commitNewStartTiming", timing?.commit?.newStartTime);
  setTimelineDebugField("finalSavedStartTiming", timing?.final?.savedStart);
  setTimelineDebugField("finalSavedTimeTiming", timing?.final?.savedTimeLabel);
  setTimelineDebugField("finalRenderedLeftTiming", timing?.final?.renderedLeft);
  setTimelineDebugField("matchHoverCommit", timing?.matches?.hoverToCommit);
  setTimelineDebugField("matchCommitSaved", timing?.matches?.commitToSaved);
  setTimelineDebugField("matchSavedRecheck", timing?.matches?.savedToRecheck);
  setTimelineDebugField("timingReverted", timing?.postRender?.revertedAfterCommit);
  setTimelineDebugField("timingPostCheckStart", timing?.postRender?.postCheckStart);
}

function logTimelineEvent(stage, event, lane, extras = {}) {
  if (!DEBUG_TIMELINE_DND) return;
  const resolvedLane = resolveTimelineLaneFromEvent(event) || lane || null;
  const item = state.items.find((entry) => entry.id === dragState.itemId) || dragState.draggedItemSnapshot || null;
  const snapshot = buildTimelineDebugSnapshot(stage, event, resolvedLane, item, {
    ...extras,
    lastEventType: event?.type || stage,
    targetLabel: describeTimelineNode(event?.target),
    currentTargetLabel: describeTimelineNode(event?.currentTarget),
    closestLaneLabel: describeTimelineNode(resolvedLane),
  });
  updateTimelineDebugPanel(snapshot);
  logTimelineDebug(stage, snapshot);
  return snapshot;
}

function buildTimelineDebugSnapshot(stage, event, lane, item, extras = {}) {
  const geometry = getTimelinePointerGeometry(event, lane);
  const sourceItem = dragState.draggedItemSnapshot || item || null;
  const targetLabel = extras.targetLabel || describeTimelineNode(event?.target);
  const currentTargetLabel = extras.currentTargetLabel || describeTimelineNode(event?.currentTarget);
  const closestLaneLabel = extras.closestLaneLabel || describeTimelineNode(resolveTimelineLaneFromEvent(event) || lane);
  const visibleLaneX = Number.isFinite(extras.visibleLaneX) ? extras.visibleLaneX : geometry.visibleLaneX;
  const timelineContentX =
    Number.isFinite(extras.timelineContentX) ? extras.timelineContentX : geometry.timelineContentX;
  const adjustedX =
    Number.isFinite(extras.adjustedX)
      ? extras.adjustedX
      : Number.isFinite(timelineContentX) && Number.isFinite(extras.pointerOffsetX)
        ? timelineContentX - extras.pointerOffsetX
        : null;
  const preSnapMinutes = Number.isFinite(extras.preSnapMinutes)
    ? extras.preSnapMinutes
    : Number.isFinite(adjustedX)
      ? timelineXToMinutes(adjustedX)
      : null;
  const snappedMinutes = Number.isFinite(extras.snappedMinutes)
    ? extras.snappedMinutes
    : Number.isFinite(preSnapMinutes)
      ? snapStartMinutes(item || { itemType: "Programme", channelId: "", category: "" }, preSnapMinutes)
      : null;
  const clampedMinutes = Number.isFinite(extras.clampedMinutes) ? extras.clampedMinutes : null;

  return {
    stage,
    lastEventType: extras.lastEventType || (event?.type || "-"),
    dropStatus: extras.dropStatus || "-",
    dropFired: extras.dropFired || "-",
    dropCommitted: extras.dropCommitted || "-",
    dropAllowed: extras.dropAllowed || dragState.dropAllowed,
    dropBlockedReason: extras.dropBlockedReason || dragState.dropBlockedReason,
    itemLabel: sourceItem ? `${sourceItem.title} (${sourceItem.id})` : "-",
    itemId: sourceItem ? sourceItem.id : "-",
    laneLabel: sourceItem ? `${getChannelName(sourceItem.channelId)} / ${sourceItem.category}` : "-",
    eventLabel:
      Number.isFinite(geometry.clientX) || Number.isFinite(geometry.clientY)
        ? `clientX=${formatTimelineDebugValue(geometry.clientX)} clientY=${formatTimelineDebugValue(geometry.clientY)}`
        : "-",
    pointerOffsetX: dragState.pointerOffsetX,
    originalMinutes: sourceItem ? getSafeStartMinutes(sourceItem.start) : null,
    originalTime: sourceItem ? sourceItem.start : "-",
    laneLeft: geometry.laneLeft,
    laneWidth: geometry.laneWidth,
    scrollLeft: geometry.shellScrollLeft,
    shellClientWidth: geometry.shellClientWidth,
    shellScrollWidth: geometry.shellScrollWidth,
    targetLabel,
    currentTargetLabel,
    closestLaneLabel,
    visibleLaneX,
    timelineContentX,
    adjustedX,
    fallbackX: Number.isFinite(extras.fallbackX) ? extras.fallbackX : null,
    xSource: extras.xSource || "-",
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    finalMinutes: Number.isFinite(extras.finalMinutes) ? extras.finalMinutes : null,
    finalTime: extras.finalTime || "-",
    renderedLeft: Number.isFinite(extras.renderedLeft) ? extras.renderedLeft : null,
    renderNote: extras.renderNote || "-",
  };
}

function handleBlockDragStart(event) {
  if (state.timelineScale !== "day") {
    event.preventDefault();
    return;
  }
  const item = state.items.find((entry) => entry.id === event.currentTarget.dataset.itemId);
  dragState.itemId = event.currentTarget.dataset.itemId;
  dragState.commitSnapshot = null;
  dragState.timingSnapshot = {
    status: "dragging",
    hover: null,
    commit: null,
    final: null,
    postRender: null,
    matches: {
      hoverToCommit: "-",
      commitToSaved: "-",
      savedToRecheck: "-",
    },
    itemId: item?.id || dragState.itemId || "-",
    itemTitle: item?.title || "-",
    originalStartMinutes: item ? getSafeStartMinutes(item.start) : null,
    originalStartTime: item?.start || "-",
    revertedAfterCommit: "no",
  };
  dragState.draggedItemSnapshot = item
    ? {
        id: item.id,
        title: item.title,
        start: item.start,
        duration: item.duration,
        channelId: item.channelId,
        category: item.category,
        itemType: item.itemType,
      }
    : null;
  const blockRect = event.currentTarget.getBoundingClientRect();
  dragState.pointerOffsetX = event.clientX - blockRect.left;
  dragState.lastClientX = event.clientX;
  dragState.lastTimelineX = null;
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", dragState.itemId || "");
  }
  event.currentTarget.classList.add("dragging");

  if (DEBUG_TIMELINE_DND) {
    dragState.lastEventType = event.type;
    dragState.lastTargetLabel = describeTimelineNode(event.target);
    dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
    dragState.lastClosestLaneLabel = "-";
    dragState.dropAllowed = "unknown";
    dragState.dropBlockedReason = "-";
    dragState.lastDropCommitted = false;
    const snapshot = buildTimelineDebugSnapshot("dragstart", event, null, item, {
      xSource: "dragstart",
      dropStatus: "DRAG ACTIVE",
      dropFired: "no",
      dropCommitted: "no",
      dropAllowed: "unknown",
      dropBlockedReason: "-",
      originalMinutes: dragState.draggedItemSnapshot ? getSafeStartMinutes(dragState.draggedItemSnapshot.start) : null,
      originalTime: dragState.draggedItemSnapshot ? dragState.draggedItemSnapshot.start : "-",
      eventLabel: `clientX=${formatTimelineDebugValue(event.clientX)} clientY=${formatTimelineDebugValue(event.clientY)}`,
      targetLabel: describeTimelineNode(event.target),
      currentTargetLabel: describeTimelineNode(event.currentTarget),
      closestLaneLabel: "-",
      renderNote: "block picked up",
    });
    updateTimelineDebugPanel(snapshot);
    logTimelineDebug("dragstart", snapshot);
  }
}

function handleBlockDragEnd(event) {
  event.currentTarget.classList.remove("dragging");
  clearLaneHighlights();
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = "-";
  dragState.itemId = null;
  dragState.pointerOffsetX = 0;
  dragState.lastClientX = null;
  dragState.lastTimelineX = null;
  dragState.draggedItemSnapshot = null;
  dragState.dropAllowed = "unknown";
  dragState.dropBlockedReason = "-";
  if (DEBUG_TIMELINE_DND) {
    logTimelineEvent("dragend", event, null, {
      dropStatus: dragState.lastDropCommitted ? "DRAG END AFTER DROP" : "DROP NEVER FIRED",
      dropFired: dragState.lastDropCommitted ? "yes" : "no",
      dropCommitted: dragState.lastDropCommitted ? "yes" : "no",
      renderNote: "drag state cleared",
    });
  }
  dragState.lastDropCommitted = false;
}

function handleLaneDragOver(event) {
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  dragState.lastClientX = event.clientX;
  dragState.lastTimelineX = getTimelineContentXFromPointer(event, event.currentTarget);
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(resolveTimelineLaneFromEvent(event) || event.currentTarget);
  dragState.dropAllowed = "yes";
  dragState.dropBlockedReason = "-";
  clearLaneHighlights();
  event.currentTarget.classList.add("drop-target");

  if (DEBUG_TIMELINE_DND) {
    const item = state.items.find((entry) => entry.id === dragState.itemId);
    const timelineX = dragState.lastTimelineX;
    const adjustedX = Number.isFinite(timelineX) ? timelineX - dragState.pointerOffsetX : null;
    const preSnapMinutes = Number.isFinite(adjustedX) ? timelineXToMinutes(adjustedX) : null;
    const snappedMinutes = Number.isFinite(preSnapMinutes) ? snapStartMinutes(item, preSnapMinutes) : null;
    const duration = item ? getSafeDuration(item.duration, item.itemType) : null;
    const clampedMinutes =
      Number.isFinite(snappedMinutes) && Number.isFinite(duration)
        ? Math.max(DAY_START, Math.min(DAY_END - duration, snappedMinutes))
        : null;
    const targetChannelId = event.currentTarget.dataset.channelId || null;
    const targetLaneCategory = event.currentTarget.dataset.category || null;
    dragState.timingSnapshot = dragState.timingSnapshot || {};
    dragState.timingSnapshot.status = "hovering";
    dragState.timingSnapshot.hover = {
      itemId: item?.id || dragState.itemId || "-",
      itemTitle: item?.title || dragState.draggedItemSnapshot?.title || "-",
      originalStartMinutes: dragState.draggedItemSnapshot ? getSafeStartMinutes(dragState.draggedItemSnapshot.start) : item ? getSafeStartMinutes(item.start) : null,
      originalStartTime: dragState.draggedItemSnapshot?.start || item?.start || "-",
      visibleLaneX: Number.isFinite(event.clientX) ? event.clientX - event.currentTarget.getBoundingClientRect().left : null,
      timelineContentX: timelineX,
      adjustedX,
      preSnapMinutes,
      snappedMinutes,
      clampedMinutes,
      targetChannelId,
      targetLaneCategory,
      targetLaneLabel: `${getChannelName(targetChannelId)} / ${targetLaneCategory}`,
    };
    console.log("[timeline-dnd] hover snapshot", {
      itemId: dragState.timingSnapshot.hover.itemId,
      originalStart: dragState.timingSnapshot.hover.originalStartTime,
      hoverSnapped: snappedMinutes,
      hoverClamped: clampedMinutes,
      targetChannelId,
      targetLaneCategory,
    });
    const snapshot = buildTimelineDebugSnapshot("dragover", event, event.currentTarget, item, {
      timelineContentX: timelineX,
      adjustedX,
      preSnapMinutes,
      snappedMinutes,
      clampedMinutes,
      dropAllowed: "yes",
      dropBlockedReason: "-",
      dropStatus: "DRAG OVER",
      dropFired: "no",
      dropCommitted: "no",
      xSource: "dragover:lastTimelineX",
      fallbackX: null,
      targetLabel: describeTimelineNode(event.target),
      currentTargetLabel: describeTimelineNode(event.currentTarget),
      closestLaneLabel: describeTimelineNode(resolveTimelineLaneFromEvent(event) || event.currentTarget),
      renderNote: "live hover preview",
    });
    updateTimelineDebugPanel(snapshot);
    logTimelineDebug("dragover", snapshot);
  }
}

function handleLaneDragLeave(event) {
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(resolveTimelineLaneFromEvent(event) || event.currentTarget);
  event.currentTarget.classList.remove("drop-target");
  if (DEBUG_TIMELINE_DND) {
    logTimelineEvent("dragleave", event, event.currentTarget, {
      dropStatus: "DRAG LEAVE",
      dropFired: "no",
      dropCommitted: "no",
      dropAllowed: dragState.dropAllowed,
      dropBlockedReason: dragState.dropBlockedReason,
      renderNote: "lane left",
    });
  }
}

function handleLaneDrop(event) {
  const lane = resolveTimelineLaneFromEvent(event) || event.currentTarget;
  commitTimelineDrop(event, lane, "lane");
}

function commitTimelineDrop(event, lane, source) {
  if (event?.__fs42TimelineDropHandled) {
    return { committed: false, dropSnapshot: null, renderSnapshot: null, duplicate: true };
  }
  if (event) event.__fs42TimelineDropHandled = true;
  const item = state.items.find((entry) => entry.id === dragState.itemId);
  const fallbackTimelineX = getTimelineContentXFromPointer(event, lane);
  const timelineX = Number.isFinite(dragState.lastTimelineX) ? dragState.lastTimelineX : fallbackTimelineX;
  const dropFired = Boolean(event?.type === "drop");
  const sourceItem = dragState.draggedItemSnapshot || item || null;
  const targetChannelId = lane?.dataset?.channelId || null;
  const targetLaneCategory = lane?.dataset?.category || null;
  const targetLaneLabel = lane ? `${getChannelName(targetChannelId)} / ${targetLaneCategory}` : "-";
  const targetDetail = [describeTimelineNode(event?.target), describeTimelineNode(event?.currentTarget), describeTimelineNode(resolveTimelineLaneFromEvent(event) || lane)].join(" | ");
  const originalStartMinutes = sourceItem ? getSafeStartMinutes(sourceItem.start) : null;
  const originalStartTime = sourceItem ? sourceItem.start : "-";
  const originalChannelId = sourceItem?.channelId || item?.channelId || null;
  const originalLaneCategory = sourceItem?.category || item?.category || null;
  const sameLane = Boolean(
    targetChannelId &&
      targetLaneCategory &&
      originalChannelId === targetChannelId &&
      originalLaneCategory === targetLaneCategory,
  );
  const timelineContentX = timelineX;
  const adjustedX = Number.isFinite(timelineX) ? timelineX - dragState.pointerOffsetX : null;
  const preSnapMinutes = Number.isFinite(adjustedX) ? timelineXToMinutes(adjustedX) : null;
  const snappedMinutes = Number.isFinite(preSnapMinutes) ? snapStartMinutes(item || sourceItem, preSnapMinutes) : null;
  const duration = item ? getSafeDuration(item.duration, item.itemType) : sourceItem ? getSafeDuration(sourceItem.duration, sourceItem.itemType) : null;
  const maxStart = Number.isFinite(duration) ? DAY_END - duration : DAY_END;
  const clampedMinutes = Number.isFinite(snappedMinutes)
    ? clampPlanningMinutes(Math.max(DAY_START, Math.min(maxStart, snappedMinutes)))
    : null;
  const plannedStartTime = minutesToTime(clampPlanningMinutes(Number.isFinite(clampedMinutes) ? clampedMinutes : DAY_START));

  dragState.timingSnapshot = dragState.timingSnapshot || {};
  dragState.timingSnapshot.status = "committing";
  dragState.timingSnapshot.commit = {
    itemId: sourceItem?.id || item?.id || "-",
    itemTitle: sourceItem?.title || item?.title || "-",
    rawX: timelineX,
    adjustedX,
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    targetChannelId,
    targetLaneCategory,
    targetLaneLabel,
    sameLane,
    newStartTime: plannedStartTime,
    newChannelId: targetChannelId,
    newLaneCategory: targetLaneCategory,
  };
  console.log("[timeline-dnd] commit snapshot", {
    itemId: dragState.timingSnapshot.commit.itemId,
    originalStart: dragState.timingSnapshot.originalStartTime,
    commitSnapped: snappedMinutes,
    commitClamped: clampedMinutes,
    plannedStart: plannedStartTime,
    targetChannelId,
    targetLaneCategory,
    sameLane,
  });

  dragState.commitSnapshot = {
    source,
    dropFired,
    dropCommitted: false,
    itemId: sourceItem?.id || item?.id || "-",
    itemTitle: sourceItem?.title || item?.title || "-",
    originalStartMinutes,
    originalStartTime,
    timelineContentX: timelineX,
    adjustedX,
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    targetChannelId,
    targetLaneCategory,
    targetLaneLabel,
    sameLane,
    targetDetail,
    targetLabel: describeTimelineNode(event?.target),
    currentTargetLabel: describeTimelineNode(event?.currentTarget),
    closestLaneLabel: describeTimelineNode(resolveTimelineLaneFromEvent(event) || lane),
    oldStart: sourceItem?.start || item?.start || "-",
    oldChannelId: originalChannelId,
    oldLaneCategory: originalLaneCategory,
    newStartTime: "-",
    newStartMinutes: null,
    newChannelId: null,
    newLaneCategory: null,
    stateUpdated: false,
    renderReflected: false,
    revertedAfterCommit: false,
    renderedLeft: null,
    renderVisibleTimeLabel: "-",
    status: "commit-pending",
  };

  if (!lane) {
    dragState.dropAllowed = "no";
    dragState.dropBlockedReason = "no lane target";
    dragState.commitSnapshot.status = "DROP FIRED BUT HAD NO VALID TARGET";
    if (DEBUG_TIMELINE_DND) {
      const blockedSnapshot = buildTimelineDebugSnapshot("drop", event, lane, item, {
        timelineContentX,
        adjustedX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        dropStatus: "DROP FIRED BUT HAD NO VALID TARGET",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "no",
        dropAllowed: "no",
        dropBlockedReason: "no lane target",
        targetLabel: dragState.commitSnapshot.targetLabel,
        currentTargetLabel: dragState.commitSnapshot.currentTargetLabel,
        closestLaneLabel: dragState.commitSnapshot.closestLaneLabel,
        renderNote: `source=${source}`,
        commitSnapshot: dragState.commitSnapshot,
      });
      updateTimelineDebugPanel(blockedSnapshot);
      logTimelineDebug("drop", blockedSnapshot);
      console.log("[timeline-dnd] commit blocked", {
        itemId: dragState.commitSnapshot.itemId,
        reason: "no lane target",
        targetDetail,
      });
    }
    return { committed: false, dropSnapshot: null, renderSnapshot: null };
  }

  if (!event?.defaultPrevented) {
    event.preventDefault();
  }
  if (event?.dataTransfer) {
    event.dataTransfer.dropEffect = "move";
  }

  if (!item) {
    dragState.dropAllowed = "yes";
    dragState.dropBlockedReason = "drag state empty";
    dragState.commitSnapshot.status = "DROP FIRED BUT DRAG STATE WAS EMPTY";
    if (DEBUG_TIMELINE_DND) {
      const blockedSnapshot = buildTimelineDebugSnapshot("drop", event, lane, null, {
        timelineContentX,
        adjustedX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        dropStatus: "DROP FIRED BUT DRAG STATE WAS EMPTY",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "no",
        dropAllowed: "yes",
        dropBlockedReason: "drag state empty",
        targetLabel: dragState.commitSnapshot.targetLabel,
        currentTargetLabel: dragState.commitSnapshot.currentTargetLabel,
        closestLaneLabel: dragState.commitSnapshot.closestLaneLabel,
        renderNote: `source=${source}`,
        commitSnapshot: dragState.commitSnapshot,
      });
      updateTimelineDebugPanel(blockedSnapshot);
      logTimelineDebug("drop", blockedSnapshot);
      console.log("[timeline-dnd] commit blocked", {
        itemId: dragState.commitSnapshot.itemId,
        reason: "drag state empty",
        targetDetail,
      });
    }
    return { committed: false, dropSnapshot: null, renderSnapshot: null };
  }

  console.log("[timeline-dnd] commit before update", {
    itemId: item.id,
    title: item.title,
    oldStart: item.start,
    oldStartMinutes: getSafeStartMinutes(item.start),
    oldChannelId: item.channelId,
    oldLaneCategory: item.category,
    newTargetChannelId: targetChannelId,
    newTargetLaneCategory: targetLaneCategory,
    sameLane,
    timelineContentX: timelineX,
    adjustedX,
    preSnapMinutes,
    snappedMinutes,
    clampedMinutes,
    targetDetail,
  });

  const relativeX = adjustedX;
  const magneticMinutes = Number.isFinite(clampedMinutes) ? clampedMinutes : DAY_START;
  const magneticStart = applyMagneticTargets(
    item,
    targetChannelId,
    targetLaneCategory,
    magneticMinutes,
  );

  item.channelId = targetChannelId;
  item.category = targetLaneCategory;
  item.duration = duration;
  item.start = minutesToTime(clampPlanningMinutes(magneticStart));
  item.slot = getSlotFromMinutes(timeToMinutes(item.start));
  dragState.timingSnapshot.commit.newStartTime = item.start;
  dragState.timingSnapshot.commit.newChannelId = item.channelId;
  dragState.timingSnapshot.commit.newLaneCategory = item.category;
  dragState.commitSnapshot.stateUpdated = true;
  dragState.commitSnapshot.newStartMinutes = timeToMinutes(item.start);
  dragState.commitSnapshot.newStartTime = item.start;
  dragState.commitSnapshot.newChannelId = item.channelId;
  dragState.commitSnapshot.newLaneCategory = item.category;
  dragState.commitSnapshot.dropCommitted = true;
  dragState.commitSnapshot.status = "DROP FIRED AND COMMITTED";
  dragState.commitSnapshot.renderVisibleTimeLabel = item.start;

  console.log("[timeline-dnd] commit after update", {
    itemId: item.id,
    newStart: item.start,
    newStartMinutes: getSafeStartMinutes(item.start),
    newChannelId: item.channelId,
    newLaneCategory: item.category,
  });

  const dropSnapshot = DEBUG_TIMELINE_DND
    ? buildTimelineDebugSnapshot("drop", event, lane, item, {
        timelineContentX,
        adjustedX: relativeX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        preSnapMinutes,
        snappedMinutes,
        clampedMinutes: magneticMinutes,
        finalMinutes: getSafeStartMinutes(item.start),
        finalTime: item.start,
        dropStatus: "DROP FIRED AND COMMITTED",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "yes",
        dropAllowed: "yes",
        dropBlockedReason: "-",
        itemId: item.id,
        targetLabel: dragState.commitSnapshot.targetLabel,
        currentTargetLabel: dragState.commitSnapshot.currentTargetLabel,
        closestLaneLabel: dragState.commitSnapshot.closestLaneLabel,
        renderNote: `source=${source}`,
        commitSnapshot: dragState.commitSnapshot,
      })
    : null;

  if (DEBUG_TIMELINE_DND && dropSnapshot) {
    updateTimelineDebugPanel(dropSnapshot);
    logTimelineDebug("drop", dropSnapshot);
  }

  clearLaneHighlights();
  persistAndRender();

  const renderedItem = state.items.find((entry) => entry.id === item.id) || null;
  const renderedLeft = renderedItem ? timelineMinutesToX(getSafeStartMinutes(renderedItem.start)) : null;
  const renderVisibleTimeLabel = renderedItem?.start || "-";
  const renderReflected = Boolean(renderedItem && renderedItem.start === item.start);
  dragState.commitSnapshot.renderReflected = renderReflected;
  dragState.commitSnapshot.renderedLeft = renderedLeft;
  dragState.commitSnapshot.renderVisibleTimeLabel = renderVisibleTimeLabel;

  const renderSnapshot = DEBUG_TIMELINE_DND
    ? buildTimelineDebugSnapshot("render", event, lane, item, {
        timelineContentX,
        adjustedX: relativeX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        preSnapMinutes,
        snappedMinutes,
        clampedMinutes: magneticMinutes,
        finalMinutes: getSafeStartMinutes(item.start),
        finalTime: item.start,
        dropStatus: "DROP FIRED AND COMMITTED",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "yes",
        dropAllowed: "yes",
        dropBlockedReason: "-",
        itemId: item.id,
        renderedLeft,
        renderNote: `saved start rendered at ${renderedLeft}px (${renderVisibleTimeLabel})`,
        commitSnapshot: dragState.commitSnapshot,
      })
    : null;

  if (DEBUG_TIMELINE_DND && renderSnapshot) {
    updateTimelineDebugPanel(renderSnapshot);
    logTimelineDebug("render", renderSnapshot);
  }

  const postCheckItem = state.items.find((entry) => entry.id === item.id) || null;
  const revertedAfterCommit = Boolean(
    !postCheckItem ||
      postCheckItem.start !== item.start ||
      postCheckItem.channelId !== item.channelId ||
      postCheckItem.category !== item.category,
  );
  const finalSavedStart = postCheckItem?.start || item.start;
  const finalSavedMinutes = postCheckItem ? getSafeStartMinutes(postCheckItem.start) : getSafeStartMinutes(item.start);
  const finalRenderedLeft = postCheckItem ? timelineMinutesToX(getSafeStartMinutes(postCheckItem.start)) : renderedLeft;
  const finalSavedMatchesCommit = Boolean(finalSavedStart === dragState.timingSnapshot.commit.newStartTime);
  const savedMatchesRecheck = Boolean(postCheckItem && postCheckItem.start === finalSavedStart);
  const hoverSnapshot = dragState.timingSnapshot.hover || null;
  const commitSnapshot = dragState.timingSnapshot.commit || null;
  const hoverMatchesCommit = Boolean(
    hoverSnapshot &&
      commitSnapshot &&
      hoverSnapshot.snappedMinutes === commitSnapshot.snappedMinutes &&
      hoverSnapshot.clampedMinutes === commitSnapshot.clampedMinutes,
  );
  dragState.timingSnapshot.final = {
    itemId: item.id,
    savedStart: finalSavedStart,
    savedMinutes: finalSavedMinutes,
    savedTimeLabel: postCheckItem?.start || item.start,
    savedChannelId: postCheckItem?.channelId || item.channelId,
    savedLaneCategory: postCheckItem?.category || item.category,
    renderedLeft: finalRenderedLeft,
    visibleTimeLabel: postCheckItem?.start || item.start,
  };
  dragState.timingSnapshot.matches = {
    hoverToCommit: hoverMatchesCommit ? "yes" : "no",
    commitToSaved: finalSavedMatchesCommit ? "yes" : "no",
    savedToRecheck: savedMatchesRecheck ? "yes" : "no",
  };
  dragState.timingSnapshot.postRender = {
    itemId: item.id,
    expectedCommittedStart: dragState.timingSnapshot.commit.newStartTime,
    postCheckStart: postCheckItem?.start || "-",
    revertedAfterCommit: revertedAfterCommit ? "yes" : "no",
  };
  dragState.timingSnapshot.revertedAfterCommit = revertedAfterCommit ? "yes" : "no";
  dragState.timingSnapshot.status = revertedAfterCommit ? "reverted" : "saved";
  console.log("[timeline-dnd] final saved snapshot", {
    itemId: item.id,
    originalStart: dragState.timingSnapshot.originalStartTime,
    hoverSnapped: hoverSnapshot?.snappedMinutes,
    hoverClamped: hoverSnapshot?.clampedMinutes,
    commitSnapped: commitSnapshot?.snappedMinutes,
    commitClamped: commitSnapshot?.clampedMinutes,
    finalSavedStart,
    finalSavedMinutes,
    revertedAfterCommit: revertedAfterCommit ? "yes" : "no",
  });
  dragState.commitSnapshot.revertedAfterCommit = revertedAfterCommit;
  dragState.commitSnapshot.stateUpdated = Boolean(postCheckItem && postCheckItem.start === item.start);
  dragState.commitSnapshot.renderReflected = Boolean(postCheckItem && postCheckItem.start === item.start);
  dragState.commitSnapshot.postCheckStart = postCheckItem?.start || "-";
  dragState.commitSnapshot.postCheckMinutes = postCheckItem ? getSafeStartMinutes(postCheckItem.start) : null;
  dragState.commitSnapshot.postCheckTargetChannelId = postCheckItem?.channelId || null;
  dragState.commitSnapshot.postCheckTargetLaneCategory = postCheckItem?.category || null;
  dragState.commitSnapshot.postCheckMatched = Boolean(postCheckItem && postCheckItem.start === item.start);
  if (revertedAfterCommit) {
    console.log("[timeline-dnd] ITEM REVERTED AFTER COMMIT", {
      itemId: item.id,
      writtenStart: item.start,
      postCheckStart: postCheckItem?.start || "-",
      writtenChannelId: item.channelId,
      postCheckChannelId: postCheckItem?.channelId || "-",
      writtenCategory: item.category,
      postCheckCategory: postCheckItem?.category || "-",
    });
  }
  console.log("[timeline-dnd] post-render recheck", {
    itemId: item.id,
    expectedCommittedStart: dragState.timingSnapshot.commit.newStartTime,
    currentStart: postCheckItem?.start || "-",
    currentStartMinutes: postCheckItem ? getSafeStartMinutes(postCheckItem.start) : null,
    currentChannelId: postCheckItem?.channelId || "-",
    currentLaneCategory: postCheckItem?.category || "-",
    renderedLeft: finalRenderedLeft,
    visibleTimeLabel: renderVisibleTimeLabel,
    revertedAfterCommit,
  });

  dragState.dropAllowed = "yes";
  dragState.dropBlockedReason = "-";
  dragState.lastDropCommitted = true;
  return { committed: true, dropSnapshot, renderSnapshot };
}

function handleTimelineDragEnterCapture(event) {
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(resolveTimelineLaneFromEvent(event));
  if (DEBUG_TIMELINE_DND) {
    logTimelineEvent("dragenter", event, resolveTimelineLaneFromEvent(event), {
      dropStatus: "DRAG ENTER",
      dropFired: "no",
      dropCommitted: "no",
      renderNote: "shell capture",
    });
  }
}

function handleTimelineDragOverCapture(event) {
  const lane = resolveTimelineLaneFromEvent(event);
  if (!lane) return;
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  dragState.dropAllowed = "yes";
  dragState.dropBlockedReason = "-";
  if (DEBUG_TIMELINE_DND) {
    logTimelineEvent("dragover", event, lane, {
      dropStatus: "DRAG OVER",
      dropFired: "no",
      dropCommitted: "no",
      dropAllowed: "yes",
      dropBlockedReason: "-",
      renderNote: "shell capture",
    });
  }
}

function handleTimelineDropCapture(event) {
  const lane = resolveTimelineLaneFromEvent(event);
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(lane);
  if (!lane) {
    dragState.dropAllowed = "no";
    dragState.dropBlockedReason = "no lane target";
    if (DEBUG_TIMELINE_DND) {
      logTimelineEvent("drop", event, null, {
        dropStatus: "DROP FIRED BUT HAD NO VALID TARGET",
        dropFired: "yes",
        dropCommitted: "no",
        dropAllowed: "no",
        dropBlockedReason: "no lane target",
        renderNote: "capture saw drop without lane",
      });
    }
    return;
  }
  commitTimelineDrop(event, lane, "shell-capture");
}

function handleTimelineDragLeaveCapture(event) {
  if (!DEBUG_TIMELINE_DND) return;
  logTimelineEvent("dragleave", event, resolveTimelineLaneFromEvent(event), {
    dropStatus: "DRAG LEAVE",
    dropFired: "no",
    dropCommitted: "no",
    renderNote: "shell capture",
  });
}

function handleTimelineDragEndCapture(event) {
  dragState.lastEventType = event.type;
  dragState.lastTargetLabel = describeTimelineNode(event.target);
  dragState.lastCurrentTargetLabel = describeTimelineNode(event.currentTarget);
  dragState.lastClosestLaneLabel = describeTimelineNode(resolveTimelineLaneFromEvent(event));
  if (DEBUG_TIMELINE_DND) {
    const snapshot = logTimelineEvent("dragend", event, resolveTimelineLaneFromEvent(event), {
      dropStatus: dragState.lastDropCommitted ? "DRAG END AFTER DROP" : "DROP NEVER FIRED",
      dropFired: dragState.lastDropCommitted ? "yes" : "no",
      dropCommitted: dragState.lastDropCommitted ? "yes" : "no",
      renderNote: "drag end observed",
    });
    if (snapshot) {
      // keep the panel readable on end state
      dragState.dropAllowed = snapshot.dropAllowed;
      dragState.dropBlockedReason = snapshot.dropBlockedReason;
    }
  }
}

function clearLaneHighlights() {
  document.querySelectorAll(".day-lane.drop-target").forEach((lane) => lane.classList.remove("drop-target"));
}

function getTimelineContentXFromPointer(event, lane) {
  const geometry = getTimelinePointerGeometry(event, lane);
  return geometry.timelineContentX;
}

function getTimelinePointerGeometry(event, lane) {
  const shell = elements.timelineShell;
  const laneRect = lane?.getBoundingClientRect?.() || null;
  const clientX = Number.isFinite(event?.clientX) ? event.clientX : dragState.lastClientX;
  const clientY = Number.isFinite(event?.clientY) ? event.clientY : null;
  const laneLeft = laneRect ? laneRect.left : null;
  const laneWidth = laneRect ? laneRect.width : null;
  const shellScrollLeft = shell?.scrollLeft ?? 0;
  const shellClientWidth = shell?.clientWidth ?? null;
  const shellScrollWidth = shell?.scrollWidth ?? null;
  const visibleLaneX = Number.isFinite(clientX) && Number.isFinite(laneLeft) ? clientX - laneLeft : null;
  const timelineContentX = Number.isFinite(visibleLaneX) ? visibleLaneX + shellScrollLeft : NaN;

  return {
    clientX,
    clientY,
    laneLeft,
    laneWidth,
    shellScrollLeft,
    shellClientWidth,
    shellScrollWidth,
    visibleLaneX,
    timelineContentX,
  };
}

function startResize(event, itemId, edge) {
  event.stopPropagation();
  event.preventDefault();
  const item = state.items.find((entry) => entry.id === itemId);
  if (!item || state.timelineScale !== "day") return;

  resizeState.itemId = itemId;
  resizeState.edge = edge;
  resizeState.startX = event.clientX;
  resizeState.startMinutes = getSafeStartMinutes(item.start);
  resizeState.duration = getSafeDuration(item.duration, item.itemType);
}

function handleResizeMove(event) {
  if (!resizeState.itemId) return;

  const item = state.items.find((entry) => entry.id === resizeState.itemId);
  if (!item) return;

  const deltaMinutes = snapMinutes(timelinePixelsToMinutes(event.clientX - resizeState.startX));
  const startMinutes = clampPlanningMinutes(resizeState.startMinutes);
  const duration = getSafeDuration(resizeState.duration, item.itemType);
  if (resizeState.edge === "right") {
    const nextDuration = Math.max(5, duration + deltaMinutes);
    item.duration = Math.min(nextDuration, DAY_END - startMinutes);
  } else {
    const latestStart = startMinutes + duration - 5;
    const nextStart = Math.max(DAY_START, Math.min(latestStart, startMinutes + deltaMinutes));
    const consumed = nextStart - startMinutes;
    item.start = minutesToTime(nextStart);
    item.duration = Math.max(5, duration - consumed);
  }

  item.duration = snapDurationForItem(item, applyMagneticDuration(item, item.duration));
  item.slot = getSlotFromMinutes(timeToMinutes(item.start));
  render();
}

function stopResize() {
  if (!resizeState.itemId) return;
  resizeState.itemId = null;
  resizeState.edge = null;
  persistAndRender();
}

function syncControls() {
  refreshChannelSelects();
  elements.viewMode.value = state.viewMode;
  elements.timelineScale.value = state.timelineScale;
  elements.colorMode.value = state.colorMode;
  elements.exportProfile.value = state.exportProfile;
  elements.dayFilter.value = state.selectedDay;
  elements.channelFilter.value = state.selectedChannelId;
  if (elements.themeMode) elements.themeMode.value = normalizeTheme(state.theme);
  renderWorkspaceChrome();
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function exportSchedule(format) {
  const rows = state.items.map((item) => ({
    title: item.title,
    itemType: item.itemType,
    category: item.category,
    channel: getChannelName(item.channelId),
    channelGroup: state.channels.find((channel) => channel.id === item.channelId)?.group || "",
    day: item.day,
    start: item.start,
    end: minutesToTime(getSafeStartMinutes(item.start) + getSafeDuration(item.duration, item.itemType)),
    duration: getSafeDuration(item.duration, item.itemType),
    slot: item.slot,
    importance: item.importance,
    watershedRestricted: item.watershedRestricted ? "Yes" : "No",
    primeTime: item.primeTime ? "Yes" : "No",
    mustRun: item.mustRun ? "Yes" : "No",
    blockGroup: item.blockGroup,
    assetCode: item.assetCode,
    notes: item.notes,
  }));

  if (format === "json") {
    if (state.exportProfile === "internal") {
      exportInternalSchedulerJson();
      return;
    }

    const validation = validateFs42NativeExport();
    if (!validation.ready) {
      renderExportReadiness();
      const message = validation.blockers.slice(0, 4).map(renderIssueLine).join("\n");
      window.alert(`FS42 station config export is blocked:\n${message}`);
      return;
    }

    exportFs42NativeStationConfigs();
    return;
  }

  const headers = Object.keys(rows[0] || {
    title: "",
    itemType: "",
    category: "",
    channel: "",
    channelGroup: "",
    day: "",
    start: "",
    end: "",
    duration: "",
    slot: "",
    importance: "",
    watershedRestricted: "",
    primeTime: "",
    mustRun: "",
    blockGroup: "",
    assetCode: "",
    notes: "",
  });
  const csv = [
    headers.join(","),
    ...rows.map((row) => headers.map((header) => escapeCsv(row[header] ?? "")).join(",")),
  ].join("\n");
  downloadFile("fs42-schedule.csv", csv, "text/csv");
}

function exportInternalSchedulerJson() {
  const validation = validateSchedule("internal");
  if (!validation.ready) {
    renderExportReadiness();
    const message = validation.blockers.slice(0, 4).map(renderIssueLine).join("\n");
    window.alert(`Internal scheduler export is blocked:\n${message}`);
    return;
  }
  const payload = {
    export_profile: "internal",
    generated_at: new Date().toISOString(),
    ready: validation.ready,
    summary: {
      blockers: validation.blockers.length,
      warnings: validation.warnings.length,
    },
    channels: state.channels,
    items: state.items,
    view: {
      workspace: state.workspace,
      viewMode: state.viewMode,
      timelineScale: state.timelineScale,
      selectedDay: state.selectedDay,
      selectedChannelId: state.selectedChannelId,
    },
  };
  downloadFile("fs42-scheduler-internal.json", JSON.stringify(payload, null, 2), "application/json");
}

function exportFs42NativeStationConfigs() {
  state.channels.forEach((channel, index) => {
    const channelItems = state.items.filter((item) => item.channelId === channel.id);
    const channelNumber = parsePositiveInteger(channel.channelNumber) || index + 1;
    const payload = { station_conf: buildFs42NativeStationConfig(channel, channelItems, channelNumber) };
    const baseName = getFs42NativeFileBase(channel, channelNumber);

    window.setTimeout(() => {
      downloadFile(`${baseName}.json`, JSON.stringify(payload, null, 2), "application/json");
    }, index * 120);
  });
}

// Authoritative FS42 station_conf builder for native exports.
function buildFs42NativeStationConfig(channel, channelItems, channelNumber) {
  const programmeItems = channelItems
    .filter((item) => !ITEM_TYPE_META[item.itemType]?.commercial)
    .sort(compareItems);
  const daySchedules = buildFs42NativeDaySchedules(programmeItems);
  const baseName = getFs42NativeFileBase(channel, channelNumber);
  const resolved = resolveFs42NativeChannelConfig(channel, baseName);

  return {
    network_name: channel.name,
    channel_number: resolved.channelNumber,
    network_type: resolved.networkType,
    content_dir: resolved.contentDir,
    commercial_dir: resolved.commercialDir,
    bump_dir: resolved.bumpDir,
    schedule_increment: resolved.scheduleIncrement,
    break_strategy: resolved.breakStrategy,
    commercial_free: resolved.commercialFree,
    break_duration: resolved.breakDuration,
    sign_off_video: resolved.signOffVideo,
    off_air_video: resolved.offAirVideo,
    standby_image: resolved.standbyImage,
    be_right_back_media: resolved.beRightBackMedia,
    logo_dir: resolved.logoDir,
    show_logo: resolved.showLogo,
    default_logo: resolved.defaultLogo,
    logo_permanent: resolved.logoPermanent,
    multi_logo: resolved.multiLogo,
    clip_shows: resolved.clipShows,
    ...daySchedules,
  };
}

function resolveFs42NativeChannelConfig(channel, baseName) {
  const channelNumber = parsePositiveInteger(channel.channelNumber) || 1;
  return {
    channelNumber,
    networkType: NETWORK_TYPES.includes(channel.networkType) ? channel.networkType : "standard",
    contentDir: String(channel.contentDir || "").trim() || `catalog/${baseName}`,
    commercialDir: String(channel.commercialDir || "").trim() || `commercial/${baseName}`,
    bumpDir: String(channel.bumpDir || "").trim() || `bump/${baseName}`,
    scheduleIncrement: parsePositiveInteger(channel.scheduleIncrement) || 30,
    breakStrategy: BREAK_STRATEGIES.includes(channel.breakStrategy) ? channel.breakStrategy : "standard",
    commercialFree: typeof channel.commercialFree === "boolean" ? channel.commercialFree : false,
    breakDuration: parsePositiveInteger(channel.breakDuration) || 120,
    signOffVideo: String(channel.signOffVideo || "").trim() || "runtime/signoff.mp4",
    offAirVideo: String(channel.offAirVideo || "").trim() || "runtime/off_air_pattern.mp4",
    standbyImage: String(channel.standbyImage || "").trim() || "runtime/standby.png",
    beRightBackMedia: String(channel.beRightBackMedia || "").trim() || "runtime/brb.png",
    logoDir: String(channel.logoDir || "").trim() || `logos/${baseName}`,
    showLogo: typeof channel.showLogo === "boolean" ? channel.showLogo : true,
    defaultLogo: String(channel.defaultLogo || "").trim() || `${baseName}.png`,
    logoPermanent: typeof channel.logoPermanent === "boolean" ? channel.logoPermanent : true,
    multiLogo: channel.multiLogoMode ? String(channel.multiLogoProfile || "").trim() : "",
    clipShows: normalizeClipShowList(channel.clipShows),
  };
}

function buildFs42NativeDaySchedules(items) {
  const schedule = {};
  DAY_NAMES.forEach((day) => {
    const dayKey = day.toLowerCase();
    const dayItems = items.filter((item) => item.day === day).sort(compareItems);
    const daySchedule = {};

    const hourGroups = new Map();
    dayItems.forEach((item) => {
      const startMinutes = parseTimeMinutes(item.start);
      if (!Number.isFinite(startMinutes)) return;
      const hour = String(Math.floor(startMinutes / 60));
      if (!hourGroups.has(hour)) hourGroups.set(hour, []);
      hourGroups.get(hour).push(item);
    });

    Array.from(hourGroups.entries())
      .sort(([leftHour], [rightHour]) => Number(leftHour) - Number(rightHour))
      .forEach(([hour, hourItems]) => {
        const slot = buildFs42NativeHourSlot(hourItems);
        if (slot) daySchedule[hour] = slot;
      });

    schedule[dayKey] = daySchedule;
  });

  return schedule;
}

function buildFs42NativeHourSlot(hourItems) {
  const programmeItems = hourItems.filter((item) => !ITEM_TYPE_META[item.itemType]?.commercial);
  if (programmeItems.length === 0) return null;

  const chosenItem = programmeItems[0];
  return {
    tags: buildFs42NativeTags(chosenItem, hourItems),
    loop: true,
  };
}

function buildFs42NativeTags(item, hourItems = []) {
  const baseTag = slugify(item.blockGroup || item.category || item.title || "programming");
  const hasWatershed = hourItems.some((entry) => entry.watershedRestricted);
  return hasWatershed ? `watershed/${baseTag}` : baseTag;
}

function getFs42NativeFileBase(channel, channelNumber) {
  return `ch${String(channelNumber).padStart(2, "0")}_${slugify(channel.name || "channel")}`;
}

function validateFs42NativePayload(payload, channel, channelNumber) {
  const issues = [];
  const stationConf = payload?.station_conf;
  const baseTitle = channel?.name || `Channel ${channelNumber}`;
  const validNetworkTypes = new Set(["standard", "web", "guide", "loop", "streaming"]);
  const validBreakStrategies = new Set(["standard", "end", "center"]);
  const requiredStringFields = ["network_name", "content_dir", "commercial_dir", "bump_dir"];
  const optionalStringFields = ["sign_off_video", "off_air_video", "standby_image", "be_right_back_media", "logo_dir", "default_logo"];

  if (!stationConf || typeof stationConf !== "object" || Array.isArray(stationConf)) {
    return [{ title: baseTitle, message: "station_conf is missing or invalid." }];
  }

  if (!Number.isInteger(stationConf.channel_number) || stationConf.channel_number <= 0) {
    issues.push({ title: baseTitle, message: "channel_number must be a positive integer." });
  }
  if (!validNetworkTypes.has(stationConf.network_type)) {
    issues.push({ title: baseTitle, message: "network_type must be standard, web, guide, loop, or streaming." });
  }
  if (!validBreakStrategies.has(stationConf.break_strategy)) {
    issues.push({ title: baseTitle, message: "break_strategy must be standard, end, or center." });
  }
  if (typeof stationConf.commercial_free !== "boolean") {
    issues.push({ title: baseTitle, message: "commercial_free must be a boolean." });
  }
  if (typeof stationConf.schedule_increment !== "number" || !Number.isInteger(stationConf.schedule_increment) || stationConf.schedule_increment <= 0) {
    issues.push({ title: baseTitle, message: "schedule_increment must be a positive integer." });
  }
  if (!Number.isInteger(stationConf.break_duration) || stationConf.break_duration <= 0) {
    issues.push({ title: baseTitle, message: "break_duration must be a positive integer." });
  }

  requiredStringFields.forEach((field) => {
    if (typeof stationConf[field] !== "string" || !stationConf[field].trim()) {
      issues.push({ title: baseTitle, message: `${field} is missing or empty.` });
    }
  });
  optionalStringFields.forEach((field) => {
    if (field in stationConf && (typeof stationConf[field] !== "string" || !stationConf[field].trim())) {
      issues.push({ title: baseTitle, message: `${field} must be a non-empty string when provided.` });
    }
  });
  ["show_logo", "logo_permanent", "multi_logo"].forEach((field) => {
    if (field === "multi_logo") {
      if (field in stationConf && typeof stationConf[field] !== "string") {
        issues.push({ title: baseTitle, message: `${field} must be a string when provided.` });
      }
      return;
    }
    if (field in stationConf && typeof stationConf[field] !== "boolean") {
      issues.push({ title: baseTitle, message: `${field} must be boolean when provided.` });
    }
  });
  if ("clip_shows" in stationConf) {
    if (!Array.isArray(stationConf.clip_shows)) {
      issues.push({ title: baseTitle, message: "clip_shows must be an array when provided." });
    } else if (stationConf.clip_shows.some((entry) => typeof entry !== "string" || !entry.trim())) {
      issues.push({ title: baseTitle, message: "clip_shows may only contain non-empty strings." });
    }
  }

  if (stationConf.break_strategy === "none") {
    issues.push({ title: baseTitle, message: "break_strategy cannot be none." });
  }

  DAY_NAMES.forEach((day) => {
    const dayKey = day.toLowerCase();
    const dayValue = stationConf[dayKey];
    if (typeof dayValue === "string") return;
    if (!dayValue || typeof dayValue !== "object" || Array.isArray(dayValue)) {
      issues.push({ title: baseTitle, message: `${dayKey} must be a day template string or an hour object.` });
      return;
    }
    Object.entries(dayValue).forEach(([hour, slot]) => {
      if (!/^(?:[0-9]|1[0-9]|2[0-3])$/.test(hour)) {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour} is not a valid hour key.` });
      }
      if (!slot || typeof slot !== "object" || Array.isArray(slot)) {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour} must be a slot object.` });
        return;
      }
      if (typeof slot.tags !== "string" && !Array.isArray(slot.tags)) {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour}.tags must be a string or string array.` });
      }
      if ("loop" in slot && typeof slot.loop !== "boolean") {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour}.loop must be boolean if present.` });
      }
      if ("sequence" in slot && typeof slot.sequence !== "string") {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour}.sequence must be a string if present.` });
      }
      if ("sequence_start" in slot && !Number.isInteger(slot.sequence_start)) {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour}.sequence_start must be an integer if present.` });
      }
      if ("sequence_end" in slot && !Number.isInteger(slot.sequence_end)) {
        issues.push({ title: baseTitle, message: `${dayKey}.${hour}.sequence_end must be an integer if present.` });
      }
    });
  });

  return issues;
}

function downloadFile(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeCsv(value) {
  const text = String(value).replace(/"/g, '""');
  return `"${text}"`;
}

function slugify(value) {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function uniqueValues(values) {
  return Array.from(new Set(values.filter(Boolean)));
}

function persistAndRender() {
  persistState();
  render();
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY) || LEGACY_KEYS.map((key) => localStorage.getItem(key)).find(Boolean);
  if (!raw) return createDefaultState();

  try {
    return normalizeState(JSON.parse(raw));
  } catch (error) {
    return createDefaultState();
  }
}

function normalizeState(parsed) {
  const fallback = createDefaultState();

  if (Array.isArray(parsed.channels) && parsed.channels.length > 0) {
    return {
      workspace: normalizeWorkspace(parsed.workspace),
      viewMode: parsed.viewMode || "timeline",
      timelineScale: parsed.timelineScale || "day",
      colorMode: parsed.colorMode || "channel",
      exportProfile: normalizeExportProfile(parsed.exportProfile),
      theme: normalizeTheme(parsed.theme),
      selectedDay: parsed.selectedDay || "Monday",
      selectedChannelId: parsed.selectedChannelId || "all",
      sectionOpenState: normalizeSectionOpenState(parsed.sectionOpenState || parsed.collapsedSections),
      channels: normalizeChannels(parsed.channels),
      items: seedStrategicOrders(normalizeItems(parsed.items || parsed.shows || [], parsed.channels)),
    };
  }

  if (Array.isArray(parsed.shows) && parsed.shows.length > 0) {
    const channels = normalizeChannels(structuredClone(DEFAULT_CHANNELS));
    const map = new Map(channels.map((channel) => [channel.name, channel.id]));
    return {
      workspace: normalizeWorkspace(parsed.workspace),
      viewMode: "timeline",
      timelineScale: "day",
      colorMode: parsed.colorMode || "channel",
      exportProfile: normalizeExportProfile(parsed.exportProfile),
      theme: normalizeTheme(parsed.theme),
      selectedDay: parsed.selectedDay || "Monday",
      selectedChannelId: "all",
      sectionOpenState: normalizeSectionOpenState(parsed.sectionOpenState || parsed.collapsedSections),
      channels,
      items: seedStrategicOrders(parsed.shows.map((show) => ({
        id: show.id || crypto.randomUUID(),
        title: show.title || "Untitled item",
        itemType: show.itemType || "Programme",
        category: show.category || ITEM_TYPE_META[show.itemType || "Programme"]?.defaultCategory || "Entertainment",
        channelId: map.get(show.channel) || channels[0].id,
        day: show.day || "Monday",
        start: Number.isFinite(parseTimeMinutes(show.start)) ? show.start : "20:00",
        duration: getSafeDuration(show.duration, show.itemType || "Programme"),
        importance: show.importance || "Medium",
        watershedRestricted: Boolean(show.watershedRestricted),
        primeTime: Boolean(show.primeTime),
        mustRun: Boolean(show.mustRun),
        slot: show.slot || getSlotFromMinutes(Number.isFinite(parseTimeMinutes(show.start)) ? parseTimeMinutes(show.start) : timeToMinutes("20:00")),
        blockGroup: show.blockGroup || "",
        assetCode: show.assetCode || "",
        weekOrder: Number.isFinite(show.weekOrder) ? show.weekOrder : null,
        monthOrder: Number.isFinite(show.monthOrder) ? show.monthOrder : null,
        planningWeek: Number.isInteger(show.planningWeek) ? show.planningWeek : null,
        notes: show.notes || "",
      }))),
    };
  }

  return fallback;
}

function normalizeWorkspace(workspace) {
  return Object.prototype.hasOwnProperty.call(WORKSPACE_META, workspace) ? workspace : "schedule";
}

function normalizeTheme(theme) {
  return theme === "light" ? "light" : "dark";
}

function generateDefaultExportCode(title) {
  const normalized = String(title || "")
    .trim()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");
  return normalized;
}

function getUniqueExportCode(baseCode, items, currentItemId = null) {
  const normalizedBase = String(baseCode || "").trim();
  if (!normalizedBase) return "";

  const usedCodes = new Set(
    (items || [])
      .filter((item) => item.id !== currentItemId)
      .map((item) => String(item.assetCode || "").trim().toUpperCase())
      .filter(Boolean),
  );

  let candidate = normalizedBase;
  let suffix = 2;
  while (usedCodes.has(candidate.toUpperCase())) {
    candidate = `${normalizedBase}-${suffix}`;
    suffix += 1;
  }

  return candidate;
}

function normalizeItems(items, channels) {
  const nameMap = new Map((channels || []).map((channel) => [channel.name, channel.id]));
  return items.map((item) => ({
    id: item.id || crypto.randomUUID(),
    title: item.title || "Untitled item",
    itemType: item.itemType || "Programme",
    category: item.category || ITEM_TYPE_META[item.itemType || "Programme"]?.defaultCategory || "Entertainment",
    channelId: item.channelId || nameMap.get(item.channel) || channels?.[0]?.id,
    day: item.day || "Monday",
    start: Number.isFinite(parseTimeMinutes(item.start)) ? item.start : "20:00",
    duration: getSafeDuration(item.duration, item.itemType),
    importance: item.importance || "Medium",
    watershedRestricted: Boolean(item.watershedRestricted),
    primeTime: Boolean(item.primeTime),
    mustRun: Boolean(item.mustRun),
    slot: item.slot || getSlotFromMinutes(Number.isFinite(parseTimeMinutes(item.start)) ? parseTimeMinutes(item.start) : timeToMinutes("20:00")),
    blockGroup: item.blockGroup || "",
    assetCode: item.assetCode || "",
    weekOrder: Number.isFinite(item.weekOrder) ? item.weekOrder : null,
    monthOrder: Number.isFinite(item.monthOrder) ? item.monthOrder : null,
    planningWeek: Number.isInteger(item.planningWeek) ? item.planningWeek : null,
    notes: item.notes || "",
  }));
}

function createOption(value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  return option;
}

function timeToMinutes(value) {
  const [hours, minutes] = value.split(":").map(Number);
  return hours * 60 + minutes;
}

function minutesToTime(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function getSlotFromMinutes(totalMinutes) {
  if (totalMinutes < 600) return "Breakfast";
  if (totalMinutes < 1080) return "Daytime";
  if (totalMinutes < 1320) return "Prime Time";
  return "Late";
}

function snapMinutes(minutes) {
  return Math.round(minutes / SNAP_MINUTES) * SNAP_MINUTES;
}

function snapStartMinutes(item, proposedStart) {
  const snapped = snapMinutes(proposedStart);
  if (ITEM_TYPE_META[item.itemType]?.commercial) {
    return applyMagneticTargets(item, item.channelId, item.category, snapped, 5, 5);
  }

  return applyMagneticTargets(item, item.channelId, item.category, snapped, 12, 30);
}

function snapDurationForItem(item, duration) {
  const step = ITEM_TYPE_META[item.itemType]?.commercial ? 5 : 30;
  const snapped = Math.max(5, Math.round(duration / step) * step);
  return ITEM_TYPE_META[item.itemType]?.commercial ? snapped : Math.max(30, snapped);
}

function applyMagneticTargets(item, channelId, category, proposedStart, threshold, preferenceStep) {
  if (!Number.isFinite(proposedStart)) {
    return DAY_START;
  }
  const targets = new Set();
  const laneItems = state.items.filter(
    (entry) =>
      entry.id !== item.id &&
      entry.channelId === channelId &&
      entry.category === category &&
      entry.day === item.day,
  );

  for (let minute = DAY_START; minute <= DAY_END; minute += 30) {
    targets.add(minute);
  }

  laneItems.forEach((entry) => {
    const start = timeToMinutes(entry.start);
    targets.add(start);
    targets.add(start + getSafeDuration(entry.duration, entry.itemType));
  });

  let closest = proposedStart;
  let closestDistance = Infinity;
  let foundTarget = false;
  targets.forEach((target) => {
    const distance = Math.abs(target - proposedStart);
    if (distance <= threshold && distance < closestDistance) {
      closest = target;
      closestDistance = distance;
      foundTarget = true;
    }
  });

  const snapped = foundTarget ? closest : Math.round(proposedStart / preferenceStep) * preferenceStep;
  return Math.max(DAY_START, Math.min(DAY_END, snapped));
}

function applyMagneticDuration(item, duration) {
  if (ITEM_TYPE_META[item.itemType]?.commercial) return duration;
  const targets = [30, 60, 90, 120, 180];
  let closest = duration;
  let closestDistance = Infinity;
  targets.forEach((target) => {
    const distance = Math.abs(target - duration);
    if (distance <= 5 && distance < closestDistance) {
      closest = target;
      closestDistance = distance;
    }
  });
  return closest;
}
