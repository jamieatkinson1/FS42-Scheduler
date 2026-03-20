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
const DAY_START = 6 * 60;
const DAY_END = 24 * 60;
const SNAP_MINUTES = 5;
const MAGNETIC_THRESHOLD = 7;
const HOUR_WIDTH = getTimelineHourWidth();
const DEBUG_TIMELINE_DND = true;

const DEFAULT_CHANNELS = [
  createChannel("FS42 Main", "Entertainment", "#ff8a5b", "Flagship entertainment feed"),
  createChannel("FS42 Stories", "Factual", "#67d4ff", "Docs and factual blocks"),
  createChannel("FS42 Action", "Sport", "#9bff9f", "Sports and event windows"),
  createChannel("FS42 Extra", "Extensions", "#ffd86b", "Catch-up and themed nights"),
];

const elements = {
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
  channelColor: document.getElementById("channelColor"),
  resetChannelForm: document.getElementById("resetChannelForm"),
  channelList: document.getElementById("channelList"),
  legend: document.getElementById("legend"),
  legendModeLabel: document.getElementById("legendModeLabel"),
  exportReadiness: document.getElementById("exportReadiness"),
  timelineHeading: document.getElementById("timelineHeading"),
  timelineSubheading: document.getElementById("timelineSubheading"),
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
  exportCsv: document.getElementById("exportCsv"),
  exportJson: document.getElementById("exportJson"),
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
  resetItemForm();
  resetChannelForm();
  render();
}

function createChannel(name, group, color, tagline) {
  return {
    id: crypto.randomUUID(),
    name,
    group,
    color,
    tagline,
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
    ...normalizeFlags(flags),
    notes,
  };
}

function createDefaultState() {
  const channels = structuredClone(DEFAULT_CHANNELS);
  return {
    viewMode: "timeline",
    timelineScale: "day",
    colorMode: "channel",
    exportProfile: "fs42-strict",
    selectedDay: "Monday",
    selectedChannelId: "all",
    channels,
    items: [
      createItem("FS42 Breakfast Live", "Programme", "Entertainment", channels[0].id, "Monday", "06:00", 180, "High", "Morning studio", "FS42-MAIN-001", {}, "Live breakfast block."),
      createItem("Breakfast Break A", "Ad Break", "Commercials", channels[0].id, "Monday", "07:28", 6, "Medium", "Breakfast ads", "FS42-ADS-101", {}, "Contains 4 spots and bumper."),
      createItem("Coffee Sponsor Bumper", "Bumper", "Continuity", channels[0].id, "Monday", "07:34", 1, "Medium", "Breakfast ads", "FS42-BUMP-007", {}, "Lead back into programme."),
      createItem("The Ad Vault", "Programme", "Lifestyle", channels[1].id, "Monday", "09:00", 60, "Medium", "Archive docs", "FS42-STO-011", {}, "Advert history strand."),
      createItem("Midday Promo Wheel", "Promo", "Continuity", channels[1].id, "Monday", "12:00", 3, "Low", "Promo junction", "FS42-PRO-031", {}, "Cross-channel promotion."),
      createItem("Matchday Reload", "Programme", "Events", channels[2].id, "Monday", "18:00", 120, "Critical", "Sport", "FS42-ACT-021", { primeTime: true, mustRun: true }, "Highest sponsorship priority."),
      createItem("Prime Time Break", "Ad Break", "Commercials", channels[2].id, "Monday", "18:55", 5, "High", "Prime ads", "FS42-ADS-202", { primeTime: true }, "Mid-match sponsored break."),
      createItem("Late Signal", "Programme", "Drama", channels[3].id, "Monday", "22:00", 90, "Low", "Overnight", "FS42-EXT-003", { watershedRestricted: true }, "Experimental and repeat-friendly slot."),
      createItem("Spotlight UK", "Programme", "Entertainment", channels[0].id, "Tuesday", "20:00", 60, "Critical", "Feature", "FS42-MAIN-010", { primeTime: true, mustRun: true }, "Signature weekly tentpole."),
    ],
  };
}

function populateStaticSelects() {
  DAY_NAMES.forEach((day) => {
    elements.dayFilter.appendChild(createOption(day, day));
    elements.day.appendChild(createOption(day, day));
  });

  CATEGORY_OPTIONS.forEach((category) => {
    elements.category.appendChild(createOption(category, category));
  });
}

function bindEvents() {
  elements.viewMode.addEventListener("change", handleControlChange);
  elements.timelineScale.addEventListener("change", handleControlChange);
  elements.colorMode.addEventListener("change", handleControlChange);
  elements.exportProfile.addEventListener("change", handleControlChange);
  elements.dayFilter.addEventListener("change", handleControlChange);
  elements.channelFilter.addEventListener("change", handleControlChange);
  elements.itemType.addEventListener("change", handleItemTypeChange);
  const markItemFormDirty = () => {
    itemFormDirty = true;
    updateItemValidationSummary();
  };
  elements.showForm.addEventListener("input", markItemFormDirty);
  elements.showForm.addEventListener("change", markItemFormDirty);

  elements.showForm.addEventListener("submit", handleItemSubmit);
  elements.resetForm.addEventListener("click", resetItemForm);
  elements.duplicateItem.addEventListener("click", () => duplicateActiveItem());
  elements.channelForm.addEventListener("submit", handleChannelSubmit);
  elements.resetChannelForm.addEventListener("click", resetChannelForm);
  elements.seedButton.addEventListener("click", () => {
    state = createDefaultState();
    syncControls();
    resetItemForm();
    resetChannelForm();
    persistAndRender();
  });
  elements.exportCsv.addEventListener("click", () => exportSchedule("csv"));
  elements.exportJson.addEventListener("click", () => exportSchedule("json"));

  if (elements.timelineShell) {
    elements.timelineShell.addEventListener("dragenter", handleTimelineDragEnterCapture, true);
    elements.timelineShell.addEventListener("dragover", handleTimelineDragOverCapture, true);
    elements.timelineShell.addEventListener("drop", handleTimelineDropCapture, true);
    elements.timelineShell.addEventListener("dragleave", handleTimelineDragLeaveCapture, true);
  }
  document.addEventListener("dragend", handleTimelineDragEndCapture, true);

  window.addEventListener("mousemove", handleResizeMove);
  window.addEventListener("mouseup", stopResize);
  window.addEventListener("keydown", handleKeyboardShortcuts);
}

function handleControlChange(event) {
  if (event.target === elements.viewMode) state.viewMode = event.target.value;
  if (event.target === elements.timelineScale) state.timelineScale = event.target.value;
  if (event.target === elements.colorMode) state.colorMode = event.target.value;
  if (event.target === elements.exportProfile) state.exportProfile = event.target.value;
  if (event.target === elements.dayFilter) state.selectedDay = event.target.value;
  if (event.target === elements.channelFilter) state.selectedChannelId = event.target.value;
  syncControls();
  persistAndRender();
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

function handleItemSubmit(event) {
  event.preventDefault();
  const item = {
    id: elements.showId.value || crypto.randomUUID(),
    title: elements.title.value.trim(),
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
    assetCode: elements.assetCode.value.trim(),
    notes: elements.notes.value.trim(),
  };

  const validation = validateItem(item, state.items, state.exportProfile);
  const blockers = validation.filter((entry) => entry.severity === "blocker");
  if (blockers.length > 0) {
    updateItemValidationSummary(item, validation);
    return;
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
  const channel = {
    id: elements.channelId.value || crypto.randomUUID(),
    name: elements.channelName.value.trim(),
    group: elements.channelGroup.value.trim(),
    color: elements.channelColor.value,
    tagline: elements.channelTagline.value.trim(),
  };

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
  elements.formHint.textContent = "Schedule programming, commercials, and bumpers";
  updateItemValidationSummary();
}

function resetChannelForm() {
  elements.channelForm.reset();
  elements.channelId.value = "";
  elements.channelColor.value = state.channels[0]?.color || "#ff8a5b";
}

function render() {
  refreshChannelSelects();
  renderLegend();
  renderPlanner();
  renderTable();
  renderShowList();
  renderChannelList();
  renderInsights();
  renderStats();
  syncExportControls();
  updateItemValidationSummary();
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

function renderDayTimeline() {
  const visibleChannels = getVisibleChannels();
  const visibleItems = getVisibleItems();
  const conflictIds = new Set(getConflicts(visibleItems));
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
      const laneItems = visibleItems
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
      lane.addEventListener("dragover", handleLaneDragOver);
      lane.addEventListener("dragleave", handleLaneDragLeave);
      lane.addEventListener("drop", handleLaneDrop);
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
        const cellItems = getStrategicItems(channel.id, category, scaleMode, index);

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
            card.style.background = getItemColor(item);
            card.innerHTML = `<strong>${item.title}</strong><small>${item.day} | ${item.start} | ${item.itemType}</small>`;
            card.addEventListener("click", () => hydrateItemForm(item.id));
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
  const width = Math.max((duration / 60) * HOUR_WIDTH, ITEM_TYPE_META[item.itemType]?.commercial ? 42 : 48);
  block.className = "schedule-block";
  block.draggable = true;
  block.dataset.itemId = item.id;
  if (elements.showId.value === item.id) block.classList.add("is-active");
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
  block.addEventListener("dragstart", handleBlockDragStart);
  block.addEventListener("dragend", handleBlockDragEnd);
  return block;
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
    row.innerHTML = `
      <div class="channel-card-main">
        <span class="channel-swatch" style="background:${channel.color}"></span>
        <div>
          <strong>${channel.name}</strong>
          <p>${channel.group || "Ungrouped"} | ${channel.tagline || "No tagline set"}</p>
        </div>
      </div>
    `;

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
    state.exportProfile === "internal" ? "Export Internal JSON" : "Export Channel JSON";
  renderExportReadiness();
}

function renderExportReadiness() {
  const validation = validateSchedule(state.exportProfile);
  const readiness = elements.exportReadiness;
  const label = state.exportProfile === "internal" ? "Internal scheduler JSON" : "FS42 JSON (strict)";
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
  const duplicate = {
    ...structuredClone(source),
    id: crypto.randomUUID(),
    title: source.title.endsWith(" Copy") ? `${source.title} 2` : `${source.title} Copy`,
    duration: sourceDuration,
    start: minutesToTime(Math.max(DAY_START, startMinutes)),
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

function hydrateChannelForm(channelId) {
  const channel = state.channels.find((entry) => entry.id === channelId);
  if (!channel) return;

  elements.channelId.value = channel.id;
  elements.channelName.value = channel.name;
  elements.channelGroup.value = channel.group || "";
  elements.channelTagline.value = channel.tagline || "";
  elements.channelColor.value = channel.color;
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

function getStrategicItems(channelId, category, scaleMode, index) {
  const laneItems = state.items
    .filter((item) => item.channelId === channelId && item.category === category)
    .sort(compareItems);

  if (scaleMode === "week") {
    return laneItems.filter((item) => DAY_NAMES.indexOf(item.day) === index);
  }

  return laneItems.filter((item) => index < 4);
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
  const strict = profile === "fs42-strict";
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
  if (strict && exportedItem && !assetCode) {
    issues.push({ severity: "blocker", code: "assetCode", message: "Export code is required for FS42 strict export." });
  } else if (!assetCode) {
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

  state.items.forEach((item) => {
    const issues = validateItem(item, state.items, profile);
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

  if (!lane) {
    dragState.dropAllowed = "no";
    dragState.dropBlockedReason = "no lane target";
    if (DEBUG_TIMELINE_DND) {
      const blockedSnapshot = buildTimelineDebugSnapshot("drop", event, lane, item, {
        timelineContentX: timelineX,
        adjustedX:
          Number.isFinite(timelineX) ? timelineX - dragState.pointerOffsetX : null,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        dropStatus: "DROP FIRED BUT HAD NO VALID TARGET",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "no",
        dropAllowed: "no",
        dropBlockedReason: "no lane target",
        renderNote: `source=${source}`,
      });
      updateTimelineDebugPanel(blockedSnapshot);
      logTimelineDebug("drop", blockedSnapshot);
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
    if (DEBUG_TIMELINE_DND) {
      const blockedSnapshot = buildTimelineDebugSnapshot("drop", event, lane, null, {
        timelineContentX: timelineX,
        adjustedX:
          Number.isFinite(timelineX) ? timelineX - dragState.pointerOffsetX : null,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        dropStatus: "DROP FIRED BUT DRAG STATE WAS EMPTY",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "no",
        dropAllowed: "yes",
        dropBlockedReason: "drag state empty",
        renderNote: `source=${source}`,
      });
      updateTimelineDebugPanel(blockedSnapshot);
      logTimelineDebug("drop", blockedSnapshot);
    }
    return { committed: false, dropSnapshot: null, renderSnapshot: null };
  }

  const duration = getSafeDuration(item.duration, item.itemType);
  const relativeX = timelineX - dragState.pointerOffsetX;
  const preSnapMinutes = timelineXToMinutes(relativeX);
  const snappedStart = snapStartMinutes(item, preSnapMinutes);
  const maxStart = DAY_END - duration;
  const clampedMinutes = clampPlanningMinutes(
    Math.max(DAY_START, Math.min(maxStart, Number.isFinite(snappedStart) ? snappedStart : DAY_START)),
  );
  const magneticStart = applyMagneticTargets(
    item,
    lane.dataset.channelId,
    lane.dataset.category,
    clampedMinutes,
  );

  item.channelId = lane.dataset.channelId;
  item.category = lane.dataset.category;
  item.duration = duration;
  item.start = minutesToTime(clampPlanningMinutes(magneticStart));
  item.slot = getSlotFromMinutes(timeToMinutes(item.start));

  const dropSnapshot = DEBUG_TIMELINE_DND
    ? buildTimelineDebugSnapshot("drop", event, lane, item, {
        timelineContentX: timelineX,
        adjustedX: relativeX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        preSnapMinutes,
        snappedMinutes: snappedStart,
        clampedMinutes,
        finalMinutes: getSafeStartMinutes(item.start),
        finalTime: item.start,
        dropStatus: "DROP FIRED AND COMMITTED",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "yes",
        dropAllowed: "yes",
        dropBlockedReason: "-",
        itemId: item.id,
        renderNote: `source=${source}`,
      })
    : null;

  if (DEBUG_TIMELINE_DND && dropSnapshot) {
    updateTimelineDebugPanel(dropSnapshot);
    logTimelineDebug("drop", dropSnapshot);
  }

  clearLaneHighlights();
  persistAndRender();

  const renderSnapshot = DEBUG_TIMELINE_DND
    ? buildTimelineDebugSnapshot("render", event, lane, item, {
        timelineContentX: timelineX,
        adjustedX: relativeX,
        fallbackX: fallbackTimelineX,
        xSource: Number.isFinite(dragState.lastTimelineX) ? "lastTimelineX" : "drop-event-fallback",
        preSnapMinutes,
        snappedMinutes: snappedStart,
        clampedMinutes,
        finalMinutes: getSafeStartMinutes(item.start),
        finalTime: item.start,
        dropStatus: "DROP FIRED AND COMMITTED",
        dropFired: dropFired ? "yes" : "no",
        dropCommitted: "yes",
        dropAllowed: "yes",
        dropBlockedReason: "-",
        itemId: item.id,
        renderedLeft: timelineMinutesToX(getSafeStartMinutes(item.start)),
        renderNote: `saved start rendered at ${timelineMinutesToX(getSafeStartMinutes(item.start))}px`,
      })
    : null;

  if (DEBUG_TIMELINE_DND && renderSnapshot) {
    updateTimelineDebugPanel(renderSnapshot);
    logTimelineDebug("render", renderSnapshot);
  }

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

    const validation = validateSchedule("fs42-strict");
    if (!validation.ready) {
      renderExportReadiness();
      const message = validation.blockers.slice(0, 4).map(renderIssueLine).join("\n");
      window.alert(`FS42 strict export is blocked:\n${message}`);
      return;
    }

    exportChannelJsons("fs42-strict");
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
      viewMode: state.viewMode,
      timelineScale: state.timelineScale,
      selectedDay: state.selectedDay,
      selectedChannelId: state.selectedChannelId,
    },
  };
  downloadFile("fs42-scheduler-internal.json", JSON.stringify(payload, null, 2), "application/json");
}

function exportChannelJsons(profile = "fs42-strict") {
  const manifests = [];
  state.channels.forEach((channel, index) => {
    const channelItems = state.items.filter((item) => item.channelId === channel.id);
    const payload = buildChannelJson(channel, channelItems, index + 1, profile);
    manifests.push({
      channel: channel.name,
      file: `${slugify(channel.name)}.json`,
      commercial: payload.station_conf.commercial,
      items: payload.station_conf.programme_count,
    });

    window.setTimeout(() => {
      downloadFile(`${slugify(channel.name)}.json`, JSON.stringify(payload, null, 2), "application/json");
    }, index * 120);
  });

  window.setTimeout(() => {
    downloadFile("fs42-channel-manifest.json", JSON.stringify(manifests, null, 2), "application/json");
  }, state.channels.length * 120 + 50);
}

function buildChannelJson(channel, channelItems, channelNumber, profile = "fs42-strict") {
  const nonCommercialItems = channelItems
    .filter((item) => !ITEM_TYPE_META[item.itemType]?.commercial)
    .sort(compareItems);
  const commercialItems = channelItems
    .filter((item) => ITEM_TYPE_META[item.itemType]?.commercial)
    .sort(compareItems);
  const daySchedules = buildScheduleByDay(nonCommercialItems);
  const clipShows = uniqueValues(
    nonCommercialItems
      .filter((item) => item.itemType === "Programme")
      .map((item) => slugify(item.blockGroup || item.title)),
  );

  return {
    station_conf: {
      network_name: channel.name,
      channel_number: channelNumber,
      network_type: "standard",
      schedule_increment: 30,
      break_strategy: commercialItems.length > 0 ? "standard" : "none",
      commercial_free: commercialItems.length === 0,
      commercial: commercialItems.length > 0,
      break_duration: commercialItems.length > 0 ? 120 : 0,
      break_count: commercialItems.length,
      export_profile: profile,
      content_dir: `catalog/${slugify(channel.name)}_catalog`,
      commercial_dir: "commercial",
      bump_dir: "bump",
      clip_shows: clipShows,
      sign_off_video: "runtime/signoff.mp4",
      off_air_video: "runtime/off_air_pattern.mp4",
      standby_image: "runtime/standby.png",
      be_right_back_media: "runtime/brb.png",
      logo_dir: "logos",
      show_logo: true,
      default_logo: `${slugify(channel.name)}.png`,
      logo_permanent: true,
      multi_logo: "multi",
      channel_group: channel.group || "",
      tagline: channel.tagline || "",
      programme_count: nonCommercialItems.filter((item) => item.itemType === "Programme").length,
      commercial_count: commercialItems.length,
      day_count: uniqueValues(nonCommercialItems.map((item) => item.day)).length,
      ...daySchedules,
    },
  };
}

function buildScheduleByDay(items) {
  const schedule = {};
  DAY_NAMES.forEach((day) => {
    const dayKey = day.toLowerCase();
    const dayItems = items.filter((item) => item.day === day).sort(compareItems);
    const daySchedule = {};

    dayItems.forEach((item) => {
      const hour = String(Math.floor(timeToMinutes(item.start) / 60)).padStart(2, "0");
      if (!daySchedule[hour]) daySchedule[hour] = [];
      daySchedule[hour].push({
        tags: scheduleTagForItem(item),
        title: item.title,
        start: item.start,
        duration: getSafeDuration(item.duration, item.itemType),
        type: item.itemType,
        category: item.category,
        code: item.assetCode || "",
        watershed_restricted: Boolean(item.watershedRestricted),
        prime_time: Boolean(item.primeTime),
        must_run: Boolean(item.mustRun),
        flags: getFlagList(item).map((flag) => flag.short),
      });
    });

    schedule[dayKey] = daySchedule;
  });

  return schedule;
}

function scheduleTagForItem(item) {
  if (item.itemType === "Programme") return slugify(item.blockGroup || item.category || item.title);
  if (item.itemType === "Ad Break") return "commercial-break";
  if (item.itemType === "Ad Spot") return "commercial-spot";
  if (item.itemType === "Bumper") return "bump";
  if (item.itemType === "Promo") return "promo";
  return slugify(item.itemType);
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
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
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
      viewMode: parsed.viewMode || "timeline",
      timelineScale: parsed.timelineScale || "day",
      colorMode: parsed.colorMode || "channel",
      exportProfile: parsed.exportProfile || "fs42-strict",
      selectedDay: parsed.selectedDay || "Monday",
      selectedChannelId: parsed.selectedChannelId || "all",
      channels: parsed.channels.map((channel) => ({
        id: channel.id || crypto.randomUUID(),
        name: channel.name,
        group: channel.group || "",
        color: channel.color || "#7d91b9",
        tagline: channel.tagline || "",
      })),
      items: normalizeItems(parsed.items || parsed.shows || [], parsed.channels),
    };
  }

  if (Array.isArray(parsed.shows) && parsed.shows.length > 0) {
    const channels = structuredClone(DEFAULT_CHANNELS);
    const map = new Map(channels.map((channel) => [channel.name, channel.id]));
    return {
      viewMode: "timeline",
      timelineScale: "day",
      colorMode: parsed.colorMode || "channel",
      exportProfile: parsed.exportProfile || "fs42-strict",
      selectedDay: parsed.selectedDay || "Monday",
      selectedChannelId: "all",
      channels,
      items: parsed.shows.map((show) => ({
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
        notes: show.notes || "",
      })),
    };
  }

  return fallback;
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
