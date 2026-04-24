const BOM_COLUMN_ALIASES = {
  projectCode: ["project code", "projectcode", "project"],
  drawingNo: ["drawing no", "drawing no.", "drawing number", "drawingno"],
  spoolNo: ["spool no", "spool no.", "spool number", "spoolno"],
  itemCode: ["item code", "item code icd", "item code (icd)", "icd", "itemcode"],
  quantity: ["total qty", "total q'ty", "qty", "quantity", "total quantity"],
  priority: ["priority", "spool priority"],
  totalWt: ["total wt", "total wt.", "weight", "total weight"],
  totalInchDia: ["total inch dia", "inch dia", "inch diameter", "total inch diameter"],
  material: ["material", "material type", "moc"]
};

const STOCK_COLUMN_ALIASES = {
  projectCode: ["project code", "projectcode", "project"],
  itemCode: ["item code", "item code icd", "item code (icd)", "icd", "itemcode"],
  location: ["location", "stock location", "store qc", "store/qc"],
  quantity: ["total qty", "total q'ty", "qty", "quantity", "total quantity"]
};

const ANALYSIS_MODES = {
  priority: {
    label: "By Priority",
    note: "Spools are analyzed by priority from lowest number to highest. Blank priorities are treated as the last priority."
  },
  weight: {
    label: "By Wt.",
    note: "Spools are analyzed by master-row Total Wt. from highest to lowest."
  },
  inchDia: {
    label: "By Inch Dia",
    note: "Spools are analyzed by master-row Total Inch Dia from highest to lowest."
  },
  best: {
    label: "Best Analysis",
    note: "Spools stay in uploaded order. If a spool is only partly coverable, its tentative allocation is released so later spools can use that stock."
  }
};

const STATUS_FULL = "100% Material Available";
const STATUS_FULL_QC = "100% Material Available but some items are at QC location";
const STATUS_PARTIAL = "Partial Material Available";
const STATUS_NONE = "No Material Available";

const EMPTY_SUMMARY_COLSPAN = 13;
const EMPTY_DETAIL_COLSPAN = 13;

const state = {
  bomRows: [],
  stockRows: [],
  bomFileName: "",
  stockFileName: "",
  bomColumns: [],
  stockColumns: [],
  analysis: null,
  statusFilter: "all",
  searchFilter: "",
  analysisMode: "priority"
};

const elements = {
  bomUpload: document.getElementById("bom-upload"),
  stockUpload: document.getElementById("stock-upload"),
  loadDemoButton: document.getElementById("load-demo"),
  runAnalysisButton: document.getElementById("run-analysis"),
  exportSummaryButton: document.getElementById("export-summary"),
  exportDetailButton: document.getElementById("export-detail"),
  resetButton: document.getElementById("reset-app"),
  bomFileInfo: document.getElementById("bom-file-info"),
  stockFileInfo: document.getElementById("stock-file-info"),
  bomColumnPreview: document.getElementById("bom-column-preview"),
  stockColumnPreview: document.getElementById("stock-column-preview"),
  analysisMode: document.getElementById("analysis-mode"),
  analysisModeNote: document.getElementById("analysis-mode-note"),
  analysisStatus: document.getElementById("analysis-status"),
  kpiTotalSpools: document.getElementById("kpi-total-spools"),
  kpiFullStore: document.getElementById("kpi-full-store"),
  kpiFullQc: document.getElementById("kpi-full-qc"),
  kpiPartial: document.getElementById("kpi-partial"),
  kpiNone: document.getElementById("kpi-none"),
  statusFilter: document.getElementById("status-filter"),
  searchFilter: document.getElementById("search-filter"),
  summaryNote: document.getElementById("summary-note"),
  summaryTableBody: document.getElementById("summary-table-body"),
  detailTitle: document.getElementById("detail-title"),
  detailNote: document.getElementById("detail-note"),
  detailTableBody: document.getElementById("detail-table-body")
};

bindEvents();
render();

function bindEvents() {
  elements.bomUpload.addEventListener("change", (event) => handleFileUpload(event, "bom"));
  elements.stockUpload.addEventListener("change", (event) => handleFileUpload(event, "stock"));
  elements.loadDemoButton.addEventListener("click", loadDemoData);
  elements.runAnalysisButton.addEventListener("click", () => runAnalysis());
  elements.exportSummaryButton.addEventListener("click", exportSummaryWorkbook);
  elements.exportDetailButton.addEventListener("click", exportDetailWorkbook);
  elements.resetButton.addEventListener("click", resetApp);
  elements.analysisMode.addEventListener("change", () => {
    state.analysisMode = elements.analysisMode.value;
    if (state.bomRows.length && state.stockRows.length) {
      runAnalysis({ silentStatus: true });
      setStatus(
        `Analysis rerun in ${ANALYSIS_MODES[state.analysisMode].label}. ${ANALYSIS_MODES[state.analysisMode].note}`
      );
      return;
    }
    state.analysis = null;
    render();
  });
  elements.statusFilter.addEventListener("change", () => {
    state.statusFilter = elements.statusFilter.value;
    render();
  });
  elements.searchFilter.addEventListener("input", () => {
    state.searchFilter = elements.searchFilter.value;
    render();
  });
}

async function handleFileUpload(event, type) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  try {
    const imported = await parseImportedFile(file);
    const { columns, rows } = imported;

    if (type === "bom") {
      const parsedBomRows = buildBomRows(rows, columns);
      state.bomRows = parsedBomRows;
      state.bomFileName = file.name;
      state.bomColumns = columns;
      if (!parsedBomRows.length) {
        setStatus(
          `Loaded BOM file ${file.name}, but no valid BOM rows were found. Keep the data in this column order: Project Code, Drawing No., Spool No., Item Code, Total Q'ty, Priority, Total Wt., Total Inch Dia, Material.`
        );
      } else {
        setStatus(
          `Loaded BOM file ${file.name} with ${parsedBomRows.length} valid row(s) across ${countUniqueSpools(parsedBomRows)} spool(s).`
        );
      }
    } else {
      const parsedStockRows = buildStockRows(rows, columns);
      state.stockRows = parsedStockRows;
      state.stockFileName = file.name;
      state.stockColumns = columns;
      if (!parsedStockRows.length) {
        setStatus(
          `Loaded inventory file ${file.name}, but no valid stock rows were found. Keep the data in this column order: Project Code, Item Code, Location, Total Q'ty.`
        );
      } else {
        setStatus(`Loaded inventory file ${file.name} with ${parsedStockRows.length} valid stock row(s).`);
      }
    }

    state.analysis = null;
  } catch (error) {
    setStatus(error.message || "The file could not be loaded.");
  } finally {
    event.target.value = "";
    render();
  }
}

async function parseImportedFile(file) {
  const extension = getFileExtension(file.name);
  let matrix = [];

  if (extension === "csv") {
    matrix = parseCsvMatrix(await file.text());
  } else {
    if (!window.XLSX) {
      throw new Error("Excel parsing is unavailable right now. Please save the file as CSV and upload it.");
    }

    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });
    const firstSheet = workbook.SheetNames[0];
    matrix = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false
    });
  }

  return normalizeImportedMatrix(matrix);
}

function normalizeImportedMatrix(matrix) {
  const safeRows = (Array.isArray(matrix) ? matrix : []).map((row) => (Array.isArray(row) ? row : [row]));
  const firstContentIndex = safeRows.findIndex(isNonEmptyRow);

  if (firstContentIndex === -1) {
    return { columns: [], rows: [] };
  }

  const headerRow = safeRows[firstContentIndex];
  const columns = headerRow.map((value, index) => sanitizeHeaderLabel(value, index));
  const rows = safeRows
    .slice(firstContentIndex + 1)
    .filter(isNonEmptyRow)
    .map((row) => row.map((value) => value ?? ""));

  return { columns, rows };
}

function buildBomRows(rows, columns) {
  const headerIndexes = buildHeaderIndexMap(columns);

  return rows
    .map((row, index) => {
      const projectCode = sanitizeCode(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.projectCode, 0));
      const drawingNo = sanitizeCode(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.drawingNo, 1));
      const spoolNo = sanitizeCode(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.spoolNo, 2));
      const itemCode = sanitizeCode(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.itemCode, 3));
      const quantity = parseQuantity(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.quantity, 4));
      const priority = parseOptionalNumber(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.priority, 5));
      const totalWt = parseOptionalNumber(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.totalWt, 6));
      const totalInchDia = parseOptionalNumber(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.totalInchDia, 7));
      const material = sanitizeText(getCellValue(row, headerIndexes, BOM_COLUMN_ALIASES.material, 8));
      const hasIdentity = Boolean(projectCode && drawingNo && spoolNo);
      const isMaster = hasIdentity && !itemCode;
      const isComponent = hasIdentity && Boolean(itemCode) && quantity > 0;

      if (!hasIdentity || (!isMaster && !isComponent)) {
        return null;
      }

      return {
        id: `bom-${index}`,
        order: index,
        projectCode,
        drawingNo,
        spoolNo,
        itemCode,
        quantity: isComponent ? quantity : 0,
        priority,
        totalWt,
        totalInchDia,
        material,
        isMaster
      };
    })
    .filter(Boolean);
}

function buildStockRows(rows, columns) {
  const headerIndexes = buildHeaderIndexMap(columns);

  return rows
    .map((row, index) => {
      const projectCode = sanitizeCode(getCellValue(row, headerIndexes, STOCK_COLUMN_ALIASES.projectCode, 0));
      const itemCode = sanitizeCode(getCellValue(row, headerIndexes, STOCK_COLUMN_ALIASES.itemCode, 1));
      const location = normalizeLocation(getCellValue(row, headerIndexes, STOCK_COLUMN_ALIASES.location, 2));
      const quantity = parseQuantity(getCellValue(row, headerIndexes, STOCK_COLUMN_ALIASES.quantity, 3));

      if (!projectCode || !itemCode || !location || quantity <= 0) {
        return null;
      }

      return {
        id: `stock-${index}`,
        projectCode,
        itemCode,
        location,
        quantity
      };
    })
    .filter(Boolean);
}

function runAnalysis(options = {}) {
  if (!state.bomRows.length || !state.stockRows.length) {
    setStatus("Please load both the BOM file and the inventory file before running the checker.");
    render();
    return;
  }

  state.analysis = analyzeAvailability(state.bomRows, state.stockRows, state.analysisMode);

  if (!options.silentStatus) {
    setStatus(
      `Analysis complete for ${state.analysis.spools.length} spool(s) using ${state.analysis.modeLabel}. Store is checked first, then QC.`
    );
  }

  render();
}

function analyzeAvailability(bomRows, stockRows, analysisMode) {
  const spools = sortSpoolsForMode(buildSpools(bomRows), analysisMode);
  const stockTotals = buildStockTotals(stockRows);
  const remaining = cloneStockMap(stockTotals);
  const analyzedSpools = spools.map((spool) =>
    allocateSpool(spool, remaining, stockTotals, {
      analysisMode
    })
  );

  return {
    mode: analysisMode,
    modeLabel: ANALYSIS_MODES[analysisMode].label,
    spools: analyzedSpools,
    counts: {
      total: analyzedSpools.length,
      fullStore: analyzedSpools.filter((spool) => spool.status === STATUS_FULL).length,
      fullQc: analyzedSpools.filter((spool) => spool.status === STATUS_FULL_QC).length,
      partial: analyzedSpools.filter((spool) => spool.status === STATUS_PARTIAL).length,
      none: analyzedSpools.filter((spool) => spool.status === STATUS_NONE).length
    }
  };
}

function buildSpools(bomRows) {
  const groupedSpools = new Map();

  bomRows.forEach((row) => {
    const key = buildSpoolKey(row.projectCode, row.drawingNo, row.spoolNo);
    if (!groupedSpools.has(key)) {
      groupedSpools.set(key, {
        key,
        order: row.order,
        projectCode: row.projectCode,
        drawingNo: row.drawingNo,
        spoolNo: row.spoolNo,
        masterRow: null,
        componentRows: []
      });
    }

    const group = groupedSpools.get(key);
    group.order = Math.min(group.order, row.order);
    if (row.isMaster && !group.masterRow) {
      group.masterRow = row;
      return;
    }
    if (!row.isMaster) {
      group.componentRows.push(row);
    }
  });

  return [...groupedSpools.values()]
    .sort((left, right) => left.order - right.order)
    .map((group) => {
      const components = mergeComponents(group.componentRows);
      const priority = pickFirstNumber([group.masterRow?.priority, ...group.componentRows.map((row) => row.priority)]);
      const totalWt = pickFirstNumber([group.masterRow?.totalWt, ...group.componentRows.map((row) => row.totalWt)]);
      const totalInchDia = pickFirstNumber([
        group.masterRow?.totalInchDia,
        ...group.componentRows.map((row) => row.totalInchDia)
      ]);
      const material = buildMaterialLabel(group.masterRow?.material, components.map((component) => component.material));

      return {
        key: group.key,
        order: group.order,
        projectCode: group.projectCode,
        drawingNo: group.drawingNo,
        spoolNo: group.spoolNo,
        priority,
        totalWt,
        totalInchDia,
        material,
        components
      };
    });
}

function mergeComponents(componentRows) {
  const merged = new Map();

  componentRows.forEach((row) => {
    if (!merged.has(row.itemCode)) {
      merged.set(row.itemCode, {
        itemCode: row.itemCode,
        quantity: 0,
        materialLabels: []
      });
    }

    const component = merged.get(row.itemCode);
    component.quantity = roundQuantity(component.quantity + row.quantity);
    if (row.material) {
      component.materialLabels.push(row.material);
    }
  });

  return [...merged.values()].map((component) => ({
    itemCode: component.itemCode,
    quantity: component.quantity,
    material: collapseLabels(component.materialLabels)
  }));
}

function sortSpoolsForMode(spools, analysisMode) {
  const sorted = [...spools];

  sorted.sort((left, right) => {
    if (analysisMode === "priority") {
      return (
        compareNullableNumbersAscending(left.priority, right.priority) ||
        left.order - right.order ||
        left.key.localeCompare(right.key)
      );
    }

    if (analysisMode === "weight") {
      return (
        compareNullableNumbersDescending(left.totalWt, right.totalWt) ||
        left.order - right.order ||
        left.key.localeCompare(right.key)
      );
    }

    if (analysisMode === "inchDia") {
      return (
        compareNullableNumbersDescending(left.totalInchDia, right.totalInchDia) ||
        left.order - right.order ||
        left.key.localeCompare(right.key)
      );
    }

    return left.order - right.order || left.key.localeCompare(right.key);
  });

  return sorted;
}

function allocateSpool(spool, remaining, stockTotals, options = {}) {
  const analysisMode = options.analysisMode || "priority";
  const workingRemaining = analysisMode === "best" ? cloneStockMap(remaining) : remaining;
  let totalRequiredQty = 0;
  let totalStoreAllocated = 0;
  let totalQcAllocated = 0;
  let totalShortQty = 0;
  let anyAllocated = false;
  let anyQcUsed = false;

  const components = spool.components.map((component) => {
    const stockKey = buildStockKey(spool.projectCode, component.itemCode);
    const pool = cloneStockEntry(workingRemaining.get(stockKey));
    const visibleTotals = cloneStockEntry(stockTotals.get(stockKey));
    const storeBeforeQty = pool.store;
    const qcBeforeQty = pool.qc;
    const storeAllocatedQty = Math.min(component.quantity, storeBeforeQty);
    const qcAllocatedQty = Math.min(roundQuantity(component.quantity - storeAllocatedQty), qcBeforeQty);
    const shortQty = roundQuantity(component.quantity - storeAllocatedQty - qcAllocatedQty);

    pool.store = roundQuantity(storeBeforeQty - storeAllocatedQty);
    pool.qc = roundQuantity(qcBeforeQty - qcAllocatedQty);
    workingRemaining.set(stockKey, pool);

    totalRequiredQty = roundQuantity(totalRequiredQty + component.quantity);
    totalStoreAllocated = roundQuantity(totalStoreAllocated + storeAllocatedQty);
    totalQcAllocated = roundQuantity(totalQcAllocated + qcAllocatedQty);
    totalShortQty = roundQuantity(totalShortQty + shortQty);
    anyAllocated = anyAllocated || storeAllocatedQty > 0 || qcAllocatedQty > 0;
    anyQcUsed = anyQcUsed || qcAllocatedQty > 0;

    return {
      itemCode: component.itemCode,
      material: component.material || spool.material,
      requiredQty: component.quantity,
      storeBeforeQty,
      storeAllocatedQty,
      qcBeforeQty,
      qcAllocatedQty,
      shortQty,
      remarks: getStandardComponentRemark({
        storeAllocatedQty,
        qcAllocatedQty,
        shortQty,
        totalQcVisible: visibleTotals.qc
      })
    };
  });

  const isFullyCovered = roundQuantity(totalShortQty) === 0;

  if (analysisMode === "best" && !isFullyCovered) {
    const rolledBackComponents = components.map((component) => {
      const hadTentativeAllocation = component.storeAllocatedQty > 0 || component.qcAllocatedQty > 0;
      return {
        ...component,
        storeAllocatedQty: 0,
        qcAllocatedQty: 0,
        shortQty: component.requiredQty,
        remarks: getBestAnalysisRemark(component, hadTentativeAllocation)
      };
    });

    return finalizeSpoolResult(spool, STATUS_NONE, rolledBackComponents);
  }

  if (analysisMode === "best") {
    copyStockMap(remaining, workingRemaining);
  }

  return finalizeSpoolResult(
    spool,
    determineSpoolStatus({
      totalShortQty,
      anyAllocated,
      anyQcUsed
    }),
    components
  );
}

function finalizeSpoolResult(spool, status, components) {
  const totals = components.reduce(
    (result, component) => ({
      requiredQty: roundQuantity(result.requiredQty + component.requiredQty),
      storeAllocatedQty: roundQuantity(result.storeAllocatedQty + component.storeAllocatedQty),
      qcAllocatedQty: roundQuantity(result.qcAllocatedQty + component.qcAllocatedQty),
      shortQty: roundQuantity(result.shortQty + component.shortQty)
    }),
    {
      requiredQty: 0,
      storeAllocatedQty: 0,
      qcAllocatedQty: 0,
      shortQty: 0
    }
  );

  return {
    key: spool.key,
    projectCode: spool.projectCode,
    drawingNo: spool.drawingNo,
    spoolNo: spool.spoolNo,
    priority: spool.priority,
    totalWt: spool.totalWt,
    totalInchDia: spool.totalInchDia,
    material: spool.material,
    status,
    componentCount: components.length,
    totalRequiredQty: totals.requiredQty,
    totalStoreAllocated: totals.storeAllocatedQty,
    totalQcAllocated: totals.qcAllocatedQty,
    totalShortQty: totals.shortQty,
    components
  };
}

function determineSpoolStatus({ totalShortQty, anyAllocated, anyQcUsed }) {
  if (roundQuantity(totalShortQty) === 0) {
    return anyQcUsed ? STATUS_FULL_QC : STATUS_FULL;
  }
  if (anyAllocated) {
    return STATUS_PARTIAL;
  }
  return STATUS_NONE;
}

function buildStockTotals(stockRows) {
  const totals = new Map();

  stockRows.forEach((row) => {
    const key = buildStockKey(row.projectCode, row.itemCode);
    if (!totals.has(key)) {
      totals.set(key, {
        store: 0,
        qc: 0
      });
    }

    const projectItemTotals = totals.get(key);
    if (row.location === "Store") {
      projectItemTotals.store = roundQuantity(projectItemTotals.store + row.quantity);
    }
    if (row.location === "QC") {
      projectItemTotals.qc = roundQuantity(projectItemTotals.qc + row.quantity);
    }
  });

  return totals;
}

function render() {
  renderAnalysisModeInfo();
  renderFileInfo();
  renderSummaryCards();
  renderSummaryTable();
  renderDetailTable();
  const hasVisibleSpools = getFilteredSpools().length > 0;
  elements.exportSummaryButton.disabled = !hasVisibleSpools;
  elements.exportDetailButton.disabled = !hasVisibleSpools;
}

function renderAnalysisModeInfo() {
  elements.analysisMode.value = state.analysisMode;
  elements.analysisModeNote.textContent = ANALYSIS_MODES[state.analysisMode].note;
}

function renderFileInfo() {
  elements.bomFileInfo.textContent = state.bomFileName
    ? `${state.bomFileName} | ${state.bomRows.length} valid row(s) across ${countUniqueSpools(state.bomRows)} spool(s)`
    : "No BOM file loaded.";
  elements.stockFileInfo.textContent = state.stockFileName
    ? `${state.stockFileName} | ${state.stockRows.length} valid inventory row(s)`
    : "No inventory file loaded.";
  renderColumnPreview(elements.bomColumnPreview, state.bomColumns);
  renderColumnPreview(elements.stockColumnPreview, state.stockColumns);
}

function renderColumnPreview(container, columns) {
  container.innerHTML = columns.length
    ? columns.map((column) => `<span class="chip">${escapeHtml(column)}</span>`).join("")
    : "";
}

function renderSummaryCards() {
  setText(elements.kpiTotalSpools, state.analysis?.counts.total || 0);
  setText(elements.kpiFullStore, state.analysis?.counts.fullStore || 0);
  setText(elements.kpiFullQc, state.analysis?.counts.fullQc || 0);
  setText(elements.kpiPartial, state.analysis?.counts.partial || 0);
  setText(elements.kpiNone, state.analysis?.counts.none || 0);
}

function renderSummaryTable() {
  if (!state.analysis) {
    elements.summaryNote.textContent = "Run the checker to see spool status.";
    elements.summaryTableBody.innerHTML = `
      <tr><td colspan="${EMPTY_SUMMARY_COLSPAN}"><div class="empty-state">No result yet. Upload files and run the analysis.</div></td></tr>
    `;
    return;
  }

  const filteredSpools = getFilteredSpools();
  elements.summaryNote.textContent = `${filteredSpools.length} spool(s) shown from ${state.analysis.spools.length} analyzed spool(s) in ${state.analysis.modeLabel}.`;

  if (!filteredSpools.length) {
    elements.summaryTableBody.innerHTML = `
      <tr><td colspan="${EMPTY_SUMMARY_COLSPAN}"><div class="empty-state">No spool matches the current filters.</div></td></tr>
    `;
    return;
  }

  elements.summaryTableBody.innerHTML = filteredSpools
    .map(
      (spool) => `
      <tr>
        <td>${escapeHtml(spool.projectCode)}</td>
        <td>${escapeHtml(spool.drawingNo)}</td>
        <td>${escapeHtml(spool.spoolNo)}</td>
        <td class="is-number">${formatDisplayNumber(spool.priority)}</td>
        <td class="is-number">${formatDisplayNumber(spool.totalWt)}</td>
        <td class="is-number">${formatDisplayNumber(spool.totalInchDia)}</td>
        <td>${escapeHtml(spool.material || "-")}</td>
        <td>${renderStatusChip(spool.status)}</td>
        <td class="is-number">${formatQuantity(spool.componentCount)}</td>
        <td class="is-number">${formatQuantity(spool.totalRequiredQty)}</td>
        <td class="is-number">${formatQuantity(spool.totalStoreAllocated)}</td>
        <td class="is-number">${formatQuantity(spool.totalQcAllocated)}</td>
        <td class="is-number">${formatQuantity(spool.totalShortQty)}</td>
      </tr>
    `
    )
    .join("");
}

function renderDetailTable() {
  if (!state.analysis) {
    elements.detailTitle.textContent = "Component detail for all shown spools";
    elements.detailNote.textContent = "Store-first allocation with QC fallback is shown component by component.";
    elements.detailTableBody.innerHTML = `
      <tr><td colspan="${EMPTY_DETAIL_COLSPAN}"><div class="empty-state">No result yet. Upload files and run the analysis.</div></td></tr>
    `;
    return;
  }

  const visibleSpools = getFilteredSpools();
  const componentRows = getVisibleComponentRows(visibleSpools);

  if (!visibleSpools.length) {
    elements.detailTitle.textContent = "Component detail for all shown spools";
    elements.detailNote.textContent = "No component rows match the current filters.";
    elements.detailTableBody.innerHTML = `
      <tr><td colspan="${EMPTY_DETAIL_COLSPAN}"><div class="empty-state">No component detail matches the current filters.</div></td></tr>
    `;
    return;
  }

  elements.detailTitle.textContent = "Component detail for all shown spools";
  elements.detailNote.textContent = `${componentRows.length} component row(s) shown across ${visibleSpools.length} spool(s).`;
  elements.detailTableBody.innerHTML = componentRows
    .map(
      (component) => `
      <tr>
        <td>${escapeHtml(component.projectCode)}</td>
        <td>${escapeHtml(component.drawingNo)}</td>
        <td>${escapeHtml(component.spoolNo)}</td>
        <td>${escapeHtml(component.material || "-")}</td>
        <td>${renderStatusChip(component.status)}</td>
        <td>${escapeHtml(component.itemCode)}</td>
        <td class="is-number">${formatQuantity(component.requiredQty)}</td>
        <td class="is-number">${formatQuantity(component.storeBeforeQty)}</td>
        <td class="is-number">${formatQuantity(component.storeAllocatedQty)}</td>
        <td class="is-number">${formatQuantity(component.qcBeforeQty)}</td>
        <td class="is-number">${formatQuantity(component.qcAllocatedQty)}</td>
        <td class="is-number">${formatQuantity(component.shortQty)}</td>
        <td>${escapeHtml(component.remarks)}</td>
      </tr>
    `
    )
    .join("");
}

function getFilteredSpools() {
  if (!state.analysis) {
    return [];
  }

  const search = normalizeHeaderName(state.searchFilter);
  return state.analysis.spools.filter((spool) => {
    const matchesStatus = state.statusFilter === "all" || spool.status === state.statusFilter;
    const haystack = normalizeHeaderName(
      [
        spool.projectCode,
        spool.drawingNo,
        spool.spoolNo,
        spool.material,
        spool.status,
        spool.priority,
        spool.totalWt,
        spool.totalInchDia,
        ...spool.components.map((component) => `${component.itemCode} ${component.material}`)
      ]
        .filter(Boolean)
        .join(" ")
    );
    const matchesSearch = !search || haystack.includes(search);
    return matchesStatus && matchesSearch;
  });
}

function getVisibleComponentRows(spools = getFilteredSpools()) {
  return spools.flatMap((spool) =>
    spool.components.map((component) => ({
      projectCode: spool.projectCode,
      drawingNo: spool.drawingNo,
      spoolNo: spool.spoolNo,
      material: component.material || spool.material,
      status: spool.status,
      itemCode: component.itemCode,
      requiredQty: component.requiredQty,
      storeBeforeQty: component.storeBeforeQty,
      storeAllocatedQty: component.storeAllocatedQty,
      qcBeforeQty: component.qcBeforeQty,
      qcAllocatedQty: component.qcAllocatedQty,
      shortQty: component.shortQty,
      remarks: component.remarks
    }))
  );
}

function exportSummaryWorkbook() {
  const filteredSpools = getFilteredSpools();
  if (!filteredSpools.length) {
    return;
  }

  downloadWorkbook(
    buildExportFileName("fabrication-readiness-by-spool"),
    "Fabrication readiness",
    [
      [
        "Project Code",
        "Drawing No.",
        "Spool No.",
        "Priority",
        "Total Wt.",
        "Total Inch Dia",
        "Material",
        "Status",
        "Components",
        "Required Qty",
        "Store Allocated",
        "QC Allocated",
        "Short Qty"
      ],
      ...filteredSpools.map((spool) => [
        spool.projectCode,
        spool.drawingNo,
        spool.spoolNo,
        spool.priority ?? "",
        spool.totalWt ?? "",
        spool.totalInchDia ?? "",
        spool.material,
        spool.status,
        spool.componentCount,
        spool.totalRequiredQty,
        spool.totalStoreAllocated,
        spool.totalQcAllocated,
        spool.totalShortQty
      ])
    ]
  );
}

function exportDetailWorkbook() {
  const componentRows = getVisibleComponentRows();
  if (!componentRows.length) {
    return;
  }

  downloadWorkbook(
    buildExportFileName("component-detail"),
    "Component detail",
    [
      [
        "Project Code",
        "Drawing No.",
        "Spool No.",
        "Material",
        "Status",
        "Item Code",
        "Required Qty",
        "Store Before",
        "Store Allocated",
        "QC Before",
        "QC Allocated",
        "Short Qty",
        "Remarks"
      ],
      ...componentRows.map((component) => [
        component.projectCode,
        component.drawingNo,
        component.spoolNo,
        component.material,
        component.status,
        component.itemCode,
        component.requiredQty,
        component.storeBeforeQty,
        component.storeAllocatedQty,
        component.qcBeforeQty,
        component.qcAllocatedQty,
        component.shortQty,
        component.remarks
      ])
    ]
  );
}

function downloadWorkbook(fileName, sheetName, rows) {
  if (!window.XLSX) {
    setStatus("Excel export is unavailable right now. Please refresh the page and try again.");
    return;
  }

  const workbook = window.XLSX.utils.book_new();
  const sheet = window.XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = buildColumnWidths(rows);
  window.XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  window.XLSX.writeFile(workbook, fileName);
  setStatus(`${fileName} downloaded successfully.`);
}

function resetApp() {
  const shouldReset =
    typeof window.confirm !== "function" ||
    window.confirm("Reset the uploaded files and all analysis results?");

  if (!shouldReset) {
    return;
  }

  state.bomRows = [];
  state.stockRows = [];
  state.bomFileName = "";
  state.stockFileName = "";
  state.bomColumns = [];
  state.stockColumns = [];
  state.analysis = null;
  state.statusFilter = "all";
  state.searchFilter = "";
  state.analysisMode = "priority";

  elements.bomUpload.value = "";
  elements.stockUpload.value = "";
  elements.statusFilter.value = "all";
  elements.searchFilter.value = "";
  elements.analysisMode.value = "priority";

  setStatus("Upload both files, then run the checker.");
  render();
}

function loadDemoData() {
  const bomColumns = [
    "Project Code",
    "Drawing No.",
    "Spool No.",
    "Item Code",
    "Total Q'ty",
    "Priority",
    "Total Wt.",
    "Total Inch Dia",
    "Material"
  ];
  const stockColumns = ["Project Code", "Item Code", "Location", "Total Q'ty"];

  const bomRows = [
    ["PRJ-01", "DWG-100", "SP-001", "", "", 1, 120, 24, "Carbon Steel"],
    ["PRJ-01", "DWG-100", "SP-001", "ICD-PIPE-8", 5, "", "", "", "Pipe"],
    ["PRJ-01", "DWG-100", "SP-001", "ICD-ELB-8", 3, "", "", "", "Elbow"],
    ["PRJ-01", "DWG-101", "SP-002", "", "", "", 95, 18, "Stainless Steel"],
    ["PRJ-01", "DWG-101", "SP-002", "ICD-PIPE-8", 4, "", "", "", "Pipe"],
    ["PRJ-01", "DWG-101", "SP-002", "ICD-FLG-8", 2, "", "", "", "Flange"],
    ["PRJ-01", "DWG-102", "SP-003", "", "", 2, 110, 22, "Carbon Steel"],
    ["PRJ-01", "DWG-102", "SP-003", "ICD-PIPE-8", 4, "", "", "", "Pipe"],
    ["PRJ-01", "DWG-102", "SP-003", "ICD-GSK-8", 1, "", "", "", "Gasket"],
    ["PRJ-02", "DWG-220", "SP-010", "", "", 1, 60, 12, "Alloy Steel"],
    ["PRJ-02", "DWG-220", "SP-010", "ICD-TEE-6", 2, "", "", "", "Tee"],
    ["PRJ-02", "DWG-220", "SP-010", "ICD-PIPE-6", 1, "", "", "", "Pipe"]
  ];

  const stockRows = [
    ["PRJ-01", "ICD-PIPE-8", "Store", 9],
    ["PRJ-01", "ICD-ELB-8", "Store", 3],
    ["PRJ-01", "ICD-FLG-8", "Store", 1],
    ["PRJ-01", "ICD-GSK-8", "Store", 1],
    ["PRJ-02", "ICD-TEE-6", "Store", 1],
    ["PRJ-02", "ICD-TEE-6", "QC", 1],
    ["PRJ-02", "ICD-PIPE-6", "QC", 1]
  ];

  state.bomRows = buildBomRows(bomRows, bomColumns);
  state.stockRows = buildStockRows(stockRows, stockColumns);
  state.bomFileName = "demo-bom.xlsx";
  state.stockFileName = "demo-stock.xlsx";
  state.bomColumns = bomColumns;
  state.stockColumns = stockColumns;
  runAnalysis({ silentStatus: true });
  setStatus(
    "Demo data loaded. Try switching between By Priority and Best Analysis to see how partial spool allocations are treated."
  );
  render();
}

function setStatus(message) {
  elements.analysisStatus.textContent = message;
}

function renderStatusChip(status) {
  return `<span class="status-pill ${getStatusClass(status)}">${escapeHtml(status)}</span>`;
}

function getStatusClass(status) {
  if (status === STATUS_FULL) {
    return "status-full";
  }
  if (status === STATUS_FULL_QC) {
    return "status-full-qc";
  }
  if (status === STATUS_PARTIAL) {
    return "status-partial";
  }
  return "status-none";
}

function getStandardComponentRemark({ storeAllocatedQty, qcAllocatedQty, shortQty, totalQcVisible }) {
  if (shortQty === 0 && qcAllocatedQty === 0) {
    return "Covered from Store";
  }
  if (shortQty === 0 && qcAllocatedQty > 0) {
    return "Store + QC together cover this item";
  }
  if (storeAllocatedQty > 0 || qcAllocatedQty > 0) {
    return totalQcVisible > 0 ? "Partly available across Store and QC" : "Partly available in Store only";
  }
  if (totalQcVisible > 0) {
    return "Visible in QC but still short overall";
  }
  return "Missing in Store and QC";
}

function getBestAnalysisRemark(component, hadTentativeAllocation) {
  if (hadTentativeAllocation) {
    return "Tentative stock released for later spools in Best Analysis";
  }
  if (component.storeBeforeQty > 0 || component.qcBeforeQty > 0) {
    return "Visible stock kept free for later spools in Best Analysis";
  }
  return "Missing in Store and QC";
}

function buildSpoolKey(projectCode, drawingNo, spoolNo) {
  return `${projectCode}@@${drawingNo}@@${spoolNo}`;
}

function buildStockKey(projectCode, itemCode) {
  return `${projectCode}@@${itemCode}`;
}

function buildHeaderIndexMap(columns) {
  return columns.reduce((indexMap, column, index) => {
    const normalized = normalizeHeaderName(column);
    if (normalized && !indexMap.has(normalized)) {
      indexMap.set(normalized, index);
    }
    return indexMap;
  }, new Map());
}

function getCellValue(row, headerIndexes, aliases, fallbackIndex) {
  const matchedIndex = findColumnIndex(headerIndexes, aliases);
  if (matchedIndex !== null) {
    return row[matchedIndex] ?? "";
  }
  return row[fallbackIndex] ?? "";
}

function findColumnIndex(headerIndexes, aliases) {
  for (const alias of aliases) {
    const normalized = normalizeHeaderName(alias);
    if (headerIndexes.has(normalized)) {
      return headerIndexes.get(normalized);
    }
  }
  return null;
}

function parseCsvMatrix(text) {
  const rows = [];
  let currentRow = [];
  let currentValue = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        currentValue += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === "," && !inQuotes) {
      currentRow.push(currentValue);
      currentValue = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      currentRow.push(currentValue);
      rows.push(currentRow);
      currentRow = [];
      currentValue = "";
      continue;
    }

    currentValue += character;
  }

  currentRow.push(currentValue);
  rows.push(currentRow);

  return rows.filter(isNonEmptyRow);
}

function normalizeHeaderName(value) {
  return `${value || ""}`
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function sanitizeHeaderLabel(value, index) {
  const cleaned = sanitizeText(index === 0 ? `${value || ""}`.replace(/^\ufeff/, "") : value);
  return cleaned || `Column ${index + 1}`;
}

function sanitizeCode(value) {
  return sanitizeText(value).toUpperCase();
}

function sanitizeText(value) {
  return `${value ?? ""}`.replace(/\s+/g, " ").trim();
}

function parseOptionalNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? roundQuantity(value) : null;
  }

  const cleaned = sanitizeText(value).replace(/,/g, "");
  if (!cleaned) {
    return null;
  }

  const numeric = Number.parseFloat(cleaned);
  return Number.isFinite(numeric) ? roundQuantity(numeric) : null;
}

function parseQuantity(value) {
  return parseOptionalNumber(value) ?? 0;
}

function normalizeLocation(value) {
  const normalized = normalizeHeaderName(value);
  if (normalized.includes("store")) {
    return "Store";
  }
  if (normalized.includes("qc")) {
    return "QC";
  }
  return "";
}

function countUniqueSpools(bomRows) {
  return new Set(bomRows.map((row) => buildSpoolKey(row.projectCode, row.drawingNo, row.spoolNo))).size;
}

function pickFirstNumber(values) {
  return values.find((value) => typeof value === "number" && Number.isFinite(value)) ?? null;
}

function buildMaterialLabel(masterMaterial, componentMaterials) {
  return collapseLabels([masterMaterial, ...componentMaterials]);
}

function collapseLabels(values) {
  const uniqueValues = [...new Set(values.map((value) => sanitizeText(value)).filter(Boolean))];
  return uniqueValues.join(", ");
}

function cloneStockMap(stockMap) {
  return new Map([...stockMap.entries()].map(([key, value]) => [key, cloneStockEntry(value)]));
}

function copyStockMap(target, source) {
  target.clear();
  source.forEach((value, key) => {
    target.set(key, cloneStockEntry(value));
  });
}

function cloneStockEntry(value) {
  return {
    store: roundQuantity(value?.store || 0),
    qc: roundQuantity(value?.qc || 0)
  };
}

function compareNullableNumbersAscending(left, right) {
  const leftMissing = left === null || left === undefined;
  const rightMissing = right === null || right === undefined;
  if (leftMissing && rightMissing) {
    return 0;
  }
  if (leftMissing) {
    return 1;
  }
  if (rightMissing) {
    return -1;
  }
  return left - right;
}

function compareNullableNumbersDescending(left, right) {
  const leftMissing = left === null || left === undefined;
  const rightMissing = right === null || right === undefined;
  if (leftMissing && rightMissing) {
    return 0;
  }
  if (leftMissing) {
    return 1;
  }
  if (rightMissing) {
    return -1;
  }
  return right - left;
}

function buildColumnWidths(rows) {
  const columnCount = rows[0]?.length || 0;
  return Array.from({ length: columnCount }, (_, columnIndex) => {
    const width = rows.reduce((maxWidth, row) => {
      const cellLength = `${row[columnIndex] ?? ""}`.length;
      return Math.max(maxWidth, Math.min(cellLength + 2, 42));
    }, 10);
    return { wch: width };
  });
}

function buildExportFileName(prefix) {
  const dateStamp = new Date().toISOString().slice(0, 10);
  return `${prefix}-${state.analysisMode}-${dateStamp}.xlsx`;
}

function formatDisplayNumber(value) {
  return value === null || value === undefined ? "-" : formatQuantity(value);
}

function getFileExtension(fileName) {
  return `${fileName}`.split(".").pop().toLowerCase();
}

function isNonEmptyRow(row) {
  return Array.isArray(row) && row.some((cell) => sanitizeText(cell) !== "");
}

function roundQuantity(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function formatQuantity(value) {
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: 3
  }).format(Number(value) || 0);
}

function setText(node, value) {
  node.textContent = `${value}`;
}

function escapeHtml(value) {
  return `${value || ""}`
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
