const BOM_COLUMN_ALIASES = {
  projectCode: ["project code", "projectcode", "project"],
  drawingNo: ["drawing no", "drawing no.", "drawing no#", "drawing number", "drawingno", "drawing no #"],
  spoolNo: ["spool no", "spool no.", "spool number", "spoolno"],
  itemCode: ["item code", "item code icd", "item code (icd)", "icd", "itemcode"],
  quantity: ["total qty", "total q'ty", "total q'ty", "qty", "quantity", "total quantity"]
};

const STOCK_COLUMN_ALIASES = {
  itemCode: ["item code", "item code icd", "item code (icd)", "icd", "itemcode"],
  location: ["location", "stock location", "store qc", "store/qc"],
  quantity: ["total qty", "total q'ty", "total q'ty", "qty", "quantity", "total quantity"]
};

const STATUS_FULL = "100% Material Available";
const STATUS_FULL_QC = "100% Material Available but some items are at QC location";
const STATUS_PARTIAL = "Partial Material Available";
const STATUS_NONE = "No Material Available";

const state = {
  bomRows: [],
  stockRows: [],
  bomFileName: "",
  stockFileName: "",
  bomColumns: [],
  stockColumns: [],
  analysis: null,
  statusFilter: "all",
  searchFilter: ""
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
  elements.runAnalysisButton.addEventListener("click", runAnalysis);
  elements.exportSummaryButton.addEventListener("click", exportSummaryWorkbook);
  elements.exportDetailButton.addEventListener("click", exportDetailWorkbook);
  elements.resetButton.addEventListener("click", resetApp);
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
    const rows = await parseImportedFile(file);
    const columns = Object.keys(rows[0] || {});

    if (type === "bom") {
      const parsedBomRows = buildBomRows(rows);
      state.bomRows = parsedBomRows;
      state.bomFileName = file.name;
      state.bomColumns = columns;
      if (!parsedBomRows.length) {
        setStatus(
          `Loaded BOM file ${file.name}, but no valid BOM rows were found. Please verify columns like Project Code, Drawing No., Spool No., Item Code (ICD), and Total Q'ty.`
        );
      } else {
        setStatus(`Loaded BOM file ${file.name} with ${state.bomRows.length} valid row(s).`);
      }
    } else {
      const parsedStockRows = buildStockRows(rows);
      state.stockRows = parsedStockRows;
      state.stockFileName = file.name;
      state.stockColumns = columns;
      if (!parsedStockRows.length) {
        setStatus(
          `Loaded inventory file ${file.name}, but no valid stock rows were found. Please verify columns like Item Code (ICD), Location, and Total Q'ty.`
        );
      } else {
        setStatus(`Loaded inventory file ${file.name} with ${state.stockRows.length} valid row(s).`);
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
  if (extension === "csv") {
    return parseCsvText(await file.text());
  }

  if (!window.XLSX) {
    throw new Error("Excel parsing is unavailable right now. Please save the file as CSV and upload it.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });
  const firstSheet = workbook.SheetNames[0];
  return window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
    defval: "",
    raw: true
  });
}

function buildBomRows(rows) {
  return rows
    .map((row, index) => {
      const projectCode = sanitizeCode(
        getValueByAliases(row, BOM_COLUMN_ALIASES.projectCode) || getValueByPosition(row, 0)
      );
      const drawingNo = sanitizeCode(
        getValueByAliases(row, BOM_COLUMN_ALIASES.drawingNo) || getValueByPosition(row, 1)
      );
      const spoolNo = sanitizeCode(
        getValueByAliases(row, BOM_COLUMN_ALIASES.spoolNo) || getValueByPosition(row, 2)
      );
      const itemCode = sanitizeCode(
        getValueByAliases(row, BOM_COLUMN_ALIASES.itemCode) || getValueByPosition(row, 3)
      );
      const quantity = parseQuantity(
        getValueByAliases(row, BOM_COLUMN_ALIASES.quantity) || getValueByPosition(row, 4)
      );

      if (!projectCode || !drawingNo || !spoolNo || !itemCode || quantity <= 0) {
        return null;
      }

      return {
        id: `bom-${index}`,
        order: index,
        projectCode,
        drawingNo,
        spoolNo,
        itemCode,
        quantity
      };
    })
    .filter(Boolean);
}

function buildStockRows(rows) {
  return rows
    .map((row, index) => {
      const itemCode = sanitizeCode(
        getValueByAliases(row, STOCK_COLUMN_ALIASES.itemCode) || getValueByPosition(row, 0)
      );
      const location = normalizeLocation(
        getValueByAliases(row, STOCK_COLUMN_ALIASES.location) || getValueByPosition(row, 1)
      );
      const quantity = parseQuantity(
        getValueByAliases(row, STOCK_COLUMN_ALIASES.quantity) || getValueByPosition(row, 2)
      );

      if (!itemCode || !location || quantity <= 0) {
        return null;
      }

      return {
        id: `stock-${index}`,
        itemCode,
        location,
        quantity
      };
    })
    .filter(Boolean);
}

function runAnalysis() {
  if (!state.bomRows.length || !state.stockRows.length) {
    setStatus("Please load both the BOM file and the inventory file before running the checker.");
    render();
    return;
  }

  state.analysis = analyzeAvailability(state.bomRows, state.stockRows);
  setStatus(
    `Analysis complete for ${state.analysis.spools.length} spool(s). Material is allocated from Store first, then QC.`
  );
  render();
}

function analyzeAvailability(bomRows, stockRows) {
  const groupedSpools = new Map();
  const stockTotals = buildStockTotals(stockRows);
  const remaining = new Map(
    [...stockTotals.entries()].map(([itemCode, totals]) => [
      itemCode,
      { store: totals.store, qc: totals.qc }
    ])
  );

  bomRows.forEach((row) => {
    const key = buildSpoolKey(row.projectCode, row.drawingNo, row.spoolNo);
    if (!groupedSpools.has(key)) {
      groupedSpools.set(key, {
        key,
        order: row.order,
        projectCode: row.projectCode,
        drawingNo: row.drawingNo,
        spoolNo: row.spoolNo,
        components: []
      });
    }

    groupedSpools.get(key).components.push({
      itemCode: row.itemCode,
      quantity: row.quantity
    });
  });

  const spools = [...groupedSpools.values()]
    .sort((left, right) => left.order - right.order)
    .map((spool) => allocateSpool(spool, remaining, stockTotals));

  return {
    spools,
    counts: {
      total: spools.length,
      fullStore: spools.filter((spool) => spool.status === STATUS_FULL).length,
      fullQc: spools.filter((spool) => spool.status === STATUS_FULL_QC).length,
      partial: spools.filter((spool) => spool.status === STATUS_PARTIAL).length,
      none: spools.filter((spool) => spool.status === STATUS_NONE).length
    }
  };
}

function allocateSpool(spool, remaining, stockTotals) {
  const mergedComponents = mergeSpoolComponents(spool.components);
  let totalRequiredQty = 0;
  let totalStoreAllocated = 0;
  let totalQcAllocated = 0;
  let totalShortQty = 0;
  let anyAllocated = false;
  let anyQcUsed = false;

  const components = mergedComponents.map((component) => {
    const pool = remaining.get(component.itemCode) || { store: 0, qc: 0 };
    const totals = stockTotals.get(component.itemCode) || { store: 0, qc: 0 };
    const storeBefore = pool.store;
    const qcBefore = pool.qc;
    const storeAllocated = Math.min(component.quantity, storeBefore);
    const afterStoreNeed = roundQuantity(component.quantity - storeAllocated);
    const qcAllocated = Math.min(afterStoreNeed, qcBefore);
    const shortQty = roundQuantity(component.quantity - storeAllocated - qcAllocated);

    pool.store = roundQuantity(storeBefore - storeAllocated);
    pool.qc = roundQuantity(qcBefore - qcAllocated);
    remaining.set(component.itemCode, pool);

    totalRequiredQty += component.quantity;
    totalStoreAllocated += storeAllocated;
    totalQcAllocated += qcAllocated;
    totalShortQty += shortQty;
    anyAllocated = anyAllocated || storeAllocated > 0 || qcAllocated > 0;
    anyQcUsed = anyQcUsed || qcAllocated > 0;

    return {
      itemCode: component.itemCode,
      requiredQty: component.quantity,
      storeBeforeQty: storeBefore,
      storeAllocatedQty: storeAllocated,
      qcBeforeQty: qcBefore,
      qcAllocatedQty: qcAllocated,
      shortQty,
      remarks: getComponentRemark(component.quantity, storeAllocated, qcAllocated, shortQty, totals.qc)
    };
  });

  let status = STATUS_NONE;
  if (roundQuantity(totalShortQty) === 0) {
    status = anyQcUsed ? STATUS_FULL_QC : STATUS_FULL;
  } else if (anyAllocated) {
    status = STATUS_PARTIAL;
  }

  return {
    key: spool.key,
    projectCode: spool.projectCode,
    drawingNo: spool.drawingNo,
    spoolNo: spool.spoolNo,
    status,
    componentCount: components.length,
    totalRequiredQty: roundQuantity(totalRequiredQty),
    totalStoreAllocated: roundQuantity(totalStoreAllocated),
    totalQcAllocated: roundQuantity(totalQcAllocated),
    totalShortQty: roundQuantity(totalShortQty),
    components
  };
}

function buildStockTotals(stockRows) {
  const totals = new Map();
  stockRows.forEach((row) => {
    if (!totals.has(row.itemCode)) {
      totals.set(row.itemCode, { store: 0, qc: 0 });
    }
    const itemTotals = totals.get(row.itemCode);
    if (row.location === "Store") {
      itemTotals.store = roundQuantity(itemTotals.store + row.quantity);
    }
    if (row.location === "QC") {
      itemTotals.qc = roundQuantity(itemTotals.qc + row.quantity);
    }
  });
  return totals;
}

function mergeSpoolComponents(components) {
  const merged = new Map();
  components.forEach((component) => {
    merged.set(component.itemCode, roundQuantity((merged.get(component.itemCode) || 0) + component.quantity));
  });
  return [...merged.entries()].map(([itemCode, quantity]) => ({ itemCode, quantity }));
}

function getComponentRemark(requiredQty, storeAllocated, qcAllocated, shortQty, totalQcVisible) {
  if (shortQty === 0 && qcAllocated === 0) {
    return "Covered from Store";
  }
  if (shortQty === 0 && qcAllocated > 0) {
    return "Store + QC together cover this item";
  }
  if (storeAllocated > 0 || qcAllocated > 0) {
    return totalQcVisible > 0
      ? "Partly available across Store and QC"
      : "Partly available in Store only";
  }
  if (totalQcVisible > 0) {
    return "Visible in QC but still short overall";
  }
  return "Missing in Store and QC";
}

function render() {
  renderFileInfo();
  renderSummaryCards();
  renderSummaryTable();
  renderDetailTable();
  const hasVisibleSpools = getFilteredSpools().length > 0;
  elements.exportSummaryButton.disabled = !hasVisibleSpools;
  elements.exportDetailButton.disabled = !hasVisibleSpools;
}

function renderFileInfo() {
  elements.bomFileInfo.textContent = state.bomFileName
    ? `${state.bomFileName} | ${state.bomRows.length} valid BOM row(s)`
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
      <tr><td colspan="9"><div class="empty-state">No result yet. Upload files and run the analysis.</div></td></tr>
    `;
    return;
  }

  const filteredSpools = getFilteredSpools();
  elements.summaryNote.textContent = `${filteredSpools.length} spool(s) shown from ${state.analysis.spools.length} analyzed spool(s).`;

  if (!filteredSpools.length) {
    elements.summaryTableBody.innerHTML = `
      <tr><td colspan="9"><div class="empty-state">No spool matches the current filters.</div></td></tr>
    `;
    return;
  }

  elements.summaryTableBody.innerHTML = filteredSpools
    .map((spool) => `
      <tr>
        <td>${escapeHtml(spool.projectCode)}</td>
        <td>${escapeHtml(spool.drawingNo)}</td>
        <td>${escapeHtml(spool.spoolNo)}</td>
        <td>${renderStatusChip(spool.status)}</td>
        <td class="is-number">${formatQuantity(spool.componentCount)}</td>
        <td class="is-number">${formatQuantity(spool.totalRequiredQty)}</td>
        <td class="is-number">${formatQuantity(spool.totalStoreAllocated)}</td>
        <td class="is-number">${formatQuantity(spool.totalQcAllocated)}</td>
        <td class="is-number">${formatQuantity(spool.totalShortQty)}</td>
      </tr>
    `)
    .join("");
}

function renderDetailTable() {
  if (!state.analysis) {
    elements.detailTitle.textContent = "Component detail for all shown spools";
    elements.detailNote.textContent = "Store-first allocation with QC fallback is shown component by component.";
    elements.detailTableBody.innerHTML = `
      <tr><td colspan="12"><div class="empty-state">No result yet. Upload files and run the analysis.</div></td></tr>
    `;
    return;
  }

  const visibleSpools = getFilteredSpools();
  const componentRows = getVisibleComponentRows(visibleSpools);

  if (!visibleSpools.length) {
    elements.detailTitle.textContent = "Component detail for all shown spools";
    elements.detailNote.textContent = "No component rows match the current filters.";
    elements.detailTableBody.innerHTML = `
      <tr><td colspan="12"><div class="empty-state">No component detail matches the current filters.</div></td></tr>
    `;
    return;
  }

  elements.detailTitle.textContent = "Component detail for all shown spools";
  elements.detailNote.textContent = `${componentRows.length} component row(s) shown across ${visibleSpools.length} spool(s).`;
  elements.detailTableBody.innerHTML = componentRows
    .map((component) => `
      <tr>
        <td>${escapeHtml(component.projectCode)}</td>
        <td>${escapeHtml(component.drawingNo)}</td>
        <td>${escapeHtml(component.spoolNo)}</td>
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
    `)
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
      `${spool.projectCode} ${spool.drawingNo} ${spool.spoolNo} ${spool.status} ${spool.components.map((component) => component.itemCode).join(" ")}`
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
    "fabrication-readiness-by-spool.xlsx",
    "Fabrication readiness",
    [
      [
        "Project Code",
        "Drawing No.",
        "Spool No.",
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
    "component-detail.xlsx",
    "Component detail",
    [
      [
        "Project Code",
        "Drawing No.",
        "Spool No.",
        "Status",
        "ICD",
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
  window.XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  window.XLSX.writeFile(workbook, fileName);
}

function resetApp() {
  state.bomRows = [];
  state.stockRows = [];
  state.bomFileName = "";
  state.stockFileName = "";
  state.bomColumns = [];
  state.stockColumns = [];
  state.analysis = null;
  state.statusFilter = "all";
  state.searchFilter = "";
  elements.statusFilter.value = "all";
  elements.searchFilter.value = "";
  setStatus("Upload both files, then run the checker.");
  render();
}

function loadDemoData() {
  state.bomRows = [
    { id: "bom-1", order: 0, projectCode: "PRJ-01", drawingNo: "DWG-100", spoolNo: "SP-001", itemCode: "ICD-PIPE-8", quantity: 8 },
    { id: "bom-2", order: 1, projectCode: "PRJ-01", drawingNo: "DWG-100", spoolNo: "SP-001", itemCode: "ICD-ELB-8", quantity: 2 },
    { id: "bom-3", order: 2, projectCode: "PRJ-01", drawingNo: "DWG-101", spoolNo: "SP-002", itemCode: "ICD-PIPE-8", quantity: 5 },
    { id: "bom-4", order: 3, projectCode: "PRJ-01", drawingNo: "DWG-101", spoolNo: "SP-002", itemCode: "ICD-FLG-8", quantity: 2 },
    { id: "bom-5", order: 4, projectCode: "PRJ-02", drawingNo: "DWG-220", spoolNo: "SP-010", itemCode: "ICD-TEE-6", quantity: 1 },
    { id: "bom-6", order: 5, projectCode: "PRJ-02", drawingNo: "DWG-220", spoolNo: "SP-010", itemCode: "ICD-PIPE-6", quantity: 4 }
  ];
  state.stockRows = [
    { id: "stock-1", itemCode: "ICD-PIPE-8", location: "Store", quantity: 10 },
    { id: "stock-2", itemCode: "ICD-ELB-8", location: "Store", quantity: 2 },
    { id: "stock-3", itemCode: "ICD-FLG-8", location: "Store", quantity: 1 },
    { id: "stock-4", itemCode: "ICD-FLG-8", location: "QC", quantity: 1 },
    { id: "stock-5", itemCode: "ICD-TEE-6", location: "QC", quantity: 1 },
    { id: "stock-6", itemCode: "ICD-PIPE-6", location: "Store", quantity: 0.5 },
    { id: "stock-7", itemCode: "ICD-PIPE-6", location: "QC", quantity: 2 }
  ];
  state.bomFileName = "demo-bom.xlsx";
  state.stockFileName = "demo-stock.xlsx";
  state.bomColumns = ["Project Code", "Drawing No.", "Spool No.", "Item Code (ICD)", "Total Q'ty"];
  state.stockColumns = ["Item Code (ICD)", "Location", "Total Q'ty"];
  state.analysis = analyzeAvailability(state.bomRows, state.stockRows);
  setStatus("Demo data loaded. Review the summary and detail tables.");
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

function buildSpoolKey(projectCode, drawingNo, spoolNo) {
  return `${projectCode}@@${drawingNo}@@${spoolNo}`;
}

function getValueByAliases(row, aliases) {
  const normalizedAliases = aliases.map((alias) => normalizeHeaderName(alias));
  for (const [key, value] of Object.entries(row || {})) {
    if (normalizedAliases.includes(normalizeHeaderName(key))) {
      return value;
    }
  }
  return "";
}

function getValueByPosition(row, position) {
  const values = Object.values(row || {});
  return values[position] ?? "";
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

function getFileExtension(fileName) {
  return `${fileName}`.split(".").pop().toLowerCase();
}

function parseCsvText(text) {
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

  const filteredRows = rows.filter((row) => row.some((cell) => `${cell}`.trim() !== ""));
  if (!filteredRows.length) {
    return [];
  }

  const headers = filteredRows[0].map((header, index) => {
    const cleaned = `${header}`.trim();
    return index === 0 ? cleaned.replace(/^\ufeff/, "") : cleaned;
  });

  return filteredRows.slice(1).map((row) => {
    const rowObject = {};
    headers.forEach((header, index) => {
      rowObject[header] = row[index] ?? "";
    });
    return rowObject;
  });
}

function normalizeHeaderName(value) {
  return `${value || ""}`
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function sanitizeCode(value) {
  return `${value || ""}`.trim().toUpperCase();
}

function parseQuantity(value) {
  if (typeof value === "number") {
    return roundQuantity(value);
  }
  const numeric = Number.parseFloat(`${value}`.replace(/,/g, ""));
  return Number.isFinite(numeric) ? roundQuantity(numeric) : 0;
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
