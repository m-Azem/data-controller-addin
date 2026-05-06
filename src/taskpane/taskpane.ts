/* global Excel, Office, OfficeRuntime */

import { idbGet, idbSet } from "../utils/db";
import { migrateFromLocalStorage } from "../services/migration";
import { exportCSVData, downloadBackup, processRestoreFile } from "../services/fileService";
import { executeInsertTable, executeInsertDropdown, executeInsertFormula, executeConvertToValues, executeRefreshFormulas } from "../services/excelService";

const APP_VERSION = "1.0.1"; // Update this number whenever you release a new version

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("capture-button").onclick = loadRangeForCapture;
    document.getElementById("finish-capture-button").onclick = finishCapture;
    document.getElementById("backup-button").onclick = backupData;
    document.getElementById("restore-button").onclick = triggerRestore;
    document.getElementById("restore-file-input").addEventListener("change", restoreData);
    document.getElementById("refresh-formulas-button").onclick = () => refreshFormulas(false);
    document.getElementById("convert-values-button").onclick = convertToValues;
    
    // Setup Tabs
    document.querySelectorAll('.tab-link').forEach(link => {
      link.addEventListener('click', switchTab);
    });

    // Setup Theme Toggle
    const themeToggle = document.getElementById("theme-toggle");
    if (themeToggle) themeToggle.addEventListener("change", toggleTheme);

    // Display App Version
    const versionDisplay = document.getElementById("app-version-display");
    if (versionDisplay) versionDisplay.innerText = APP_VERSION;

    await migrateFromLocalStorage(); // Migrate old data if present
    loadSettings();
    renderDashboard();
  }
});

export function switchTab(event: any) {
  document.querySelectorAll('.tab-link').forEach(link => link.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

  const target = event.currentTarget as HTMLElement;
  target.classList.add('active');
  const tabId = target.getAttribute('data-tab');
  if (tabId) {
    document.getElementById(tabId)?.classList.add('active');
  }
}

export async function toggleTheme(event: any) {
  const isDark = event.target.checked;
  if (isDark) document.body.classList.add('theme-dark');
  else document.body.classList.remove('theme-dark');
  await idbSet("DC_THEME", isDark ? "dark" : "light");
}

export async function loadSettings() {
  const theme = await idbGet("DC_THEME");
  const toggle = document.getElementById("theme-toggle") as HTMLInputElement;
  if (theme === "dark") {
    document.body.classList.add("theme-dark");
    if (toggle) toggle.checked = true;
  }
}

let pendingCaptureData: { name: string, headers: string[], dataRows: any[][] } | null = null;

export async function loadRangeForCapture() {
  const status = document.getElementById("status-text");
  try {
    await Excel.run(async (context) => {
      const dataName = (document.getElementById("data-name") as HTMLInputElement).value;

      if (!dataName || dataName.trim() === "") {
        throw new Error("Please enter a data table name.");
      }

      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      if (values.length < 2) {
        throw new Error("Select a range with headers and data.");
      }

      pendingCaptureData = {
        name: dataName,
        headers: values[0],
        dataRows: values.slice(1)
      };

      const idSelect = document.getElementById("id-column-select") as HTMLSelectElement;
      idSelect.innerHTML = "";
      pendingCaptureData.headers.forEach((header, index) => {
         const opt = document.createElement("option");
         opt.value = String(index);
         opt.text = header || `Column ${index + 1}`;
         idSelect.appendChild(opt);
      });

      const step2Div = document.getElementById("capture-step-2");
      if (step2Div) step2Div.style.display = "block";
      
      if (status) {
        status.innerText = "Range loaded. Please select the ID column and click Finish.";
        status.style.color = "blue";
      }
    });
  } catch (error) {
    if (status) {
      status.innerText = "Error: " + error.message;
      status.style.color = "red";
    }
    const step2Div = document.getElementById("capture-step-2");
    if (step2Div) step2Div.style.display = "none";
    pendingCaptureData = null;
  }
}

export async function finishCapture() {
  const status = document.getElementById("status-text");
  try {
    if (!pendingCaptureData) {
      throw new Error("No pending data to capture. Please load the range again.");
    }

    const idSelect = document.getElementById("id-column-select") as HTMLSelectElement;
    const idIndex = parseInt(idSelect.value, 10);
    
    const originalHeaders = pendingCaptureData.headers;
    const dataRows = pendingCaptureData.dataRows;
    const dataName = pendingCaptureData.name;

    if (isNaN(idIndex) || idIndex < 0 || idIndex >= originalHeaders.length) {
      throw new Error("Invalid ID Column selected.");
    }

    const idColumn = originalHeaders[idIndex];
    if (!idColumn || String(idColumn).trim() === "") {
      throw new Error("The selected ID column must have a valid header.");
    }

    // Reorder fields
    const fields = [...originalHeaders];
    fields.splice(idIndex, 1);
    fields.unshift(idColumn);

    let store: any = {};
    const existingStore = await idbGet("DC_STORE");
    if (existingStore) {
      store = JSON.parse(existingStore);
    }

    let rev = 1;
    let history = {};
    if (store[dataName]) {
      rev = store[dataName].revision || 1;
      history = store[dataName].history || {};
    }

    const dataSet = {
      dataTableName: dataName,
      revision: rev,
      history: history,
      fields: fields,
      records: dataRows.map((row: any[]) => {
        let record: any = {};
        originalHeaders.forEach((header, index) => {
          record[header] = row[index];
        });
        return record;
      })
    };

    store[dataName] = dataSet;
    await idbSet("DC_STORE", JSON.stringify(store));
    
    if (status) {
      status.innerText = `Saved ${dataRows.length} records in ${dataName}. (ID: "${idColumn}")`;
      status.style.color = "green";
    }

    // Reset UI
    const step2Div = document.getElementById("capture-step-2");
    if (step2Div) step2Div.style.display = "none";
    (document.getElementById("data-name") as HTMLInputElement).value = "";
    pendingCaptureData = null;

    renderDashboard();
    refreshFormulas(true);
  } catch (error) {
    if (status) {
      status.innerText = "Error: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function replaceTableData(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      if (values.length < 2) {
        throw new Error("Select a range with headers and data.");
      }

      const headers = values[0];
      const dataRows = values.slice(1);

      const idColumn = headers[0];
      if (!idColumn || String(idColumn).trim() === "") {
        throw new Error("The first column must have a valid header to serve as the Record ID.");
      }

      let store: any = {};
      const existingStore = await idbGet("DC_STORE");
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(`Data table '${dataTableName}' not found.`);
      }

      const dataSet = store[dataTableName];
      const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
      const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
      const isLatest = selectedRev === (dataSet.revision || 1);

      let targetDataSet = dataSet;
      if (!isLatest) {
        if (!dataSet.history) dataSet.history = {};
        if (!dataSet.history[selectedRev]) dataSet.history[selectedRev] = {};
        targetDataSet = dataSet.history[selectedRev];
      }

      targetDataSet.fields = headers;
      targetDataSet.records = dataRows.map((row: any[]) => {
        let record: any = {};
        headers.forEach((header, index) => {
          record[header] = row[index];
        });
        return record;
      });

      await idbSet("DC_STORE", JSON.stringify(store));
      
      if (status) {
        status.innerText = isLatest ? `Replaced '${dataTableName}' with ${dataRows.length} new records. (ID: "${idColumn}")` : `Replaced Rev ${selectedRev} of '${dataTableName}' with ${dataRows.length} records.`;
        status.style.color = "green";
      }
      renderDashboard();
      refreshFormulas(true);
    });
  } catch (error) {
    if (status) {
      status.innerText = "Error replacing data: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function renderDashboard() {
  const list = document.getElementById("data-table-list");
  if (!list) return;
  
  list.innerHTML = "";
  const storedData = await idbGet("DC_STORE");
  const defaultSelect = document.getElementById("default-data-table-select") as HTMLSelectElement;

  if (!storedData) {
    list.innerHTML = "<li>No data tables stored yet.</li>";
    if (defaultSelect) defaultSelect.innerHTML = "";
    return;
  }

  let store = JSON.parse(storedData);
  let needsSave = false;

  // Cleanup: migrate old single-data-table format if present
  if (store.collectionName && typeof store.collectionName === "string") {
    const oldName = store.collectionName;
    if (!store[oldName]) {
      store[oldName] = {
        dataTableName: store.collectionName,
        fields: store.fields,
        records: store.records
      };
    }
    delete store.collectionName;
    delete store.entityName;
    delete store.fields;
    delete store.records;
    needsSave = true;
  }

  // Remove any remaining invalid keys
  for (const key of Object.keys(store)) {
    if (!store[key] || typeof store[key] !== "object" || Array.isArray(store[key])) {
      delete store[key];
      needsSave = true;
    }
  }

  if (needsSave) {
    await idbSet("DC_STORE", JSON.stringify(store));
  }

  const keys = Object.keys(store);
  if (keys.length === 0) {
    list.innerHTML = "<li>No data tables stored yet.</li>";
    if (defaultSelect) defaultSelect.innerHTML = "";
    return;
  }

  // Setup Default Data Table dropdown
  if (defaultSelect) {
    const defaultRevSelect = document.getElementById("default-revision-select") as HTMLSelectElement;
    defaultSelect.innerHTML = "";

    keys.forEach(key => {
      const opt = document.createElement("option");
      opt.value = key;
      opt.text = key;
      defaultSelect.appendChild(opt);
    });
    
    let currentDefault = await idbGet("DC_DEFAULT_DATA_TABLE");
    if (!currentDefault || !keys.includes(currentDefault)) {
      currentDefault = keys[0];
      await idbSet("DC_DEFAULT_DATA_TABLE", currentDefault);
    }
    defaultSelect.value = currentDefault;
    
    const updateDefaultRevOptions = async (tableName: string) => {
      if (!defaultRevSelect) return;
      defaultRevSelect.innerHTML = "";
      const tData = store[tableName];
      if (tData) {
        const maxRev = tData.revision || 1;
        for (let i = maxRev; i >= 1; i--) {
          if (i === maxRev || (tData.history && tData.history[i])) {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.text = `Rev ${i}${i === maxRev ? " (Latest)" : ""}`;
            defaultRevSelect.appendChild(opt);
          }
        }
      }
      let currentDefaultRev = await idbGet("DC_DEFAULT_REVISION");
      if (currentDefaultRev && Array.from(defaultRevSelect.options).some(o => o.value === currentDefaultRev)) {
        defaultRevSelect.value = currentDefaultRev;
      } else {
        defaultRevSelect.selectedIndex = 0;
        await idbSet("DC_DEFAULT_REVISION", defaultRevSelect.value);
      }
    };

    await updateDefaultRevOptions(currentDefault);

    defaultSelect.onchange = async () => {
      await idbSet("DC_DEFAULT_DATA_TABLE", defaultSelect.value);
      await updateDefaultRevOptions(defaultSelect.value);
    };

    if (defaultRevSelect) {
      defaultRevSelect.onchange = async () => {
        await idbSet("DC_DEFAULT_REVISION", defaultRevSelect.value);
      };
    }
  }

  keys.forEach(key => {
    const dataTable = store[key];
    const count = dataTable.records ? dataTable.records.length : 0;
    const rev = dataTable.revision || 1;
    
    const li = document.createElement("li");
    li.className = "data-table-item";
    
    // Header (Clickable to expand/collapse)
    const header = document.createElement("div");
    header.className = "data-table-header";
    header.innerHTML = `<div><span style="display:inline-block; width:15px;">&#9654;</span> <strong>${key}</strong></div> <span style="font-weight: normal; font-size: 12px; color: #666;">(Rev ${rev} | ${count} records)</span>`;
    
    // Details container (Hidden by default)
    const details = document.createElement("div");
    details.className = "data-table-details";
    details.style.display = "none";
    details.style.marginTop = "16px";
    
    header.onclick = () => {
      const isHidden = details.style.display === "none";
      details.style.display = isHidden ? "block" : "none";
      header.innerHTML = `<div><span style="display:inline-block; width:15px;">${isHidden ? '&#9660;' : '&#9654;'}</span> <strong>${key}</strong></div> <span style="font-weight: normal; font-size: 12px; color: #666;">(Rev ${rev} | ${count} records)</span>`;
    };

    const insertTableBtn = document.createElement("button");
    insertTableBtn.innerHTML = '<i class="ms-Icon ms-Icon--Table" style="margin-right:8px;"></i> Insert Table';
    insertTableBtn.className = "ms-Button ms-Button--primary";
    insertTableBtn.style.width = "100%";
    insertTableBtn.onclick = () => insertTable(key);
    
    const replaceBtn = document.createElement("button");
    replaceBtn.innerHTML = '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Data';
    replaceBtn.className = "ms-Button ms-Button--default";
    replaceBtn.style.width = "100%";
    replaceBtn.onclick = () => replaceTableData(key);

    const snapshotBtn = document.createElement("button");
    snapshotBtn.innerHTML = '<i class="ms-Icon ms-Icon--Camera" style="margin-right:8px;"></i> Snapshot';
    snapshotBtn.className = "ms-Button ms-Button--default";
    snapshotBtn.style.width = "100%";
    snapshotBtn.onclick = () => createSnapshot(key);

    const clearHistoryBtn = document.createElement("button");
    clearHistoryBtn.innerHTML = '<i class="ms-Icon ms-Icon--Broom" style="margin-right:8px;"></i> Clear History';
    clearHistoryBtn.className = "ms-Button ms-Button--default";
    clearHistoryBtn.style.width = "100%";
    clearHistoryBtn.onclick = () => clearTableHistory(key);

    const deleteBtn = document.createElement("button");
    deleteBtn.innerHTML = '<i class="ms-Icon ms-Icon--Delete" style="margin-right:8px;"></i> Delete Table';
    deleteBtn.className = "ms-Button ms-Button--default";
    deleteBtn.style.color = "red";
    deleteBtn.style.width = "100%";
    deleteBtn.onclick = () => deleteDataTable(key);
    
    // 0. Target Revision (Top of details)
    const revSelectContainer = document.createElement("div");
    revSelectContainer.style.marginBottom = "12px";
    revSelectContainer.innerHTML = `<label style="font-size:12px; font-weight:600; color:var(--text-color); margin-right:8px;">Target Revision:</label>`;

    const revSelect = document.createElement("select");
    revSelect.id = `fb-rev-${key}`;
    revSelect.className = "ms-Dropdown-title";
    revSelect.style.width = "auto";
    revSelect.style.display = "inline-block";
    for (let i = rev; i >= 1; i--) {
      if (i === rev || (dataTable.history && dataTable.history[i])) {
        const opt = document.createElement("option");
        opt.value = String(i);
        opt.text = `Revision ${i}${i === rev ? " (Latest)" : ""}`;
        revSelect.appendChild(opt);
      }
    }
    revSelectContainer.appendChild(revSelect);
    details.appendChild(revSelectContainer);

    const idInputContainer = document.createElement("div");
    idInputContainer.style.position = "relative";
    idInputContainer.style.marginBottom = "5px";

    const idInput = document.createElement("input");
    idInput.id = `fb-id-${key}`;
    idInput.type = "text";
    idInput.placeholder = "Search for Record ID...";
    idInput.className = "ms-TextField-field";
    idInput.autocomplete = "off";
    
    const idResults = document.createElement("div");
    idResults.className = "search-results-dropdown";
    idResults.style.display = "none";

    idInput.oninput = () => {
      const searchTerm = idInput.value.toLowerCase();
      idResults.innerHTML = "";
      if (!searchTerm) {
        idResults.style.display = "none";
        return;
      }

      const selectedRev = parseInt(revSelect.value);
      let targetDataSet = dataTable;
      if (selectedRev !== rev && dataTable.history && dataTable.history[selectedRev]) {
        targetDataSet = dataTable.history[selectedRev];
      }

      if (targetDataSet.records && targetDataSet.fields && targetDataSet.fields.length > 0) {
        const idField = targetDataSet.fields[0];
        const matches = targetDataSet.records.filter((record: any) => 
          record[idField] !== undefined && String(record[idField]).toLowerCase().includes(searchTerm)
        ).slice(0, 50);

        if (matches.length > 0) {
          matches.forEach((record: any) => {
            const resultItem = document.createElement("div");
            resultItem.className = "search-result-item";
            resultItem.innerText = String(record[idField]);
            resultItem.onmousedown = () => {
              idInput.value = String(record[idField]);
              idResults.style.display = "none";
            };
            idResults.appendChild(resultItem);
          });
          idResults.style.display = "block";
        } else {
          idResults.style.display = "none";
        }
      }
    };

    idInput.onblur = () => { setTimeout(() => { idResults.style.display = 'none'; }, 150); };
    idInput.onfocus = () => { if (idInput.value) { idInput.dispatchEvent(new Event('input')); } };

    idInputContainer.appendChild(idInput);
    
    const fieldSelect = document.createElement("select");
    fieldSelect.id = `fb-field-${key}`;
    fieldSelect.className = "ms-Dropdown-title";
    fieldSelect.style.marginBottom = "8px";
    
    const populateFields = () => {
      const selectedRev = parseInt(revSelect.value);
      let targetDataSet = dataTable;
      if (selectedRev !== rev && dataTable.history && dataTable.history[selectedRev]) {
        targetDataSet = dataTable.history[selectedRev];
      }
      fieldSelect.innerHTML = "";
      if (targetDataSet.fields) {
        targetDataSet.fields.forEach((field: string) => {
          const opt = document.createElement("option");
          opt.value = field;
          opt.text = field;
          fieldSelect.appendChild(opt);
        });
      }
    };
    populateFields();
    
    revSelect.onchange = () => {
      populateFields();
      idInput.value = "";
      idResults.style.display = "none";
      
      const isLatest = parseInt(revSelect.value) === rev;
      replaceBtn.innerHTML = isLatest ? '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Data' : '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Rev Data';
      snapshotBtn.innerHTML = isLatest ? '<i class="ms-Icon ms-Icon--Camera" style="margin-right:8px;"></i> Snapshot' : '<i class="ms-Icon ms-Icon--Undo" style="margin-right:8px;"></i> Restore as Active';
      deleteBtn.innerHTML = isLatest ? '<i class="ms-Icon ms-Icon--Delete" style="margin-right:8px;"></i> Delete Table' : '<i class="ms-Icon ms-Icon--Delete" style="margin-right:8px;"></i> Delete Revision';
    };
    
    const insertFormulaBtn = document.createElement("button");
    insertFormulaBtn.innerHTML = '<i class="ms-Icon ms-Icon--Calculator" style="margin-right:8px;"></i> Insert DC.GET Formula';
    insertFormulaBtn.className = "ms-Button ms-Button--primary";
    insertFormulaBtn.style.width = "100%";
    insertFormulaBtn.onclick = () => insertFormulaFor(key);
    
    const insertDropdownBtn = document.createElement("button");
    insertDropdownBtn.innerHTML = '<i class="ms-Icon ms-Icon--Dropdown" style="margin-right:8px;"></i> Insert Headers Dropdown';
    insertDropdownBtn.className = "ms-Button ms-Button--default";
    insertDropdownBtn.style.width = "100%";
    insertDropdownBtn.onclick = () => insertDropdown(key);

    const editBtn = document.createElement("button");
    editBtn.innerHTML = '<i class="ms-Icon ms-Icon--Edit" style="margin-right:8px;"></i> Edit Record';
    editBtn.className = "ms-Button ms-Button--default";
    editBtn.style.width = "100%";
    editBtn.onclick = () => loadRecordForEdit(key, idInput.value);

    const exportCSVBtn = document.createElement("button");
    exportCSVBtn.innerHTML = '<i class="ms-Icon ms-Icon--Download" style="margin-right:8px;"></i> Export CSV';
    exportCSVBtn.className = "ms-Button ms-Button--default";
    exportCSVBtn.style.width = "100%";
    exportCSVBtn.onclick = () => exportCSV(key);

    // 1. EXCEL INTEGRATION GROUP
    const excelGroup = document.createElement("div");
    excelGroup.className = "action-group";
    excelGroup.innerHTML = `<div class="action-group-title">Excel Integration</div>`;
    excelGroup.appendChild(idInputContainer);
    excelGroup.appendChild(fieldSelect);
    
    const excelBtns = document.createElement("div");
    excelBtns.className = "button-group";
    excelBtns.style.display = "flex";
    excelBtns.style.flexDirection = "column";
    excelBtns.style.gap = "8px";
    excelBtns.appendChild(insertFormulaBtn);
    excelBtns.appendChild(insertDropdownBtn);
    excelBtns.appendChild(insertTableBtn);
    excelGroup.appendChild(excelBtns);
    details.appendChild(excelGroup);

    // 2. DATA MANAGEMENT GROUP
    const dataGroup = document.createElement("div");
    dataGroup.className = "action-group";
    dataGroup.innerHTML = `<div class="action-group-title">Data Management</div>`;
    const dataBtns = document.createElement("div");
    dataBtns.className = "button-group";
    dataBtns.style.display = "flex";
    dataBtns.style.flexDirection = "column";
    dataBtns.style.gap = "8px";
    dataBtns.appendChild(replaceBtn);
    dataBtns.appendChild(editBtn);
    dataBtns.appendChild(exportCSVBtn);
    dataGroup.appendChild(dataBtns);
    details.appendChild(dataGroup);

    // 3. VERSION CONTROL GROUP
    const versionGroup = document.createElement("div");
    versionGroup.className = "action-group";
    versionGroup.innerHTML = `<div class="action-group-title">Version Control</div>`;
    const versionBtns = document.createElement("div");
    versionBtns.className = "button-group";
    versionBtns.style.display = "flex";
    versionBtns.style.flexDirection = "column";
    versionBtns.style.gap = "8px";
    versionBtns.appendChild(snapshotBtn);
    versionBtns.appendChild(clearHistoryBtn);
    versionBtns.appendChild(deleteBtn);
    versionGroup.appendChild(versionBtns);
    details.appendChild(versionGroup);
    
    const editFormDiv = document.createElement("div");
    editFormDiv.id = `edit-form-container-${key}`;
    editFormDiv.style.display = "none";
    details.appendChild(editFormDiv);

    li.appendChild(header);
    li.appendChild(details);
    list.appendChild(li);
  });
}

export async function deleteDataTable(dataTableName: string) {
  const storedData = await idbGet("DC_STORE");
  if (storedData) {
    let store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) return;
    
    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    const isLatest = selectedRev === (dataSet.revision || 1);

    if (isLatest) {
      delete store[dataTableName];
    } else {
      if (dataSet.history && dataSet.history[selectedRev]) {
        delete dataSet.history[selectedRev];
      }
    }

    await idbSet("DC_STORE", JSON.stringify(store));
    renderDashboard();
    refreshFormulas(true);
    
    const status = document.getElementById("status-text");
    if (status) {
        status.innerText = isLatest ? `Deleted data table: ${dataTableName}` : `Deleted Revision ${selectedRev} of ${dataTableName}`;
        status.style.color = "blue";
    }
  }
}

export async function clearTableHistory(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet("DC_STORE");
    if (storedData) {
      let store = JSON.parse(storedData);
      if (store[dataTableName]) {
        store[dataTableName].history = {};
        await idbSet("DC_STORE", JSON.stringify(store));
        renderDashboard();
        if (status) {
          status.innerText = `Cleared old revisions for ${dataTableName}. Space reclaimed.`;
          status.style.color = "green";
        }
      }
    }
  } catch (error) {
    if (status) {
      status.innerText = "Error clearing history: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function createSnapshot(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    let store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) return;

    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    const isLatest = selectedRev === (dataSet.revision || 1);

    const currentRev = dataSet.revision || 1;
    dataSet.history = dataSet.history || {};

    if (isLatest) {
      // Deep copy the current records to history
      dataSet.history[currentRev] = {
        fields: dataSet.fields,
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      
      // Increment revision
      dataSet.revision = currentRev + 1;

      if (status) {
          status.innerText = `Locked Rev ${currentRev}. Current is now Rev ${dataSet.revision}`;
          status.style.color = "green";
      }
    } else {
      // Restore old revision as current
      dataSet.history[currentRev] = {
        fields: dataSet.fields,
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      dataSet.fields = JSON.parse(JSON.stringify(dataSet.history[selectedRev].fields));
      dataSet.records = JSON.parse(JSON.stringify(dataSet.history[selectedRev].records));
      dataSet.revision = currentRev + 1;

      if (status) {
          status.innerText = `Restored Rev ${selectedRev} as new active Rev ${dataSet.revision}.`;
          status.style.color = "green";
      }
    }

    await idbSet("DC_STORE", JSON.stringify(store));
    renderDashboard();
    if (!isLatest) refreshFormulas(true);
  } catch (error) {
    if (status) {
      if (error.name === "QuotaExceededError" || (error.message && error.message.includes("exceeded the quota"))) {
        status.innerText = "Storage limit reached! Please use 'Clear History' or delete tables to free up space.";
      } else {
        status.innerText = "Error creating snapshot: " + error.message;
      }
      status.style.color = "red";
    }
  }
}

export async function loadRecordForEdit(dataTableName: string, id: string) {
  const status = document.getElementById("status-text");
  if (!id) {
      if (status) { status.innerText = "Please select a Record ID to edit."; status.style.color = "red"; }
      return;
  }

  const storedData = await idbGet("DC_STORE");
  if (!storedData) return;
  const store = JSON.parse(storedData);
  const dataSet = store[dataTableName];
  if (!dataSet) return;

  const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
  const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
  
  let targetDataSet = dataSet;
  if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
  }

  const idField = targetDataSet.fields[0];
  const record = targetDataSet.records.find((r: any) => String(r[idField]) === String(id));

  if (!record) {
      if (status) { status.innerText = "Record not found."; status.style.color = "red"; }
      return;
  }

  const container = document.getElementById(`edit-form-container-${dataTableName}`);
  if (!container) return;

  container.innerHTML = `<h4 style="margin-top:0; color: var(--primary-color);">Editing: ${id}</h4>`;
  container.className = "edit-form-container";
  container.style.display = "block";

  targetDataSet.fields.forEach((field: string) => {
      const fieldDiv = document.createElement("div");
      fieldDiv.className = "ms-TextField";
      fieldDiv.style.marginBottom = "8px";

      const label = document.createElement("label");
      label.className = "ms-Label";
      label.innerText = field;

      const input = document.createElement("input");
      input.type = "text";
      input.value = record[field] !== undefined ? record[field] : "";
      input.id = `edit-input-${dataTableName}-${field}`;
      input.className = "ms-TextField-field";
      if (field === idField) input.disabled = true; // Prevent changing the ID

      fieldDiv.appendChild(label);
      fieldDiv.appendChild(input);
      container.appendChild(fieldDiv);
  });

  const saveBtn = document.createElement("button");
  saveBtn.innerHTML = '<i class="ms-Icon ms-Icon--Save" style="margin-right:8px;"></i> Save Changes';
  saveBtn.className = "ms-Button ms-Button--primary";
  saveBtn.style.width = "100%";
  saveBtn.onclick = () => saveRecordChanges(dataTableName, id, targetDataSet.fields, selectedRev);

  const cancelBtn = document.createElement("button");
  cancelBtn.innerHTML = '<i class="ms-Icon ms-Icon--Cancel" style="margin-right:8px;"></i> Cancel';
  cancelBtn.className = "ms-Button ms-Button--default";
  cancelBtn.style.width = "100%";
  cancelBtn.onclick = () => { container.style.display = "none"; };

  const editBtnGroup = document.createElement("div");
  editBtnGroup.className = "button-group";
  editBtnGroup.style.display = "flex";
  editBtnGroup.style.flexDirection = "column";
  editBtnGroup.style.gap = "8px";
  editBtnGroup.style.marginTop = "12px";
  editBtnGroup.appendChild(saveBtn);
  editBtnGroup.appendChild(cancelBtn);
  container.appendChild(editBtnGroup);
}

export async function saveRecordChanges(dataTableName: string, id: string, fields: string[], selectedRev?: number) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    let store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) return;

    let targetDataSet = dataSet;
    if (selectedRev && selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
        targetDataSet = dataSet.history[selectedRev];
    }

    const idField = targetDataSet.fields[0];
    const recordIndex = targetDataSet.records.findIndex((r: any) => String(r[idField]) === String(id));

    if (recordIndex === -1) return;

    fields.forEach(field => {
        const input = document.getElementById(`edit-input-${dataTableName}-${field}`) as HTMLInputElement;
        if (input && !input.disabled) {
            targetDataSet.records[recordIndex][field] = input.value;
        }
    });

    await idbSet("DC_STORE", JSON.stringify(store));

    const container = document.getElementById(`edit-form-container-${dataTableName}`);
    if (container) container.style.display = "none";

    if (status) {
        status.innerText = `Record updated. Refreshing Excel...`;
        status.style.color = "green";
    }

    refreshFormulas(true);
  } catch (error) {
    if (status) {
      if (error.name === "QuotaExceededError" || (error.message && error.message.includes("exceeded the quota"))) {
        status.innerText = "Storage limit reached! Cannot save changes until you clear some space.";
      } else {
        status.innerText = "Error saving changes: " + error.message;
      }
      status.style.color = "red";
    }
  }
}

export async function exportCSV(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) throw new Error("No data found.");

    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) throw new Error("Data table not found.");

    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;

    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
    }

    if (!targetDataSet || !targetDataSet.fields || !targetDataSet.records) {
      throw new Error("No data found for this revision.");
    }

    await exportCSVData(dataTableName, selectedRev, targetDataSet);

    if (status) {
      status.innerText = `Exported ${dataTableName} (Rev ${selectedRev}) to CSV.`;
      status.style.color = "green";
    }
  } catch (error) {
    if (status) {
      status.innerText = "Error exporting CSV: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function refreshFormulas(silent: boolean = false) {
  const status = document.getElementById("status-text");
  try {
    if (status && !silent) status.innerText = "Refreshing dashboard & formulas... Please wait.";

    if (!silent) {
      await renderDashboard();
    }

    const count = await executeRefreshFormulas();
    if (status && !silent) {
      status.innerText = `Refreshed dashboard and ${count} DC formulas.`;
      status.style.color = "green";
    }
  } catch (error) {
    if (status && !silent) {
      status.innerText = "Error: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function convertToValues() {
  const status = document.getElementById("status-text");
  try {
    if (status) status.innerText = "Converting... Please wait.";
    const count = await executeConvertToValues();
    if (status) {
      status.innerText = `Converted ${count} DC formulas to values.`;
      status.style.color = "green";
    }
  } catch (error) {
    if (status) {
      status.innerText = "Error: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function backupData() {
  try {
    await downloadBackup();

    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = "Backup downloaded successfully.";
      status.style.color = "green";
    }
  } catch (error) {
    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = "Backup error: " + error.message;
      status.style.color = "red";
    }
  }
}

export function triggerRestore() {
  document.getElementById("restore-file-input").click();
}

export async function restoreData(event: any) {
  const status = document.getElementById("status-text");
  try {
    await processRestoreFile(event);
    if (status) { status.innerText = "Data restored successfully."; status.style.color = "green"; }
    renderDashboard();
    refreshFormulas(true);
  } catch (error) {
    if (status) { status.innerText = "Restore error: " + error.message; status.style.color = "red"; }
  }
}

export async function insertDropdown(dataTableName: string) {
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) throw new Error("No data found.");
    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
    }
    if (!targetDataSet || !targetDataSet.fields) throw new Error("No headers found.");

    await executeInsertDropdown(targetDataSet.fields);
    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = `Inserted headers dropdown for ${dataTableName}`;
      status.style.color = "green";
    }
  } catch (error) {
    console.error(error);
    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = "Error inserting dropdown: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function insertTable(dataTableName: string) {
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
    }
    if (!targetDataSet || !targetDataSet.fields || !targetDataSet.records || targetDataSet.records.length === 0) {
      throw new Error("No data found for this data table.");
    }

    await executeInsertTable(targetDataSet.fields, targetDataSet.records);
    const status = document.getElementById("status-text");
    if (status) {
       status.innerText = `Inserted table for ${dataTableName}`;
       status.style.color = "green";
    }
  } catch (error) {
    const status = document.getElementById("status-text");
    if (status) {
       status.innerText = "Error inserting table: " + error.message;
       status.style.color = "red";
    }
  }
}

export async function insertFormulaFor(dataTableName: string) {
  try {
    const idInput = document.getElementById(`fb-id-${dataTableName}`) as HTMLInputElement;
    const fieldSelect = document.getElementById(`fb-field-${dataTableName}`) as HTMLSelectElement;
    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    if (!idInput || !fieldSelect) return;

    const id = idInput.value;
    const field = fieldSelect.value;
    if (!id) throw new Error("Please fill the ID to insert formula.");

    const storedData = await idbGet("DC_STORE");
    let store: any = {};
    if (storedData) store = JSON.parse(storedData);
    const dataSet = store[dataTableName] || {};
    
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    const isLatest = selectedRev === (dataSet.revision || 1);

    await executeInsertFormula(id, field, dataTableName, selectedRev, isLatest);
  } catch (error) {
    const status = document.getElementById("status-text");
    if (status) {
       status.innerText = error.message;
       status.style.color = "red";
    }
  }
}