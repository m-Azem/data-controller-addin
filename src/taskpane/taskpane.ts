/* global Excel, Office, OfficeRuntime */

import { idbGet, idbSet } from "../utils/db";
import { migrateFromLocalStorage } from "../services/migration";
import { exportCSVData, downloadBackup, processRestoreFile } from "../services/fileService";
import { executeInsertTable, executeInsertDropdown, executeConvertToValues, executeRefreshFormulas } from "../services/excelService";

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
    document.getElementById("settings-button").onclick = toggleSettings;
    
    // Setup Capture Accordion
    const captureAccordion = document.getElementById("capture-accordion");
    const captureContent = document.getElementById("capture-content");
    if (captureAccordion && captureContent) {
      captureAccordion.addEventListener("click", () => {
        captureAccordion.classList.toggle("active");
        captureContent.classList.toggle("show");
      });
    }

    // Setup Formula Builder Accordion
    const formulaAccordion = document.getElementById("formula-builder-accordion");
    const formulaContent = document.getElementById("formula-builder-content");
    if (formulaAccordion && formulaContent) {
      formulaAccordion.addEventListener("click", () => {
        formulaAccordion.classList.toggle("active");
        formulaContent.classList.toggle("show");
      });
    }

    document.getElementById("formula-select")?.addEventListener("change", renderFormulaBuilder);
    document.getElementById("insert-built-formula-button").onclick = insertBuiltFormula;

    // Setup Theme Toggle
    const themeToggle = document.getElementById("theme-toggle");
    if (themeToggle) themeToggle.addEventListener("change", toggleTheme);

    // Display App Version
    const versionDisplay = document.getElementById("app-version-display");
    if (versionDisplay) versionDisplay.innerText = APP_VERSION;

    await migrateFromLocalStorage(); // Migrate old data if present
    loadSettings();
    renderDashboard();
    renderFormulaBuilder();
  }
});

export function customPrompt(title: string, message: string, defaultValue: string = "", autocompleteOptions?: string[]): Promise<string | null> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-modal");
      const titleEl = document.getElementById("modal-title");
      const messageEl = document.getElementById("modal-message");
      const inputEl = document.getElementById("modal-input") as HTMLInputElement;
      const btnOk = document.getElementById("modal-ok");
      const btnCancel = document.getElementById("modal-cancel");

      if (!modal || !titleEl || !messageEl || !inputEl || !btnOk || !btnCancel) {
          resolve(null);
          return;
      }

      let dataList = document.getElementById("modal-datalist") as HTMLDataListElement;
      if (!dataList) {
          dataList = document.createElement("datalist");
          dataList.id = "modal-datalist";
          inputEl.parentNode?.appendChild(dataList);
      }
      dataList.innerHTML = "";
      if (autocompleteOptions && autocompleteOptions.length > 0) {
          inputEl.setAttribute("list", "modal-datalist");
          autocompleteOptions.forEach(opt => {
              const option = document.createElement("option");
              option.value = opt;
              dataList.appendChild(option);
          });
      } else {
          inputEl.removeAttribute("list");
      }

      titleEl.innerText = title;
      messageEl.innerText = message;
      inputEl.value = defaultValue;
      modal.style.display = "flex";
      inputEl.focus();
      
      inputEl.onkeydown = (e) => { if (e.key === "Enter") btnOk.click(); };

      const cleanup = () => {
          modal.style.display = "none";
          btnOk.onclick = null;
          btnCancel.onclick = null;
          inputEl.onkeydown = null;
      };

      btnOk.onclick = () => { cleanup(); resolve(inputEl.value); };
      btnCancel.onclick = () => { cleanup(); resolve(null); };
  });
}

export function customConfirm(title: string, message: string, confirmText: string = "Yes"): Promise<boolean> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-confirm-modal");
      const titleEl = document.getElementById("confirm-modal-title");
      const messageEl = document.getElementById("confirm-modal-message");
      const btnYes = document.getElementById("confirm-modal-yes");
      const btnNo = document.getElementById("confirm-modal-no");

      if (!modal || !titleEl || !messageEl || !btnYes || !btnNo) {
          resolve(false);
          return;
      }

      titleEl.innerText = title;
      messageEl.innerText = message;
      btnYes.innerText = confirmText;
      modal.style.display = "flex";

      const cleanup = () => {
          modal.style.display = "none";
          btnYes.onclick = null;
          btnNo.onclick = null;
      };

      btnYes.onclick = () => { cleanup(); resolve(true); };
      btnNo.onclick = () => { cleanup(); resolve(false); };
  });
}

export function customSortPrompt(title: string, message: string, items: string[]): Promise<string[] | null> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-sort-modal");
      const titleEl = document.getElementById("sort-modal-title");
      const messageEl = document.getElementById("sort-modal-message");
      const listEl = document.getElementById("sort-modal-list");
      const btnOk = document.getElementById("sort-modal-ok");
      const btnCancel = document.getElementById("sort-modal-cancel");

      if (!modal || !titleEl || !messageEl || !listEl || !btnOk || !btnCancel) {
          resolve(null);
          return;
      }

      titleEl.innerText = title;
      messageEl.innerText = message;
      listEl.innerHTML = "";

      items.forEach((item) => {
          const li = document.createElement("li");
          li.className = "sortable-item";
          li.draggable = true;
          li.innerHTML = `<i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888;"></i><span>${item}</span>`;
          li.dataset.value = item;
          
          li.addEventListener("dragstart", () => li.classList.add("dragging"));
          li.addEventListener("dragend", () => li.classList.remove("dragging"));
          
          listEl.appendChild(li);
      });
      
      const handleDragOver = (e: DragEvent) => {
          e.preventDefault();
          const dragging = listEl.querySelector('.dragging') as HTMLElement;
          if (!dragging) return;
          
          const siblings = [...listEl.querySelectorAll('.sortable-item:not(.dragging)')] as HTMLElement[];
          const nextSibling = siblings.find(sibling => {
              const box = sibling.getBoundingClientRect();
              return (e.clientY - box.top - box.height / 2) < 0;
          });
          
          if (nextSibling) {
              listEl.insertBefore(dragging, nextSibling);
          } else {
              listEl.appendChild(dragging);
          }
      };
      
      listEl.addEventListener("dragover", handleDragOver);

      modal.style.display = "flex";

      const cleanup = () => {
          modal.style.display = "none";
          btnOk.onclick = null;
          btnCancel.onclick = null;
          listEl.removeEventListener("dragover", handleDragOver);
      };

      btnOk.onclick = () => {
          const newOrder = Array.from(listEl.children).map((li) => (li as HTMLElement).dataset.value as string);
          cleanup();
          resolve(newOrder);
      };
      btnCancel.onclick = () => { cleanup(); resolve(null); };
  });
}

export interface FormField {
    id: string;
    label: string;
    type: 'text' | 'select' | 'autocomplete';
    options?: string[];
    value?: string;
    disabled?: boolean;
}

export function customFormPrompt(title: string, message: string, fields: FormField[]): Promise<Record<string, string> | null> {
    return new Promise((resolve) => {
        const modal = document.getElementById("custom-form-modal");
        const titleEl = document.getElementById("form-modal-title");
        const messageEl = document.getElementById("form-modal-message");
        const inputsContainer = document.getElementById("form-modal-inputs");
        const btnOk = document.getElementById("form-modal-ok");
        const btnCancel = document.getElementById("form-modal-cancel");

        if (!modal || !titleEl || !messageEl || !inputsContainer || !btnOk || !btnCancel) {
            resolve(null);
            return;
        }

        titleEl.innerText = title;
        messageEl.innerText = message;
        inputsContainer.innerHTML = "";

        const inputElements: { id: string, el: HTMLInputElement | HTMLSelectElement }[] = [];

        fields.forEach(f => {
            const fieldDiv = document.createElement("div");
            fieldDiv.className = "ms-TextField";
            fieldDiv.style.marginBottom = "12px";

            const label = document.createElement("label");
            label.className = "ms-Label";
            label.innerText = f.label;
            fieldDiv.appendChild(label);

            if (f.type === 'select') {
                const select = document.createElement("select");
                select.className = "ms-Dropdown-title";
                if (f.disabled) select.disabled = true;
                (f.options || []).forEach(opt => {
                    const option = document.createElement("option");
                    option.value = opt;
                    option.text = opt;
                    select.appendChild(option);
                });
                if (f.value !== undefined) select.value = f.value;
                fieldDiv.appendChild(select);
                inputElements.push({ id: f.id, el: select });
            } else if (f.type === 'autocomplete') {
                const input = document.createElement("input");
                input.type = "text";
                input.className = "ms-TextField-field";
                if (f.disabled) input.disabled = true;
                if (f.value !== undefined) input.value = f.value;
                
                const dataListId = `datalist-${f.id}-${Date.now()}`;
                input.setAttribute("list", dataListId);
                
                const dataList = document.createElement("datalist");
                dataList.id = dataListId;
                (f.options || []).forEach(opt => {
                    const option = document.createElement("option");
                    option.value = opt;
                    dataList.appendChild(option);
                });
                
                fieldDiv.appendChild(input);
                fieldDiv.appendChild(dataList);
                inputElements.push({ id: f.id, el: input });
            } else {
                const input = document.createElement("input");
                input.type = "text";
                input.className = "ms-TextField-field";
                if (f.disabled) input.disabled = true;
                if (f.value !== undefined) input.value = f.value;
                fieldDiv.appendChild(input);
                inputElements.push({ id: f.id, el: input });
            }

            inputsContainer.appendChild(fieldDiv);
        });

        modal.style.display = "flex";

        const cleanup = () => {
            modal.style.display = "none";
            btnOk.onclick = null;
            btnCancel.onclick = null;
        };

        btnOk.onclick = () => {
            const result: Record<string, string> = {};
            inputElements.forEach(item => {
                result[item.id] = item.el.value;
            });
            cleanup();
            resolve(result);
        };
        btnCancel.onclick = () => { cleanup(); resolve(null); };
    });
}

const formulaDefinitions: Record<string, { id: string, label: string, required: boolean }[]> = {
    "GET": [
        { id: "param-id", label: "Record ID", required: false },
        { id: "param-fieldName", label: "Field Name", required: false },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ],
    "SEARCH": [
        { id: "param-searchField", label: "Search Field", required: true },
        { id: "param-searchValue", label: "Search Value", required: true },
        { id: "param-returnField", label: "Return Field", required: true },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ],
    "FILTER": [
        { id: "param-searchField", label: "Search Field", required: true },
        { id: "param-searchValue", label: "Search Value", required: true },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false },
        { id: "param-exactMatch", label: "Exact Match (TRUE/FALSE)", required: false }
    ],
    "SUM": [
        { id: "param-sumField", label: "Sum Field", required: true },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ],
    "SUMIFS": [
        { id: "param-sumField", label: "Sum Field", required: true },
        { id: "param-criteriaField", label: "Criteria Field", required: true },
        { id: "param-criteriaValue", label: "Criteria Value", required: true },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ],
    "JOIN": [
        { id: "param-baseTableName", label: "Base Table Name", required: true },
        { id: "param-foreignKeyField", label: "Foreign Key Field", required: true },
        { id: "param-foreignTableName", label: "Foreign Table Name", required: true },
        { id: "param-foreignReturnField", label: "Foreign Return Field", required: true }
    ]
};

export async function renderFormulaBuilder() {
    const select = document.getElementById("formula-select") as HTMLSelectElement;
    const container = document.getElementById("formula-inputs-container");
    if (!select || !container) return;

    const formula = select.value;
    const fields = formulaDefinitions[formula] || [];

    container.innerHTML = "";

    const storedData = await idbGet("DC_STORE");
    let store: any = {};
    let tables: string[] = [];
    if (storedData) {
        store = JSON.parse(storedData);
        tables = Object.keys(store);
    }
    let defaultTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    if (!defaultTable && tables.length > 0) defaultTable = tables[0];

    const inputRefs: { id: string, el: HTMLInputElement | HTMLSelectElement, datalist?: HTMLDataListElement }[] = [];

    const getTargetTableForInput = (inputId: string): string => {
        if (formula === "JOIN") {
            if (inputId === "param-foreignReturnField") {
                const fInput = document.getElementById("param-foreignTableName") as HTMLInputElement;
                return fInput && fInput.value.trim() !== "" ? fInput.value.trim() : "";
            } else {
                const bInput = document.getElementById("param-baseTableName") as HTMLInputElement;
                return bInput && bInput.value.trim() !== "" ? bInput.value.trim() : defaultTable;
            }
        } else {
            const tInput = document.getElementById("param-dataTableName") as HTMLInputElement;
            return tInput && tInput.value.trim() !== "" ? tInput.value.trim() : defaultTable;
        }
    };

    const updateDataLists = () => {
        inputRefs.forEach(ref => {
            const targetEl = ref.datalist || ref.el;
            const currentVal = ref.el.value;
            targetEl.innerHTML = "";
            
            if (ref.el.tagName === "SELECT") {
                const emptyOpt = document.createElement("option");
                emptyOpt.value = "";
                emptyOpt.text = "Select...";
                targetEl.appendChild(emptyOpt);
            }

            if (ref.id.toLowerCase().includes("tablename")) {
                tables.forEach(t => {
                    const opt = document.createElement("option");
                    opt.value = t;
                    targetEl.appendChild(opt);
                });
                return;
            }

            if (ref.id === "param-exactMatch") {
                ["TRUE", "FALSE"].forEach(t => {
                    const opt = document.createElement("option");
                    opt.value = t;
                    opt.text = t;
                    targetEl.appendChild(opt);
                });
                if (ref.el.tagName === "SELECT" && ["TRUE", "FALSE"].includes(currentVal)) ref.el.value = currentVal;
                return;
            }

            const tName = getTargetTableForInput(ref.id);
            const dataSet = store[tName];
            if (!dataSet) return;

            let options: string[] = [];
            if (ref.id.toLowerCase().includes("field")) {
                options = dataSet.fields || [];
            } else if (ref.id === "param-id") {
                const idField = dataSet.idField || dataSet.fields[0];
                options = (dataSet.records || []).map((r: any) => String(r[idField]));
            }

            // Remove duplicates and empty strings
            options = [...new Set(options)].filter(o => o !== undefined && o !== null && String(o).trim() !== "");

            options.forEach(opt => {
                const optionElement = document.createElement("option");
                optionElement.value = opt;
                if (ref.el.tagName === "SELECT") optionElement.text = opt;
                targetEl.appendChild(optionElement);
            });

            if (ref.el.tagName === "SELECT" && options.includes(currentVal)) {
                ref.el.value = currentVal;
            }
        });
    };

    fields.forEach(f => {
        const fieldDiv = document.createElement("div");
        fieldDiv.className = "ms-TextField";
        fieldDiv.style.marginBottom = "8px";

        const label = document.createElement("label");
        label.className = "ms-Label";
        label.innerHTML = `${f.label} ${f.required ? '<span style="color:#d13438;">*</span>' : '<span style="color:#888;font-weight:normal;">(Optional)</span>'}`;

        const isSelect = f.id.toLowerCase().includes("field") || f.id === "param-exactMatch";
        let inputEl: HTMLInputElement | HTMLSelectElement;
        let dataList: HTMLDataListElement | undefined;

        if (isSelect) {
            inputEl = document.createElement("select");
            inputEl.id = f.id;
            inputEl.className = "ms-Dropdown-title";
            inputRefs.push({ id: f.id, el: inputEl });
        } else {
            inputEl = document.createElement("input");
            inputEl.type = "text";
            inputEl.id = f.id;
            inputEl.className = "ms-TextField-field";
            inputEl.autocomplete = "off";

            const listId = `datalist-${f.id}`;
            inputEl.setAttribute("list", listId);
            
            dataList = document.createElement("datalist");
            dataList.id = listId;

            inputRefs.push({ id: f.id, el: inputEl, datalist: dataList });

            if (f.id.toLowerCase().includes("tablename")) {
                inputEl.addEventListener("input", updateDataLists);
                inputEl.addEventListener("change", updateDataLists);
            }
        }

        fieldDiv.appendChild(label);
        fieldDiv.appendChild(inputEl);
        if (dataList) fieldDiv.appendChild(dataList);
        container.appendChild(fieldDiv);
    });

    updateDataLists();
}

export async function insertBuiltFormula() {
    const status = document.getElementById("status-text");
    try {
        const select = document.getElementById("formula-select") as HTMLSelectElement;
        const formula = select.value;
        const fields = formulaDefinitions[formula] || [];

        let args: string[] = [];

        for (let i = 0; i < fields.length; i++) {
            const f = fields[i];
            const input = document.getElementById(f.id) as HTMLInputElement | HTMLSelectElement;
            let val = input.value.trim();

            if (f.required && val === "") {
                throw new Error(`'${f.label}' is required.`);
            }

            if (val === "") {
                args.push("");
            } else {
                if (val.toUpperCase() === "TRUE" || val.toUpperCase() === "FALSE") {
                    args.push(val.toUpperCase());
                } else if (!isNaN(Number(val))) {
                    args.push(val);
                } else if (val.startsWith('"') && val.endsWith('"')) {
                    args.push(val);
                } else if (/^('?[^'!]+'?!)?\$?[A-Za-z]+\$?[0-9]+(:\$?[A-Za-z]+\$?[0-9]+)?$/.test(val)) {
                    args.push(val.toUpperCase()); // Preserve cell references
                } else {
                    args.push(`"${val}"`); // Auto-quote simple text
                }
            }
        }

        // Remove trailing empty arguments for cleaner formula
        while (args.length > 0 && args[args.length - 1] === "") {
            args.pop();
        }

        const formulaStr = `=DC.${formula}(${args.join(", ")})`;

        await Excel.run(async (context) => {
            const cell = context.workbook.getActiveCell();
            cell.formulas = [[formulaStr]];
            await context.sync();
        });

        if (status) {
            status.innerText = `Inserted formula: ${formulaStr}`;
            status.style.color = "green";
        }

    } catch (error) {
        if (status) {
            status.innerText = error.message;
            status.style.color = "red";
        }
    }
}

export function toggleSettings() {
  const settingsView = document.getElementById('settings-view');
  const mainView = document.getElementById('main-view');
  const settingsBtn = document.getElementById('settings-button');
  const isSettingsOpen = settingsView?.style.display === 'block';

  settingsBtn?.classList.remove('active');

  if (isSettingsOpen) {
    settingsView.style.display = 'none';
    mainView.style.display = 'block';
  } else {
    settingsView.style.display = 'block';
    mainView.style.display = 'none';
    settingsBtn?.classList.add('active');
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

    const fields = [...originalHeaders];

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
      idField: idColumn,
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

    const captureAccordion = document.getElementById("capture-accordion");
    const captureContent = document.getElementById("capture-content");
    if (captureAccordion) captureAccordion.classList.remove("active");
    if (captureContent) captureContent.classList.remove("show");

    await idbSet("DC_DEFAULT_DATA_TABLE", dataName);
    await idbSet("DC_DEFAULT_REVISION", String(rev));

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

      let store: any = {};
      const existingStore = await idbGet("DC_STORE");
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(`Data table '${dataTableName}' not found.`);
      }

      const dataSet = store[dataTableName];
      const existingIdField = dataSet.idField || dataSet.fields[0];

      if (!headers.includes(existingIdField)) {
        throw new Error(`The selected range must include the original ID column: '${existingIdField}'.`);
      }

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
        status.innerText = isLatest ? `Replaced '${dataTableName}' with ${dataRows.length} new records. (ID: "${existingIdField}")` : `Replaced Rev ${selectedRev} of '${dataTableName}' with ${dataRows.length} records.`;
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

export async function captureNewRevision(dataTableName: string) {
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

      let store: any = {};
      const existingStore = await idbGet("DC_STORE");
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(`Data table '${dataTableName}' not found.`);
      }

      const dataSet = store[dataTableName];
      const existingIdField = dataSet.idField || dataSet.fields[0];

      if (!headers.includes(existingIdField)) {
        throw new Error(`The selected range must include the original ID column: '${existingIdField}'.`);
      }

      const currentRev = dataSet.revision || 1;
      dataSet.history = dataSet.history || {};
      dataSet.history[currentRev] = {
        fields: [...dataSet.fields],
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      
      dataSet.revision = currentRev + 1;
      dataSet.fields = headers;
      dataSet.records = dataRows.map((row: any[]) => {
        let record: any = {};
        headers.forEach((header, index) => { record[header] = row[index]; });
        return record;
      });

      await idbSet("DC_STORE", JSON.stringify(store));
      await idbSet("DC_DEFAULT_REVISION", String(dataSet.revision));
      
      if (status) { status.innerText = `Captured Rev ${dataSet.revision} for '${dataTableName}' with ${dataRows.length} records.`; status.style.color = "green"; }
      renderDashboard();
      refreshFormulas(true);
    });
  } catch (error) {
    if (status) { status.innerText = "Error capturing new revision: " + error.message; status.style.color = "red"; }
  }
}

export async function renderDashboard() {
  const list = document.getElementById("data-table-list");
  if (!list) return;
  
  list.innerHTML = "";
  const storedData = await idbGet("DC_STORE");
  const defaultSelect = document.getElementById("default-data-table-select") as HTMLSelectElement;
  const defaultRevSelect = document.getElementById("default-revision-select") as HTMLSelectElement;

  if (!storedData) {
    list.innerHTML = "<li>No data tables stored yet.</li>";
    if (defaultSelect) defaultSelect.innerHTML = "";
    if (defaultRevSelect) defaultRevSelect.innerHTML = "";
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
    } else if (!store[key].idField && store[key].fields && store[key].fields.length > 0) {
      const origFields = store[key].history && store[key].history[1] ? store[key].history[1].fields : store[key].fields;
      store[key].idField = origFields[0];
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
    if (defaultRevSelect) defaultRevSelect.innerHTML = "";
    return;
  }

  // Setup Default Data Table dropdown
  if (defaultSelect) {
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

      const localRevSelect = document.getElementById(`fb-rev-${tableName}`) as HTMLSelectElement;
      if (localRevSelect && localRevSelect.value !== defaultRevSelect.value) {
          localRevSelect.value = defaultRevSelect.value;
          localRevSelect.dispatchEvent(new Event('change'));
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
        
        const localRevSelect = document.getElementById(`fb-rev-${defaultSelect.value}`) as HTMLSelectElement;
        if (localRevSelect && localRevSelect.value !== defaultRevSelect.value) {
            localRevSelect.value = defaultRevSelect.value;
            localRevSelect.dispatchEvent(new Event('change'));
        }
      };
    }
  }

  keys.forEach(key => {
    const dataTable = store[key];
    const count = dataTable.records ? dataTable.records.length : 0;
    const rev = dataTable.revision || 1;
    
    const li = document.createElement("li");
    li.style.listStyle = "none";
    li.style.marginBottom = "16px";
    
    const header = document.createElement("button");
    header.className = "accordion";
    header.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; width:100%;">
          <div><i class="ms-Icon ms-Icon--Table" style="margin-right: 8px;"></i> ${key}</div>
          <div style="font-size:12px; font-weight:normal; opacity:0.8;">Rev ${rev} &bull; ${count} rows</div>
      </div>
    `;
    
    const details = document.createElement("div");
    details.className = "accordion-content";
    
    header.onclick = () => {
      header.classList.toggle("active");
      details.classList.toggle("show");
    };

    const insertTableBtn = document.createElement("button");
    insertTableBtn.innerHTML = '<i class="ms-Icon ms-Icon--Table" style="margin-right:8px;"></i> Insert Table';
    insertTableBtn.className = "ms-Button ms-Button--default w-100";
    insertTableBtn.onclick = () => insertTable(key);
    
    const replaceBtn = document.createElement("button");
    replaceBtn.innerHTML = '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Current Version';
    replaceBtn.className = "ms-Button ms-Button--default w-100";
    replaceBtn.onclick = () => replaceTableData(key);

    const snapshotBtn = document.createElement("button");
    snapshotBtn.innerHTML = '<i class="ms-Icon ms-Icon--Camera" style="margin-right:8px;"></i> Snapshot';
    snapshotBtn.className = "ms-Button ms-Button--primary w-100";
    snapshotBtn.onclick = () => createSnapshot(key);

    const deleteVersionBtn = document.createElement("button");
    deleteVersionBtn.innerHTML = '<i class="ms-Icon ms-Icon--RemoveEvent" style="margin-right:8px;"></i> Delete Version';
    deleteVersionBtn.className = "ms-Button ms-Button--danger w-100";
    deleteVersionBtn.onclick = () => deleteCurrentVersion(key);

    const deleteBtn = document.createElement("button");
    deleteBtn.innerHTML = '<i class="ms-Icon ms-Icon--Delete" style="margin-right:8px;"></i> Delete Table';
    deleteBtn.className = "ms-Button ms-Button--danger w-100";
    deleteBtn.onclick = () => deleteDataTable(key);
    
    const exportCSVBtn = document.createElement("button");
    exportCSVBtn.innerHTML = '<i class="ms-Icon ms-Icon--Download" style="margin-right:8px;"></i> Export CSV';
    exportCSVBtn.className = "ms-Button ms-Button--default w-100";
    exportCSVBtn.onclick = () => exportCSV(key);

        const insertDropdownBtn = document.createElement("button");
        insertDropdownBtn.innerHTML = '<i class="ms-Icon ms-Icon--Dropdown" style="margin-right:8px;"></i> Headers Dropdown';
        insertDropdownBtn.className = "ms-Button ms-Button--default w-100";
        insertDropdownBtn.onclick = () => insertDropdown(key);

        const editBtn = document.createElement("button");
        editBtn.innerHTML = '<i class="ms-Icon ms-Icon--Edit" style="margin-right:8px;"></i> Edit Record';
        editBtn.className = "ms-Button ms-Button--primary w-100";

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
        
        if (defaultSelect && defaultSelect.value === key && defaultRevSelect) {
            if (Array.from(revSelect.options).some(o => o.value === defaultRevSelect.value)) {
                revSelect.value = defaultRevSelect.value;
            }
        }

        revSelectContainer.appendChild(revSelect);

        editBtn.onclick = () => loadRecordForEdit(key);
        
        revSelect.onchange = async () => {
          const isLatest = parseInt(revSelect.value) === rev;
          replaceBtn.innerHTML = isLatest ? '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Current Version' : '<i class="ms-Icon ms-Icon--Sync" style="margin-right:8px;"></i> Replace Rev Data';
          snapshotBtn.innerHTML = isLatest ? '<i class="ms-Icon ms-Icon--Camera" style="margin-right:8px;"></i> Snapshot' : '<i class="ms-Icon ms-Icon--Undo" style="margin-right:8px;"></i> Restore as Active';
          
          const fastAcc = accordions.find(a => a.theme === 'fast');
          if (fastAcc) {
              fastAcc.title = `Fast Tools (Rev ${revSelect.value})`;
              const titleSpan = fastAcc.header.querySelector("span:not(.icon)");
              if (titleSpan) titleSpan.innerHTML = fastAcc.title;
          }

          if (defaultSelect && defaultSelect.value === key && defaultRevSelect) {
              if (defaultRevSelect.value !== revSelect.value) {
                  defaultRevSelect.value = revSelect.value;
                  await idbSet("DC_DEFAULT_REVISION", revSelect.value);
              }
          }
        };

        const accordions: { header: HTMLElement, content: HTMLElement, title: string, theme: 'default' | 'danger' | 'fast' }[] = [];

        // Accordion Builder Helper
        const buildAccordion = (title: string, elements: HTMLElement[], theme: 'default' | 'danger' | 'fast' = 'default', defaultOpen: boolean = false) => {
            const accContainer = document.createElement("div");
            accContainer.style.marginBottom = "8px";
            
            let headerClass = "inner-accordion-header";
            let iconColor = "var(--primary-color)";
            if (theme === 'danger') { headerClass += " danger"; iconColor = "#d13438"; }
            else if (theme === 'fast') { headerClass += " fast"; iconColor = "#0078d4"; }
            if (defaultOpen) { headerClass += " open"; }
            
            const accHeader = document.createElement("button");
            accHeader.className = headerClass;
            
            const iconSpan = `<span class="icon" style="display:inline-block; width:15px; color: ${iconColor};">${defaultOpen ? '&#9660;' : '&#9654;'}</span>`;
            accHeader.innerHTML = `${iconSpan}<span>${title}</span>`;
            
            const accContent = document.createElement("div");
            accContent.className = "inner-accordion-content flex-col";
            accContent.style.display = defaultOpen ? "flex" : "none";
            
            elements.forEach(el => accContent.appendChild(el));
            
            const accObj = { header: accHeader, content: accContent, title, theme };
            accordions.push(accObj);

            accHeader.onclick = () => {
                const isHidden = accContent.style.display === "none";
                
                // Collapse all others
                accordions.forEach(acc => {
                    acc.content.style.display = "none";
                    acc.header.classList.remove("open");
                    let cColor = "var(--primary-color)";
                    if (acc.theme === 'danger') cColor = "#d13438";
                    else if (acc.theme === 'fast') cColor = "#0078d4";
                    const accIcon = `<span class="icon" style="display:inline-block; width:15px; color: ${cColor};">&#9654;</span>`;
                    acc.header.innerHTML = `${accIcon}<span>${acc.title}</span>`;
                });

                // Open if it was hidden
                if (isHidden) {
                    accContent.style.display = "flex";
                    accHeader.classList.add("open");
                    const openIcon = `<span class="icon" style="display:inline-block; width:15px; color: ${iconColor};">&#9660;</span>`;
                    accHeader.innerHTML = `${openIcon}<span>${accObj.title}</span>`;
                }
            };
            
            accContainer.appendChild(accHeader);
            accContainer.appendChild(accContent);
            return accContainer;
        };

        const addColumnBtn = document.createElement("button");
        addColumnBtn.innerHTML = '<i class="ms-Icon ms-Icon--Add" style="margin-right:8px;"></i> Add Column';
        addColumnBtn.className = "ms-Button ms-Button--default w-100";
        addColumnBtn.onclick = () => addColumn(key);

        const resortColumnsBtn = document.createElement("button");
        resortColumnsBtn.innerHTML = '<i class="ms-Icon ms-Icon--Sort" style="margin-right:8px;"></i> Resort Columns';
        resortColumnsBtn.className = "ms-Button ms-Button--default w-100";
        resortColumnsBtn.onclick = () => resortColumns(key);

        const captureRevBtn = document.createElement("button");
        captureRevBtn.innerHTML = '<i class="ms-Icon ms-Icon--Camera" style="margin-right:8px;"></i> Capture New Revision';
        captureRevBtn.className = "ms-Button ms-Button--primary w-100";
        captureRevBtn.onclick = () => captureNewRevision(key);

        // 1. Fast Tools
        const secFastTools = buildAccordion(`Fast Tools (Rev ${revSelect.value})`, [
            insertDropdownBtn,
            insertTableBtn,
            editBtn
        ], 'fast', true);

        // 2. Version & Data Options
        const sec1 = buildAccordion("Data & Version Options", [
            revSelectContainer,
            captureRevBtn,
            replaceBtn,
            addColumnBtn,
            resortColumnsBtn,
            snapshotBtn,
            exportCSVBtn
        ]);

        // 3. Delete
        const sec4 = buildAccordion("Delete Operations", [
            deleteVersionBtn,
            deleteBtn
        ], 'danger');

        details.appendChild(secFastTools);
        details.appendChild(sec1);
        details.appendChild(sec4);

    li.appendChild(header);
    li.appendChild(details);
    list.appendChild(li);
  });
  
  renderFormulaBuilder();
}

export async function deleteDataTable(dataTableName: string) {
  const confirmed = await customConfirm("Delete Entire Table", `Are you sure you want to permanently delete the table '${dataTableName}' and all its history? This action cannot be undone.`, "Yes, Delete Everything");
  if (!confirmed) return;

  const storedData = await idbGet("DC_STORE");
  if (storedData) {
    let store = JSON.parse(storedData);
    if (store[dataTableName]) {
      delete store[dataTableName];
    }

    const defaultTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    if (defaultTable === dataTableName) {
      await idbSet("DC_DEFAULT_DATA_TABLE", "");
      await idbSet("DC_DEFAULT_REVISION", "");
    }

    await idbSet("DC_STORE", JSON.stringify(store));
    renderDashboard();
    refreshFormulas(true);
    
    const status = document.getElementById("status-text");
    if (status) {
        status.innerText = `Deleted entire data table: ${dataTableName}`;
        status.style.color = "blue";
    }
  }
}

export async function deleteCurrentVersion(dataTableName: string) {
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

    const msg = isLatest ? `Are you sure you want to delete the current version of '${dataTableName}' and rollback to the previous revision?` : `Are you sure you want to delete historical Rev ${selectedRev} from '${dataTableName}'?`;
    const confirmed = await customConfirm("Confirm Deletion", msg, "Yes, Delete");
    if (!confirmed) return;

    if (isLatest) {
      const historyKeys = dataSet.history ? Object.keys(dataSet.history).map(Number).sort((a, b) => b - a) : [];
      if (historyKeys.length > 0) {
        const highestOldRev = historyKeys[0];
        dataSet.fields = dataSet.history[highestOldRev].fields;
        dataSet.records = dataSet.history[highestOldRev].records;
        dataSet.revision = highestOldRev;
        delete dataSet.history[highestOldRev];

        if (status) {
          status.innerText = `Deleted current version. Rolled back to Rev ${highestOldRev}.`;
          status.style.color = "green";
        }
      } else {
        delete store[dataTableName];
        if (status) {
          status.innerText = `Deleted the only version. Table '${dataTableName}' removed.`;
          status.style.color = "blue";
        }
      }
    } else {
      if (dataSet.history && dataSet.history[selectedRev]) {
        delete dataSet.history[selectedRev];
        if (status) {
          status.innerText = `Deleted historical Rev ${selectedRev} from '${dataTableName}'.`;
          status.style.color = "green";
        }
      }
    }

    await idbSet("DC_STORE", JSON.stringify(store));

    const defaultTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    if (defaultTable === dataTableName) {
        let defaultRev = await idbGet("DC_DEFAULT_REVISION");
        if (defaultRev === String(selectedRev)) {
            await idbSet("DC_DEFAULT_REVISION", store[dataTableName] ? String(store[dataTableName].revision || 1) : "1");
        }
    }

    renderDashboard();
    refreshFormulas(true);
  } catch (error) {
    if (status) {
      status.innerText = "Error deleting version: " + error.message;
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
    
    // Automatically set the new revision as default!
    await idbSet("DC_DEFAULT_REVISION", String(dataSet.revision));

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

export async function addColumn(dataTableName: string) {
  const status = document.getElementById("status-text");
  const colName = await customPrompt("Add Column", `Enter a new column name to add to '${dataTableName}':`);
  
  if (!colName || colName.trim() === "") return;

  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    let store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) return;

    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    const isLatest = selectedRev === (dataSet.revision || 1);

    if (!isLatest) {
      if (status) { status.innerText = "Cannot add column to a historical revision."; status.style.color = "red"; }
      return;
    }

    if (dataSet.fields.includes(colName)) {
      if (status) { status.innerText = "Column already exists."; status.style.color = "red"; }
      return;
    }

    // Create history backup before adding the column
    dataSet.history = dataSet.history || {};
    dataSet.history[dataSet.revision] = {
      fields: [...dataSet.fields],
      records: JSON.parse(JSON.stringify(dataSet.records))
    };

    dataSet.fields.push(colName);
    dataSet.records.forEach((r: any) => { r[colName] = ""; });
    dataSet.revision += 1;

    await idbSet("DC_STORE", JSON.stringify(store));
    await idbSet("DC_DEFAULT_REVISION", String(dataSet.revision));
    renderDashboard();
    refreshFormulas(true);
    
    if (status) {
      status.innerText = `Added column '${colName}'. Current is Rev ${dataSet.revision}.`;
      status.style.color = "green";
    }
  } catch (error) {
    if (status) {
      status.innerText = "Error: " + error.message;
      status.style.color = "red";
    }
  }
}

export async function resortColumns(dataTableName: string) {
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

    if (!isLatest) {
      if (status) { status.innerText = "Cannot resort columns of a historical revision."; status.style.color = "red"; }
      return;
    }

    const newFields = await customSortPrompt("Resort Columns", `Drag and drop the columns below to reorder them for '${dataTableName}'.`, dataSet.fields);
    
    if (!newFields) return; // User cancelled

    // If the order hasn't changed, do nothing
    if (dataSet.fields.join(",") === newFields.join(",")) return;
    
    // Validation: must have the exact same elements
    const oldFieldsSorted = [...dataSet.fields].sort().join(",");
    const newFieldsSorted = [...newFields].sort().join(",");
    
    if (oldFieldsSorted !== newFieldsSorted) {
      if (status) { status.innerText = "Invalid column list."; status.style.color = "red"; }
      return;
    }

    // Create history backup before resorting
    dataSet.history = dataSet.history || {};
    dataSet.history[dataSet.revision] = {
      fields: [...dataSet.fields],
      records: JSON.parse(JSON.stringify(dataSet.records))
    };

    dataSet.fields = newFields;
    dataSet.revision += 1;

    await idbSet("DC_STORE", JSON.stringify(store));
    renderDashboard();
    refreshFormulas(true);
    
    if (status) {
      status.innerText = `Columns resorted successfully. Current is Rev ${dataSet.revision}.`;
      status.style.color = "green";
    }
  } catch (error) {
    if (status) { status.innerText = "Error: " + error.message; status.style.color = "red"; }
  }
}

export async function loadRecordForEdit(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet("DC_STORE");
    if (!storedData) return;
    let store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) return;

    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;

    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
        targetDataSet = dataSet.history[selectedRev];
    }

    const idField = dataSet.idField || targetDataSet.fields[0];
    const allIds = targetDataSet.records.map((r: any) => String(r[idField]));

    const id = await customPrompt("Edit Record", `Enter the Record ID to edit in '${dataTableName}':`, "", allIds);
    if (!id || id.trim() === "") return;

    const record = targetDataSet.records.find((r: any) => String(r[idField]) === String(id));

    if (!record) {
        if (status) { status.innerText = `Record '${id}' not found.`; status.style.color = "red"; }
        return;
    }

    const formFields: FormField[] = targetDataSet.fields.map((field: string) => ({
        id: field,
        label: field,
        type: 'text',
        value: record[field] !== undefined ? String(record[field]) : "",
        disabled: field === idField
    }));

    const editResult = await customFormPrompt(`Editing Record: ${id}`, "Update the values below:", formFields);
    if (!editResult) return;

    const recordIndex = targetDataSet.records.findIndex((r: any) => String(r[idField]) === String(id));
    if (recordIndex === -1) return;

    targetDataSet.fields.forEach((field: string) => {
        if (field !== idField && editResult[field] !== undefined) {
            targetDataSet.records[recordIndex][field] = editResult[field];
        }
    });

    await idbSet("DC_STORE", JSON.stringify(store));

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