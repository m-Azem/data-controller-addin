/* global Excel, Office, OfficeRuntime */

import { idbGet, idbSet } from "../utils/db";
import { migrateFromLocalStorage } from "../services/migration";
import { exportCSVData, downloadBackup, processRestoreFile } from "../services/fileService";
import { executeInsertTable, executeInsertDropdown, executeConvertToValues, executeRefreshFormulas } from "../services/excelService";

const APP_VERSION = "1.0.1"; // Update this number whenever you release a new version

const IDB_KEYS = {
    STORE: "DC_STORE",
    THEME: "DC_THEME",
    DEFAULT_TABLE: "DC_DEFAULT_DATA_TABLE",
    DEFAULT_REVISION: "DC_DEFAULT_REVISION"
};

interface DCRecord {
  __DC_ID__: string;
  [key: string]: any;
}

interface DataSetVersion {
  idField: string;
  fields: string[];
  records: DCRecord[];
}

interface DataSet extends DataSetVersion {
  dataTableName: string;
  revision: number;
  history: { [rev: number]: DataSetVersion };
}

interface Store {
  [dataTableName: string]: DataSet;
}

let activeFormulaInput: HTMLInputElement | null = null;

export function showStatus(message: string, type: "success" | "error" | "info" = "info") {
    const status = document.getElementById("status-text");
    if (!status) return;
    status.innerText = message;
    if (type === "success") status.style.color = "green";
    else if (type === "error") status.style.color = "red";
    else status.style.color = "blue";
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("capture-button").onclick = loadRangeForCapture;
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
        if (!formulaContent.classList.contains("show")) {
            activeFormulaInput = null;
        }
      });
    }

    document.getElementById("formula-select")?.addEventListener("change", renderFormulaBuilder);
    document.getElementById("formula-select")?.addEventListener("focus", () => { activeFormulaInput = null; });
    document.getElementById("formula-select")?.addEventListener("mousedown", () => { activeFormulaInput = null; });
    document.getElementById("insert-built-formula-button").onclick = insertBuiltFormula;
    document.getElementById("insert-built-formula-button")?.addEventListener("focus", () => { activeFormulaInput = null; });
    document.getElementById("insert-built-formula-button")?.addEventListener("mousedown", () => { activeFormulaInput = null; });

    document.getElementById("clear-formula-button")?.addEventListener("click", clearFormulaForm);
    document.getElementById("clear-formula-button")?.addEventListener("focus", () => { activeFormulaInput = null; });
    document.getElementById("clear-formula-button")?.addEventListener("mousedown", () => { activeFormulaInput = null; });

    // Setup Theme Toggle
    const themeToggle = document.getElementById("theme-toggle");
    if (themeToggle) themeToggle.addEventListener("change", toggleTheme);

    // Display App Version
    const versionDisplay = document.getElementById("app-version-display");
    if (versionDisplay) versionDisplay.innerText = APP_VERSION;

    // Handle automatic cell reference insertion for Formula Builder
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
        const formulaContent = document.getElementById("formula-builder-content");
        const isVisible = formulaContent && formulaContent.classList.contains("show");
        if (isVisible && activeFormulaInput && document.body.contains(activeFormulaInput)) {
            try {
                await Excel.run(async (context) => {
                    const range = context.workbook.getSelectedRange();
                    range.load("address");
                    await context.sync();
                    let address = range.address;
                    if (address.includes("!")) address = address.split("!")[1];
                    if (address.includes(":")) address = address.split(":")[0];
                    activeFormulaInput.value = address.replace(/\$/g, "");
                });
            } catch (e) {
                // Silently ignore errors, e.g., when selection is not a range.
            }
        }
    });

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

export function customManageColumnsPrompt(title: string, message: string, items: string[], idField: string): Promise<{changes: {oldName: string, newName: string}[], saveAsNewRevision: boolean} | null> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-sort-modal");
      const titleEl = document.getElementById("sort-modal-title");
      const messageEl = document.getElementById("sort-modal-message");
      const listEl = document.getElementById("sort-modal-list");
      const btnDeleteSelected = document.getElementById("sort-modal-delete-selected");
      const btnSaveCurrent = document.getElementById("sort-modal-save-current");
      const btnSaveNew = document.getElementById("sort-modal-save-new");
      const btnCancel = document.getElementById("sort-modal-cancel");

      if (!modal || !titleEl || !messageEl || !listEl || !btnDeleteSelected || !btnSaveCurrent || !btnSaveNew || !btnCancel) {
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
          li.dataset.original = item;
          
          const isId = item === idField;
          
          li.innerHTML = `
              <div style="display:flex; align-items:center; width:100%;">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-right: 8px;" ${isId ? 'disabled title="Cannot delete ID column"' : ''} />
                <input type="text" class="ms-TextField-field" value="${item}" style="flex: 1; padding: 2px 8px;" />
                ${isId ? '<span style="font-size:10px; color:#0078d4; margin-left:8px; font-weight:bold;" title="Primary ID Column">(ID)</span>' : ''}
              </div>
          `;
          
          // Prevent drag when selecting text inside the input
          const inputEl = li.querySelector("input");
          if (inputEl) {
              inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
          }
          
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
          btnDeleteSelected.onclick = null;
          btnSaveCurrent.onclick = null;
          btnSaveNew.onclick = null;
          btnCancel.onclick = null;
          listEl.removeEventListener("dragover", handleDragOver);
      };

      const getChanges = () => {
          return Array.from(listEl.children).map((li) => {
              const original = (li as HTMLElement).dataset.original as string;
              const newName = ((li as HTMLElement).querySelector("input") as HTMLInputElement).value.trim();
              return { oldName: original, newName: newName || original };
          });
      };

      btnDeleteSelected.onclick = () => {
          const checkboxes = listEl.querySelectorAll('.col-delete-checkbox:checked');
          checkboxes.forEach(cb => {
              const li = cb.closest('.sortable-item');
              if (li) li.remove();
          });
      };

      btnSaveCurrent.onclick = () => {
          cleanup();
          resolve({ changes: getChanges(), saveAsNewRevision: false });
      };
      btnSaveNew.onclick = () => {
          cleanup();
          resolve({ changes: getChanges(), saveAsNewRevision: true });
      };
      btnCancel.onclick = () => { cleanup(); resolve(null); };
  });
}

export function customDataSummaryPrompt(
  title: string,
  message: string,
  headers: string[],
  dataRows: (string | number | boolean)[][],
  defaultIdField: string = ""
): Promise<{ idField: string, fields: string[], records: any[] } | null> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-summary-modal");
      const titleEl = document.getElementById("summary-modal-title");
      const messageEl = document.getElementById("summary-modal-message");
      const idSelect = document.getElementById("summary-modal-id-select") as HTMLSelectElement;
      const errorEl = document.getElementById("summary-modal-error");
      const listEl = document.getElementById("summary-modal-list");
      const btnDeleteSelected = document.getElementById("summary-modal-delete-selected");
      const btnOk = document.getElementById("summary-modal-ok");
      const btnCancel = document.getElementById("summary-modal-cancel");
      const btnNext = document.getElementById("summary-modal-next") as HTMLButtonElement;
      const btnBack = document.getElementById("summary-modal-back") as HTMLButtonElement;
      const step1Div = document.getElementById("summary-step-1");
      const step2Div = document.getElementById("summary-step-2");
      const recordsCountEl = document.getElementById("summary-records-count");
      const idColumnEl = document.getElementById("summary-id-column");
      const finalColumnsEl = document.getElementById("summary-final-columns");


      if (!modal || !titleEl || !messageEl || !idSelect || !errorEl || !listEl || !btnDeleteSelected || !btnOk || !btnCancel || !btnNext || !btnBack || !step1Div || !step2Div || !recordsCountEl || !idColumnEl || !finalColumnsEl) {
          resolve(null); return;
      }

      titleEl.innerText = title;
      messageEl.innerText = message;
      errorEl.style.display = "none";
      
      idSelect.innerHTML = "";
      headers.forEach((h, i) => {
          const opt = document.createElement("option");
          opt.value = h;
          opt.text = h || `Column ${i + 1}`;
          idSelect.appendChild(opt);
      });
      if (defaultIdField && headers.includes(defaultIdField)) idSelect.value = defaultIdField;

      const renderBadges = () => {
          const idName = idSelect.value;
          Array.from(listEl.children).forEach(li => {
              const orig = (li as HTMLElement).dataset.original;
              let badge = li.querySelector('.id-badge');
              if (orig === idName) {
                  if (!badge) {
                      badge = document.createElement("span");
                      badge.className = "id-badge";
                      badge.style.cssText = "font-size:10px; color:#0078d4; margin-left:8px; font-weight:bold;";
                      badge.innerText = "(ID)";
                      li.querySelector('div')?.appendChild(badge);
                  }
              } else {
                  if (badge) badge.remove();
              }
          });
      };
      idSelect.onchange = renderBadges;

      listEl.innerHTML = "";
      headers.forEach((item) => {
          const li = document.createElement("li");
          li.className = "sortable-item";
          li.draggable = true;
          li.dataset.original = item;
          li.innerHTML = `
              <div style="display:flex; align-items:center; width:100%;">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-right: 8px;" />
                <input type="text" class="ms-TextField-field" value="${item}" style="flex: 1; padding: 2px 8px;" />
              </div>
          `;
          const inputEl = li.querySelector("input[type='text']");
          if (inputEl) inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
          li.addEventListener("dragstart", () => li.classList.add("dragging"));
          li.addEventListener("dragend", () => li.classList.remove("dragging"));
          listEl.appendChild(li);
      });
      renderBadges();
      
      const handleDragOver = (e: DragEvent) => {
          e.preventDefault();
          const dragging = listEl.querySelector('.dragging') as HTMLElement;
          if (!dragging) return;
          const siblings = [...listEl.querySelectorAll('.sortable-item:not(.dragging)')] as HTMLElement[];
          const nextSibling = siblings.find(sibling => {
              const box = sibling.getBoundingClientRect();
              return (e.clientY - box.top - box.height / 2) < 0;
          });
          if (nextSibling) listEl.insertBefore(dragging, nextSibling);
          else listEl.appendChild(dragging);
      };
      listEl.addEventListener("dragover", handleDragOver);
      modal.style.display = "flex";

      let validationData: { idField: string, fields: string[], records: any[] } | null = null;

      const cleanup = () => {
          modal.style.display = "none";
          btnDeleteSelected.onclick = null;
          btnNext.onclick = null;
          btnBack.onclick = null;
          btnOk.onclick = null;
          btnCancel.onclick = null;
          idSelect.onchange = null;
          listEl.removeEventListener("dragover", handleDragOver);
      };

      btnDeleteSelected.onclick = () => {
          const checkboxes = listEl.querySelectorAll('.col-delete-checkbox:checked');
          checkboxes.forEach(cb => {
              const li = cb.closest('.sortable-item');
              if (li && (li as HTMLElement).dataset.original !== idSelect.value) li.remove();
              else if (li) { errorEl.innerText = `Cannot delete the Primary ID column ('${idSelect.value}').`; errorEl.style.display = "block"; }
          });
      };

      btnNext.onclick = () => {
          errorEl.style.display = "none";
          const selectedIdOrigName = idSelect.value;
          const idLi = Array.from(listEl.children).find(li => (li as HTMLElement).dataset.original === selectedIdOrigName);
          if (!idLi) { errorEl.innerText = "The Primary ID column must be present in the table."; errorEl.style.display = "block"; return; }

          const idIndexOrig = headers.indexOf(selectedIdOrigName);
          const idSet = new Set();
          for (let i = 0; i < dataRows.length; i++) {
              const val = String(dataRows[i][idIndexOrig]);
              if (!val || val.trim() === "") { errorEl.innerText = `Row ${i + 1} has an empty ID. IDs must be non-empty.`; errorEl.style.display = "block"; return; }
              if (idSet.has(val)) { errorEl.innerText = `Duplicate ID found: '${val}'. All IDs must be unique.`; errorEl.style.display = "block"; return; }
              idSet.add(val);
          }
          
          const newFieldsList = Array.from(listEl.children).map(li => ((li as HTMLElement).querySelector("input[type='text']") as HTMLInputElement).value.trim());
          const uniqueFields = new Set(newFieldsList);
          if (uniqueFields.size !== newFieldsList.length) { errorEl.innerText = "All column names must be unique."; errorEl.style.display = "block"; return; }
          
          const columnsInfo = Array.from(listEl.children).map(li => ({
              oIdx: headers.indexOf((li as HTMLElement).dataset.original as string),
              newName: ((li as HTMLElement).querySelector("input[type='text']") as HTMLInputElement).value.trim()
          }));

          const finalRecords = dataRows.map(row => {
              const rec: any = {};
              columnsInfo.forEach(col => { rec[col.newName] = row[col.oIdx]; });
              rec.__DC_ID__ = String(row[idIndexOrig]);
              return rec;
          });

          validationData = {
              idField: ((idLi as HTMLElement).querySelector("input[type='text']") as HTMLInputElement).value.trim(),
              fields: newFieldsList,
              records: finalRecords
          };

          step1Div.style.display = "none";
          step2Div.style.display = "block";
          btnDeleteSelected.style.display = "none";
          btnNext.style.display = "none";
          btnBack.style.display = "block";
          btnOk.style.display = "block";
          
          if (recordsCountEl) recordsCountEl.innerText = String(finalRecords.length); 
          if (idColumnEl) idColumnEl.innerText = validationData.idField;
          if (finalColumnsEl) finalColumnsEl.innerText = newFieldsList.join(", ");
      };

      btnBack.onclick = () => {
          step1Div.style.display = "block";
          step2Div.style.display = "none";
          btnDeleteSelected.style.display = "block";
          btnNext.style.display = "block";
          btnBack.style.display = "none";
          btnOk.style.display = "none";
          validationData = null;
      };
      btnOk.onclick = () => { cleanup(); resolve(validationData); };
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
    ],
    "SORT": [
        { id: "param-sortField", label: "Sort Field", required: true },
        { id: "param-ascending", label: "Ascending (TRUE/FALSE)", required: false },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ]
};

export async function renderFormulaBuilder() {
    const select = document.getElementById("formula-select") as HTMLSelectElement;
    const container = document.getElementById("formula-inputs-container");
    if (!select || !container) return;

    const formula = select.value;
    const fields = formulaDefinitions[formula] || [];

    container.innerHTML = "";

    const storedData = await idbGet(IDB_KEYS.STORE);
    let store: Store = {};
    let tables: string[] = [];
    if (storedData) {
        store = JSON.parse(storedData);
        tables = Object.keys(store);
    }
    let defaultTable = await idbGet(IDB_KEYS.DEFAULT_TABLE);
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

            if (ref.id === "param-exactMatch" || ref.id === "param-ascending") {
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
                options = (dataSet.records || []).map((r: any) => String(r.__DC_ID__));
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

        const isSelect = f.id.toLowerCase().includes("field") || f.id === "param-exactMatch" || f.id === "param-ascending";
        let inputEl: HTMLInputElement | HTMLSelectElement;
        let dataList: HTMLDataListElement | undefined;

        if (isSelect) {
            inputEl = document.createElement("select");
            inputEl.id = f.id;
            inputEl.className = "ms-Dropdown-title";
            inputRefs.push({ id: f.id, el: inputEl });
            inputEl.addEventListener("focus", () => { activeFormulaInput = null; });
            inputEl.addEventListener("mousedown", () => { activeFormulaInput = null; });
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

            inputEl.addEventListener("focus", () => { activeFormulaInput = inputEl as HTMLInputElement; });
            inputEl.addEventListener("mousedown", () => { activeFormulaInput = inputEl as HTMLInputElement; });

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
        showStatus(error instanceof Error ? error.message : String(error), "error");
    }
}

export function clearFormulaForm() {
    const container = document.getElementById("formula-inputs-container");
    if (!container) return;
    
    const inputs = container.querySelectorAll<HTMLInputElement | HTMLSelectElement>("input, select");
    inputs.forEach((el) => {
        el.value = "";
    });
    activeFormulaInput = null;
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

export async function toggleTheme(event: Event) {
  const isDark = (event.target as HTMLInputElement).checked;
  if (isDark) document.body.classList.add('theme-dark');
  else document.body.classList.remove('theme-dark');
  await idbSet(IDB_KEYS.THEME, isDark ? "dark" : "light");
}

export async function loadSettings() {
  const theme = await idbGet(IDB_KEYS.THEME);
  const toggle = document.getElementById("theme-toggle") as HTMLInputElement;
  if (theme === "dark") {
    document.body.classList.add("theme-dark");
    if (toggle) toggle.checked = true;
  }
}
 
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

          const summary = await customDataSummaryPrompt(
              "Capture New Table",
              `Review and map columns for '${dataName}' (${values.length - 1} rows detected).`,
              values[0],
              values.slice(1)
          );
          if (!summary) return;

          let store: Store = {};
          const existingStore = await idbGet(IDB_KEYS.STORE);
          if (existingStore) store = JSON.parse(existingStore);
          
          let rev = 1;
          let history = {};
          if (store[dataName]) {
              rev = store[dataName].revision || 1;
              history = store[dataName].history || {};
          }

          store[dataName] = {
              dataTableName: dataName,
              idField: summary.idField,
              revision: rev,
              history: history,
              fields: summary.fields,
              records: summary.records
          };

          await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
          if (status) { status.innerText = `Saved ${summary.records.length} records in ${dataName}.`; status.style.color = "green"; }

          (document.getElementById("data-name") as HTMLInputElement).value = "";

    const captureAccordion = document.getElementById("capture-accordion");
    const captureContent = document.getElementById("capture-content");
    if (captureAccordion) captureAccordion.classList.remove("active");
    if (captureContent) captureContent.classList.remove("show");

    await idbSet(IDB_KEYS.DEFAULT_TABLE, dataName);
    await idbSet(IDB_KEYS.DEFAULT_REVISION, String(rev));

    await renderDashboard();
    await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus(error.message, "error");
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

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
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

      const existingIdField = targetDataSet.idField || dataSet.idField || dataSet.fields[0];

      const summary = await customDataSummaryPrompt(
          "Replace Version",
          `Review replacing data for '${dataTableName}' (${values.length - 1} rows detected).`,
          values[0],
          values.slice(1),
          existingIdField
      );
      if (!summary) return;

      targetDataSet.fields = summary.fields;
      targetDataSet.records = summary.records;
      targetDataSet.idField = summary.idField;

      await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
      
      if (status) {
        status.innerText = isLatest ? `Replaced '${dataTableName}' with ${summary.records.length} new records. (ID: "${summary.idField}")` : `Replaced Rev ${selectedRev} of '${dataTableName}' with ${summary.records.length} records.`;
        status.style.color = "green";
      }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus("Error replacing data: " + error.message, "error");
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

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(`Data table '${dataTableName}' not found.`);
      }

      const dataSet = store[dataTableName];
      const existingIdField = dataSet.idField || dataSet.fields[0];

      const summary = await customDataSummaryPrompt(
          "Capture New Revision",
          `Review data for new revision of '${dataTableName}' (${values.length - 1} rows detected).`,
          values[0],
          values.slice(1),
          existingIdField
      );
      if (!summary) return;

      const currentRev = dataSet.revision || 1;
      dataSet.history = dataSet.history || {};
      dataSet.history[currentRev] = {
        idField: dataSet.idField,
        fields: [...dataSet.fields],
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      
      dataSet.revision = currentRev + 1;
      dataSet.idField = summary.idField;
      dataSet.fields = summary.fields;
      dataSet.records = summary.records;

      await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
      const currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);
      if (currentDefault === dataTableName) {
          await idbSet(IDB_KEYS.DEFAULT_REVISION, String(dataSet.revision));
      }
      
      if (status) { status.innerText = `Captured Rev ${dataSet.revision} for '${dataTableName}' with ${dataRows.length} records.`; status.style.color = "green"; }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus("Error capturing new revision: " + error.message, "error");
  }
}

export async function renderDashboard() {
  const list = document.getElementById("data-table-list");
  if (!list) return;
  
  list.innerHTML = "";
  const storedData = await idbGet(IDB_KEYS.STORE);
  const defaultSelect = document.getElementById("default-data-table-select") as HTMLSelectElement;
  const defaultRevSelect = document.getElementById("default-revision-select") as HTMLSelectElement;

  if (!storedData) {
    list.innerHTML = "<li>No data tables stored yet.</li>";
    if (defaultSelect) defaultSelect.innerHTML = "";
    if (defaultRevSelect) defaultRevSelect.innerHTML = "";
    return;
  }

  let store: Store = JSON.parse(storedData);
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

  // NEW: Ensure __DC_ID__ exists on all records for the internal mark
  const keys = Object.keys(store);
  keys.forEach(key => {
      const dataTable = store[key];
      if (dataTable.records && dataTable.records.length > 0 && !dataTable.records[0].hasOwnProperty('__DC_ID__')) {
          const idF = dataTable.idField || dataTable.fields[0];
          dataTable.records.forEach((r: any) => { r.__DC_ID__ = String(r[idF]); });
          needsSave = true;
      }
      if (dataTable.history) {
          Object.values(dataTable.history).forEach((h: any) => {
              if (h.records && h.records.length > 0 && !h.records[0].hasOwnProperty('__DC_ID__')) {
                  const hIdF = h.idField || dataTable.idField || h.fields[0];
                  h.records.forEach((r: any) => { r.__DC_ID__ = String(r[hIdF]); });
                  needsSave = true;
              }
          });
      }
  });

  if (needsSave) {
    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
  }

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
    
    let currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);
    if (!currentDefault || !keys.includes(currentDefault)) {
      currentDefault = keys[0];
      await idbSet(IDB_KEYS.DEFAULT_TABLE, currentDefault);
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
      let currentDefaultRev = await idbGet(IDB_KEYS.DEFAULT_REVISION);
      if (currentDefaultRev && Array.from(defaultRevSelect.options).some(o => o.value === currentDefaultRev)) {
        defaultRevSelect.value = currentDefaultRev;
      } else {
        defaultRevSelect.selectedIndex = 0;
        await idbSet(IDB_KEYS.DEFAULT_REVISION, defaultRevSelect.value);
      }

      const localRevSelect = document.getElementById(`fb-rev-${tableName}`) as HTMLSelectElement;
      if (localRevSelect && localRevSelect.value !== defaultRevSelect.value) {
          localRevSelect.value = defaultRevSelect.value;
          localRevSelect.dispatchEvent(new Event('change'));
      }
    };

    await updateDefaultRevOptions(currentDefault);

    defaultSelect.onchange = async () => {
      await idbSet(IDB_KEYS.DEFAULT_TABLE, defaultSelect.value);
      const tData = store[defaultSelect.value];
      if (tData) {
          await idbSet(IDB_KEYS.DEFAULT_REVISION, String(tData.revision || 1));
      }
      await updateDefaultRevOptions(defaultSelect.value);
    };

    if (defaultRevSelect) {
      defaultRevSelect.onchange = async () => {
        await idbSet(IDB_KEYS.DEFAULT_REVISION, defaultRevSelect.value);
        
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
                  await idbSet(IDB_KEYS.DEFAULT_REVISION, revSelect.value);
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
        addColumnBtn.innerHTML = '<i class="ms-Icon ms-Icon--Add" style="margin-right:8px;"></i> Add Empty Column';
        addColumnBtn.className = "ms-Button ms-Button--default w-100";
        addColumnBtn.onclick = () => addColumn(key);

        const editColumnsBtn = document.createElement("button");
        editColumnsBtn.innerHTML = '<i class="ms-Icon ms-Icon--Sort" style="margin-right:8px;"></i> Edit Columns';
        editColumnsBtn.className = "ms-Button ms-Button--default w-100"; 
        editColumnsBtn.onclick = () => manageColumns(key);

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
            editColumnsBtn,
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

  const storedData = await idbGet(IDB_KEYS.STORE);
  if (storedData) {
    let store: Store = JSON.parse(storedData);
    if (store[dataTableName]) {
      delete store[dataTableName];
    }

    const defaultTable = await idbGet(IDB_KEYS.DEFAULT_TABLE);
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
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    let store: Store = JSON.parse(storedData);
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
        dataSet.idField = dataSet.history[highestOldRev].idField || dataSet.history[highestOldRev].fields[0];
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

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));

    const defaultTable = await idbGet(IDB_KEYS.DEFAULT_TABLE);
    if (defaultTable === dataTableName) {
        let defaultRev = await idbGet(IDB_KEYS.DEFAULT_REVISION);
        if (defaultRev === String(selectedRev)) {
            await idbSet("DC_DEFAULT_REVISION", store[dataTableName] ? String(store[dataTableName].revision || 1) : "1");
        }
    }

    renderDashboard();
    refreshFormulas(true);
  } catch (error: any) {
    showStatus("Error deleting version: " + error.message, "error");
  }
}

export async function createSnapshot(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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
        idField: dataSet.idField,
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
        idField: dataSet.idField,
        fields: dataSet.fields,
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      dataSet.fields = JSON.parse(JSON.stringify(dataSet.history[selectedRev].fields));
      dataSet.records = JSON.parse(JSON.stringify(dataSet.history[selectedRev].records));
      dataSet.idField = dataSet.history[selectedRev].idField || dataSet.fields[0];
      dataSet.revision = currentRev + 1;

      if (status) {
          status.innerText = `Restored Rev ${selectedRev} as new active Rev ${dataSet.revision}.`;
          status.style.color = "green";
      }
    }

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    
    const currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);
    if (currentDefault === dataTableName) {
        await idbSet(IDB_KEYS.DEFAULT_REVISION, String(dataSet.revision));
    }

    renderDashboard();
    if (!isLatest) refreshFormulas(true);
  } catch (error: any) {
    if (error.name === "QuotaExceededError" || (error.message && error.message.includes("exceeded the quota"))) {
        showStatus("Storage limit reached! Please use 'Clear History' or delete tables to free up space.", "error");
    } else {
        showStatus("Error creating snapshot: " + error.message, "error");
    }
  }
}

export async function addColumn(dataTableName: string) {
  const status = document.getElementById("status-text");
  const colName = await customPrompt("Add Column", `Enter a new column name to add to '${dataTableName}':`);
  
  if (!colName || colName.trim() === "") return;

  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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
      idField: dataSet.idField,
      fields: [...dataSet.fields],
      records: JSON.parse(JSON.stringify(dataSet.records))
    };

    dataSet.fields.push(colName);
    dataSet.records.forEach((r: any) => { r[colName] = ""; });
    dataSet.revision += 1;

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    const currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);
    if (currentDefault === dataTableName) {
        await idbSet(IDB_KEYS.DEFAULT_REVISION, String(dataSet.revision));
    }
    renderDashboard();
    refreshFormulas(true);
    
    if (status) {
      status.innerText = `Added column '${colName}'. Current is Rev ${dataSet.revision}.`;
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus("Error adding column: " + error.message, "error");
  }
}

export async function manageColumns(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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

    const idField = dataSet.idField || dataSet.fields[0];
    const promptResult = await customManageColumnsPrompt("Edit Columns", `Drag to reorder or click to rename columns for '${dataTableName}'.`, dataSet.fields, idField);
    
    if (!promptResult) return; // User cancelled

    const columnChanges = promptResult.changes;
    const newFields = columnChanges.map(c => c.newName);
    
    // Check for duplicates
    const uniqueFields = new Set(newFields);
    if (uniqueFields.size !== newFields.length) {
        if (status) { status.innerText = "Column names must be unique."; status.style.color = "red"; }
        return;
    }

    // If nothing changed
    if (dataSet.fields.join(",") === newFields.join(",") && columnChanges.every(c => c.oldName === c.newName) && columnChanges.length === dataSet.fields.length) return;

    const droppedFields = dataSet.fields.filter((f: string) => !columnChanges.find(c => c.oldName === f));

    if (promptResult.saveAsNewRevision) {
        // Create history backup before resorting
        dataSet.history = dataSet.history || {};
        dataSet.history[dataSet.revision] = {
            idField: dataSet.idField,
            fields: [...dataSet.fields],
            records: JSON.parse(JSON.stringify(dataSet.records))
        };
        dataSet.revision += 1;
    }

    dataSet.fields = newFields;
    
    const hasRenames = columnChanges.some(c => c.oldName !== c.newName) || droppedFields.length > 0;
    if (hasRenames) {
        dataSet.records.forEach((r: any) => {
            columnChanges.forEach(c => {
                if (c.oldName !== c.newName) {
                    r[c.newName] = r[c.oldName];
                    delete r[c.oldName];
                }
            });
            droppedFields.forEach((df: string) => {
                delete r[df];
            });
        });
        
        const idChange = columnChanges.find(c => c.oldName === dataSet.idField);
        if (idChange && idChange.oldName !== idChange.newName) {
            dataSet.idField = idChange.newName;
        }
    }

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    if (promptResult.saveAsNewRevision) {
        const currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);
        if (currentDefault === dataTableName) {
            await idbSet(IDB_KEYS.DEFAULT_REVISION, String(dataSet.revision));
        }
    }
    renderDashboard();
    refreshFormulas(true);
    
    if (status) {
      status.innerText = `Columns updated successfully. Current is Rev ${dataSet.revision}.`;
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus("Error: " + error.message, "error");
  }
}

export async function loadRecordForEdit(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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

    const idField = targetDataSet.idField || dataSet.idField || targetDataSet.fields[0];
    const allIds = targetDataSet.records.map((r: any) => String(r.__DC_ID__));

    const id = await customPrompt("Edit Record", `Enter the Record ID to edit in '${dataTableName}':`, "", allIds);
    if (!id || id.trim() === "") return;

    const record = targetDataSet.records.find((r: any) => String(r.__DC_ID__) === String(id));

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

    const recordIndex = targetDataSet.records.findIndex((r: any) => String(r.__DC_ID__) === String(id));
    if (recordIndex === -1) return;

    targetDataSet.fields.forEach((field: string) => {
        if (field !== idField && editResult[field] !== undefined) {
            targetDataSet.records[recordIndex][field] = editResult[field];
        }
    });

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));

    if (status) {
        status.innerText = `Record updated. Refreshing Excel...`;
        status.style.color = "green";
    }

    await refreshFormulas(true);
  } catch (error: any) {
    if (error.name === "QuotaExceededError" || (error.message && error.message.includes("exceeded the quota"))) {
        showStatus("Storage limit reached! Cannot save changes until you clear some space.", "error");
    } else {
        showStatus("Error saving changes: " + error.message, "error");
    }
  }
}

export async function exportCSV(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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
  } catch (error: any) {
    showStatus("Error exporting CSV: " + error.message, "error");
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
  } catch (error: any) {
    if (!silent) showStatus(error.message, "error");
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
  } catch (error: any) {
    showStatus(error.message, "error");
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
  } catch (error: any) {
    showStatus("Backup error: " + error.message, "error");
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
    await renderDashboard();
    await refreshFormulas(true);
  } catch (error: any) {
    showStatus("Restore error: " + error.message, "error");
  }
}

export async function insertDropdown(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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
  } catch (error: any) {
    console.error(error);
    showStatus("Error inserting dropdown: " + error.message, "error");
  }
}

export async function insertTable(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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
  } catch (error: any) {
    showStatus("Error inserting table: " + error.message, "error");
  }
}