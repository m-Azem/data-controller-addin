/* global Excel, Office, OfficeRuntime */

import { idbGet, idbSet } from "../utils/db";
import { migrateFromLocalStorage } from "../services/migration";
import { exportCSVData, downloadBackup, processRestoreFile } from "../services/fileService";
import { executeInsertTable, executeInsertDropdown, executeConvertToValues, executeRefreshFormulas } from "../services/excelService";

const APP_VERSION = "1.0.1"; // Update this number whenever you release a new version

// Translation Dictionary
const TRANSLATIONS: Record<string, Record<string, string>> = {
    en: {
        app_title: "Data Controller",
        status_label: "Status:",
        default_table: "Default Table (Formulas)",
        revision: "Revision",
        workspaces: "Workspaces",
        formula_builder: "Visual Formula Builder",
        global_variables: "Global Variables",
        add_variable: "Add Variable",
        backup_restore: "Data Backup & Restore",
        backup_data: "Backup Data",
        restore_data: "Restore Data",
        theme_appearance: "Theme Appearance",
        dark_theme: "Enable Dark Theme",
        language: "Language / اللغة",
        help_doc: "Help & Documentation",
        help_btn: "Open Help Center",
        about: "About",
        insert_btn: "Insert",
        clear_btn: "Clear",
        status_ready: "Ready",
        global_relations: "Global Relations",
        manage_relations: "Add / Manage Relations",
        manage_workspaces: "Manage Workspaces",
        add_workspace: "Add Workspace"
    },
    ar: { // Keeping Arabic dictionary for future use, even if UI selection is hidden
        app_title: "متحكم البيانات",
        status_label: "الحالة:",
        default_table: "الجدول الافتراضي (للمعادلات)",
        revision: "النسخة",
        workspaces: "مساحات العمل",
        formula_builder: "منشئ المعادلات المرئي",
        global_variables: "المتغيرات العامة",
        add_variable: "إضافة متغير",
        backup_restore: "النسخ الاحتياطي والاستعادة",
        backup_data: "نسخ احتياطي",
        restore_data: "استعادة البيانات",
        theme_appearance: "المظهر",
        dark_theme: "تفعيل الوضع الداكن",

        manage_workspaces: "Manage Workspaces",
        add_workspace: "Add Workspace"
    },
    ar: {
        app_title: "متحكم البيانات",
        status_label: "الحالة:",
        default_table: "الجدول الافتراضي (للمعادلات)",
        revision: "النسخة",
        workspaces: "مساحات العمل",
        formula_builder: "منشئ المعادلات المرئي",
        global_variables: "المتغيرات العامة",
        add_variable: "إضافة متغير",
        backup_restore: "النسخ الاحتياطي والاستعادة",
        backup_data: "نسخ احتياطي",
        restore_data: "استعادة البيانات",
        theme_appearance: "المظهر",
        dark_theme: "تفعيل الوضع الداكن",
        language: "Language / اللغة",
        help_doc: "المساعدة والوثائق",
        help_btn: "فتح مركز المساعدة",
        about: "حول التطبيق",
        insert_btn: "إدراج",
        clear_btn: "مسح",
        status_ready: "جاهز",
        global_relations: "العلاقات العامة",
        manage_relations: "إضافة / إدارة العلاقات"
    }
};

let currentLang = "en";

const IDB_KEYS = {
    STORE: "DC_STORE",
    THEME: "DC_THEME",
    LANGUAGE: "DC_LANGUAGE",
    DEFAULT_TABLE: "DC_DEFAULT_DATA_TABLE",
    DEFAULT_REVISION: "DC_DEFAULT_REVISION",
    VARIABLES: "DC_VARIABLES",
    WORKSPACES_ORDER: "DC_WORKSPACES_ORDER",
    TABLES_ORDER: "DC_TABLES_ORDER"
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
  family: string;
  revision: number;
  history: { [rev: number]: DataSetVersion };
  relations?: { subTable: string; foreignKey: string }[];
  calculatedFields?: Record<string, string>;
}

interface Store {
  [dataTableName: string]: DataSet;
}

let activeFormulaInput: HTMLInputElement | null = null;
let isSwitchingRecord = false;

function scrollToTarget(element: HTMLElement) {
    setTimeout(() => {
        const offset = 80; // Leaves space for the fixed top header
        const distance = element.getBoundingClientRect().top - offset;
        
        const mainContainer = document.querySelector('.ms-welcome__main');
        if (mainContainer) {
            mainContainer.scrollBy({ top: distance, behavior: 'smooth' });
        }
        window.scrollBy({ top: distance, behavior: 'smooth' });
    }, 50);
}

function toggleAccordion(header: HTMLElement, content: HTMLElement, levelClass: string) {
    const isActive = header.classList.contains("active");
    
    if (!isActive) {
        // Close all other accordions of the same level
        const others = document.querySelectorAll(`.${levelClass}.active`);
        others.forEach(other => {
            other.classList.remove("active");
            const otherContent = other.nextElementSibling as HTMLElement;
            if (otherContent && otherContent.classList.contains("accordion-content")) {
                otherContent.classList.remove("show");
            }
            if (other.id === "formula-builder-accordion") {
                activeFormulaInput = null;
            }
        });
        
        // Open this one
        header.classList.add("active");
        content.classList.add("show");
        
        // Scroll into view
        scrollToTarget(header);
    } else {
        // Close this one
        header.classList.remove("active");
        content.classList.remove("show");
        if (header.id === "formula-builder-accordion") {
            activeFormulaInput = null;
        }
    }
}

export function t(key: string): string {
    return TRANSLATIONS[currentLang]?.[key] || TRANSLATIONS["en"][key] || key;
}

export function applyTranslations() {
    document.body.dir = currentLang === "ar" ? "rtl" : "ltr";
    document.querySelectorAll("[data-i18n]").forEach(el => {
        const key = el.getAttribute("data-i18n");
        if (key && TRANSLATIONS[currentLang] && TRANSLATIONS[currentLang][key]) {
            (el as HTMLElement).innerText = TRANSLATIONS[currentLang][key];
        }
    });
}

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
    document.getElementById("backup-button").onclick = backupData;
    document.getElementById("restore-button").onclick = triggerRestore;
    document.getElementById("restore-file-input").addEventListener("change", restoreData);
    document.getElementById("refresh-formulas-button").onclick = () => refreshFormulas(false);
    document.getElementById("convert-values-button").onclick = convertToValues;
    document.getElementById("settings-button").onclick = toggleSettings;
    const manageWsBtn = document.getElementById("manage-workspaces-btn");
    if (manageWsBtn) manageWsBtn.onclick = manageWorkspaces;

    
    const editorSaveBtn = document.getElementById("editor-save-btn");
    if (editorSaveBtn) editorSaveBtn.onclick = () => saveGridEditor(true);
    const editorCancelBtn = document.getElementById("editor-cancel-btn");
    if (editorCancelBtn) editorCancelBtn.onclick = cancelGridEditor;

    // Setup Formula Builder Accordion
    const formulaAccordion = document.getElementById("formula-builder-accordion");
    const formulaContent = document.getElementById("formula-builder-content");
    if (formulaAccordion && formulaContent) {
      formulaAccordion.classList.add("accordion-top-level");
      formulaAccordion.addEventListener("click", () => {
        toggleAccordion(formulaAccordion, formulaContent, "accordion-top-level");
      });
    }

    // Setup Global Variables Accordion
    const varAccordion = document.getElementById("variables-accordion");
    const varContent = document.getElementById("variables-content");
    if (varAccordion && varContent) {
      varAccordion.classList.add("accordion-top-level");
      varAccordion.addEventListener("click", () => {
        toggleAccordion(varAccordion, varContent, "accordion-top-level");
      });
    }
    const addVarBtn = document.getElementById("add-variable-btn");
    if (addVarBtn) addVarBtn.onclick = manageVariable;

    // Setup Global Relations Accordion
    const relAccordion = document.getElementById("relations-accordion");
    const relContent = document.getElementById("relations-content");
    if (relAccordion && relContent) {
      relAccordion.classList.add("accordion-top-level");
      relAccordion.addEventListener("click", () => {
        toggleAccordion(relAccordion, relContent, "accordion-top-level");
      });
    }
    const manageRelBtn = document.getElementById("manage-global-relations-btn");
    if (manageRelBtn) manageRelBtn.onclick = async () => {
        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        const tables = Object.keys(store);
        const res = await customFormPrompt("Manage Relations", "Select a table to manage its relations:", [{ id: "table", label: "Main Data Table", type: "select", options: tables }], "Next");
        if (res && res.table) manageRelations(res.table);
    };

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

        // Live Sub-table Tracking (Auto-Switching)
        try {
            await Excel.run(async (context) => {
                const worksheets = context.workbook.worksheets;
                worksheets.load("items/name");
                await context.sync();
                
                let subEditorSheet: Excel.Worksheet | null = null;
                for (const sheet of worksheets.items) {
                    const pProp = sheet.customProperties.getItemOrNullObject("SheetPurpose");
                    pProp.load("value");
                    await context.sync();
                    if (!pProp.isNullObject && pProp.value === "SubDataEditor") {
                        subEditorSheet = sheet;
                        break;
                    }
                }
                if (!subEditorSheet) return;

                const activeSheet = context.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");
                await context.sync();
                if (activeSheet.name === subEditorSheet.name) return;

                const editorTableProp = subEditorSheet.customProperties.getItemOrNullObject("EditingTable");
                const currentFilterValueProp = subEditorSheet.customProperties.getItemOrNullObject("FilterValue");
                const editorColsProp = subEditorSheet.customProperties.getItemOrNullObject("EditorColumns");
                const mainTableProp = subEditorSheet.customProperties.getItemOrNullObject("MainTable");
                editorTableProp.load("value");
                currentFilterValueProp.load("value");
                editorColsProp.load("value");
                mainTableProp.load("value");
                await context.sync();
                if (editorTableProp.isNullObject || editorTableProp.value === "") return;
                const editorTargetTable = editorTableProp.value;
                const mainTableValue = !mainTableProp.isNullObject ? mainTableProp.value : "";

                const range = context.workbook.getSelectedRange();
                const tables = range.getTables();
                tables.load("items/name");
                await context.sync();
                if (tables.items.length === 0) return;

                const selectedTable = tables.items[0];
                const storedData = await idbGet(IDB_KEYS.STORE);
                if (!storedData) return;
                const store = JSON.parse(storedData);
                let matchedMainTable = mainTableValue;
                if (!matchedMainTable) {
                    for (const key of Object.keys(store)) {
                        if (selectedTable.name.includes(key.replace(/\s+/g, ""))) {
                            matchedMainTable = key;
                            break;
                        }
                    }
                }
                if (!matchedMainTable) return;
                
                const mainDataSet = store[matchedMainTable];
                const relation = (mainDataSet.relations || []).find((r: any) => r.subTable === editorTargetTable);
                // If there is no relation, or if the selected table is not the main table for the active sub-editor, do nothing.
                if (!relation || matchedMainTable !== mainTableValue) return;

                const idField = mainDataSet.idField || mainDataSet.fields[0];
                const idColumn = selectedTable.columns.getItemOrNullObject(idField);
                const tableRange = selectedTable.getRange();
                idColumn.load("index");
                range.load("rowIndex");
                tableRange.load("rowIndex");
                await context.sync();
                if (idColumn.isNullObject) return;

                const rowIdxInTable = range.rowIndex - tableRange.rowIndex;
                if (rowIdxInTable <= 0) return;

                const cell = selectedTable.getDataBodyRange().getCell(rowIdxInTable - 1, idColumn.index);
                cell.load("values");
                await context.sync();

                const newId = String(cell.values[0][0]);
                const currentId = currentFilterValueProp.isNullObject ? "" : currentFilterValueProp.value;

                console.log(`[Live Sync] Checking switch: newId=${newId}, currentId=${currentId}`);
                if (newId && newId !== currentId && newId.trim() !== "") {
                    if (isSwitchingRecord) return;
                    isSwitchingRecord = true;
                    const editorCols = (editorColsProp.value || "").split(",").filter(c => c);
                    setTimeout(async () => {
                        try {
                            console.log(`[Live Sync] Executing switch to ${newId}`);
                            await switchGridEditorRecord(newId, relation.foreignKey, editorTargetTable, editorCols, subEditorSheet!.name, matchedMainTable);
                        } catch (err) {
                            console.error("[Live Sync] Switch error:", err);
                        } finally {
                            isSwitchingRecord = false;
                        }
                    }, 10);
                }
            });
        } catch (e) { console.error("[Live Sync] Setup/Detection error:", e); }
    });

    await migrateFromLocalStorage(); // Migrate old data if present
    loadSettings();
    renderDashboard();
    renderFormulaBuilder();

    // Setup Worksheet Activation Listener
    try {
        await Excel.run(async (context) => {
            context.workbook.worksheets.onActivated.add(handleSheetActivation);
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const purposeProp = sheet.customProperties.getItemOrNullObject("SheetPurpose");
            const tableProp = sheet.customProperties.getItemOrNullObject("EditingTable");
            const mainTableProp = sheet.customProperties.getItemOrNullObject("MainTable");
            purposeProp.load("value");
            tableProp.load("value");
            mainTableProp.load("value");
            await context.sync();
            const isLive = !mainTableProp.isNullObject && mainTableProp.value !== "";
            if (!purposeProp.isNullObject && purposeProp.value === "SubDataEditor") showEditorView(!tableProp.isNullObject ? tableProp.value : "Unknown", isLive);
        });
    } catch (e) { console.error("Sheet listener error:", e); }
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

export function customManageListPrompt(title: string, message: string, items: string[], addBtnText: string, allowRename: boolean = true): Promise<{ original: string, newName: string, isDeleted: boolean, isNew: boolean }[] | null> {
    return new Promise((resolve) => {
        const modal = document.getElementById("custom-manage-list-modal");
        const titleEl = document.getElementById("manage-list-title");
        const messageEl = document.getElementById("manage-list-message");
        const listEl = document.getElementById("manage-list-items");
        const btnAdd = document.getElementById("manage-list-add-btn");
        const btnSave = document.getElementById("manage-list-save-btn");
        const btnCancel = document.getElementById("manage-list-cancel-btn");

        if (!modal || !titleEl || !messageEl || !listEl || !btnAdd || !btnSave || !btnCancel) {
            resolve(null);
            return;
        }

        titleEl.innerText = title;
        messageEl.innerText = message;
        listEl.innerHTML = "";
        btnAdd.innerHTML = `<i class="ms-Icon ms-Icon--Add" style="margin-right:8px;"></i>${addBtnText}`;

        const createLi = (original: string, isNew: boolean) => {
            const li = document.createElement("li");
            li.className = "sortable-item";
            li.draggable = true;
            li.dataset.original = original;
            li.dataset.isNew = String(isNew);
            li.dataset.isDeleted = "false";
            
            li.innerHTML = `
                <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: #fff;">
                    <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888; cursor: grab;"></i>
                    <input type="text" class="ms-TextField-field list-name-input" value="${original}" placeholder="Name" style="flex: 1; padding: 4px 8px; font-weight: 600;" ${(!allowRename && !isNew) ? 'disabled' : ''} />
                    <button type="button" class="icon-btn list-delete-btn" style="margin-left: 8px; color: #d13438;" title="Delete">
                        <i class="ms-Icon ms-Icon--Delete"></i>
                    </button>
                </div>
            `;
            const inputEl = li.querySelector("input[type='text']") as HTMLInputElement;
            inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
            
            const deleteBtn = li.querySelector(".list-delete-btn") as HTMLButtonElement;
            deleteBtn.onclick = async () => {
                const itemName = inputEl.value || "this item";
                const confirmed = await customConfirm("Mark for Deletion", `Are you sure you want to delete '${itemName}'? It will be permanently removed when you save changes.`, "Yes, delete");
                if (confirmed) {
                    li.dataset.isDeleted = "true";
                    li.style.display = "none";
                }
            };
            
            li.addEventListener("dragstart", () => li.classList.add("dragging"));
            li.addEventListener("dragend", () => li.classList.remove("dragging"));
            return li;
        };

        items.forEach(item => {
            listEl.appendChild(createLi(item, false));
        });

        btnAdd.onclick = () => {
            const li = createLi("", true);
            listEl.appendChild(li);
            li.scrollIntoView({ behavior: 'smooth' });
            const input = li.querySelector("input[type='text']") as HTMLInputElement;
            if (input) input.focus();
        };

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

        const cleanup = () => {
            modal.style.display = "none";
            btnAdd.onclick = null;
            btnSave.onclick = null;
            btnCancel.onclick = null;
            listEl.removeEventListener("dragover", handleDragOver);
        };

        btnSave.onclick = () => {
            const results = Array.from(listEl.children).map((li) => {
                const original = (li as HTMLElement).dataset.original as string;
                const isNew = (li as HTMLElement).dataset.isNew === "true";
                const newName = ((li as HTMLElement).querySelector(".list-name-input") as HTMLInputElement).value.trim();
                const isDeleted = (li as HTMLElement).dataset.isDeleted === "true";
                return { original, newName, isDeleted, isNew };
            }).filter(r => !(r.isNew && r.newName === "") && !(r.isNew && r.isDeleted));
            
            cleanup();
            resolve(results);
        };

        btnCancel.onclick = () => { cleanup(); resolve(null); };
    });
}

export function customManageColumnsPrompt(title: string, message: string, items: string[], idField: string, calcFields: Record<string, string> = {}, varNames: string[] = [], tablesWithFields: Record<string, string[]> = {}): Promise<{changes: {oldName: string, newName: string, formula: string}[], saveAsNewRevision: boolean} | null> {
  return new Promise((resolve) => {
      const modal = document.getElementById("custom-sort-modal");
      const titleEl = document.getElementById("sort-modal-title");
      const messageEl = document.getElementById("sort-modal-message");
      const listEl = document.getElementById("sort-modal-list");
      const btnAddColumn = document.getElementById("sort-modal-add-column");
      const btnDeleteSelected = document.getElementById("sort-modal-delete-selected");
      const btnSaveCurrent = document.getElementById("sort-modal-save-current");
      const btnSaveNew = document.getElementById("sort-modal-save-new");
      const btnCancel = document.getElementById("sort-modal-cancel");

      if (!modal || !titleEl || !messageEl || !listEl || !btnAddColumn || !btnDeleteSelected || !btnSaveCurrent || !btnSaveNew || !btnCancel) {
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
          const formula = calcFields[item] || "";
          
          li.innerHTML = `
                      <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: #fff;">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-right: 8px;" ${isId ? 'disabled title="Cannot delete ID column"' : ''} />
                  <input type="text" class="ms-TextField-field col-name-input" value="${item}" placeholder="Column Name" style="flex: 1; padding: 4px 8px; font-weight: 600;" />
                  ${isId ? '<span style="font-size:10px; color:#0078d4; margin-left:8px; font-weight:bold;" title="Primary ID Column">(ID)</span>' : ''}
                  <input type="hidden" class="col-formula-hidden" value="${formula.replace(/"/g, '&quot;')}" />
                  <button type="button" class="icon-btn edit-formula-btn" style="margin-left: 8px; color: ${formula ? '#0078d4' : '#888'};" title="${formula ? 'Edit Formula: ' + formula.replace(/"/g, '&quot;') : 'Add Formula'}">
                      <i class="ms-Icon ms-Icon--Variable"></i>
                  </button>
              </div>
          `;
          
          const inputs = li.querySelectorAll("input[type='text']");
          inputs.forEach(inputEl => {
              inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
          });
          
          const formulaBtn = li.querySelector(".edit-formula-btn") as HTMLButtonElement;
          const formulaHidden = li.querySelector(".col-formula-hidden") as HTMLInputElement;
          const nameInput = li.querySelector(".col-name-input") as HTMLInputElement;

          formulaBtn.addEventListener("click", async () => {
              const currentFields = Array.from(listEl.children).map(child => (child.querySelector(".col-name-input") as HTMLInputElement).value.trim()).filter(Boolean);
              const colName = nameInput.value.trim() || item || "New Column";
              
              modal.style.display = "none";
              const res = await customFormPrompt(`Formula for '${colName}'`, "Define calculated formula (leave empty to remove):", [
                  { id: "vFormula", label: "Formula Definition", type: "formula", varsList: varNames, tablesWithFields: tablesWithFields, fieldsList: currentFields, value: formulaHidden.value }
              ]);
              modal.style.display = "flex";

              if (res !== null) {
                  const newFormula = res.vFormula || "";
                  formulaHidden.value = newFormula;
                  if (newFormula) {
                      formulaBtn.style.color = "#0078d4";
                      formulaBtn.title = "Edit Formula: " + newFormula.replace(/"/g, '&quot;');
                  } else {
                      formulaBtn.style.color = "#888";
                      formulaBtn.title = "Add Formula";
                  }
              }
          })
          
          li.addEventListener("dragstart", () => li.classList.add("dragging"));
          li.addEventListener("dragend", () => li.classList.remove("dragging"));
          
          listEl.appendChild(li);
      });
      
      btnAddColumn.onclick = () => {
          const li = document.createElement("li");
          li.className = "sortable-item";
          li.draggable = true;
          li.dataset.original = "";
          
          li.innerHTML = `
              <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: #fff;">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-right: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-right: 8px;" />
                  <input type="text" class="ms-TextField-field col-name-input" value="" placeholder="New Column Name" style="flex: 1; padding: 4px 8px; font-weight: 600;" />
                  <input type="hidden" class="col-formula-hidden" value="" />
                  <button type="button" class="icon-btn edit-formula-btn" style="margin-left: 8px; color: #888;" title="Add Formula">
                      <i class="ms-Icon ms-Icon--Variable"></i>
                  </button>
              </div>
          `;
          const inputs = li.querySelectorAll("input[type='text']");
          inputs.forEach(inputEl => { inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); }); });

          const formulaBtn = li.querySelector(".edit-formula-btn") as HTMLButtonElement;
          const formulaHidden = li.querySelector(".col-formula-hidden") as HTMLInputElement;
          const nameInput = li.querySelector(".col-name-input") as HTMLInputElement;

          formulaBtn.addEventListener("click", async () => {
              const currentFields = Array.from(listEl.children).map(child => (child.querySelector(".col-name-input") as HTMLInputElement).value.trim()).filter(Boolean);
              const colName = nameInput.value.trim() || "New Column";
              
              modal.style.display = "none";
              const res = await customFormPrompt(`Formula for '${colName}'`, "Define calculated formula (leave empty to remove):", [
                  { id: "vFormula", label: "Formula Definition", type: "formula", varsList: varNames, tablesWithFields: tablesWithFields, fieldsList: currentFields, value: formulaHidden.value }
              ]);
              modal.style.display = "flex";

              if (res !== null) {
                  const newFormula = res.vFormula || "";
                  formulaHidden.value = newFormula;
                  if (newFormula) {
                      formulaBtn.style.color = "#0078d4";
                      formulaBtn.title = "Edit Formula: " + newFormula.replace(/"/g, '&quot;');
                  } else {
                      formulaBtn.style.color = "#888";
                      formulaBtn.title = "Add Formula";
                  }
              }
          });

          li.addEventListener("dragstart", () => li.classList.add("dragging"));
          li.addEventListener("dragend", () => li.classList.remove("dragging"));
          listEl.appendChild(li);
          li.scrollIntoView({ behavior: 'smooth' });
      };

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
          btnAddColumn.onclick = null;
          btnDeleteSelected.onclick = null;
          btnSaveCurrent.onclick = null;
          btnSaveNew.onclick = null;
          btnCancel.onclick = null;
          listEl.removeEventListener("dragover", handleDragOver);
      };

      const getChanges = () => {
          return Array.from(listEl.children).map((li) => {
              const original = (li as HTMLElement).dataset.original as string;
              const newName = ((li as HTMLElement).querySelector(".col-name-input") as HTMLInputElement).value.trim();
              const formula = ((li as HTMLElement).querySelector(".col-formula-hidden") as HTMLInputElement).value.trim();
              return { oldName: original, newName: newName || original, formula: formula };
          }).filter(c => c.newName !== "");
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
      
      step1Div.style.display = "block";
      step2Div.style.display = "none";
      btnDeleteSelected.style.display = "block";
      btnNext.style.display = "block";
      btnBack.style.display = "none";
      btnOk.style.display = "none";
      
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

export async function applyCalculatedFields(records: any[], fields: string[], calculatedFields: Record<string, string> | undefined, store?: Store, dataTableName?: string) {
    if (!calculatedFields || Object.keys(calculatedFields).length === 0) return records;
    
    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};

    const evaluateVar = (vName: string, visited: Set<string>): any => {
        if (visited.has(vName)) throw new Error("Circular reference detected");
        visited.add(vName);
        const vForm = variables[vName];
        if (!vForm) return 0;
        const vDC = {
            SUM: (t: string, c: string) => store?.[t]?.records?.reduce((a:number, b:any) => a + (Number(b[c])||0), 0) || 0,
            COUNT: (t: string) => store?.[t]?.records?.length || 0,
            VAR: (v: string) => evaluateVar(v, new Set(visited))
        };
        const func = new Function('store', 'DC', `return ${vForm};`);
        return func(store, vDC);
    };

    const DC = {
        SUM: (subTable: string, sumCol: string, fkCol?: string, fkValue?: any) => {
            if (!store || !store[subTable]) return 0;
            let total = 0;
            store[subTable].records.forEach((r: any) => {
                if (!fkCol || String(r[fkCol]) === String(fkValue)) {
                    total += Number(r[sumCol]) || 0;
                }
            });
            return total;
        },
        COUNT: (subTable: string, fkCol?: string, fkValue?: any) => {
            if (!store || !store[subTable]) return 0;
            let count = 0;
            store[subTable].records.forEach((r: any) => {
                if (!fkCol || String(r[fkCol]) === String(fkValue)) count++;
            });
            return count;
        },
        VAR: (varName: string) => {
            try {
                return evaluateVar(varName, new Set());
            } catch(e) { return 0; }
        }
    };

    const compiledFormulas: Record<string, Function> = {};
    for (const [col, formula] of Object.entries(calculatedFields)) {
        if (!fields.includes(col)) {
            fields.push(col);
        }
        let jsFormula = formula.replace(/\[([^\]]+)\]/g, '(record["$1"] ?? "")');
        try {
            compiledFormulas[col] = new Function('record', 'store', 'DC', `return ${jsFormula};`);
        } catch (e) {
            console.error(`Invalid formula for ${col}:`, e);
        }
    }

    records.forEach(record => {
        for (const [col, func] of Object.entries(compiledFormulas)) {
            try {
                record[col] = func(record, store, DC);
            } catch (e) {
                record[col] = "ERROR";
            }
        }
    });

    return records;
}

export interface FormField {
    id: string;
    label: string;
    type: 'text' | 'select' | 'autocomplete' | 'checkboxes' | 'formula';
    options?: string[];
    value?: string;
    disabled?: boolean;
    fieldsList?: string[];
    varsList?: string[];
    tablesWithFields?: Record<string, string[]>;
    dependsOn?: string;
    optionsMap?: Record<string, string[]>;
}

export function customFormPrompt(title: string, message: string, fields: FormField[], confirmText: string = "Save"): Promise<Record<string, string> | null> {
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
        btnOk.innerText = confirmText;
        inputsContainer.innerHTML = "";

        const inputElements: { id: string, el: HTMLInputElement | HTMLSelectElement | HTMLDivElement, type: string }[] = [];

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
                inputElements.push({ id: f.id, el: select, type: f.type });
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
                inputElements.push({ id: f.id, el: input, type: f.type });
            } else if (f.type === 'checkboxes') {
                const cbContainer = document.createElement("div");
                cbContainer.style.maxHeight = "150px";
                cbContainer.style.overflowY = "auto";
                cbContainer.style.border = "1px solid var(--border-color)";
                cbContainer.style.padding = "8px";
                cbContainer.style.borderRadius = "2px";
                (f.options || []).forEach(opt => {
                    const lbl = document.createElement("label");
                    lbl.style.display = "block";
                    lbl.style.marginBottom = "4px";
                    const cb = document.createElement("input");
                    cb.type = "checkbox";
                    cb.value = opt;
                    cb.checked = true;
                    lbl.appendChild(cb);
                    lbl.appendChild(document.createTextNode(" " + opt));
                    cbContainer.appendChild(lbl);
                });
                fieldDiv.appendChild(cbContainer);
                inputElements.push({ id: f.id, el: cbContainer, type: 'checkboxes' });
            } else if (f.type === 'formula') {
                const wrap = document.createElement("div");
                wrap.style.border = "1px solid var(--border-color)";
                wrap.style.borderRadius = "4px";
                wrap.style.overflow = "hidden";
                wrap.style.backgroundColor = "#fdfdfd";
                
                let fieldsHtml = "";
                if (f.fieldsList && f.fieldsList.length > 0) {
                    fieldsHtml = `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center;">
                        <span style="font-size:9px; font-weight:bold; color:#888;">FIELDS:</span>
                        ${f.fieldsList.map(i => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#e1dfdd; border:1px solid #ccc; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="[${i}]">${i}</button>`).join('')}
                    </div>`;
                }
                
                const tablesWithFields = f.tablesWithFields || {};

                const toolbarHtml = `
                <div style="background: #f3f2f1; padding: 6px; border-bottom: 1px solid var(--border-color); display: flex; flex-direction: column; gap: 6px;">
                    <div style="display:flex; gap:4px;">
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" + ">+</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" - ">-</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" * ">*</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" / ">/</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" ( ">(</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:#fff; border:1px solid #ccc; border-radius:2px; cursor:pointer;" data-op=" ) ">)</button>
                    </div>
                    <div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;">
                        <span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">FUNCS:</span>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="DC.SUM('Table', 'Col')">SUM()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="DC.COUNT('Table')">COUNT()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="DC.VAR('VarName')">VAR()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op=" ( condition ? true_val : false_val ) ">IF()</button>
                    </div>
                    ${(f.varsList && f.varsList.length > 0) ? `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;"><span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">VARS:</span>${f.varsList.map(v => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#c7e0f4; border:1px solid #99c9ef; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="DC.VAR('${v}')">${v}</button>`).join('')}</div>` : ''}
                    ${Object.keys(tablesWithFields).length > 0 ? `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;"><span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">TABLES:</span>${Object.keys(tablesWithFields).map(t => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#fce1cb; border:1px solid #f9c79f; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="'${t}'">${t}</button>`).join('')}</div>` : ''}
                    <div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;">
                        <span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">FIELDS:</span>
                        <select class="field-table-selector" style="font-size:10px; padding:2px; max-width:100px; border:1px solid #ccc; border-radius:2px; flex-shrink:0;">
                            ${(f.fieldsList && f.fieldsList.length > 0) ? '<option value="_CURRENT_">Current Table</option>' : '<option value="">-- Select Table --</option>'}
                            ${Object.keys(tablesWithFields).map(t => `<option value="${t}">${t}</option>`).join('')}
                        </select>
                        <span class="dynamic-fields-container" style="display:flex; gap:4px; flex-shrink:0;">
                            ${(f.fieldsList || []).map(i => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#e1dfdd; border:1px solid #ccc; border-radius:2px; cursor:pointer; white-space:nowrap;" data-op="[${i}]">${i}</button>`).join('')}
                        </span>
                    </div>
                </div>
                `;
                
                wrap.innerHTML = toolbarHtml;

                const input = document.createElement("textarea");
                input.className = "ms-TextField-field";
                input.style.width = "100%";
                input.style.border = "none";
                input.style.padding = "6px 8px";
                input.style.fontSize = "11px";
                input.style.minHeight = "44px";
                input.style.resize = "vertical";
                input.style.fontFamily = "monospace";
                input.placeholder = "Type or click above to build formula...";
                if (f.value !== undefined) input.value = f.value;
                
                wrap.appendChild(input);

                const appendText = (txt: string) => { 
                    const start = input.selectionStart || 0;
                    const end = input.selectionEnd || 0;
                    const val = input.value;
                    input.value = val.substring(0, start) + txt + val.substring(end);
                    input.selectionStart = input.selectionEnd = start + txt.length;
                    input.focus(); 
                };
        
                const opBtns = wrap.querySelectorAll(".col-insert-op");
                opBtns.forEach(btn => {
                    btn.addEventListener("click", () => appendText((btn as HTMLElement).dataset.op || ""));
                });

                const fieldSelect = wrap.querySelector(".field-table-selector") as HTMLSelectElement;
                const dynamicContainer = wrap.querySelector(".dynamic-fields-container") as HTMLElement;
                if (fieldSelect && dynamicContainer) {
                    fieldSelect.addEventListener("change", () => {
                        dynamicContainer.innerHTML = "";
                        const isCurrent = fieldSelect.value === "_CURRENT_";
                        const fieldsToShow = isCurrent ? (f.fieldsList || []) : (tablesWithFields[fieldSelect.value] || []);
                        fieldsToShow.forEach(fld => {
                            const btn = document.createElement("button");
                            btn.type = "button";
                            btn.className = "col-insert-op";
                            btn.style.cssText = "font-size:10px; padding:2px 6px; background:#e1dfdd; border:1px solid #ccc; border-radius:2px; cursor:pointer; white-space:nowrap;";
                            const opVal = isCurrent ? `[${fld}]` : `'${fld}'`;
                            btn.dataset.op = opVal;
                            btn.innerText = fld;
                            btn.addEventListener("click", () => appendText(opVal, input));
                            dynamicContainer.appendChild(btn);
                        });
                    });
                }

                fieldDiv.appendChild(wrap);
                inputElements.push({ id: f.id, el: input, type: 'text' });
            } else {
                const input = document.createElement("input");
                input.type = "text";
                input.className = "ms-TextField-field";
                if (f.disabled) input.disabled = true;
                if (f.value !== undefined) input.value = f.value;
                fieldDiv.appendChild(input);
                inputElements.push({ id: f.id, el: input, type: f.type });
            }

            inputsContainer.appendChild(fieldDiv);
        });

        // Setup dependencies
        fields.forEach(f => {
            if (f.dependsOn && f.optionsMap) {
                const parentInput = inputElements.find(i => i.id === f.dependsOn);
                const childInput = inputElements.find(i => i.id === f.id);
                if (parentInput && childInput) {
                    const updateChild = () => {
                        const pVal = parentInput.el.value;
                        const opts = f.optionsMap![pVal] || [];
                        if (childInput.type === 'select') {
                            const sel = childInput.el as HTMLSelectElement;
                            sel.innerHTML = "";
                            const emptyOpt = document.createElement("option");
                            emptyOpt.value = "-- None --";
                            emptyOpt.text = "-- None --";
                            sel.appendChild(emptyOpt);
                            opts.forEach(o => {
                                const opt = document.createElement("option");
                                opt.value = o;
                                opt.text = o;
                                sel.appendChild(opt);
                            });
                        }
                    };
                    parentInput.el.addEventListener("change", updateChild);
                    updateChild(); // Initial population
                }
            }
        });

        modal.style.display = "flex";

        const cleanup = () => {
            modal.style.display = "none";
            btnOk.onclick = null;
            btnCancel.onclick = null;
        };

        btnOk.onclick = () => {
            const result: Record<string, string> = {};
            let hasValidationErrors = false;
            inputElements.forEach(item => {
                if (item.type === 'checkboxes') {
                    const checked = Array.from((item.el as HTMLDivElement).querySelectorAll('input:checked')).map(cb => (cb as HTMLInputElement).value);
                    result[item.id] = checked.join(",");
                } else {
                    result[item.id] = (item.el as HTMLInputElement).value;
                }
            });
            if (hasValidationErrors) return; // Allow custom handling logic by caller later
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
        { id: "param-foreignKeyField", label: "Link Column", required: true },
        { id: "param-foreignTableName", label: "Target Table Name", required: true },
        { id: "param-foreignReturnField", label: "Target Return Field", required: true }
    ],
    "SORT": [
        { id: "param-sortField", label: "Sort Field", required: true },
        { id: "param-ascending", label: "Ascending (TRUE/FALSE)", required: false },
        { id: "param-dataTableName", label: "Data Table Name", required: false },
        { id: "param-rev", label: "Revision", required: false }
    ],
    "VAR": [
        { id: "param-varName", label: "Variable Name", required: true }
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

    const updateDataLists = async () => {
        const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
        const vars = vStoreRaw ? JSON.parse(vStoreRaw) : {};
        const varKeys = Object.keys(vars);

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

            let options: string[] = [];

            if (ref.id === "param-varName") {
                options = varKeys;
            } else {
                const tName = getTargetTableForInput(ref.id);
                const dataSet = store[tName];
                if (!dataSet) return;

                if (ref.id.toLowerCase().includes("field")) {
                    options = dataSet.fields || [];
                } else if (ref.id === "param-id") {
                    options = (dataSet.records || []).map((r: any) => String(r.__DC_ID__));
                }
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
  
  const lang = await idbGet(IDB_KEYS.LANGUAGE);
  const langSelect = document.getElementById("language-select") as HTMLSelectElement;
  if (lang && langSelect) {
      langSelect.value = lang;
      currentLang = lang;
  }
  applyTranslations();
}
 
export async function loadRangeForCapture(dataName: string, familyName: string, parentName: string) {
  const status = document.getElementById("status-text");
  try {
    await Excel.run(async (context) => {
      if (!dataName || dataName.trim() === "") {
        throw new Error("Please enter a data table name.");
      }

      const range = context.workbook.getSelectedRange();
      range.load(["values", "formulas", "columnIndex"]);
      await context.sync();
      const colStartIndex = range.columnIndex;

      const values = range.values;
      const formulas = range.formulas;
      if (values.length < 2) {
        throw new Error("Select a range with headers and data.");
      }

      let detectedFormulas: { header: string, formula: string }[] = [];
      if (formulas.length > 1) {
          values[0].forEach((h, i) => {
              if (typeof formulas[1][i] === 'string' && formulas[1][i].startsWith('=')) {
                  detectedFormulas.push({ header: String(h), formula: formulas[1][i] });
              }
          });
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
          let relations: any[] | undefined = undefined;
          let calculatedFields: Record<string, string> | undefined = undefined;
          if (store[dataName]) {
              rev = store[dataName].revision || 1;
              history = store[dataName].history || {};
              relations = store[dataName].relations;
              calculatedFields = store[dataName].calculatedFields;
          }

          store[dataName] = {
              dataTableName: dataName,
              family: familyName,
              idField: summary.idField,
              revision: rev,
              history: history,
              fields: summary.fields,
              records: summary.records
          };
          if (relations) store[dataName].relations = relations;
          if (calculatedFields) store[dataName].calculatedFields = calculatedFields;

          await applyCalculatedFields(store[dataName].records, store[dataName].fields, store[dataName].calculatedFields, store, dataName);

          await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
          if (status) { status.innerText = `Saved ${summary.records.length} records in ${dataName}.`; status.style.color = "green"; }

          if (parentName && store[parentName]) {
              const res = await customFormPrompt("Link Sub-table", `Which column in '${dataName}' links to '${parentName}'?`, [
                  { id: "foreignKey", label: "Link Column", type: "select", options: summary.fields }
              ]);
              if (res && res.foreignKey) {
                  store[parentName].relations = store[parentName].relations || [];
                  store[parentName].relations.push({ subTable: dataName, foreignKey: res.foreignKey });
                  await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
              }
          }

    await idbSet(IDB_KEYS.DEFAULT_TABLE, dataName);
    await idbSet(IDB_KEYS.DEFAULT_REVISION, String(rev));

    await renderDashboard();
    await refreshFormulas(true);

    if (detectedFormulas.length > 0) {
            const colLetterToIndex = (letter: string) => {
                let index = 0;
                for (let i = 0; i < letter.length; i++) {
                    index = index * 26 + (letter.charCodeAt(i) - 64);
                }
                return index - 1;
            };
            
            let mappedCalcs: Record<string, { excel: string, dc: string }> = {};
            detectedFormulas.forEach(df => {
                let dcFormula = df.formula.substring(1); // Remove '='
                dcFormula = dcFormula.replace(/\$?[A-Z]+\$?[0-9]+/gi, (match) => {
                    const colMatch = match.match(/[A-Z]+/i);
                    if (colMatch) {
                        const colIdx = colLetterToIndex(colMatch[0].toUpperCase());
                        const relativeIdx = colIdx - colStartIndex;
                        if (relativeIdx >= 0 && relativeIdx < summary.fields.length) {
                            return `[${summary.fields[relativeIdx]}]`;
                        }
                    }
                    return match;
                });
                mappedCalcs[df.header] = { excel: df.formula, dc: dcFormula };
            });
            
            const tablesWithFields: Record<string, string[]> = {};
            Object.keys(store).forEach(t => tablesWithFields[t] = store[t].fields || []);

            const reviewFields: FormField[] = detectedFormulas.map(df => ({
                id: df.header,
                label: `${df.header} (Detected: ${mappedCalcs[df.header].excel})`,
                type: 'formula',
                fieldsList: summary.fields,
                tablesWithFields: tablesWithFields,
                value: mappedCalcs[df.header].dc
            }));
            
            const reviewRes = await customFormPrompt("Review Detected Formulas", "We converted your Excel formulas. Adjust them below using the visual formula builder, or clear the text to skip:", reviewFields, "Save Formulas");
            
            if (reviewRes) {
                let finalCalcs: Record<string, string> = {};
                Object.keys(reviewRes).forEach(k => {
                    if (reviewRes[k].trim() !== "") finalCalcs[k] = reviewRes[k].trim();
                });
                if (Object.keys(finalCalcs).length > 0) {
                    store[dataName].calculatedFields = { ...(store[dataName].calculatedFields || {}), ...finalCalcs };
                    await applyCalculatedFields(store[dataName].records, store[dataName].fields, store[dataName].calculatedFields, store, dataName);
                    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
                    await renderDashboard();
                    await refreshFormulas(true);
                    showStatus(`Mapped formulas for: ${Object.keys(finalCalcs).join(", ")}`, "success");
                }
            }
    }
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

      await applyCalculatedFields(targetDataSet.records, targetDataSet.fields, dataSet.calculatedFields, store, dataTableName);

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

      await applyCalculatedFields(dataSet.records, dataSet.fields, dataSet.calculatedFields, store, dataTableName);

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
  const list = document.getElementById("workspaces-container");
  if (!list) return;
  
  list.innerHTML = "";
  const storedData = await idbGet(IDB_KEYS.STORE);
  const defaultSelect = document.getElementById("default-data-table-select") as HTMLSelectElement;
  const defaultRevSelect = document.getElementById("default-revision-select") as HTMLSelectElement;

  if (!storedData) {
    list.innerHTML = "<div style='font-size: 12px; color: #666; padding: 8px;'>No data tables stored yet.</div>";
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
    list.innerHTML = "<div style='font-size: 12px; color: #666; padding: 8px;'>No data tables stored yet.</div>";
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
  
  // Populate Capture Parent Select
  const captureParentSelect = document.getElementById("data-parent") as HTMLSelectElement;
  if (captureParentSelect) {
      captureParentSelect.innerHTML = `<option value="">-- None --</option>`;
      keys.forEach(key => {
          const opt = document.createElement("option");
          opt.value = key;
          opt.text = key;
          captureParentSelect.appendChild(opt);
      });
  }

   const families: Record<string, string[]> = {};
  keys.forEach(key => {
     const fam = store[key].family || 'Public';
     if (!families[fam]) families[fam] = [];
     families[fam].push(key);
  });

  let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
  let wsOrder: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];
  
  const orderedFamilies: string[] = [];
  wsOrder.forEach(f => {
      if (families[f] || wsOrder.includes(f)) orderedFamilies.push(f);
  });
  Object.keys(families).forEach(f => {
      if (!orderedFamilies.includes(f)) orderedFamilies.push(f);
  });

  let tbOrderRaw = await idbGet(IDB_KEYS.TABLES_ORDER);
  let tbOrder: string[] = tbOrderRaw ? JSON.parse(tbOrderRaw) : [];

  // Populate Workspace Datalist
  const familyList = document.getElementById("family-list") as HTMLDataListElement;
  if (familyList) {
      familyList.innerHTML = "";
      Object.keys(families).forEach(fam => {
          const opt = document.createElement("option");
          opt.value = fam;
          familyList.appendChild(opt);
      });
  }

  // Collect all relations
  const allRelations: { main: string, sub: string, fk: string }[] = [];
  keys.forEach(key => {
      const dataTable = store[key];
      if (dataTable.relations) {
          dataTable.relations.forEach(r => {
              allRelations.push({ main: key, sub: r.subTable, fk: r.foreignKey });
          });
      }
  });

  list.innerHTML = "";

  // Render Global Relations List
  const relList = document.getElementById("global-relations-list");
  if (relList) {
      relList.innerHTML = "";
      if (allRelations.length > 0) {
          allRelations.sort((a, b) => a.main.localeCompare(b.main) || a.sub.localeCompare(b.sub));
          allRelations.forEach(rel => {
              const li = document.createElement("li");
              li.style.listStyle = "none";
              li.style.marginBottom = "8px";
              li.style.display = "flex";
              li.style.justifyContent = "space-between";
              li.style.alignItems = "center";
              li.style.background = "var(--border-color)";
              li.style.padding = "8px";
              li.style.borderRadius = "4px";
  
              li.innerHTML = `
                  <div style="flex:1; min-width: 0; margin-right: 8px;">
                      <div style="font-weight:bold; font-size:13px; color:var(--primary-color); margin-bottom: 2px;">
                          <i class="ms-Icon ms-Icon--Table"></i> ${rel.main} 
                          <i class="ms-Icon ms-Icon--Forward" style="font-size:10px; margin:0 4px;"></i> 
                          <i class="ms-Icon ms-Icon--Table"></i> ${rel.sub}
                      </div>
                      <div style="font-size:11px; color:var(--text-color); opacity:0.8;">Link Column: ${rel.fk}</div>
                  </div>
              `;
  
              const actionDiv = document.createElement("div");
              actionDiv.style.display = "flex";
              actionDiv.style.gap = "4px";
              actionDiv.style.flexShrink = "0";
  
              const openBtn = document.createElement("button");
              openBtn.className = "icon-btn";
              openBtn.style.color = "#0078d4";
              openBtn.title = "Open Split Editor";
              openBtn.innerHTML = '<i class="ms-Icon ms-Icon--SplitObject"></i>';
              openBtn.onclick = () => openRelationEditor(rel.main, rel.sub, rel.fk);
  
              const deleteBtn = document.createElement("button");
              deleteBtn.className = "icon-btn";
              deleteBtn.style.color = "#d13438";
              deleteBtn.title = "Delete Link";
              deleteBtn.innerHTML = '<i class="ms-Icon ms-Icon--Delete"></i>';
              deleteBtn.onclick = async () => {
                  const confirm = await customConfirm("Delete Link", `Are you sure you want to permanently remove the link between '${rel.main}' and '${rel.sub}'?`);
                  if (!confirm) return;
                  const storedData = await idbGet(IDB_KEYS.STORE);
                  if (storedData) {
                      const store = JSON.parse(storedData);
                      if (store[rel.main] && store[rel.main].relations) {
                          store[rel.main].relations = store[rel.main].relations.filter((r: any) => r.subTable !== rel.sub || r.foreignKey !== rel.fk);
                          await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
                          renderDashboard();
                          showStatus("Link removed successfully.", "success");
                      }
                  }
              };
  
              actionDiv.appendChild(openBtn);
              actionDiv.appendChild(deleteBtn);
              li.appendChild(actionDiv);
              relList.appendChild(li);
          });
      } else {
          relList.innerHTML = `<li style="font-size: 12px; color: #666;">No relations defined yet.</li>`;
      }
  }

  orderedFamilies.forEach(fam => {
    const famCard = document.createElement("div");
    famCard.className = "workspace-card";
    famCard.style.marginBottom = "8px";

    const famHeader = document.createElement("button");
    famHeader.className = "accordion accordion-top-level";
    famHeader.style.backgroundColor = "var(--neutral-lighter, #f3f2f1)";
    famHeader.style.border = "1px solid var(--border-color)";
    famHeader.style.borderRadius = "4px";
    famHeader.style.padding = "10px 14px";

    const famHeaderContent = document.createElement("div");
    famHeaderContent.style.display = "flex";
    famHeaderContent.style.justifyContent = "space-between";
    famHeaderContent.style.alignItems = "center";
    famHeaderContent.style.width = "100%";

    const titleSpan = document.createElement("span");
    titleSpan.innerHTML = `<i class="ms-Icon ms-Icon--FabricFolder" style="margin-right: 8px; color: var(--primary-color);"></i> <span style="font-weight: 600; font-size: 14px;">${fam}</span>`;
    famHeaderContent.appendChild(titleSpan);
    
    const manageTablesBtn = document.createElement("button");
    manageTablesBtn.className = "ms-Button ms-Button--default";
    manageTablesBtn.style.minWidth = "auto";
    manageTablesBtn.style.padding = "0 8px";
    manageTablesBtn.style.height = "24px";
    manageTablesBtn.style.lineHeight = "24px";
    manageTablesBtn.innerHTML = `<span class="ms-Button-label" style="font-size: 11px;"><i class="ms-Icon ms-Icon--Settings" style="margin-right: 4px;"></i>Tables</span>`;
    manageTablesBtn.onclick = (e) => { e.stopPropagation(); manageWorkspaceTables(fam); };
    famHeaderContent.appendChild(manageTablesBtn);

    famHeader.appendChild(famHeaderContent);

    const famContent = document.createElement("div");
    famContent.className = "accordion-content";
    famContent.style.padding = "0 12px 12px 12px";

    famHeader.onclick = () => {
        toggleAccordion(famHeader, famContent, "accordion-top-level");
    };

    famCard.appendChild(famHeader);
    famCard.appendChild(famContent);
    list.appendChild(famCard);

    const tablesInFam = families[fam] || [];
    tablesInFam.sort((a, b) => {
        const idxA = tbOrder.indexOf(a);
        const idxB = tbOrder.indexOf(b);
        if (idxA === -1 && idxB === -1) return a.localeCompare(b);
        if (idxA === -1) return 1;
        if (idxB === -1) return -1;
        return idxA - idxB;
    }).forEach(key => {
    const dataTable = store[key];
    const count = dataTable.records ? dataTable.records.length : 0;
    const rev = dataTable.revision || 1;
    
    const tableCard = document.createElement("div");
    tableCard.style.marginTop = "12px";
    
    const header = document.createElement("button");
    header.className = "accordion accordion-table-level";
    header.style.padding = "8px 12px";
    header.style.border = "1px solid var(--border-color)";
    header.style.borderRadius = "4px";
    header.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; width:100%;">
          <div><i class="ms-Icon ms-Icon--Table" style="margin-right: 8px;"></i> ${key}</div>
          <div style="font-size:12px; font-weight:normal; opacity:0.8;">Rev ${rev} &bull; ${count} rows</div>
      </div>
    `;
    
    const details = document.createElement("div");
    details.className = "accordion-content";
    
    header.onclick = () => {
      toggleAccordion(header, details, "accordion-table-level");
    };

        const createGridBtn = (icon: string, text: string, colorCls: string, onClick: () => void) => {
            const btn = document.createElement("button");
            btn.className = `action-grid-btn ${colorCls}`;
            btn.style.display = "flex";
            btn.style.flexDirection = "column";
            btn.style.alignItems = "center";
            btn.style.justifyContent = "center";
            btn.style.padding = "8px";
            btn.style.border = "1px solid var(--border-color)";
            btn.style.borderRadius = "4px";
            btn.style.backgroundColor = "var(--surface-color, #fff)";
            btn.style.cursor = "pointer";
            btn.style.minHeight = "60px";
            btn.style.textAlign = "center";
            btn.style.boxShadow = "0 1px 2px rgba(0,0,0,0.05)";
            btn.style.transition = "all 0.15s ease-in-out";
            
            btn.onmouseover = () => { btn.style.backgroundColor = "var(--neutral-lighter, #f3f2f1)"; };
            btn.onmouseout = () => { btn.style.backgroundColor = "var(--surface-color, #fff)"; };
            
            let iconColor = "var(--text-color)";
            if (colorCls === 'primary') iconColor = "#0078d4";
            else if (colorCls === 'danger') iconColor = "#d13438";

            btn.innerHTML = `<i class="ms-Icon ms-Icon--${icon}" style="font-size: 16px; margin-bottom: 6px; color: ${iconColor};"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${text}</span>`;
            btn.onclick = onClick;
            return btn;
        };

        const insertTableBtn = createGridBtn("Table", "Insert to Sheet", "default", () => insertTable(key));
        const replaceBtn = createGridBtn("Sync", "Replace Version", "default", () => replaceTableData(key));
        const appendBtn = createGridBtn("Add", "Append Data", "default", () => appendTableData(key));
        const snapshotBtn = createGridBtn("Camera", "Snapshot", "primary", () => createSnapshot(key));
        const deleteVersionBtn = createGridBtn("RemoveEvent", "Del Version", "danger", () => deleteCurrentVersion(key));
        const deleteBtn = createGridBtn("Delete", "Del Table", "danger", () => deleteDataTable(key));
        const exportCSVBtn = createGridBtn("Download", "Export CSV", "default", () => exportCSV(key));
        const insertDropdownBtn = createGridBtn("Dropdown", "Headers Dropdown", "default", () => insertDropdown(key));
        const editBtn = createGridBtn("Edit", "Form Editor", "primary", () => loadRecordForEdit(key));
        const gridEditBtn = createGridBtn("GridViewSmall", "Grid Editor", "primary", () => openGridEditor(key));
        const duplicateRecordBtn = createGridBtn("Copy", "Clone Record", "default", () => duplicateRecordPrompt(key));
        const editColumnsBtn = createGridBtn("Sort", "Manage Columns", "default", () => manageColumns(key));
        const moveWorkspaceBtn = createGridBtn("FabricFolder", "Move Workspace", "default", () => moveTableWorkspace(key));
        const captureRevBtn = createGridBtn("Camera", "New Revision", "primary", () => captureNewRevision(key));
        const cloneSubBtn = createGridBtn("Copy", "Clone Sub-records", "default", () => cloneSubRecordsPrompt(key));

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

        revSelect.onchange = async () => {
          const isLatest = parseInt(revSelect.value) === rev;
          replaceBtn.innerHTML = `<i class="ms-Icon ms-Icon--Sync" style="font-size: 16px; margin-bottom: 6px; color: var(--text-color);"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${isLatest ? 'Replace Version' : 'Replace Rev Data'}</span>`;
          snapshotBtn.innerHTML = `<i class="ms-Icon ms-Icon--${isLatest ? 'Camera' : 'Undo'}" style="font-size: 16px; margin-bottom: 6px; color: #0078d4;"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${isLatest ? 'Snapshot' : 'Restore as Active'}</span>`;

          if (defaultSelect && defaultSelect.value === key && defaultRevSelect) {
              if (defaultRevSelect.value !== revSelect.value) {
                  defaultRevSelect.value = revSelect.value;
                  await idbSet(IDB_KEYS.DEFAULT_REVISION, revSelect.value);
              }
          }
        };

        const accordions: { header: HTMLElement, content: HTMLElement, title: string, theme: 'default' | 'danger' | 'fast', isGrid: boolean }[] = [];

        // Accordion Builder Helper
        const buildAccordion = (title: string, elements: HTMLElement[], theme: 'default' | 'danger' | 'fast' = 'default', defaultOpen: boolean = false, isGrid: boolean = true) => {
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
            if (isGrid) {
                accContent.className = "inner-accordion-content";
                accContent.style.display = defaultOpen ? "grid" : "none";
                accContent.style.gridTemplateColumns = "1fr 1fr";
                accContent.style.gap = "8px";
                accContent.style.paddingTop = "8px";
            } else {
                accContent.className = "inner-accordion-content flex-col";
                accContent.style.display = defaultOpen ? "flex" : "none";
            }
            
            elements.forEach(el => accContent.appendChild(el));
            
            const accObj = { header: accHeader, content: accContent, title, theme, isGrid };
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
                    accContent.style.display = accObj.isGrid ? "grid" : "flex";
                    accHeader.classList.add("open");
                    const openIcon = `<span class="icon" style="display:inline-block; width:15px; color: ${iconColor};">&#9660;</span>`;
                    accHeader.innerHTML = `${openIcon}<span>${accObj.title}</span>`;
                    
                    scrollToTarget(accHeader);
                }
            };
            
            accContainer.appendChild(accHeader);
            accContainer.appendChild(accContent);
            return accContainer;
        };

        const secEdit = buildAccordion("Data Entry & Views", [
            gridEditBtn,
            editBtn,
            insertTableBtn,
            insertDropdownBtn
        ], 'fast', true, true);

        const schemaElements = [
            editColumnsBtn,
            duplicateRecordBtn,
            appendBtn,
            moveWorkspaceBtn,
            exportCSVBtn
        ];
        if (dataTable.relations && dataTable.relations.length > 0) {
            schemaElements.splice(2, 0, cloneSubBtn);
        }

        const secSchema = buildAccordion("Schema & Operations", schemaElements, 'default', false, true);

        const secVersion = buildAccordion("Versioning & Danger Zone", [
            captureRevBtn,
            replaceBtn,
            snapshotBtn,
            deleteVersionBtn,
            deleteBtn
        ], 'danger', false, true);

        details.appendChild(revSelectContainer);
        details.appendChild(secEdit);
        details.appendChild(secSchema);
        details.appendChild(secVersion);

    tableCard.appendChild(header);
    tableCard.appendChild(details);
    famContent.appendChild(tableCard);
    });
    
    if (tablesInFam.length === 0) {
        const emptyMsg = document.createElement("div");
        emptyMsg.style.fontSize = "12px";
        emptyMsg.style.color = "#666";
        emptyMsg.style.paddingTop = "8px";
        emptyMsg.innerText = "No tables in this workspace.";
        famContent.appendChild(emptyMsg);
    }
  });
  
  renderFormulaBuilder();
  await renderVariables();
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

    // Remove references in other tables
    for (const key of Object.keys(store)) {
        if (store[key].relations) {
            store[key].relations = store[key].relations.filter((r: any) => r.subTable !== dataTableName);
        }
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

    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
    const varNames = Object.keys(variables);
    const tablesWithFields: Record<string, string[]> = {};
    Object.keys(store).forEach(t => tablesWithFields[t] = store[t].fields || []);

    const idField = dataSet.idField || dataSet.fields[0];
    const promptResult = await customManageColumnsPrompt("Edit Columns", `Drag to reorder, add new columns, rename them, or attach Calculated Formulas.`, dataSet.fields, idField, dataSet.calculatedFields || {}, varNames, tablesWithFields);
    
    if (!promptResult) return; // User cancelled

    const columnChanges = promptResult.changes;
    const newFields = columnChanges.map(c => c.newName);
    
    // Check for duplicates
    const uniqueFields = new Set(newFields);
    if (uniqueFields.size !== newFields.length) {
        if (status) { status.innerText = "Column names must be unique."; status.style.color = "red"; }
        return;
    }

    let newCalcFields: Record<string, string> = {};
    columnChanges.forEach(c => {
        if (c.formula) newCalcFields[c.newName] = c.formula;
    });

    // If nothing changed
    const oldCalcsStr = JSON.stringify(dataSet.calculatedFields || {});
    const newCalcsStr = JSON.stringify(newCalcFields);
    if (dataSet.fields.join(",") === newFields.join(",") && columnChanges.every(c => c.oldName === c.newName) && columnChanges.length === dataSet.fields.length && oldCalcsStr === newCalcsStr) return;

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
    
    const hasRenamesOrAdds = columnChanges.some(c => c.oldName !== c.newName || c.oldName === "") || droppedFields.length > 0;
    if (hasRenamesOrAdds) {
        dataSet.records.forEach((r: any) => {
            columnChanges.forEach(c => {
                if (c.oldName === "") {
                    if (r[c.newName] === undefined) r[c.newName] = "";
                } else if (c.oldName !== c.newName) {
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

    dataSet.calculatedFields = newCalcFields;
    await applyCalculatedFields(dataSet.records, dataSet.fields, dataSet.calculatedFields, store, dataTableName);

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

     const formFields: FormField[] = targetDataSet.fields.map((field: string) => {
        const isCalc = dataSet.calculatedFields && dataSet.calculatedFields[field];
        return {
            id: field,
            label: field + (isCalc ? ' (Calculated)' : ''),
            type: 'text',
            value: record[field] !== undefined ? String(record[field]) : "",
            disabled: field === idField || !!isCalc
        };
    });

    const editResult = await customFormPrompt(`Editing Record: ${id}`, "Update the values below:", formFields);
    if (!editResult) return;

    const recordIndex = targetDataSet.records.findIndex((r: any) => String(r.__DC_ID__) === String(id));
    if (recordIndex === -1) return;

    targetDataSet.fields.forEach((field: string) => {
        if (field !== idField && editResult[field] !== undefined) {
            const isCalc = dataSet.calculatedFields && dataSet.calculatedFields[field];
            if (!isCalc) {
                targetDataSet.records[recordIndex][field] = editResult[field];
            }
        }
    });

    await applyCalculatedFields(targetDataSet.records, targetDataSet.fields, dataSet.calculatedFields, store, dataTableName);
    for (const tName of Object.keys(store)) {
        if (tName !== dataTableName) await applyCalculatedFields(store[tName].records, store[tName].fields, store[tName].calculatedFields, store, tName);
    }
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

    const res = await customFormPrompt("Insert Table", "Select columns to insert:", [
        { id: "cols", label: "Columns", type: "checkboxes", options: targetDataSet.fields }
    ], "Insert");
    if (!res || !res.cols) return;
    const selectedCols = res.cols.split(",");
    const filteredRecords = targetDataSet.records.map((r: any) => {
       const rec: any = {};
       selectedCols.forEach(c => {
           if (dataSet.calculatedFields && dataSet.calculatedFields[c]) {
               let formula = dataSet.calculatedFields[c];
               let excelFormula = "=" + formula.replace(/\[([^\]]+)\]/g, '[@[$1]]').replace(/'/g, '"');
               rec[c] = excelFormula;
           } else {
               rec[c] = r[c];
           }
       });
       return rec;
    });

    await Excel.run(async (context) => {
        const cell = context.workbook.getActiveCell();
        const headers = selectedCols;
        const rows = filteredRecords.map((r: any) => headers.map(h => r[h] !== undefined ? r[h] : ""));
        
        const fullRange = cell.getResizedRange(rows.length, headers.length - 1);
        fullRange.values = [headers, ...rows.map(() => headers.map(() => ""))];
        
        const table = context.workbook.tables.add(fullRange, true);
        table.name = `DCTable_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
        table.style = "TableStyleLight1";
        await context.sync();
        
        if (rows.length > 0) {
            table.getDataBodyRange().formulas = rows;
            try {
                await context.sync();
            } catch (e: any) {
                table.getDataBodyRange().values = rows.map(r => r.map((c: any) => typeof c === 'string' && c.startsWith('=') ? `'${c}` : c));
                await context.sync();
                throw new Error("Some formulas contained Javascript syntax not supported by Excel and were inserted as text.");
            }
        }
    });

    const status = document.getElementById("status-text");
    if (status) {
       status.innerText = `Inserted table for ${dataTableName}`;
       status.style.color = "green";
    }
  } catch (error: any) {
    showStatus("Error inserting table: " + error.message, "error");
  }
}

export async function handleSheetActivation(event: Excel.WorksheetActivatedEventArgs) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(event.worksheetId);
            const purposeProp = sheet.customProperties.getItemOrNullObject("SheetPurpose");
            const tableProp = sheet.customProperties.getItemOrNullObject("EditingTable");
            const mainTableProp = sheet.customProperties.getItemOrNullObject("MainTable");
            purposeProp.load("value");
            tableProp.load("value");
            mainTableProp.load("value");
            await context.sync();

            if (!purposeProp.isNullObject && (purposeProp.value === "SubDataEditor" || purposeProp.value === "MainDataEditor")) {
                const isLive = !mainTableProp.isNullObject && mainTableProp.value !== "";
                showEditorView(!tableProp.isNullObject ? tableProp.value : "Unknown", isLive || purposeProp.value === "MainDataEditor");
            } else {
                hideEditorView();
            }
        });
    } catch (error) {
        console.error("Sheet activation error:", error);
    }
}

export function showEditorView(tableName: string, isLiveSync: boolean = false) {
    document.getElementById("main-view")!.style.display = "none";
    document.getElementById("settings-view")!.style.display = "none";
    const editorView = document.getElementById("editor-view");
    if (editorView) editorView.style.display = "block";
    const nameEl = document.getElementById("editor-table-name");
    if (nameEl) nameEl.innerText = tableName;
    const indicator = document.getElementById("live-sync-indicator");
    if (indicator) indicator.style.display = isLiveSync ? "block" : "none";
}

export function hideEditorView() {
    const editorView = document.getElementById("editor-view");
    if (editorView) editorView.style.display = "none";
    const settingsBtn = document.getElementById('settings-button');
    if (!settingsBtn?.classList.contains('active')) {
        document.getElementById("main-view")!.style.display = "block";
    }
}

export async function openRelationEditor(mainTableName: string, subTableName: string, foreignKey: string) {
    const status = document.getElementById("status-text");
    try {
        if (status) { status.innerText = `Loading Split Editor...`; status.style.color = "blue"; }
        const saved = await saveGridEditor(false); // ensure any open editor saves
        if (!saved) {
             showStatus("Please fix the errors in your current editor before opening a new one.", "error");
             return;
        }
        
        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        const mainDataSet = store[mainTableName];
        const subDataSet = store[subTableName];
        if (!mainDataSet || !subDataSet) return;

        const mainRevSelect = document.getElementById(`fb-rev-${mainTableName}`) as HTMLSelectElement;
        const mainSelectedRev = mainRevSelect ? parseInt(mainRevSelect.value) : mainDataSet.revision || 1;
        let targetMainData = mainDataSet;
        if (mainSelectedRev !== (mainDataSet.revision || 1) && mainDataSet.history && mainDataSet.history[mainSelectedRev]) {
            targetMainData = mainDataSet.history[mainSelectedRev];
        }

        const subRevSelect = document.getElementById(`fb-rev-${subTableName}`) as HTMLSelectElement;
        const subSelectedRev = subRevSelect ? parseInt(subRevSelect.value) : subDataSet.revision || 1;
        let targetSubData = subDataSet;
        if (subSelectedRev !== (subDataSet.revision || 1) && subDataSet.history && subDataSet.history[subSelectedRev]) {
            targetSubData = subDataSet.history[subSelectedRev];
        }

        const idField = targetMainData.idField || targetMainData.fields[0];
        const firstId = targetMainData.records && targetMainData.records.length > 0 ? targetMainData.records[0][idField] : "NEW";

        await Excel.run(async (context) => {
            context.application.suspendApiCalculationUntilNextSync();
            const sheets = context.workbook.worksheets;
            const s1 = sheets.getItemOrNullObject("DC_Grid_Editor");
            const s2 = sheets.getItemOrNullObject("DC_Main_Editor");
            const s3 = sheets.getItemOrNullObject("DC_Sub_Editor");
            await context.sync();
            if (!s1.isNullObject) s1.delete();

            // Build Main Sheet
            let newMain: Excel.Worksheet;
            if (!s2.isNullObject) {
                newMain = s2;
                newMain.tables.load("items");
            } else {
                newMain = sheets.add("DC_Main_Editor");
                newMain.tabColor = "#d13438";
            }

            // Build Sub Sheet
            let newSub: Excel.Worksheet;
            if (!s3.isNullObject) {
                newSub = s3;
                newSub.tables.load("items");
            } else {
                newSub = sheets.add("DC_Sub_Editor");
                newSub.tabColor = "#d13438";
            }

            if (!s2.isNullObject || !s3.isNullObject) {
                await context.sync(); // Load all items at once
            }
            if (!s2.isNullObject) {
                newMain.tables.items.forEach(t => t.convertToRange());
                newMain.getRange().clear();
            }
            if (!s3.isNullObject) {
                newSub.tables.items.forEach(t => t.convertToRange());
                newSub.getRange().clear();
            }

            newMain.customProperties.add("SheetPurpose", "MainDataEditor");
            newMain.customProperties.add("EditingTable", mainTableName);
            newMain.customProperties.add("EditingRev", String(mainSelectedRev));
            newMain.customProperties.add("FilterField", "");
            newMain.customProperties.add("FilterValue", "");
            newMain.customProperties.add("MainTable", "");
            newMain.customProperties.add("EditorColumns", "");

            let mHeaders = targetMainData.fields;
            let mRows = targetMainData.records.map((r: any) => mHeaders.map((h: string) => r[h] !== undefined ? r[h] : ""));
            if (mRows.length === 0) mRows.push(mHeaders.map(() => ""));
            const mMatrix = [mHeaders, ...mRows];
            const mRange = newMain.getRange("A1").getResizedRange(mMatrix.length - 1, mHeaders.length - 1);
            mRange.values = mMatrix;
            const mTable = newMain.tables.add(mRange, true);
            mTable.name = `DCEditor_${Date.now()}_Main`;
            mTable.style = "TableStyleLight1";
            mTable.showBandedRows = false;
            newSub.customProperties.add("SheetPurpose", "SubDataEditor");
            newSub.customProperties.add("EditingTable", subTableName);
            newSub.customProperties.add("EditingRev", String(subSelectedRev));
            newSub.customProperties.add("FilterField", foreignKey);
            newSub.customProperties.add("FilterValue", String(firstId));
            newSub.customProperties.add("MainTable", mainTableName);
            newSub.customProperties.add("EditorColumns", "");

            let sHeaders = targetSubData.fields;
            if (foreignKey && !sHeaders.includes(foreignKey)) sHeaders.push(foreignKey);
            let sRecords = targetSubData.records.filter((r: any) => String(r[foreignKey]) === String(firstId));
            let sRows = sRecords.map((r: any) => sHeaders.map((h: string) => r[h] !== undefined ? r[h] : ""));
            if (sRows.length === 0) {
                const emptyRow = sHeaders.map(() => "");
                const fIndex = sHeaders.indexOf(foreignKey);
                if (fIndex > -1) emptyRow[fIndex] = String(firstId);
                sRows.push(emptyRow);
            }
            const sMatrix = [sHeaders, ...sRows];
            const sRange = newSub.getRange("A1").getResizedRange(sMatrix.length - 1, sHeaders.length - 1);
            sRange.values = sMatrix;
            const sTable = newSub.tables.add(sRange, true);
            sTable.name = `DCEditor_${Date.now()}_Sub`;
            sTable.style = "TableStyleLight1";
            sTable.showBandedRows = false;

            const fIndex = sHeaders.indexOf(foreignKey);
            if (fIndex > -1) {
                const colRange = sTable.columns.getItemAt(fIndex).getDataBodyRange();
                colRange.format.fill.color = "#f3f2f1";
                colRange.format.font.color = "#a6a6a6";
                colRange.dataValidation.rule = {
                    list: { inCellDropDown: true, source: String(firstId).includes(",") ? `"${firstId}"` : String(firstId) }
                };
            }

            newMain.activate();
            await context.sync();
        });

        showEditorView(mainTableName, true);
        
        showStatus(`Opened Split Editor for ${mainTableName} and ${subTableName}. Arrange windows side-by-side.`, "success");
    } catch (e: any) {
        showStatus("Error opening split editor: " + e.message, "error");
    }
}

export async function openGridEditor(dataTableName: string, filterField?: string, filterValue?: string, mainTableName?: string, selectedColumns?: string[], targetSheetName: string = "DC_Grid_Editor", targetPurpose: string = "SubDataEditor", activateSheet: boolean = true) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
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

    await Excel.run(async (context) => {
        context.application.suspendApiCalculationUntilNextSync();
        const sheets = context.workbook.worksheets;
        
        let oldSheet: Excel.Worksheet;
        // Clean up Split Editor sheets if standard Grid Editor is launched
        if (targetSheetName === "DC_Grid_Editor") {
            const s2 = sheets.getItemOrNullObject("DC_Main_Editor");
            const s3 = sheets.getItemOrNullObject("DC_Sub_Editor");
            oldSheet = sheets.getItemOrNullObject(targetSheetName);
            await context.sync();
            if (!s2.isNullObject) s2.delete();
            if (!s3.isNullObject) s3.delete();
        } else {
            oldSheet = sheets.getItemOrNullObject(targetSheetName);
            await context.sync();
        }

        let editorSheet: Excel.Worksheet;
        if (!oldSheet.isNullObject) {
            editorSheet = oldSheet;
            editorSheet.tables.load("items");
            await context.sync();
            editorSheet.tables.items.forEach(t => t.convertToRange());
            editorSheet.getRange().clear();
        } else {
            editorSheet = sheets.add(targetSheetName);
            editorSheet.tabColor = "#d13438"; 
        }
        
        editorSheet.customProperties.add("SheetPurpose", targetPurpose);
        editorSheet.customProperties.add("EditingTable", dataTableName);
        editorSheet.customProperties.add("EditingRev", String(selectedRev));
        editorSheet.customProperties.add("FilterField", filterField || "");
        editorSheet.customProperties.add("FilterValue", filterValue || "");
        editorSheet.customProperties.add("MainTable", mainTableName || "");
        editorSheet.customProperties.add("EditorColumns", selectedColumns ? selectedColumns.join(",") : "");

        let headers = targetDataSet.fields;
        if (selectedColumns && selectedColumns.length > 0) {
            headers = headers.filter((h: string) => selectedColumns.includes(h));
            if (filterField && !headers.includes(filterField)) headers.push(filterField);
            const idField = targetDataSet.idField || targetDataSet.fields[0];
            if (!headers.includes(idField)) {
                headers.unshift(idField);
            }
        }
        let recordsToLoad = targetDataSet.records;
        if (filterField && filterValue) {
            recordsToLoad = recordsToLoad.filter((r: any) => String(r[filterField]) === String(filterValue));
        }
        
        let rows = recordsToLoad.map((r: any) => headers.map((h: string) => r[h] !== undefined ? r[h] : ""));
        
        if (rows.length === 0) {
            const emptyRow = headers.map(() => "");
            if (filterField && filterValue) {
                const fIndex = headers.indexOf(filterField);
                if (fIndex > -1) emptyRow[fIndex] = filterValue;
            }
            rows.push(emptyRow);
        }

        const dataMatrix = [headers, ...rows];
        const startCell = editorSheet.getRange("A1");
        const range = startCell.getResizedRange(dataMatrix.length - 1, headers.length - 1);
        range.values = dataMatrix;

        const table = editorSheet.tables.add(range, true);
        table.name = `DCEditor_${Date.now()}`;
        table.style = "TableStyleLight1";
        table.showBandedRows = false;
        
        // Highlight Foreign Key column in grey for visual protection
        if (filterField && filterValue) {
            const fIndex = headers.indexOf(filterField);
            if (fIndex > -1) {
                const colRange = table.columns.getItemAt(fIndex).getDataBodyRange();
                colRange.format.fill.color = "#f3f2f1";
                colRange.format.font.color = "#a6a6a6";
                colRange.dataValidation.rule = {
                    list: {
                        inCellDropDown: true,
                        source: filterValue.includes(",") ? `"${filterValue}"` : filterValue
                    }
                };
            }
        }

        if (activateSheet) {
            editorSheet.activate();
        }
        await context.sync();
        
        const msg = filterField ? `Opened Sub-table (${filterField}: ${filterValue})` : `Opened '${dataTableName}'`;
        if (status) { status.innerText = `${msg} in Grid Editor.`; status.style.color = "blue"; }
    });
  } catch (error: any) {
      showStatus("Error opening Grid Editor: " + error.message, "error");
  }
}

export async function saveGridEditor(closeEditor: boolean | Event = true, specificSheetName?: string): Promise<boolean> {
  const status = document.getElementById("status-text");
  const shouldClose = typeof closeEditor === "boolean" ? closeEditor : true;
  try {
    await Excel.run(async (context) => {
        const editorSheetNames = (typeof specificSheetName === "string" && specificSheetName.trim() !== "") ? [specificSheetName] : ["DC_Grid_Editor", "DC_Main_Editor", "DC_Sub_Editor"];
        let savedCount = 0;
        const sheetsToDelete: Excel.Worksheet[] = [];
        
        const loadedSheetsInfo = editorSheetNames.map(name => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            const purposeProp = sheet.customProperties.getItemOrNullObject("SheetPurpose");
            const tableProp = sheet.customProperties.getItemOrNullObject("EditingTable");
            const revProp = sheet.customProperties.getItemOrNullObject("EditingRev");
            const filterFieldProp = sheet.customProperties.getItemOrNullObject("FilterField");
            const filterValueProp = sheet.customProperties.getItemOrNullObject("FilterValue");
            const editorColsProp = sheet.customProperties.getItemOrNullObject("EditorColumns");
            
            purposeProp.load("value");
            tableProp.load("value");
            revProp.load("value");
            filterFieldProp.load("value");
            filterValueProp.load("value");
            editorColsProp.load("value");
            sheet.tables.load("count");
            
            return { sheet, purposeProp, tableProp, revProp, filterFieldProp, filterValueProp, editorColsProp };
        });
        await context.sync(); // One massive sync for all metadata!

        const sheetsWithData: any[] = [];
        for (const item of loadedSheetsInfo) {
            if (item.sheet.isNullObject) continue;
            if (!item.purposeProp.isNullObject && (item.purposeProp.value === "SubDataEditor" || item.purposeProp.value === "MainDataEditor")) {
                if (item.sheet.tables.count === 0) continue;
                const table = item.sheet.tables.getItemAt(0);
                const range = table.getRange();
                range.load("values");
                sheetsWithData.push({ ...item, range });
            }
        }
        if (sheetsWithData.length > 0) await context.sync();

        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        let storeHasChanges = false;

        for (const item of sheetsWithData) {
            const dataTableName = item.tableProp.value;
            const editingRev = parseInt(item.revProp.value);
            const filterField = (!item.filterFieldProp.isNullObject && item.filterFieldProp.value) ? String(item.filterFieldProp.value).trim() : "";
            const filterValue = (!item.filterValueProp.isNullObject && item.filterValueProp.value) ? String(item.filterValueProp.value).trim() : "";
            const editorCols = (!item.editorColsProp.isNullObject && item.editorColsProp.value) ? String(item.editorColsProp.value).split(",").filter(c => c.trim() !== "") : [];
            const isSubTableMode = filterField !== "" && filterValue !== "";
            
            const values = item.range.values;
            if (values.length < 1) continue;
            const headers = values[0];
            const dataRows = values.slice(1);

            const dataSet = store[dataTableName];
            if (!dataSet) continue;

            let targetDataSet = dataSet;
            if (editingRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[editingRev]) {
                targetDataSet = dataSet.history[editingRev];
            }
            
            const idField = targetDataSet.idField || headers[0];
            const idIndex = headers.indexOf(idField);
            if (idIndex === -1) continue;

            const idSet = new Set();
            let finalRecords: any[] = [];
            for (let i = 0; i < dataRows.length; i++) {
                const row = dataRows[i];
                
                let isBlank = true;
                for (let c = 0; c < row.length; c++) {
                    if (isSubTableMode && headers[c] === filterField) continue; // Ignore pre-filled foreign keys
                    if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
                        isBlank = false;
                        break;
                    }
                }
                if (isBlank) continue; // skip blank rows cleanly
                
                if (isSubTableMode) {
                    const fIndex = headers.indexOf(filterField);
                    if (fIndex > -1) row[fIndex] = filterValue; // Enforce foreign key protection
                }
                
                const val = String(row[idIndex]).trim();
                if (!val || val.trim() === "") throw new Error(`Row ${i + 1} has an empty ID in ${dataTableName}.`);
                if (idSet.has(val)) throw new Error(`Duplicate ID found: '${val}' in ${dataTableName}.`);
                idSet.add(val);
                
                const existingRecord = targetDataSet.records.find((r: any) => String(r[idField]).trim() === val);
                 const rec: any = (editorCols.length > 0 && existingRecord) ? { ...existingRecord } : {};
                headers.forEach((h: string, j: number) => { rec[h] = row[j]; });
                rec.__DC_ID__ = val;
                finalRecords.push(rec);
            }

            let isModified = false;
            let oldRecordsToCompare = targetDataSet.records;
            let normFilter = "";
            
            if (isSubTableMode) {
                normFilter = String(filterValue).toLowerCase();
                oldRecordsToCompare = targetDataSet.records.filter((r: any) => String(r[filterField]).trim().toLowerCase() === normFilter);
            }

            if (oldRecordsToCompare.length !== finalRecords.length) {
                isModified = true;
            } else {
                for (let i = 0; i < oldRecordsToCompare.length; i++) {
                    const oRec = oldRecordsToCompare[i];
                    const nRec = finalRecords[i];
                    const allKeys = Array.from(new Set([...Object.keys(oRec), ...Object.keys(nRec)]));
                    for (const k of allKeys) {
                        let v1 = oRec[k];
                        let v2 = nRec[k];
                        if (v1 === undefined || v1 === null) v1 = "";
                        if (v2 === undefined || v2 === null) v2 = "";
                        if (String(v1).trim() !== String(v2).trim()) {
                            isModified = true;
                            break;
                        }
                    }
                    if (isModified) break;
                }
            }
            
            if (isModified) {
                if (editorCols.length === 0) targetDataSet.fields = headers;
                
                if (isSubTableMode) {
                    const otherRecords = targetDataSet.records.filter((r: any) => String(r[filterField]).trim().toLowerCase() !== normFilter);
                    targetDataSet.records = [...otherRecords, ...finalRecords];
                } else {
                    targetDataSet.records = [...finalRecords];
                }
                
                // Update store reference first so formulas evaluate against the new data
                if (editingRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[editingRev]) store[dataTableName].history[editingRev].records = targetDataSet.records;
                else store[dataTableName].records = targetDataSet.records;
                
                await applyCalculatedFields(targetDataSet.records, targetDataSet.fields, dataSet.calculatedFields, store, dataTableName);
                
                storeHasChanges = true;
                savedCount += finalRecords.length;
            }

            if (shouldClose) {
                sheetsToDelete.push(item.sheet);
            }
        }
        
        if (storeHasChanges) {
            // Recalculate dependencies in all other tables
            for (const tName of Object.keys(store)) {
                if (!sheetsWithData.find(item => item.tableProp.value === tName)) await applyCalculatedFields(store[tName].records, store[tName].fields, store[tName].calculatedFields, store, tName);
            }
            await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
        }

        if (shouldClose) {
            sheetsToDelete.forEach(s => s.delete());
            await context.sync();
            hideEditorView();
            await renderDashboard();
        }
        
        if (status) { 
            if (storeHasChanges) {
                status.innerText = `Saved ${savedCount} records.`; 
                status.style.color = "green"; 
            } else {
                if (!shouldClose) status.innerText = `No changes detected.`; 
            }
        }
        
        await refreshFormulas(true);
    });
    return true;
  } catch (error: any) { 
      showStatus("Error saving Grid Editor: " + error.message, "error"); 
      return false;
  }
}

export async function switchGridEditorRecord(newId: string, foreignKey: string, targetTable: string, selectedColumns: string[], subSheetName: string = "DC_Grid_Editor", mainTableName: string = "") {
    const status = document.getElementById("status-text");
    try {
        if (status) { status.innerText = `Auto-syncing... Switching to Record ${newId}`; status.style.color = "blue"; }
        const saved = await saveGridEditor(false, subSheetName); // Save ONLY the sub-sheet without closing
        if (!saved) {
            return;
        }

        // Recreate the sheet cleanly in the background without stealing focus
        await openGridEditor(targetTable, foreignKey, newId, mainTableName, selectedColumns, subSheetName, "SubDataEditor", false);

        if (status) { status.innerText = `Switched Sub-table to Record ${newId}`; status.style.color = "green"; }
    } catch (error: any) { 
        console.error("[Live Sync] Error switching record:", error);
        showStatus("Error switching record: " + error.message, "error"); 
    }
}

export async function cancelGridEditor() {
  try {
    await Excel.run(async (context) => {
        const editorSheetNames = ["DC_Grid_Editor", "DC_Main_Editor", "DC_Sub_Editor"];
        const loadedSheets = editorSheetNames.map(name => context.workbook.worksheets.getItemOrNullObject(name));
        await context.sync();
        
        for (const sheet of loadedSheets) {
            if (!sheet.isNullObject) {
                sheet.delete();
            }
        }
        await context.sync();
        hideEditorView();
    });
  } catch (error: any) { showStatus("Error canceling Grid Editor: " + error.message, "error"); }
}

export async function manageRelations(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    const existingRelations = dataSet.relations || [];
    
    const allTables = Object.keys(store).filter(t => t !== dataTableName);
    const relationStrings = existingRelations.map((r: any) => `${r.subTable} (Link: ${r.foreignKey})`);

    const tablesWithFields: Record<string, string[]> = {};
    allTables.forEach(t => tablesWithFields[t] = store[t].fields || []);

    const fields: FormField[] = [];
    if (relationStrings.length > 0) {
        fields.push({ id: "keepRelations", label: "Existing Relations (Uncheck to remove)", type: "checkboxes", options: relationStrings });
    }
    
    fields.push({ id: "newSubTable", label: "Add New: Target Table", type: "select", options: ["-- None --", ...allTables] });
    fields.push({ id: "newForeignKey", label: "Add New: Link Column", type: "select", options: ["-- None --"], dependsOn: "newSubTable", optionsMap: tablesWithFields });

    const res = await customFormPrompt("Manage Relations", `Manage relations for '${dataTableName}':`, fields);
    
    if (!res) return;

    let finalRelations: any[] = [];
    if (relationStrings.length > 0 && res.keepRelations !== undefined) {
        const kept = res.keepRelations.split(",");
        finalRelations = existingRelations.filter((r: any) => kept.includes(`${r.subTable} (Link: ${r.foreignKey})`));
    }

    if (res.newSubTable && res.newSubTable !== "-- None --" && res.newForeignKey && res.newForeignKey.trim() !== "" && res.newForeignKey !== "-- None --") {
        if (!finalRelations.find(r => r.subTable === res.newSubTable)) {
            finalRelations.push({ subTable: res.newSubTable, foreignKey: res.newForeignKey.trim() });
        }
    }

    dataSet.relations = finalRelations;
    
    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    renderDashboard();
    showStatus(`Relations updated for '${dataTableName}'`, "success");
    if (res.newSubTable && res.newSubTable !== "-- None --" && res.newForeignKey && res.newForeignKey.trim() !== "" && res.newForeignKey !== "-- None --") {
            const addRollup = await customConfirm("Add Rollup", `Would you like to add a Calculated Field in '${dataTableName}' to summarize data from '${res.newSubTable}'?`, "Yes");
            if (addRollup) {
                const subFields = store[res.newSubTable]?.fields || [];
                const rollupRes = await customFormPrompt("Rollup Field", "Define the rollup:", [
                    { id: "fieldName", label: "New Field Name (e.g. Total Cost)", type: "text" },
                    { id: "type", label: "Aggregation Type", type: "select", options: ["SUM", "COUNT"] },
                    { id: "col", label: "Column to Aggregate (for SUM)", type: "select", options: ["-- None --", ...subFields] }
                ]);
                if (rollupRes && rollupRes.fieldName) {
                    const idField = dataSet.idField || dataSet.fields[0];
                    let formula = "";
                    if (rollupRes.type === "SUM" && rollupRes.col !== "-- None --") {
                        formula = `DC.SUM('${res.newSubTable}', '${rollupRes.col}', '${res.newForeignKey}', [${idField}])`;
                    } else if (rollupRes.type === "COUNT") {
                        formula = `DC.COUNT('${res.newSubTable}', '${res.newForeignKey}', [${idField}])`;
                    }
                    
                    if (formula) {
                        dataSet.calculatedFields = dataSet.calculatedFields || {};
                        dataSet.calculatedFields[rollupRes.fieldName] = formula;
                        await applyCalculatedFields(dataSet.records, dataSet.fields, dataSet.calculatedFields, store, dataTableName);
                        await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
                        renderDashboard();
                        showStatus(`Added rollup field '${rollupRes.fieldName}'`, "success");
                    }
                }
            }
        }

  } catch (error: any) { showStatus("Error managing relations: " + error.message, "error"); }
}


export async function appendTableData(dataTableName: string) {
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

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
      if (existingStore) store = JSON.parse(existingStore);

      const dataSet = store[dataTableName];
      if (!dataSet) throw new Error(`Data table '${dataTableName}' not found.`);

      const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
      const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
      if (selectedRev !== (dataSet.revision || 1)) {
          throw new Error("Cannot append data to a historical revision.");
      }

      const summary = await customDataSummaryPrompt(
          "Append Data",
          `Review data to append to '${dataTableName}' (${values.length - 1} rows detected).`,
          values[0],
          values.slice(1),
          dataSet.idField || dataSet.fields[0]
      );
      if (!summary) return;

      const existingIds = new Set(dataSet.records.map((r: any) => String(r.__DC_ID__)));
      const newRecords: any[] = [];

      for (const rec of summary.records) {
          const id = String(rec.__DC_ID__);
          if (existingIds.has(id)) {
              throw new Error(`Duplicate ID found: '${id}' already exists in '${dataTableName}'.`);
          }
          const finalRec: any = { __DC_ID__: id };
          dataSet.fields.forEach(f => {
              finalRec[f] = rec[f] !== undefined ? rec[f] : "";
          });
          newRecords.push(finalRec);
      }

      dataSet.history = dataSet.history || {};
      dataSet.history[dataSet.revision] = {
        idField: dataSet.idField,
        fields: [...dataSet.fields],
        records: JSON.parse(JSON.stringify(dataSet.records))
      };
      
      await applyCalculatedFields(newRecords, dataSet.fields, dataSet.calculatedFields, store, dataTableName);
      dataSet.records = [...dataSet.records, ...newRecords];
      dataSet.revision += 1;

      await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
      
      if (status) {
        status.innerText = `Appended ${newRecords.length} records. Current is Rev ${dataSet.revision}.`;
        status.style.color = "green";
      }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus("Error appending data: " + error.message, "error");
  }
}

export async function cloneSubRecordsPrompt(mainTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const mainDataSet = store[mainTableName];
    if (!mainDataSet || !mainDataSet.relations || mainDataSet.relations.length === 0) {
        showStatus("No relations defined for this table.", "error");
        return;
    }

    const allMainIds = mainDataSet.records.map((r: any) => String(r.__DC_ID__));
    const relationsOptions = mainDataSet.relations.map((r: any) => `${r.subTable} (FK: ${r.foreignKey})`);

    const res1 = await customFormPrompt("Clone Sub-records", `Select the target sub-table to clone:`, [
        { id: "relation", label: "Sub-table Relation", type: "select", options: relationsOptions }
    ]);
    if (!res1 || !res1.relation) return;

    const selectedRel = mainDataSet.relations.find((r: any) => res1.relation.includes(r.subTable));
    if (!selectedRel) return;

    const res2 = await customFormPrompt(`Clone from '${selectedRel.subTable}'`, `Select the Source ID and Target ID(s):`, [
        { id: "sourceId", label: "Source Record ID (Copy FROM)", type: "autocomplete", options: allMainIds },
        { id: "targetIds", label: "Target Record IDs (Copy TO)", type: "checkboxes", options: allMainIds }
    ]);
    if (!res2 || !res2.sourceId || !res2.targetIds) return;

    const targetIds = res2.targetIds.split(",");
    if (targetIds.length === 0 || targetIds.includes("")) return;

    const subDataSet = store[selectedRel.subTable];
    if (!subDataSet) return;

    const sourceRecords = subDataSet.records.filter((r: any) => String(r[selectedRel.foreignKey]) === String(res2.sourceId));
    if (sourceRecords.length === 0) {
        showStatus(`No sub-records found for Source ID '${res2.sourceId}'.`, "error");
        return;
    }

    const newRecords: any[] = [];
    targetIds.forEach(tId => {
        if (tId === res2.sourceId) return; // skip self
        sourceRecords.forEach((sr: any) => {
            const cloned = { ...sr };
            cloned.__DC_ID__ = `CLONED_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
            cloned[selectedRel.foreignKey] = tId; 
            const idF = subDataSet.idField || subDataSet.fields[0];
            if (idF) cloned[idF] = cloned.__DC_ID__;
            newRecords.push(cloned);
        });
    });

    if (newRecords.length === 0) return;

    subDataSet.history = subDataSet.history || {};
    subDataSet.history[subDataSet.revision] = {
        idField: subDataSet.idField,
        fields: [...subDataSet.fields],
        records: JSON.parse(JSON.stringify(subDataSet.records))
    };
    subDataSet.revision += 1;
    await applyCalculatedFields(newRecords, subDataSet.fields, subDataSet.calculatedFields, store, selectedRel.subTable);
    subDataSet.records = [...subDataSet.records, ...newRecords];

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    await renderDashboard();
    await refreshFormulas(true);

    showStatus(`Successfully cloned ${sourceRecords.length} sub-records to ${targetIds.length} target(s).`, "success");
  } catch (error: any) {
    showStatus("Error cloning sub-records: " + error.message, "error");
  }
}

export async function moveTableWorkspace(tableName: string) {
    try {
        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        const dataSet = store[tableName];
        if (!dataSet) return;

        const familiesSet = new Set<string>();
        Object.keys(store).forEach(k => familiesSet.add(store[k].family || 'Public'));
        const familiesList = Array.from(familiesSet);

        const res = await customFormPrompt("Move Workspace", `Select or type a new workspace for '${tableName}':`, [
            { id: "newWorkspace", label: "Workspace", type: "autocomplete", options: familiesList, value: dataSet.family || 'Public' }
        ]);

        if (!res || !res.newWorkspace || res.newWorkspace.trim() === "") return;
        const newWorkspace = res.newWorkspace.trim();
        if (newWorkspace === (dataSet.family || 'Public')) return; // No change

        dataSet.family = newWorkspace;
        await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
        renderDashboard();
        showStatus(`Moved '${tableName}' to '${newWorkspace}'.`, "success");

    } catch (error: any) {
        showStatus("Error moving table: " + error.message, "error");
    }
}

export async function manageVariable() {
    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
    const varNames = Object.keys(variables);
    
    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};
    const tablesWithFields: Record<string, string[]> = {};
    Object.keys(store).forEach(t => tablesWithFields[t] = store[t].fields || []);
    const allFields = Array.from(new Set(Object.values(store).flatMap((t: any) => t.fields || []))) as string[];
    
    const res = await customFormPrompt("Add Variable", "Define a global variable:", [
        { id: "vName", label: "Variable Name", type: "text" },
        { id: "vFormula", label: "Formula Definition", type: "formula", varsList: varNames, tablesWithFields: tablesWithFields }
    ]);

    if (res) {
        if (!res.vName || res.vName.trim() === "") {
            showStatus("Variable Name is required.", "error");
            return;
        }
        if (!res.vFormula || res.vFormula.trim() === "") {
            showStatus("Formula is required.", "error");
            return;
        }
        try {
            // Validation Layer: Test compiling and executing the formula
            const evaluateVar = (vName: string, visited: Set<string>): any => {
                if (visited.has(vName)) throw new Error("Loop Detected");
                visited.add(vName);
                if (vName === res.vName) throw new Error("Variable cannot reference itself");
                const vForm = variables[vName];
                if (!vForm) return 0;
                const vDC = {
                    SUM: (t: string, c: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + (Number(b[c])||0) : a, 0) || 0,
                    COUNT: (t: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + 1 : a, 0) || 0,
                    VAR: (v: string) => evaluateVar(v, new Set(visited))
                };
                const func = new Function('store', 'DC', `return ${vForm};`);
                return func(store, vDC);
            };
            const DC = {
                SUM: (t: string, c: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + (Number(b[c])||0) : a, 0) || 0,
                COUNT: (t: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + 1 : a, 0) || 0,
                VAR: (v: string) => evaluateVar(v, new Set([res.vName]))
            };
            const testFunc = new Function('store', 'DC', `return ${res.vFormula};`);
            testFunc(store, DC); // Run test execution

            variables[res.vName] = res.vFormula;
            await idbSet(IDB_KEYS.VARIABLES, JSON.stringify(variables));
            renderVariables();
            showStatus(`Variable '${res.vName}' saved.`, "success");
        } catch (error: any) {
            showStatus(`Invalid Formula for '${res.vName}': ${error.message}`, "error");
        }
    }
}

export async function addWorkspace() {
    try {
        const newName = await customPrompt("Add Workspace", "Enter a name for the new workspace:");
        if (!newName || newName.trim() === "") return;

        const storedData = await idbGet(IDB_KEYS.STORE);
        const store = storedData ? JSON.parse(storedData) : {};
        const allTableKeys = Object.keys(store);

        const res = await customFormPrompt(`Create First Table`, `Create the first table in the '${newName}' workspace.`, [
            { id: "tableName", label: "Data Table Name", type: "text" },
            { id: "parentTable", label: "Parent Table (Optional Link)", type: "select", options: ["-- None --", ...allTableKeys] }
        ]);
        if (res && res.tableName) {
            await loadRangeForCapture(res.tableName, newName, res.parentTable === "-- None --" ? "" : res.parentTable);
        }
    } catch (error: any) {
        showStatus("Error adding workspace: " + error.message, "error");
    }
}

export async function renderVariables() {
    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
    const list = document.getElementById("variables-list");
    if (!list) return;
    list.innerHTML = "";

    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};

    const evaluateVar = (vName: string, visited: Set<string>): any => {
        if (visited.has(vName)) throw new Error("Loop Detected");
        visited.add(vName);
        const vForm = variables[vName];
        if (!vForm) return 0;
        const DC = {
            SUM: (t: string, c: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + (Number(b[c])||0) : a, 0) || 0,
            COUNT: (t: string, fk?: string, fkv?: any) => store[t]?.records?.reduce((a:number, b:any) => (!fk || String(b[fk])===String(fkv)) ? a + 1 : a, 0) || 0,
            VAR: (v: string) => evaluateVar(v, new Set(visited))
        };
        const func = new Function('store', 'DC', `return ${vForm};`);
        return func(store, DC);
    };

    for (const [vName, vFormula] of Object.entries(variables)) {
        let result: any = "ERROR";
        try {
            result = evaluateVar(vName, new Set());
        } catch(e: any) { result = e.message === "Loop Detected" ? "LOOP ERROR" : "ERROR"; }

        const li = document.createElement("li");
        li.style.listStyle = "none";
        li.style.marginBottom = "8px";
        li.style.display = "flex";
        li.style.justifyContent = "space-between";
        li.style.alignItems = "center";
        li.style.background = "var(--border-color)";
        li.style.padding = "8px";
        li.style.borderRadius = "4px";

        li.innerHTML = `
            <div style="flex:1; min-width: 0; margin-right: 8px;">
                <div style="font-weight:bold; font-size:13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${vName}">${vName}</div>
                <div style="font-size:11px; color:var(--text-color); opacity:0.8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title='${vFormula.replace(/'/g, "&#39;")}'>${vFormula}</div>
            </div>
            <div style="color:green; font-weight:bold; margin-right: 8px; font-size:14px; max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0;" title="${result}">${result}</div>
        `;

        const actionDiv = document.createElement("div");
        actionDiv.style.display = "flex";
        actionDiv.style.gap = "4px";
        actionDiv.style.flexShrink = "0";

        const insertBtn = document.createElement("button");
        insertBtn.className = "icon-btn";
        insertBtn.style.color = "#0078d4";
        insertBtn.innerHTML = `<i class="ms-Icon ms-Icon--Insert"></i>`;
        insertBtn.title = "Insert to Sheet";
        insertBtn.onclick = async () => {
            const formulaStr = `=DC.VAR("${vName}")`;
            try {
                await Excel.run(async (context) => {
                    const cell = context.workbook.getActiveCell();
                    cell.formulas = [[formulaStr]];
                    await context.sync();
                });
                showStatus(`Inserted variable '${vName}' to sheet.`, "success");
            } catch (error: any) {
                showStatus("Error inserting variable: " + error.message, "error");
            }
        };

        const delBtn = document.createElement("button");
        delBtn.className = "icon-btn";
        delBtn.style.color = "#d13438";
        delBtn.innerHTML = `<i class="ms-Icon ms-Icon--Delete"></i>`;
        delBtn.title = "Delete Variable";
        delBtn.onclick = async () => {
            delete variables[vName];
            await idbSet(IDB_KEYS.VARIABLES, JSON.stringify(variables));
            renderVariables();
        };
        
        actionDiv.appendChild(insertBtn);
        actionDiv.appendChild(delBtn);
        li.appendChild(actionDiv);
        list.appendChild(li);
    }
}

export async function duplicateRecordPrompt(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const mainDataSet = store[dataTableName];
    if (!mainDataSet) return;

    const allMainIds = mainDataSet.records.map((r: any) => String(r.__DC_ID__));

    const res = await customFormPrompt("Duplicate Record", `Select the record to duplicate and provide a new ID:`, [
        { id: "sourceId", label: "Source Record ID", type: "autocomplete", options: allMainIds },
        { id: "newId", label: "New Record ID", type: "text" }
    ]);

    if (!res || !res.sourceId || !res.newId) return;

    const sourceId = res.sourceId.trim();
    const newId = res.newId.trim();

    if (sourceId === "" || newId === "") return;
    if (allMainIds.includes(newId)) {
        showStatus(`Record ID '${newId}' already exists in '${dataTableName}'.`, "error");
        return;
    }

    const sourceRecord = mainDataSet.records.find((r: any) => String(r.__DC_ID__) === sourceId);
    if (!sourceRecord) {
        showStatus(`Source record '${sourceId}' not found.`, "error");
        return;
    }

    // Clone main record
    const clonedRecord = { ...sourceRecord };
    clonedRecord.__DC_ID__ = newId;
    const idF = mainDataSet.idField || mainDataSet.fields[0];
    clonedRecord[idF] = newId;

    mainDataSet.history = mainDataSet.history || {};
    mainDataSet.history[mainDataSet.revision] = {
        idField: mainDataSet.idField,
        fields: [...mainDataSet.fields],
        records: JSON.parse(JSON.stringify(mainDataSet.records))
    };
    mainDataSet.revision += 1;
    mainDataSet.records.push(clonedRecord);
    await applyCalculatedFields([clonedRecord], mainDataSet.fields, mainDataSet.calculatedFields, store, dataTableName);

    // Clone sub-records
    const relations = mainDataSet.relations || [];
    let clonedSubRecordsCount = 0;

    for (const rel of relations) {
        const subDataSet = store[rel.subTable];
        if (!subDataSet) continue;

        const subRecordsToClone = subDataSet.records.filter((r: any) => String(r[rel.foreignKey]) === sourceId);
        if (subRecordsToClone.length === 0) continue;

        const newSubRecords: any[] = [];
        subRecordsToClone.forEach((sr: any) => {
            const cloned = { ...sr };
            cloned.__DC_ID__ = `CLONED_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
            cloned[rel.foreignKey] = newId;
            const subIdF = subDataSet.idField || subDataSet.fields[0];
            if (subIdF) cloned[subIdF] = cloned.__DC_ID__;
            newSubRecords.push(cloned);
        });

        subDataSet.history = subDataSet.history || {};
        subDataSet.history[subDataSet.revision] = {
            idField: subDataSet.idField,
            fields: [...subDataSet.fields],
            records: JSON.parse(JSON.stringify(subDataSet.records))
        };
        subDataSet.revision += 1;
        await applyCalculatedFields(newSubRecords, subDataSet.fields, subDataSet.calculatedFields, store, rel.subTable);
        subDataSet.records = [...subDataSet.records, ...newSubRecords];
        clonedSubRecordsCount += newSubRecords.length;
    }

    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    await renderDashboard();
    await refreshFormulas(true);

    showStatus(`Successfully duplicated record '${sourceId}' to '${newId}' along with ${clonedSubRecordsCount} sub-records.`, "success");
  } catch (error: any) {
    showStatus("Error duplicating record: " + error.message, "error");
  }
}

export async function manageWorkspaces() {
    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};
    
    let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
    let familiesList: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];
    
    const currentFamilies = new Set<string>();
    Object.keys(store).forEach(k => currentFamilies.add(store[k].family || 'Public'));

    familiesList = familiesList.filter(f => currentFamilies.has(f));
    currentFamilies.forEach(f => { if (!familiesList.includes(f)) familiesList.push(f); });

    const res = await customManageListPrompt("Manage Workspaces", "Drag to reorder, edit to rename. Click the trash icon to mark for deletion.", familiesList, "Add Workspace", true);

    if (!res) return;

    let hasChanges = false;
    let newOrder: string[] = [];

    for (const r of res) {
        if (r.isDeleted) {
            if (!r.isNew) {
                const ws = r.original;
                const tablesInWs = Object.keys(store).filter(k => (store[k].family || 'Public') === ws);
                if (tablesInWs.length > 0) {
                    const confirm = await customConfirm("Delete Workspace", `Are you sure you want to permanently delete '${ws}' and ALL its ${tablesInWs.length} tables?`, "Yes, Delete Everything");
                    if (!confirm) {
                        newOrder.push(ws); // keep it if they cancelled
                        continue;
                    }
                    tablesInWs.forEach(t => delete store[t]);
                    for (const key of Object.keys(store)) {
                        if (store[key].relations) {
                            store[key].relations = store[key].relations.filter((rel: any) => !tablesInWs.includes(rel.subTable));
                        }
                    }
                }
                hasChanges = true;
            }
        } else {
            let finalName = r.original;
            if (r.isNew) {
                finalName = r.newName;
                hasChanges = true;
            } else if (r.newName !== r.original) {
                const oldName = r.original;
                finalName = r.newName;
                Object.keys(store).forEach(key => {
                    if ((store[key].family || 'Public') === oldName) {
                        store[key].family = finalName;
                        hasChanges = true;
                    }
                });
            }
            if (!newOrder.includes(finalName)) newOrder.push(finalName);
        }
    }

    await idbSet(IDB_KEYS.WORKSPACES_ORDER, JSON.stringify(newOrder));
    if (hasChanges) {
        await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    }
    renderDashboard();
}

export async function manageWorkspaceTables(fam: string) {
    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};
    
    let tableOrderRaw = await idbGet(IDB_KEYS.TABLES_ORDER);
    let tableOrder: string[] = tableOrderRaw ? JSON.parse(tableOrderRaw) : [];
    
    const tablesInWs = Object.keys(store).filter(k => (store[k].family || 'Public') === fam);
    
    tablesInWs.sort((a, b) => {
        const idxA = tableOrder.indexOf(a);
        const idxB = tableOrder.indexOf(b);
        if (idxA === -1 && idxB === -1) return a.localeCompare(b);
        if (idxA === -1) return 1;
        if (idxB === -1) return -1;
        return idxA - idxB;
    });

    const res = await customManageListPrompt(`Manage Tables: ${fam}`, "Drag to reorder. Click the trash icon to mark for deletion. (Renaming disabled)", tablesInWs, "Add New Table", false);
    if (!res) return;

    let hasChanges = false;
    let newOrder: string[] = tableOrder.filter(t => !tablesInWs.includes(t));

    for (const r of res) {
        if (r.isDeleted) {
            if (!r.isNew) {
                const t = r.original;
                const confirm = await customConfirm("Delete Table", `Permanently delete table '${t}'?`, "Yes, Delete");
                if (confirm) {
                    delete store[t];
                    for (const key of Object.keys(store)) {
                        if (store[key].relations) {
                            store[key].relations = store[key].relations.filter((rel: any) => rel.subTable !== t);
                        }
                    }
                    hasChanges = true;
                } else {
                    newOrder.push(t);
                }
            }
        } else {
            if (r.isNew) {
                setTimeout(() => addTableToWorkspacePrompt(fam, r.newName), 400);
            } else {
                newOrder.push(r.original);
            }
        }
    }

    await idbSet(IDB_KEYS.TABLES_ORDER, JSON.stringify(newOrder));
    if (hasChanges) {
        await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    }
    renderDashboard();
}

export async function addTableToWorkspacePrompt(fam: string, defaultName: string = "") {
    try {
        const storedData = await idbGet(IDB_KEYS.STORE);
        const store = storedData ? JSON.parse(storedData) : {};
        const allTableKeys = Object.keys(store);

        const res = await customFormPrompt(`Add Table`, `Create a new table in the '${fam}' workspace.`, [
            { id: "tableName", label: "Data Table Name", type: "text", value: defaultName },
            { id: "parentTable", label: "Parent Table (Optional Link)", type: "select", options: ["-- None --", ...allTableKeys] }
        ]);
        if (res && res.tableName) {
            await loadRangeForCapture(res.tableName, fam, res.parentTable === "-- None --" ? "" : res.parentTable);
        }
    } catch (error: any) {
        showStatus("Error adding table: " + error.message, "error");
    }
}

export async function addTableGlobal() {
    try {
        const storedData = await idbGet(IDB_KEYS.STORE);
        const store = storedData ? JSON.parse(storedData) : {};
        const allTableKeys = Object.keys(store);
        
        const familiesSet = new Set<string>();
        allTableKeys.forEach(k => familiesSet.add(store[k].family || 'Public'));
        const familiesList = Array.from(familiesSet);
        if (!familiesList.includes("Public")) familiesList.push("Public");

        const res = await customFormPrompt(`Capture New Table`, `Select a range in Excel with headers and data, provide a name, and select a workspace.`, [
            { id: "tableName", label: "Data Table Name", type: "text" },
            { id: "workspace", label: "Workspace", type: "autocomplete", options: familiesList, value: "Public" },
            { id: "parentTable", label: "Parent Table (Optional Link)", type: "select", options: ["-- None --", ...allTableKeys] }
        ]);
        
        if (res && res.tableName && res.tableName.trim() !== "") {
            const ws = (res.workspace && res.workspace.trim() !== "") ? res.workspace.trim() : "Public";
            await loadRangeForCapture(res.tableName.trim(), ws, res.parentTable === "-- None --" ? "" : res.parentTable);
        }
    } catch (error: any) {
        showStatus("Error adding table: " + error.message, "error");
    }
}