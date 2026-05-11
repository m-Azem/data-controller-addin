import { idbGet, idbSet } from "../utils/db";

export async function exportCSVData(dataTableName: string, selectedRev: number, targetDataSet: any) {
  const escapeField = (field: any) => {
    if (field === null || field === undefined) return "";
    const str = String(field);
    if (str.includes('"') || str.includes(',') || str.includes('\n')) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const headersStr = targetDataSet.fields.map(escapeField).join(",");
  const rowsStr = targetDataSet.records.map((record: any) => 
    targetDataSet.fields.map((f: string) => escapeField(record[f])).join(",")
  ).join("\n");
  const csvContent = `${headersStr}\n${rowsStr}`;

  const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${dataTableName}_rev${selectedRev}_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function downloadBackup() {
  const storedData = await idbGet("DC_STORE");
  const variablesData = await idbGet("DC_VARIABLES");
  
  if (!storedData || storedData === "{}") throw new Error("No data to backup.");
  
  const backupObject = {
      type: "DataControllerBackup",
      version: 2,
      store: JSON.parse(storedData),
      variables: variablesData ? JSON.parse(variablesData) : {}
  };

  const blob = new Blob([JSON.stringify(backupObject, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `DataController_Backup_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function processRestoreFile(event: any): Promise<void> {
  return new Promise((resolve, reject) => {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return reject(new Error("No file selected."));
    
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = async (e: any) => {
      try {
        const content = e.target.result;
        const parsed = JSON.parse(content);
        if (typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("Invalid backup file format.");
        
        if (parsed.type === "DataControllerBackup" && parsed.store) {
            // New Format (Includes Variables)
            await idbSet("DC_STORE", JSON.stringify(parsed.store));
            if (parsed.variables) {
                await idbSet("DC_VARIABLES", JSON.stringify(parsed.variables));
            }
        } else {
            // Legacy Format (Store only)
            await idbSet("DC_STORE", JSON.stringify(parsed));
        }
        
        input.value = ""; // Reset input
        resolve();
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Error reading file."));
    reader.readAsText(file);
  });
}