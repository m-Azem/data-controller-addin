/* global Office */
import { idbGet } from "../utils/db";

export async function getTableData(dataTableName?: string, rev?: number): Promise<{ records: any[], fields: string[], error?: string }> {
  try {
    await Office.onReady(); // Ensure Shared Runtime is fully initialized

    const defaultDataTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    const targetDataTable = dataTableName && String(dataTableName).trim() !== "" ? String(dataTableName) : defaultDataTable;

    let finalRev = rev;
    if ((rev === undefined || rev === null) && targetDataTable === defaultDataTable) {
      const defRevStr = await idbGet("DC_DEFAULT_REVISION");
      if (defRevStr) finalRev = parseInt(defRevStr);
    }

    if (!targetDataTable) return { records: [], fields: [], error: "Error: No data table specified." };

    const storedData = await idbGet("DC_STORE");
    if (!storedData) return { records: [], fields: [], error: "Error: No data captured." };

    const store = JSON.parse(storedData);
    const dataSet = store[targetDataTable];
    if (!dataSet) return { records: [], fields: [], error: "Error: Data table not found." };

    let records = dataSet.records;
    let fields = dataSet.fields;

    if (finalRev !== undefined && finalRev !== null && finalRev !== (dataSet.revision || 1)) {
      if (dataSet.history && dataSet.history[finalRev]) {
        records = dataSet.history[finalRev].records;
        fields = dataSet.history[finalRev].fields;
      } else {
        return { records: [], fields: [], error: "Error: Revision not found." };
      }
    }

    return { records, fields };
  } catch (error) {
    return { records: [], fields: [], error: `DB Error: ${error.message || error}` };
  }
}