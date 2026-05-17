/* global Office */
import { idbGet } from "../utils/db";
import { t } from "../taskpane/taskpane";

let parsedStoreCache: { raw: string | null, data: any } = { raw: null, data: null };
let parsedVariablesCache: { raw: string | null, data: any } = { raw: null, data: null };

async function getParsedStore() {
    const raw = await idbGet("DC_STORE");
    if (raw === parsedStoreCache.raw && parsedStoreCache.data) {
        return parsedStoreCache.data;
    }
    const data = raw ? JSON.parse(raw) : {};
    parsedStoreCache = { raw, data };
    return data;
}

async function getParsedVariables() {
    const raw = await idbGet("DC_VARIABLES");
    if (raw === parsedVariablesCache.raw && parsedVariablesCache.data) {
        return parsedVariablesCache.data;
    }
    const data = raw ? JSON.parse(raw) : {};
    parsedVariablesCache = { raw, data };
    return data;
}

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

    if (!targetDataTable) return { records: [], fields: [], error: t("no_data_table_specified") };

    const store = await getParsedStore();
    if (!store || Object.keys(store).length === 0) return { records: [], fields: [], error: t("no_data_captured") };

    const dataSet = store[targetDataTable];
    if (!dataSet) return { records: [], fields: [], error: t("data_table_not_found") };

    let records = dataSet.records;
    let fields = dataSet.fields;

    if (finalRev !== undefined && finalRev !== null && finalRev !== (dataSet.revision || 1)) {
      if (dataSet.history && dataSet.history[finalRev]) {
        records = dataSet.history[finalRev].records;
        fields = dataSet.history[finalRev].fields;
      } else {
        return { records: [], fields: [], error: t("revision_not_found") };
      }
    }

    return { records, fields };
  } catch (error: any) {
    return { records: [], fields: [], error: `${t("db_error_prefix")}${error.message || error}` };
  }
}

export async function getGlobalVariable(varName: string): Promise<any> {
  try {
      await Office.onReady();
      const variables = await getParsedVariables();
      const formula = variables[varName];
      if (!formula) return t("variable_not_found");

      const store = await getParsedStore();
      
      const evaluateVar = (vName: string, visited: Set<string>): any => {
          if (visited.has(vName)) throw new Error(t("circular_reference_detected"));
          visited.add(vName);
          const vForm = variables[vName];
          if (!vForm) return 0;
          const DC = {
              SUM: (subTable: string, sumCol: string, fkCol?: string, fkValue?: any) => {
                  if (!store || !store[subTable]) return 0;
                  let total = 0;
                  store[subTable].records.forEach((r: any) => { if (!fkCol || String(r[fkCol]) === String(fkValue)) total += Number(r[sumCol]) || 0; });
                  return total;
              },
              COUNT: (subTable: string, fkCol?: string, fkValue?: any) => {
                  if (!store || !store[subTable]) return 0;
                  let count = 0;
                  store[subTable].records.forEach((r: any) => { if (!fkCol || String(r[fkCol]) === String(fkValue)) count++; });
                  return count;
              },
              VAR: (v: string) => evaluateVar(v, new Set(visited))
          };
          const func = new Function('store', 'DC', `return ${vForm};`);
          return func(store, DC);
      };

      return evaluateVar(varName, new Set());
  } catch (error: any) {
      return `${t("data_service_error_prefix")}${error.message || error}`;
  }
}