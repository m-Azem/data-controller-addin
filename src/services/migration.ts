/* global OfficeRuntime */
import { idbGet, idbSet } from "../utils/db";

export async function migrateFromLocalStorage() {
  const isMigrated = await idbGet("MIGRATED_FROM_OS");
  if (!isMigrated) {
    const store = await OfficeRuntime.storage.getItem("DC_STORE");
    if (store) {
      await idbSet("DC_STORE", store);
    }
    const defaultTable = await OfficeRuntime.storage.getItem("DC_DEFAULT_DATA_TABLE");
    if (defaultTable) {
      await idbSet("DC_DEFAULT_DATA_TABLE", defaultTable);
    }
    const theme = await OfficeRuntime.storage.getItem("DC_THEME");
    if (theme) await idbSet("DC_THEME", theme);
    
    await idbSet("MIGRATED_FROM_OS", "true");
  }
}