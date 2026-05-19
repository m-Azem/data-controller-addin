/* global Excel, Office, OfficeRuntime */

import { idbGet, idbSet } from "../utils/db";
import { migrateFromLocalStorage } from "../services/migration";
import { exportCSVData, downloadBackup, processRestoreFile } from "../services/fileService";
import { executeInsertTable, executeInsertDropdown, executeConvertToValues, executeRefreshFormulas } from "../services/excelService";
import { TRANSLATIONS } from "../utils/translations";

const APP_VERSION = "1.0.2"; // Update this number whenever you release a new version

// Translation Dictionary
// const TRANSLATIONS: Record<string, Record<string, string>> = {
//     en: {
//         app_title: "Data Controller",
//         status_label: "Status:",
//         default_table: "Default Table (Formulas)",
//         revision: "Revision",
//         workspaces: "Workspaces",
//         formula_builder: "Visual Formula Builder",
//         refresh_btn_title: "Refresh Dashboard & Formulas",
//         convert_btn_title: "Convert DC Formulas to Values",
//         formula_builder_desc: "Select a function and fill in the parameters. Cell references (e.g. <code>A1</code>) are supported.",
//         select_function: "Select Function",
//         global_variables: "Global Variables",
//         global_variables_desc: "Define standalone variables using table data (e.g., <code>DC.SUM('Table', 'Col')</code>).",
//         add_variable: "Add Variable",
//         backup_restore: "Data Backup & Restore",
//         backup_restore_desc: "Download a backup of your data or restore it from a file.",
//         backup_data: "Backup Data",
//         restore_data: "Restore Data",
//         theme_appearance: "Theme Appearance",
//         dark_theme: "Enable Dark Theme",
//         language: "Language / اللغة",
//         help_doc: "Help & Documentation",
//         help_btn: "Open Help Center",
//         about: "About",
//         insert_btn: "Insert",
//         clear_btn: "Clear",
//         status_ready: "Ready",
//         global_relations: "Global Relations",
//         global_relations_desc: "Manage structural links between your tables.",
//         manage_relations: "Add / Manage Relations",
//         manage_workspaces: "Manage Workspaces",
//         add_workspace: "Add Workspace",
//         settings: "Settings",
//         help_desc: "Need assistance or want to learn how to use Data Controller?",
//         about_app_ver: "App Version:",
//         close_settings: "Close Settings",
//         tables_btn: "Tables",
//         target_revision: "Target Revision:",
//         insert_sheet: "Insert to Sheet",
//         replace_version: "Replace Version",
//         append_data: "Append Data",
//         snapshot: "Snapshot",
//         del_version: "Del Version",
//         del_table: "Del Table",
//         export_csv: "Export CSV",
//         headers_dropdown: "Headers Dropdown",
//         form_editor: "Form Editor",
//         grid_editor: "Grid Editor",
//         clone_record: "Clone Record",
//         manage_columns: "Manage Columns",
//         move_workspace: "Move Workspace",
//         new_revision: "New Revision",
//         clone_sub_records: "Clone Sub-records",
//         data_entry_views: "Data Entry & Views",
//         schema_operations: "Schema & Operations",
//         versioning_danger_zone: "Versioning & Danger Zone",
//         no_tables_in_ws: "No tables in this workspace.",
//         no_tables: "No data tables stored yet.",
//         no_workspaces: "No workspaces or data tables stored yet.",
//         manage_ws_title: "Manage Workspaces",
//         manage_ws_desc: "Drag to reorder, edit to rename. Click the trash icon to mark for deletion.",
//         manage_tb_title: "Manage Tables: {0}",
//         manage_tb_desc: "Drag to reorder. Click the trash icon to mark for deletion. (Renaming disabled)",
//         add_new_table: "Add New Table",
//         del_table_title: "Delete Entire Table",
//         del_table_msg: "Are you sure you want to permanently delete the table '{0}' and all its history? This action cannot be undone.",
//         del_table_confirm: "Yes, Delete Everything",
//         del_version_title: "Confirm Deletion",
//         del_version_latest_msg: "Are you sure you want to delete the current version of '{0}' and rollback to the previous revision?",
//         del_version_hist_msg: "Are you sure you want to delete historical Rev {0} from '{1}'?",
//         del_version_confirm: "Yes, Delete",
//         snapshot_locked_msg: "Locked Rev {0}. Current is now Rev {1}",
//         snapshot_restored_msg: "Restored Rev {0} as new active Rev {1}.",
//         edit_cols_title: "Edit Columns",
//         edit_cols_msg: "Drag to reorder, add new columns, rename them, or attach Calculated Formulas.",
//         edit_record_title: "Edit Record",
//         edit_record_msg: "Enter the Record ID to edit in '{0}':",
//         edit_record_update_msg: "Update the values below:",
//         edit_record_not_found: "Record '{0}' not found.",
//         export_csv_msg: "Exported {0} (Rev {1}) to CSV.",
//         insert_dropdown_msg: "Inserted headers dropdown for {0}",
//         insert_table_title: "Insert Table",
//         insert_table_msg: "Select columns to insert:",
//         insert_table_confirm: "Insert",
//         insert_table_success: "Inserted table for {0}",
//         split_editor_loading: "Loading Split Editor...",
//         split_editor_fix_err: "Please fix the errors in your current editor before opening a new one.",
//         split_editor_success: "Opened Split Editor for {0} and {1}. Arrange windows side-by-side.",
//         grid_editor_opened_sub: "Opened Sub-table ({0}: {1})",
//         grid_editor_opened: "Opened '{0}'",
//         manage_rel_title: "Manage Relations",
//         manage_rel_select_table: "Select a table to manage its relations:",
//         manage_rel_manage_msg: "Manage relations for '{0}':",
//         manage_rel_keep: "Existing Relations (Uncheck to remove)",
//         manage_rel_add_target: "Add New: Target Table",
//         manage_rel_add_link: "Add New: Link Column",
//         manage_rel_none: "-- None --",
//         append_data_title: "Append Data",
//         append_data_msg: "Review data to append to '{0}' ({1} rows detected).",
//         clone_sub_title: "Clone Sub-records",
//         clone_sub_select_target: "Select the target sub-table to clone:",
//         clone_sub_relation: "Sub-table Relation",
//         clone_from: "Clone from '{0}'",
//         clone_select_ids: "Select the Source ID and Target ID(s):",
//         clone_source_id: "Source Record ID (Copy FROM)",
//         clone_target_ids: "Target Record IDs (Copy TO)",
//         move_ws_title: "Move Workspace",
//         move_ws_msg: "Select or type a new workspace for '{0}':",
//         add_var_title: "Add Variable",
//         add_var_msg: "Define a global variable:",
//         add_var_name: "Variable Name",
//         add_var_formula: "Formula Definition",
//         add_ws_title: "Add Workspace",
//         add_ws_msg: "Enter a name for the new workspace:",
//         add_ws_added_title: "Workspace Added",
//         add_ws_added_msg: "The workspace '{0}' was created. Do you want to capture a new data table for it now?",
//         add_ws_capture_btn: "Yes, Capture Table",
//         dup_record_title: "Duplicate Record",
//         dup_record_msg: "Select the record to duplicate and provide a new ID:",
//         dup_record_new_id: "New Record ID",
//         delete_ws_title: "Delete Workspace",
//         delete_ws_msg: "Are you sure you want to permanently delete '{0}' and ALL its {1} tables?",
//         delete_tb_title: "Delete Table",
//         delete_tb_msg: "Permanently delete table '{0}'?",
//         add_tb_title: "Add Table",
//         add_tb_msg: "Create a new table in the '{0}' workspace.",
//         add_tb_name: "Data Table Name",
//         add_tb_parent: "Parent Table (Optional Link)",
//         capture_tb_title: "Capture New Table",
//         capture_tb_msg: "Review and map columns for '{0}' ({1} rows detected).",
//         select_range_msg: "Please select a range of cells in Excel that includes your headers and data, then click Next.",
//         link_sub_title: "Link Sub-table",
//         link_sub_msg: "Which column in '{0}' links to '{1}'?",
//         link_col: "Link Column",
//         review_formulas_title: "Review Detected Formulas",
//         review_formulas_msg: "We converted your Excel formulas. Adjust them below using the visual formula builder, or clear the text to skip:",
//         save_formulas: "Save Formulas",
//         replace_ver_title: "Replace Version",
//         replace_ver_msg: "Review replacing data for '{0}' ({1} rows detected).",
//         capture_rev_title: "Capture New Revision",
//         capture_rev_msg: "Review data for new revision of '{0}' ({1} rows detected).",
//         mark_for_deletion: "Mark for Deletion",
//         mark_deletion_msg: "Are you sure you want to delete '{0}'? It will be permanently removed when you save changes.",
//         delete_link_title: "Delete Link",
//         delete_link_msg: "Are you sure you want to permanently remove the link between '{0}' and '{1}'?",
//         link_removed_success: "Link removed successfully.",
//         add_rollup_title: "Add Rollup",
//         add_rollup_msg: "Would you like to add a Calculated Field in '{0}' to summarize data from '{1}'?",
//         rollup_field_title: "Rollup Field",
//         rollup_field_msg: "Define the rollup:",
//         rollup_field_name: "New Field Name (e.g. Total Cost)",
//         rollup_type: "Aggregation Type",
//         rollup_col: "Column to Aggregate (for SUM)",
//         no_relations: "No relations defined yet.",
//         rev_latest: "Rev {0} (Latest)",
//         rev_history: "Rev {0}",
//         rows: "rows",
//         "Data Table Name": "Data Table Name",
//         "Revision": "Revision",
//         "Source Record ID": "Source Record ID",
//         formula_for: "Formula for '{0}'",
//         define_calc_formula: "Define calculated formula (leave empty to remove):",
//         yes_btn: "Yes",
//         funcs_label: "FUNCS:",
//         vars_label: "VARS:",
//         tables_label: "TABLES:",
//         fields_label: "Fields:",
//         main_data_table: "Main Data Table",
//         next_btn: "Next",
//         cancel_btn: "Cancel",
//         save_changes_btn: "Save Changes",
//         add_column_btn: "Add Column",
//         delete_selected_btn: "Delete Selected",
//         save_current_rev_btn: "Save Current Revision",
//         save_new_rev_btn: "Save as New Revision",
//         back_btn: "Back",
//         optional_label: "(Optional)",
//         wizard_step_indicator: "Step {0} / {1}",
//         ok_btn: "OK",
//         edit_record_title_with_id: "Editing Record: {0}",
//         columns_label: "Columns",
//         workspace_label: "Workspace",
//         source_record_id_label: "Source Record ID",
//         manage_btn: "Manage",
//         refresh_btn: "Refresh",
//         editor_mode_desc: "Make your changes directly in the spreadsheet. Tip: Go to View > New Window > Arrange All to edit this table side-by-side with your main data. Click Save when you are finished to sync the data.",
//         live_sync_active: "<i class=\"ms-Icon ms-Icon--SyncOccurence\"></i> Live Sync Active",
//         editing_title: "You are currently editing table ",
//         error_general: "Error: ",
//         no_data_table_specified: "Error: No data table specified.",
//         no_data_captured: "Error: No data captured.",
//         data_table_not_found: "Error: Data table not found.",
//         revision_not_found: "Error: Revision not found.",
//         db_error_prefix: "DB Error: ",
//         variable_not_found: "Error: Variable not found",
//         circular_reference_detected: "Circular reference detected",
//         data_service_error_prefix: "Error: ",
//         no_data_to_backup: "No data to backup.",
//         no_file_selected: "No file selected.",
//         invalid_backup_file_format: "Invalid backup file format.",
//         error_reading_file: "Error reading file.",
//         not_found: "Not Found",
//         field_error: "Field Error",
//         func_error_prefix: "Func Error: ",
//         base_table_error_prefix: "Base Table Error: ",
//         foreign_table_error_prefix: "Foreign Table Error: ",
//         na_value: "N/A",
//         sort_field_not_found: "Sort Field '{0}' not found",
//         select_option: "Select...",
//         true_text: "TRUE",
//         false_text: "FALSE",
//         field_is_required: "'{0}' is required.",
//         inserted_formula: "Inserted formula: {0}",
//         column_generic: "Column {0}",
//         name_placeholder: "Name",
//         id_badge: "(ID)",
//         cannot_delete_id_column_msg: "Cannot delete the Primary ID column ('{0}').",
//         table_name_required_error: "Table Name is required.",
//         table_already_exists_error: "Table '{0}' already exists.",
//         workspace_name_required_error: "Workspace name is required.",
//         primary_id_column_missing_error: "Primary ID column is missing.",
//         row_has_empty_id_error: "Row {0} has an empty ID.",
//         duplicate_id_found_error: "Duplicate ID found: '{0}'.",
//         all_column_names_must_be_unique_error: "All column names must be unique.",
//         link_column_required_error: "Link Column is required.",
//         select_range_for_table_creation: "Select a range with headers and data to create a table.",
//         error_during_capture: "Error during cell capture.",
//         select_cell_to_capture: "Select a cell to capture its reference...",
//         public_workspace: "Public",
//         saved_records_in_table: "Saved {0} records in {1}.",
//         formula_detected_label: "{0} (Detected: {1})",
//         mapped_formulas_for: "Mapped formulas for: {0}",
//         select_range_with_headers_and_data: "Select a range with headers and data.",
//         data_table_not_found_error: "Data table '{0}' not found.",
//         replaced_table_with_records: "Replaced '{0}' with {1} new records. (ID: '{2}')",
//         replaced_rev_of_table_with_records: "Replaced Rev {0} of '{1}' with {2} records.",
//         error_replacing_data: "Error replacing data: ",
//         captured_rev_for_table: "Captured Rev {0} for '{1}' with {2} records.",
//         error_capturing_new_revision: "Error capturing new revision: ",
//         no_data_tables_stored_yet: "No data tables stored yet.",
//         no_workspaces_or_tables_stored_yet: "No workspaces or data tables stored yet.",
//         rev_latest_history: "Rev {0}{1}",
//         rev_latest_suffix: " (Latest)",
//         link_column: "Link Column",
//         open_split_editor_btn: "Open Split Editor",
//         delete_link_btn: "Delete Link",
//         rev_count_rows: "Rev {0} • {1} rows",
//         revision_latest_history: "Revision {0}{1}",
//         restore_snapshot: "Restore",
//         no_tables_in_this_workspace: "No tables in this workspace.",
//         deleted_entire_data_table: "Deleted entire data table: {0}",
//         deleted_current_version_rolled_back: "Deleted current version. Rolled back to Rev {0}.",
//         deleted_only_version_table_removed: "Deleted the only version. Table '{0}' removed.",
//         deleted_historical_rev: "Deleted historical Rev {0} from '{1}'.",
//         error_deleting_version: "Error deleting version: ",
//         locked_rev_current_is: "Locked Rev {0}. Current is now Rev {1}",
//         restored_rev_as_new_active: "Restored Rev {0} as new active Rev {1}.",
//         storage_limit_reached: "Storage limit reached! Please clear space.",
//         error_creating_snapshot: "Error creating snapshot: ",
//         cannot_resort_historical_revision: "Cannot resort columns of a historical revision.",
//         column_names_must_be_unique: "Column names must be unique.",
//         columns_updated_successfully: "Columns updated successfully. Current is Rev {0}.",
//         enter_record_id_to_edit: "Enter the Record ID to edit in '{0}':",
//         record_not_found_error: "Record '{0}' not found.",
//         calculated_label: " (Calculated)",
//         record_updated_refreshing_excel: "Record updated. Refreshing Excel...",
//         storage_limit_reached_cannot_save: "Storage limit reached! Cannot save changes.",
//         error_saving_changes: "Error saving changes: ",
//         no_data_found_error: "No data found.",
//         no_data_found_for_revision_error: "No data found for this revision.",
//         exported_table_to_csv: "Exported {0} (Rev {1}) to CSV.",
//         error_exporting_csv: "Error exporting CSV: ",
//         refreshing_dashboard_formulas: "Refreshing dashboard & formulas... Please wait.",
//         refreshed_dashboard_formulas: "Refreshed dashboard and {0} DC formulas.",
//         converting_please_wait: "Converting... Please wait.",
//         converted_formulas_to_values: "Converted {0} DC formulas to values.",
//         backup_downloaded_successfully: "Backup downloaded successfully.",
//         backup_error: "Backup error: ",
//         data_restored_successfully: "Data restored successfully.",
//         restore_error: "Restore error: ",
//         no_headers_found_error: "No headers found.",
//         inserted_headers_dropdown: "Inserted headers dropdown for {0}",
//         error_inserting_dropdown: "Error inserting dropdown: ",
//         no_data_found_for_table_error: "No data found for this data table.",
//         select_columns_to_insert: "Select columns to insert:",
//         excel_formula_js_syntax_error: "Formulas contained invalid syntax and were inserted as text.",
//         inserted_table_for: "Inserted table for {0}",
//         error_inserting_table: "Error inserting table: ",
//         unknown_table_name: "Unknown",
//         loading_split_editor: "Loading Split Editor...",
//         fix_errors_before_new_editor: "Please fix the errors in your current editor before opening a new one.",
//         new_record_id_placeholder: "NEW",
//         opened_split_editor_arrange_windows: "Opened Split Editor for {0} and {1}. Arrange windows side-by-side.",
//         error_opening_split_editor: "Error opening split editor: ",
//         opened_sub_table_in_editor: "Opened Sub-table ({0}: {1})",
//         opened_table_in_grid_editor: "Opened '{0}' in Grid Editor.",
//         error_opening_grid_editor: "Error opening Grid Editor: ",
//         row_has_empty_id_in_table_error: "Row {0} has an empty ID in {1}.",
//         duplicate_id_found_in_table_error: "Duplicate ID found: '{0}' in {1}.",
//         saved_records: "Saved {0} records.",
//         no_changes_detected: "No changes detected.",
//         error_saving_grid_editor: "Error saving Grid Editor: ",
//         auto_syncing_switching_to_record: "Auto-syncing... Switching to Record {0}",
//         switched_sub_table_to_record: "Switched Sub-table to Record {0}",
//         error_switching_record: "Error switching record: ",
//         error_canceling_grid_editor: "Error canceling Grid Editor: ",
//         relation_link_suffix: "{0} (Link: {1})",
//         relations_updated_for: "Relations updated for '{0}'",
//         define_the_rollup: "Define the rollup:",
//         added_rollup_field: "Added rollup field '{0}'",
//         error_managing_relations: "Error managing relations: ",
//         cannot_append_to_historical_revision: "Cannot append data to a historical revision.",
//         review_data_to_append: "Review data to append to '{0}' ({1} rows detected).",
//         duplicate_id_already_exists_error: "Duplicate ID found: '{0}' already exists in '{1}'.",
//         appended_records_current_is: "Appended {0} records. Current is Rev {1}.",
//         error_appending_data: "Error appending data: ",
//         no_relations_defined_for_table: "No relations defined for this table.",
//         relation_fk_suffix: "{0} (FK: {1})",
//         select_target_sub_table_to_clone: "Select the target sub-table to clone:",
//         clone_from_table: "Clone from '{0}'",
//         select_source_target_ids: "Select the Source ID and Target ID(s):",
//         no_sub_records_found: "No sub-records found for Source ID '{0}'.",
//         cloned_id_prefix: "CLONED_",
//         successfully_cloned_sub_records: "Successfully cloned {0} sub-records to {1} target(s).",
//         error_cloning_sub_records: "Error cloning sub-records: ",
//         select_or_type_new_workspace: "Select or type a new workspace for '{0}':",
//         moved_table_to_workspace: "Moved '{0}' to '{1}'.",
//         error_moving_table: "Error moving table: ",
//         variable_name_required: "Variable Name is required.",
//         formula_required: "Formula is required.",
//         loop_detected: "Loop Detected",
//         variable_cannot_reference_itself: "Variable cannot reference itself",
//         variable_saved: "Variable '{0}' saved.",
//         invalid_formula_for_variable: "Invalid Formula for '{0}': {1}",
//         error_adding_workspace: "Error adding workspace: ",
//         error_text: "ERROR",
//         loop_error: "LOOP ERROR",
//         insert_to_sheet: "Insert to Sheet",
//         inserted_variable_to_sheet: "Inserted variable '{0}' to sheet.",
//         error_inserting_variable: "Error inserting variable: ",
//         delete_variable: "Delete Variable",
//         record_id_already_exists: "Record ID '{0}' already exists in '{1}'.",
//         source_record_not_found: "Source record '{0}' not found.",
//         successfully_duplicated_record: "Successfully duplicated record '{0}' to '{1}' along with {2} sub-records.",
//         error_duplicating_record: "Error duplicating record: ",
//         param_record_id: "Record ID",
//         param_field_name: "Field Name",
//         param_data_table_name: "Data Table Name",
//         param_revision: "Revision",
//         param_search_field: "Search Field",
//         param_search_value: "Search Value",
//         param_return_field: "Return Field",
//         param_exact_match: "Exact Match (TRUE/FALSE)",
//         param_sum_field: "Sum Field",
//         param_criteria_field: "Criteria Field",
//         param_criteria_value: "Criteria Value",
//         param_base_table_name: "Base Table Name",
//         param_link_column: "Link Column",
//         param_target_table_name: "Target Table Name",
//         param_target_return_field: "Target Return Field",
//         param_sort_field: "Sort Field",
//         param_ascending: "Ascending (TRUE/FALSE)",
//         param_variable_name: "Variable Name",
//         save_btn: "Save",
//         loading_text: "Loading...",
//         col_name_placeholder: "Column Name",
//         new_col_name_placeholder: "New Column Name",
//         new_column_default: "New Column",
//         formula_def_label: "Formula Definition",
//         edit_formula_tooltip: "Edit Formula: {0}",
//         add_formula_tooltip: "Add Formula",
//         cannot_delete_id_col_tooltip: "Cannot delete ID column",
//         primary_id_col_tooltip: "Primary ID Column",
//         id_col_must_be_present: "The Primary ID column must be present in the table.",
//         row_empty_id_strict: "Row {0} has an empty ID. IDs must be non-empty.",
//         duplicate_id_strict: "Duplicate ID found: '{0}'. All IDs must be unique.",
//         this_item: "this item",
//         delete_tooltip: "Delete",
//         current_table_opt: "Current Table",
//         select_table_opt: "-- Select Table --",
//         formula_builder_placeholder: "Type or click above to build formula...",
//         table_name_label: "Table Name",
//         table_placeholder: "e.g. Customers",
//         workspace_placeholder: "Select or type new",
//         parent_link_label: "Parent Table (Optional Link)",
//         ready_to_save: "Ready to Save",
//         save_table_btn: "Save Table",
//         add_new_column_btn: "Add New Column",
//         save_current_btn: "Save to Current Version",
//         save_new_btn: "Save to New Version",
//         grid_editor_mode: "Grid Editor Mode",
//         currently_editing: "You are currently editing",
//         grid_editor_tip: "Make your changes directly in the spreadsheet.<br/>Tip: Go to View > New Window > Arrange All to edit this table side-by-side with your main data.<br/>Click Save when you are finished to sync the data.",
//         save_sync_btn: "Save & Sync Changes",
//         cancel_close_btn: "Cancel & Close",
//         input_title: "Input",
//         primary_id_column: "Primary ID Column (Must be unique)",
//         columns_drag_label: "Columns (Drag to reorder, uncheck to drop, click to rename)",
//         validation_passed: "Validation Passed",
//         records_count: "Records:",
//         id_column_name: "ID Column:",
//         final_columns: "Final Columns:",
//         next_validate_btn: "Next (Validate)",
//         back_mapping_btn: "Back to Mapping",
//         confirm_save_btn: "Confirm & Save",
//         link_col_fk: "Link Column (Foreign Key)",
//         select_link_col_1: "Select the column in this new table that links to records in",
//         yes_delete_btn: "Yes, Delete",
//         add_new_btn: "Add New"
//     },
//     ar: {
//         app_title: "مدير البيانات",
//         status_label: "الحالة:",
//         default_table: "الجدول الافتراضي (للمعادلات)",
//         revision: "النسخة",
//         workspaces: "مساحات العمل",
//         formula_builder: "منشئ المعادلات المرئي",
//         refresh_btn_title: "تحديث لوحة المعلومات والمعادلات",
//         convert_btn_title: "تحويل معادلات DC إلى قيم",
//         formula_builder_desc: "حدد دالة واملأ المعلمات. مراجع الخلايا (مثل <code>A1</code>) مدعومة.",
//         select_function: "تحديد الدالة",
//         global_variables: "المتغيرات العامة",
//         global_variables_desc: "حدد متغيرات مستقلة باستخدام بيانات الجدول (مثل <code>DC.SUM('Table', 'Col')</code>).",
//         add_variable: "إضافة متغير",
//         backup_restore: "النسخ الاحتياطي والاستعادة",
//         backup_restore_desc: "تنزيل نسخة احتياطية من بياناتك أو استعادتها من ملف.",
//         backup_data: "نسخ احتياطي",
//         restore_data: "استعادة البيانات",
//         theme_appearance: "المظهر",
//         dark_theme: "تفعيل الوضع الداكن",
//         language: "Language / اللغة",
//         help_doc: "المساعدة والوثائق",
//         help_btn: "فتح مركز المساعدة",
//         about: "حول التطبيق",
//         insert_btn: "إدراج",
//         clear_btn: "مسح",
//         status_ready: "جاهز",
//         global_relations: "العلاقات العامة",
//         global_relations_desc: "إدارة الروابط الهيكلية بين جداولك.",
//         manage_relations: "إضافة / إدارة العلاقات",
//         manage_workspaces: "إدارة مساحات العمل",
//         add_workspace: "إضافة مساحة عمل",
//         settings: "الإعدادات",
//         help_desc: "هل تحتاج إلى المساعدة أو تريد معرفة كيفية استخدام متحكم البيانات؟",
//         about_app_ver: "إصدار التطبيق:",
//         close_settings: "إغلاق الإعدادات",
//         tables_btn: "الجداول",
//         target_revision: "النسخة المستهدفة:",
//         insert_sheet: "إدراج في الورقة",
//         replace_version: "استبدال النسخة",
//         append_data: "إلحاق بيانات",
//         snapshot: "لقطة",
//         del_version: "حذف النسخة",
//         del_table: "حذف الجدول",
//         export_csv: "تصدير CSV",
//         headers_dropdown: "قائمة العناوين",
//         form_editor: "محرر النموذج",
//         grid_editor: "محرر الشبكة",
//         clone_record: "استنساخ السجل",
//         manage_columns: "إدارة الأعمدة",
//         move_workspace: "نقل مساحة العمل",
//         new_revision: "نسخة جديدة",
//         clone_sub_records: "استنساخ السجلات الفرعية",
//         data_entry_views: "إدخال البيانات والعروض",
//         schema_operations: "المخطط والعمليات",
//         versioning_danger_zone: "إدارة النسخ ومنطقة الخطر",
//         no_tables_in_ws: "لا توجد جداول في مساحة العمل هذه.",
//         no_tables: "لم يتم حفظ أي جداول بيانات بعد.",
//         no_workspaces: "لم يتم حفظ مساحات عمل أو جداول بيانات بعد.",
//         manage_ws_title: "إدارة مساحات العمل",
//         manage_ws_desc: "اسحب لإعادة الترتيب، أو حرر لإعادة التسمية. انقر على أيقونة سلة المهملات للحذف.",
//         manage_tb_title: "إدارة الجداول: {0}",
//         manage_tb_desc: "اسحب لإعادة الترتيب. انقر على أيقونة سلة المهملات للحذف. (إعادة التسمية معطلة)",
//         add_new_table: "إضافة جدول جديد",
//         del_table_title: "حذف الجدول بالكامل",
//         del_table_msg: "هل أنت متأكد أنك تريد حذف الجدول '{0}' وجميع نسخه نهائياً؟ هذا الإجراء لا يمكن التراجع عنه.",
//         del_table_confirm: "نعم، احذف كل شيء",
//         del_version_title: "تأكيد الحذف",
//         del_version_latest_msg: "هل أنت متأكد أنك تريد حذف النسخة الحالية من '{0}' والعودة إلى النسخة السابقة؟",
//         del_version_hist_msg: "هل أنت متأكد أنك تريد حذف النسخة التاريخية {0} من '{1}'؟",
//         del_version_confirm: "نعم، احذف",
//         snapshot_locked_msg: "تم قفل النسخة {0}. النسخة الحالية الآن هي {1}",
//         snapshot_restored_msg: "تمت استعادة النسخة {0} كنسخة نشطة جديدة {1}.",
//         edit_cols_title: "تحرير الأعمدة",
//         edit_cols_msg: "اسحب لإعادة الترتيب، أضف أعمدة جديدة، أعد تسميتها، أو قم بإرفاق معادلات محسوبة.",
//         edit_record_title: "تعديل سجل",
//         edit_record_msg: "أدخل معرف السجل المراد تعديله في '{0}':",
//         edit_record_update_msg: "قم بتحديث القيم أدناه:",
//         edit_record_not_found: "لم يتم العثور على السجل '{0}'.",
//         export_csv_msg: "تم تصدير {0} (النسخة {1}) إلى CSV.",
//         insert_dropdown_msg: "تم إدراج قائمة عناوين لـ {0}",
//         insert_table_title: "إدراج جدول",
//         insert_table_msg: "حدد الأعمدة المراد إدراجها:",
//         insert_table_confirm: "إدراج",
//         insert_table_success: "تم إدراج الجدول لـ {0}",
//         split_editor_loading: "جاري تحميل محرر الانقسام...",
//         split_editor_fix_err: "يرجى إصلاح الأخطاء في المحرر الحالي قبل فتح محرر جديد.",
//         split_editor_success: "تم فتح محرر الانقسام لـ {0} و {1}. قم بترتيب النوافذ جنباً إلى جنب.",
//         grid_editor_opened_sub: "تم فتح الجدول الفرعي ({0}: {1})",
//         grid_editor_opened: "تم فتح '{0}'",
//         manage_rel_title: "إدارة العلاقات",
//         manage_rel_select_table: "حدد جدولاً لإدارة علاقاته:",
//         manage_rel_manage_msg: "إدارة علاقات '{0}':",
//         manage_rel_keep: "العلاقات الحالية (قم بإلغاء التحديد للإزالة)",
//         manage_rel_add_target: "إضافة جديد: الجدول المستهدف",
//         manage_rel_add_link: "إضافة جديد: عمود الربط",
//         manage_rel_none: "-- لا شيء --",
//         append_data_title: "إلحاق بيانات",
//         append_data_msg: "راجع البيانات المراد إلحاقها بـ '{0}' (تم اكتشاف {1} صفوف).",
//         clone_sub_title: "استنساخ السجلات الفرعية",
//         clone_sub_select_target: "حدد الجدول الفرعي المستهدف للاستنساخ:",
//         clone_sub_relation: "علاقة الجدول الفرعي",
//         clone_from: "استنساخ من '{0}'",
//         clone_select_ids: "حدد المعرف المصدر والمعرف(ات) الهدف:",
//         clone_source_id: "معرف السجل المصدر (نسخ من)",
//         clone_target_ids: "معرفات السجل الهدف (نسخ إلى)",
//         move_ws_title: "نقل مساحة العمل",
//         move_ws_msg: "حدد أو اكتب مساحة عمل جديدة لـ '{0}':",
//         add_var_title: "إضافة متغير",
//         add_var_msg: "قم بتعريف متغير عام:",
//         add_var_name: "اسم المتغير",
//         add_var_formula: "تعريف المعادلة",
//         add_ws_title: "إضافة مساحة عمل",
//         add_ws_msg: "أدخل اسماً لمساحة العمل الجديدة:",
//         add_ws_added_title: "تمت إضافة مساحة العمل",
//         add_ws_added_msg: "تم إنشاء مساحة العمل '{0}'. هل تريد التقاط جدول بيانات جديد لها الآن؟",
//         add_ws_capture_btn: "نعم، التقاط الجدول",
//         dup_record_title: "استنساخ سجل",
//         dup_record_msg: "حدد السجل المراد استنساخه وقدم معرفاً جديداً:",
//         dup_record_new_id: "معرف السجل الجديد",
//         delete_ws_title: "حذف مساحة العمل",
//         delete_ws_msg: "هل أنت متأكد أنك تريد حذف '{0}' نهائياً وجميع جداولها البالغ عددها {1}؟",
//         delete_tb_title: "حذف الجدول",
//         delete_tb_msg: "هل تريد حذف الجدول '{0}' نهائياً؟",
//         add_tb_title: "إضافة جدول",
//         add_tb_msg: "إنشاء جدول جديد في مساحة العمل '{0}'.",
//         add_tb_name: "اسم جدول البيانات",
//         add_tb_parent: "الجدول الأصل (رابط اختياري)",
//         capture_tb_title: "التقاط جدول جديد",
//         capture_tb_msg: "راجع وقم بتعيين الأعمدة لـ '{0}' (تم اكتشاف {1} صفوف).",
//         select_range_msg: "يرجى تحديد نطاق من الخلايا في Excel يتضمن العناوين والبيانات الخاصة بك، ثم انقر فوق التالي.",
//         link_sub_title: "ربط جدول فرعي",
//         link_sub_msg: "أي عمود في '{0}' يرتبط بـ '{1}'؟",
//         link_col: "عمود الربط",
//         review_formulas_title: "مراجعة المعادلات المكتشفة",
//         review_formulas_msg: "قمنا بتحويل معادلات Excel الخاصة بك. قم بضبطها أدناه باستخدام منشئ المعادلات المرئي، أو امسح النص للتخطي:",
//         save_formulas: "حفظ المعادلات",
//         replace_ver_title: "استبدال النسخة",
//         replace_ver_msg: "راجع استبدال البيانات لـ '{0}' (تم اكتشاف {1} صفوف).",
//         capture_rev_title: "التقاط نسخة جديدة",
//         capture_rev_msg: "راجع بيانات النسخة الجديدة من '{0}' (تم اكتشاف {1} صفوف).",
//         mark_for_deletion: "تحديد للحذف",
//         mark_deletion_msg: "هل أنت متأكد أنك تريد حذف '{0}'؟ ستتم إزالته نهائياً عند حفظ التغييرات.",
//         delete_link_title: "حذف الرابط",
//         delete_link_msg: "هل أنت متأكد أنك تريد الإزالة النهائية للرابط بين '{0}' و '{1}'؟",
//         link_removed_success: "تمت إزالة الرابط بنجاح.",
//         add_rollup_title: "إضافة حقل تجميعي",
//         add_rollup_msg: "هل ترغب في إضافة حقل محسوب في '{0}' لتلخيص البيانات من '{1}'؟",
//         rollup_field_title: "حقل تجميعي",
//         rollup_field_msg: "قم بتعريف التجميع:",
//         rollup_field_name: "اسم الحقل الجديد (مثل التكلفة الإجمالية)",
//         rollup_type: "نوع التجميع",
//         rollup_col: "العمود المراد تجميعه (للـ SUM)",
//         no_relations: "لا توجد علاقات محددة بعد.",
//         rev_latest: "النسخة {0} (الأحدث)",
//         rev_history: "النسخة {0}",
//         rows: "صفوف",
//         "Data Table Name": "اسم جدول البيانات",
//         "Revision": "النسخة",
//         "Source Record ID": "معرف السجل المصدر",
//         formula_for: "معادلة لـ '{0}'",
//         define_calc_formula: "قم بتعريف المعادلة المحسوبة (اتركه فارغاً للإزالة):",
//         yes_btn: "نعم",
//         funcs_label: "الدوال:",
//         vars_label: "المتغيرات:",
//         tables_label: "الجداول:",
//         fields_label: "الحقول:",
//         main_data_table: "جدول البيانات الأساسي",
//         next_btn: "التالي",
//         cancel_btn: "إلغاء",
//         save_changes_btn: "حفظ التغييرات",
//         add_column_btn: "إضافة عمود",
//         delete_selected_btn: "حذف المحدد",
//         save_current_rev_btn: "حفظ في النسخة الحالية",
//         save_new_rev_btn: "حفظ كنسخة جديدة",
//         back_btn: "رجوع",
//         optional_label: "(اختياري)",
//         wizard_step_indicator: "الخطوة {0} / {1}",
//         ok_btn: "موافق",
//         edit_record_title_with_id: "تعديل السجل: {0}",
//         columns_label: "الأعمدة",
//         workspace_label: "مساحة العمل",
//         source_record_id_label: "معرف السجل المصدر",
//         manage_btn: "إدارة",
//         refresh_btn: "تحديث",
//         editor_mode_desc: "قم بإجراء تغييراتك مباشرة في ورقة العمل. تلميح: انتقل إلى عرض > نافذة جديدة > ترتيب الكل لتعديل هذا الجدول جنباً إلى جنب مع بياناتك الأساسية. انقر فوق حفظ عند الانتهاء لمزامنة البيانات.",
//         live_sync_active: "<i class=\"ms-Icon ms-Icon--SyncOccurence\"></i> المزامنة المباشرة نشطة",
//         editing_title: "أنت تقوم حالياً بتعديل الجدول ",
//         error_general: "خطأ: ",
//         no_data_table_specified: "خطأ: لم يتم تحديد جدول بيانات.",
//         no_data_captured: "خطأ: لم يتم التقاط بيانات.",
//         data_table_not_found: "خطأ: لم يتم العثور على جدول البيانات.",
//         revision_not_found: "خطأ: لم يتم العثور على النسخة.",
//         db_error_prefix: "خطأ في قاعدة البيانات: ",
//         variable_not_found: "خطأ: المتغير غير موجود",
//         circular_reference_detected: "تم اكتشاف مرجع دائري",
//         data_service_error_prefix: "خطأ: ",
//         no_data_to_backup: "لا توجد بيانات للنسخ الاحتياطي.",
//         no_file_selected: "لم يتم تحديد ملف.",
//         invalid_backup_file_format: "تنسيق ملف النسخ الاحتياطي غير صالح.",
//         error_reading_file: "خطأ في قراءة الملف.",
//         not_found: "غير موجود",
//         field_error: "خطأ في الحقل",
//         func_error_prefix: "خطأ في الدالة: ",
//         base_table_error_prefix: "خطأ في الجدول الأساسي: ",
//         foreign_table_error_prefix: "خطأ في الجدول الأجنبي: ",
//         na_value: "غير متوفر",
//         sort_field_not_found: "حقل الفرز '{0}' غير موجود",
//         select_option: "اختر...",
//         true_text: "صحيح",
//         false_text: "خطأ",
//         field_is_required: "'{0}' مطلوب.",
//         inserted_formula: "تم إدراج المعادلة: {0}",
//         column_generic: "عمود {0}",
//         name_placeholder: "الاسم",
//         id_badge: "(معرف)",
//         cannot_delete_id_column_msg: "لا يمكن حذف عمود المعرف الأساسي ('{0}').",
//         table_name_required_error: "اسم الجدول مطلوب.",
//         table_already_exists_error: "الجدول '{0}' موجود بالفعل.",
//         workspace_name_required_error: "اسم مساحة العمل مطلوب.",
//         primary_id_column_missing_error: "عمود المعرف الأساسي مفقود.",
//         row_has_empty_id_error: "الصف {0} يحتوي على معرف فارغ.",
//         duplicate_id_found_error: "تم العثور على معرف مكرر: '{0}'.",
//         all_column_names_must_be_unique_error: "يجب أن تكون جميع أسماء الأعمدة فريدة.",
//         link_column_required_error: "عمود الربط مطلوب.",
//         select_range_for_table_creation: "حدد نطاقاً يحتوي على عناوين وبيانات لإنشاء جدول.",
//         error_during_capture: "خطأ أثناء التقاط الخلية.",
//         select_cell_to_capture: "حدد خلية لالتقاط مرجعها...",
//         public_workspace: "عام",
//         saved_records_in_table: "تم حفظ {0} سجل في {1}.",
//         formula_detected_label: "{0} (تم اكتشافه: {1})",
//         mapped_formulas_for: "تم تعيين المعادلات لـ: {0}",
//         select_range_with_headers_and_data: "حدد نطاقاً يحتوي على عناوين وبيانات.",
//         data_table_not_found_error: "لم يتم العثور على جدول البيانات '{0}'.",
//         replaced_table_with_records: "تم استبدال '{0}' بـ {1} سجل جديد. (المعرف: '{2}')",
//         replaced_rev_of_table_with_records: "تم استبدال النسخة {0} من '{1}' بـ {2} سجل.",
//         error_replacing_data: "خطأ في استبدال البيانات: ",
//         captured_rev_for_table: "تم التقاط النسخة {0} لـ '{1}' مع {2} سجل.",
//         error_capturing_new_revision: "خطأ في التقاط نسخة جديدة: ",
//         no_data_tables_stored_yet: "لم يتم حفظ أي جداول بيانات بعد.",
//         no_workspaces_or_tables_stored_yet: "لم يتم حفظ مساحات عمل أو جداول بيانات بعد.",
//         rev_latest_history: "النسخة {0}{1}",
//         rev_latest_suffix: " (الأحدث)",
//         link_column: "عمود الربط",
//         open_split_editor_btn: "فتح المحرر المنقسم",
//         delete_link_btn: "حذف الرابط",
//         rev_count_rows: "النسخة {0} • {1} صفوف",
//         revision_latest_history: "النسخة {0}{1}",
//         restore_snapshot: "استعادة",
//         no_tables_in_this_workspace: "لا توجد جداول في مساحة العمل هذه.",
//         deleted_entire_data_table: "تم حذف جدول البيانات بالكامل: {0}",
//         deleted_current_version_rolled_back: "تم حذف النسخة الحالية. تم التراجع إلى النسخة {0}.",
//         deleted_only_version_table_removed: "تم حذف النسخة الوحيدة. تمت إزالة الجدول '{0}'.",
//         deleted_historical_rev: "تم حذف النسخة التاريخية {0} من '{1}'.",
//         error_deleting_version: "خطأ في حذف النسخة: ",
//         locked_rev_current_is: "تم قفل النسخة {0}. النسخة الحالية الآن هي {1}",
//         restored_rev_as_new_active: "تمت استعادة النسخة {0} كنسخة نشطة جديدة {1}.",
//         storage_limit_reached: "تم الوصول إلى حد التخزين! يرجى إخلاء بعض المساحة.",
//         error_creating_snapshot: "خطأ في إنشاء لقطة: ",
//         cannot_resort_historical_revision: "لا يمكن إعادة ترتيب أعمدة لنسخة تاريخية.",
//         column_names_must_be_unique: "يجب أن تكون أسماء الأعمدة فريدة.",
//         columns_updated_successfully: "تم تحديث الأعمدة بنجاح. النسخة الحالية هي {0}.",
//         enter_record_id_to_edit: "أدخل معرف السجل المراد تعديله في '{0}':",
//         record_not_found_error: "السجل '{0}' غير موجود.",
//         calculated_label: " (محسوب)",
//         record_updated_refreshing_excel: "تم تحديث السجل. جاري تحديث Excel...",
//         storage_limit_reached_cannot_save: "تم الوصول إلى حد التخزين! لا يمكن حفظ التغييرات.",
//         error_saving_changes: "خطأ في حفظ التغييرات: ",
//         no_data_found_error: "لم يتم العثور على بيانات.",
//         no_data_found_for_revision_error: "لم يتم العثور على بيانات لهذه النسخة.",
//         exported_table_to_csv: "تم تصدير {0} (النسخة {1}) إلى CSV.",
//         error_exporting_csv: "خطأ في تصدير CSV: ",
//         refreshing_dashboard_formulas: "جاري تحديث لوحة المعلومات والمعادلات... يرجى الانتظار.",
//         refreshed_dashboard_formulas: "تم تحديث لوحة المعلومات و {0} من معادلات DC.",
//         converting_please_wait: "جاري التحويل... يرجى الانتظار.",
//         converted_formulas_to_values: "تم تحويل {0} معادلة إلى قيم.",
//         backup_downloaded_successfully: "تم تنزيل النسخة الاحتياطية بنجاح.",
//         backup_error: "خطأ في النسخ الاحتياطي: ",
//         data_restored_successfully: "تمت استعادة البيانات بنجاح.",
//         restore_error: "خطأ في الاستعادة: ",
//         no_headers_found_error: "لم يتم العثور على عناوين.",
//         inserted_headers_dropdown: "تم إدراج قائمة عناوين لـ {0}",
//         error_inserting_dropdown: "خطأ في إدراج القائمة المنسدلة: ",
//         no_data_found_for_table_error: "لم يتم العثور على بيانات لجدول البيانات هذا.",
//         select_columns_to_insert: "حدد الأعمدة لإدراجها:",
//         excel_formula_js_syntax_error: "تحتوي المعادلات على صيغة غير صالحة وتم إدراجها كنص.",
//         inserted_table_for: "تم إدراج جدول لـ {0}",
//         error_inserting_table: "خطأ في إدراج الجدول: ",
//         unknown_table_name: "غير معروف",
//         loading_split_editor: "جاري تحميل محرر الانقسام...",
//         fix_errors_before_new_editor: "يرجى إصلاح الأخطاء في المحرر الحالي قبل فتح محرر جديد.",
//         new_record_id_placeholder: "جديد",
//         opened_split_editor_arrange_windows: "تم فتح محرر الانقسام لـ {0} و {1}. قم بترتيب النوافذ جنباً إلى جنب.",
//         error_opening_split_editor: "خطأ في فتح محرر الانقسام: ",
//         opened_sub_table_in_editor: "تم فتح الجدول الفرعي ({0}: {1})",
//         opened_table_in_grid_editor: "تم فتح '{0}' في محرر الشبكة.",
//         error_opening_grid_editor: "خطأ في فتح محرر الشبكة: ",
//         row_has_empty_id_in_table_error: "الصف {0} يحتوي على معرف فارغ في {1}.",
//         duplicate_id_found_in_table_error: "تم العثور على معرف مكرر: '{0}' في {1}.",
//         saved_records: "تم حفظ {0} سجل.",
//         no_changes_detected: "لم يتم اكتشاف تغييرات.",
//         error_saving_grid_editor: "خطأ في حفظ محرر الشبكة: ",
//         auto_syncing_switching_to_record: "مزامنة تلقائية... جاري التبديل إلى السجل {0}",
//         switched_sub_table_to_record: "تم التبديل في الجدول الفرعي إلى السجل {0}",
//         error_switching_record: "خطأ في تبديل السجل: ",
//         error_canceling_grid_editor: "خطأ في إلغاء محرر الشبكة: ",
//         relation_link_suffix: "{0} (رابط: {1})",
//         relations_updated_for: "تم تحديث العلاقات لـ '{0}'",
//         define_the_rollup: "حدد التجميع:",
//         added_rollup_field: "تمت إضافة حقل التجميع '{0}'",
//         error_managing_relations: "خطأ في إدارة العلاقات: ",
//         cannot_append_to_historical_revision: "لا يمكن إلحاق بيانات بنسخة تاريخية.",
//         review_data_to_append: "راجع البيانات المراد إلحاقها بـ '{0}' (تم اكتشاف {1} صفوف).",
//         duplicate_id_already_exists_error: "تم العثور على معرف مكرر: '{0}' موجود بالفعل في '{1}'.",
//         appended_records_current_is: "تم إلحاق {0} سجل. النسخة الحالية هي {1}.",
//         error_appending_data: "خطأ في إلحاق البيانات: ",
//         no_relations_defined_for_table: "لا توجد علاقات محددة لهذا الجدول.",
//         relation_fk_suffix: "{0} (مفتاح أجنبي: {1})",
//         select_target_sub_table_to_clone: "حدد الجدول الفرعي المستهدف للاستنساخ:",
//         clone_from_table: "استنساخ من '{0}'",
//         select_source_target_ids: "حدد المعرف المصدر والمعرف(ات) الهدف:",
//         no_sub_records_found: "لم يتم العثور على سجلات فرعية للمعرف المصدر '{0}'.",
//         cloned_id_prefix: "نسخة_",
//         successfully_cloned_sub_records: "تم بنجاح استنساخ {0} سجل فرعي إلى {1} هدف.",
//         error_cloning_sub_records: "خطأ في استنساخ السجلات الفرعية: ",
//         select_or_type_new_workspace: "حدد أو اكتب مساحة عمل جديدة لـ '{0}':",
//         moved_table_to_workspace: "تم نقل '{0}' إلى '{1}'.",
//         error_moving_table: "خطأ في نقل الجدول: ",
//         variable_name_required: "اسم المتغير مطلوب.",
//         formula_required: "المعادلة مطلوبة.",
//         loop_detected: "تم اكتشاف حلقة",
//         variable_cannot_reference_itself: "لا يمكن للمتغير الإشارة إلى نفسه",
//         variable_saved: "تم حفظ المتغير '{0}'.",
//         invalid_formula_for_variable: "معادلة غير صالحة لـ '{0}': {1}",
//         error_adding_workspace: "خطأ في إضافة مساحة عمل: ",
//         error_text: "خطأ",
//         loop_error: "خطأ حلقة",
//         insert_to_sheet: "إدراج في الورقة",
//         inserted_variable_to_sheet: "تم إدراج المتغير '{0}' في الورقة.",
//         error_inserting_variable: "خطأ في إدراج المتغير: ",
//         delete_variable: "حذف المتغير",
//         record_id_already_exists: "معرف السجل '{0}' موجود بالفعل في '{1}'.",
//         source_record_not_found: "لم يتم العثور على السجل المصدر '{0}'.",
//         successfully_duplicated_record: "تم بنجاح استنساخ السجل '{0}' إلى '{1}' مع {2} سجل فرعي.",
//         error_duplicating_record: "خطأ في استنساخ السجل: ",
//         param_record_id: "معرف السجل",
//         param_field_name: "اسم الحقل",
//         param_data_table_name: "اسم جدول البيانات",
//         param_revision: "النسخة",
//         param_search_field: "حقل البحث",
//         param_search_value: "قيمة البحث",
//         param_return_field: "حقل الإرجاع",
//         param_exact_match: "تطابق تام (TRUE/FALSE)",
//         param_sum_field: "حقل الجمع",
//         param_criteria_field: "حقل المعيار",
//         param_criteria_value: "قيمة المعيار",
//         param_base_table_name: "اسم الجدول الأساسي",
//         param_link_column: "عمود الربط",
//         param_target_table_name: "اسم الجدول المستهدف",
//         param_target_return_field: "حقل الإرجاع المستهدف",
//         param_sort_field: "حقل الفرز",
//         param_ascending: "تصاعدي (TRUE/FALSE)",
//         param_variable_name: "اسم المتغير",
//         save_btn: "حفظ",
//         loading_text: "جاري التحميل...",
//         col_name_placeholder: "اسم العمود",
//         new_col_name_placeholder: "اسم العمود الجديد",
//         new_column_default: "عمود جديد",
//         formula_def_label: "تعريف المعادلة",
//         edit_formula_tooltip: "تعديل المعادلة: {0}",
//         add_formula_tooltip: "إضافة معادلة",
//         cannot_delete_id_col_tooltip: "لا يمكن حذف عمود المعرف",
//         primary_id_col_tooltip: "عمود المعرف الأساسي",
//         id_col_must_be_present: "يجب أن يكون عمود المعرف الأساسي موجوداً في الجدول.",
//         row_empty_id_strict: "الصف {0} يحتوي على معرف فارغ. يجب ألا تكون المعرفات فارغة.",
//         duplicate_id_strict: "تم العثور على معرف مكرر: '{0}'. يجب أن تكون جميع المعرفات فريدة.",
//         this_item: "هذا العنصر",
//         delete_tooltip: "حذف",
//         current_table_opt: "الجدول الحالي",
//         select_table_opt: "-- حدد الجدول --",
//         formula_builder_placeholder: "اكتب أو انقر أعلاه لبناء المعادلة...",
//         table_name_label: "اسم الجدول",
//         table_placeholder: "مثل: العملاء",
//         workspace_placeholder: "حدد أو اكتب جديداً",
//         parent_link_label: "الجدول الأصل (رابط اختياري)",
//         ready_to_save: "جاهز للحفظ",
//         save_table_btn: "حفظ الجدول",
//         add_new_column_btn: "إضافة عمود جديد",
//         save_current_btn: "حفظ في النسخة الحالية",
//         save_new_btn: "حفظ كنسخة جديدة",
//         grid_editor_mode: "وضع محرر الشبكة",
//         currently_editing: "أنت تقوم حالياً بتعديل",
//         grid_editor_tip: "قم بإجراء تغييراتك مباشرة في ورقة العمل.<br/>تلميح: انتقل إلى عرض > نافذة جديدة > ترتيب الكل لتعديل هذا الجدول جنباً إلى جنب مع بياناتك الأساسية.<br/>انقر فوق حفظ عند الانتهاء لمزامنة البيانات.",
//         save_sync_btn: "حفظ ومزامنة التغييرات",
//         cancel_close_btn: "إلغاء وإغلاق",
//         input_title: "إدخال",
//         primary_id_column: "عمود المعرف الأساسي (يجب أن يكون فريداً)",
//         columns_drag_label: "الأعمدة (اسحب لإعادة الترتيب، قم بإلغاء التحديد للحذف، انقر لإعادة التسمية)",
//         validation_passed: "تم اجتياز التحقق",
//         records_count: "السجلات:",
//         id_column_name: "عمود المعرف:",
//         final_columns: "الأعمدة النهائية:",
//         next_validate_btn: "التالي (تحقق)",
//         back_mapping_btn: "العودة إلى التعيين",
//         confirm_save_btn: "تأكيد وحفظ",
//         link_col_fk: "عمود الربط (مفتاح أجنبي)",
//         select_link_col_1: "حدد العمود في هذا الجدول الجديد الذي يرتبط بالسجلات في",
//         yes_delete_btn: "نعم، احذف",
//         add_new_btn: "إضافة جديد"
//     }
// };

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

let isSwitchingRecord = false;
let isCapturingForFormula = false;
let formulaCaptureTarget: HTMLInputElement | null = null;
let originalFormulaCellAddress: string | null = null;

export function showCaptureMessage(title: string, message: string, onCaptureActive?: () => void) {
    const modal = document.getElementById("custom-form-modal");
    const titleEl = document.getElementById("form-modal-title");
    const messageEl = document.getElementById("form-modal-message");
    const btnYes = document.getElementById("form-modal-ok");
    const btnNo = document.getElementById("form-modal-cancel");

    if (modal && titleEl && messageEl && btnYes && btnNo) {
        titleEl.innerText = title;
        titleEl.className = "mt-0 color-primary";
        messageEl.innerText = message;

        if (onCaptureActive) {
            btnYes.style.display = "flex";
            btnYes.innerText = t("current_cell_btn");
            btnYes.setAttribute("data-i18n", "current_cell_btn");
            btnYes.style.backgroundColor = "green";
            btnYes.style.borderColor = "green";
            btnYes.style.color = "white";
            btnYes.onclick = onCaptureActive;
        } else {
            btnYes.style.display = "none";
        }
        btnNo.style.display = "flex";
        btnNo.innerText = t("cancel_btn");
        btnNo.setAttribute("data-i18n", "cancel_btn");
        if (btnYes.parentElement) btnYes.parentElement.style.display = "flex";
        
        btnNo.onclick = () => {
            isCapturingForFormula = false;
            formulaCaptureTarget = null;
            originalFormulaCellAddress = null;
            hideCaptureMessage();
            showStatus(t("status_ready"), "info");
        };
        
        modal.style.display = "flex";
    }
}

export function hideCaptureMessage() {
    const modal = document.getElementById("custom-form-modal");
    const titleEl = document.getElementById("form-modal-title");
    const btnYes = document.getElementById("form-modal-ok");
    const btnNo = document.getElementById("form-modal-cancel");
    
    if (modal && btnYes && btnNo) {
        modal.style.display = "none";
        if (titleEl) titleEl.className = "mt-0 color-danger";
        btnYes.style.display = "";
        btnYes.style.backgroundColor = "";
        btnYes.style.borderColor = "";
        btnYes.style.color = "";
        btnYes.setAttribute("data-i18n", "yes_delete_btn");
        btnNo.setAttribute("data-i18n", "cancel_btn");
        btnNo.style.display = "";
        if (btnYes.parentElement) btnYes.parentElement.style.display = "flex";
        btnNo.onclick = null;
    }
}

async function executeCellCapture() {
    if (!isCapturingForFormula || !formulaCaptureTarget) return;

    const targetInput = formulaCaptureTarget;
    const originalAddress = originalFormulaCellAddress;

    // Reset state immediately
    isCapturingForFormula = false;
    formulaCaptureTarget = null;
    originalFormulaCellAddress = null;
    
    hideCaptureMessage(); // Close the dialog immediately

    try {
        await Excel.run(async (context) => {
            const selection = context.workbook.getSelectedRange();
            selection.load("address");
            await context.sync();
            
            let capturedAddress = selection.address;
            if (capturedAddress.includes("!")) capturedAddress = capturedAddress.split("!")[1];
            if (capturedAddress.includes(":")) capturedAddress = capturedAddress.split(":")[0];
            
            targetInput.value = capturedAddress.replace(/\$/g, "");
            targetInput.dispatchEvent(new Event('input', { bubbles: true }));
            targetInput.dispatchEvent(new Event('change', { bubbles: true }));
            targetInput.style.backgroundColor = "var(--hover-bg)";
            setTimeout(() => { if (targetInput) targetInput.style.backgroundColor = ""; }, 300);

            if (originalAddress) {
                let sheetName = "";
                let rangeAddress = originalAddress;
                if (originalAddress.includes("!")) {
                    const lastBang = originalAddress.lastIndexOf("!");
                    sheetName = originalAddress.substring(0, lastBang);
                    rangeAddress = originalAddress.substring(lastBang + 1);
                    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
                        sheetName = sheetName.substring(1, sheetName.length - 1);
                    }
                }
                if (sheetName) {
                    const targetSheet = context.workbook.worksheets.getItem(sheetName);
                    targetSheet.activate();
                    targetSheet.getRange(rangeAddress).select();
                } else {
                    context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress).select();
                }
                await context.sync();
            }
            
            showStatus(t("status_ready"), "info");
        });
    } catch (e: any) {
        showStatus(t("error_during_capture") + (e.message ? " " + e.message : ""), "error");
    }
}

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
    }
}

export function t(key: string, ...args: (string | number)[]): string {
    let str = TRANSLATIONS[currentLang]?.[key] || TRANSLATIONS["en"][key] || key;
    args.forEach((arg, i) => {
        str = str.replace(new RegExp(`\\{${i}\\}`, 'g'), String(arg));
    });
    return str;
}

export function applyTranslations() {
    try {
        document.body.dir = currentLang === "ar" ? "rtl" : "ltr";
        
        // Safely replace text without destroying child elements (like spans or icons)
        document.querySelectorAll("[data-i18n]").forEach(el => {
            const key = el.getAttribute("data-i18n");
            if (key && TRANSLATIONS[currentLang] && TRANSLATIONS[currentLang][key]) {
                const translatedText = t(key);
                (el as HTMLElement).innerHTML = translatedText;
            }
        });

        document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
            const key = el.getAttribute("data-i18n-placeholder");
            if (key) {
                (el as HTMLInputElement).placeholder = t(key);
            }
        });
        
        document.querySelectorAll("[data-i18n-title]").forEach(el => {
            const key = el.getAttribute("data-i18n-title");
            if (key) {
                (el as HTMLElement).title = t(key);
            }
        });

        const statusText = document.getElementById("status-text");
        if (statusText) {
            if (statusText.innerText.trim() === "Ready" || statusText.innerText.trim() === "جاهز" || statusText.innerText.trim() === "") {
                statusText.innerText = t("status_ready");
            }
        }
    } catch (e) {
        console.error("Translation apply error:", e);
    }
}

export function showStatus(message: string, type: "success" | "error" | "info" = "info") {
    const status = document.getElementById("status-text");
    if (!status) return;
    status.innerText = message;
    if (type === "success") status.style.color = "green";
    else if (type === "error") status.style.color = "red";
    else status.style.color = "blue";
}

// Catch unhandled synchronous exceptions
window.onerror = (message, source, lineno, colno, error) => {
    console.error("[Global Error]", message, error);
    showStatus(t("error_general") + (error?.message || message), "error");
};

// Catch unhandled asynchronous promise rejections
window.addEventListener("unhandledrejection", (event) => {
    console.error("[Unhandled Rejection]", event.reason);
    showStatus(t("error_general") + (event.reason?.message || event.reason || "Unhandled Promise Rejection"), "error");
});

let globalEventsBound = false; // Prevents double-binding global document handlers

Office.onReady(async (info) => {
  // Allow UI bindings to execute in standard web browsers for UI testing, as well as inside Excel
  if (info.host === Office.HostType.Excel || !info.host) {
    // Safe event bindings: Using .onclick prevents double-firing if Office.onReady runs multiple times
    const backupBtn = document.getElementById("backup-button");
    if (backupBtn) backupBtn.onclick = backupData;

    const restoreBtn = document.getElementById("restore-button");
    if (restoreBtn) restoreBtn.onclick = triggerRestore;

    const restoreFileInput = document.getElementById("restore-file-input");
    if (restoreFileInput) restoreFileInput.onchange = restoreData;

    const refreshBtn = document.getElementById("refresh-formulas-button");
    if (refreshBtn) refreshBtn.onclick = () => refreshFormulas(false);

    const convertBtn = document.getElementById("convert-values-button");
    if (convertBtn) convertBtn.onclick = convertToValues;

    const settingsBtn = document.getElementById("settings-button");
    if (settingsBtn) settingsBtn.onclick = toggleSettings;
    
    const manageWsBtn = document.getElementById("manage-workspaces-btn");
    if (manageWsBtn) manageWsBtn.onclick = manageWorkspaces;

    const addWsBtnGlobal = document.getElementById("add-workspace-btn-global");
    if (addWsBtnGlobal) addWsBtnGlobal.onclick = addWorkspace;

    const closeSettingsBtn = document.getElementById("close-settings-btn");
    if (closeSettingsBtn) closeSettingsBtn.onclick = toggleSettings;

    const helpBtn = document.getElementById("help-button") || document.getElementById("help-btn") || document.querySelector('[data-i18n="help_btn"]');
    if (helpBtn) {
        (helpBtn as HTMLElement).onclick = () => window.open("https://github.com/m-Azem/data-controller-addin", "_blank");
    }
    
    const editorSaveBtn = document.getElementById("editor-save-btn");
    if (editorSaveBtn) editorSaveBtn.onclick = () => saveGridEditor(true);
    const editorCancelBtn = document.getElementById("editor-cancel-btn");
    if (editorCancelBtn) editorCancelBtn.onclick = cancelGridEditor;

    // Setup Formula Builder Accordion
    const formulaAccordion = document.getElementById("formula-builder-accordion");
    const formulaContent = document.getElementById("formula-builder-content");
    if (formulaAccordion && formulaContent) {
      formulaAccordion.classList.add("accordion-top-level");
      formulaAccordion.onclick = () => {
        toggleAccordion(formulaAccordion, formulaContent, "accordion-top-level");
      };
    }

    // Setup Global Variables Accordion
    const varAccordion = document.getElementById("variables-accordion");
    const varContent = document.getElementById("variables-content");
    if (varAccordion && varContent) {
      varAccordion.classList.add("accordion-top-level");
      varAccordion.onclick = () => {
        toggleAccordion(varAccordion, varContent, "accordion-top-level");
      };
    }
    const addVarBtn = document.getElementById("add-variable-btn");
    if (addVarBtn) addVarBtn.onclick = manageVariable;

    // Setup Global Relations Accordion
    const relAccordion = document.getElementById("relations-accordion");
    const relContent = document.getElementById("relations-content");
    if (relAccordion && relContent) {
      relAccordion.classList.add("accordion-top-level");
      relAccordion.onclick = () => {
        toggleAccordion(relAccordion, relContent, "accordion-top-level");
      };
    }
    const manageRelBtn = document.getElementById("manage-global-relations-btn");
    if (manageRelBtn) manageRelBtn.onclick = async () => {
        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        const tables = Object.keys(store);
        const res = await customFormPrompt(t("manage_rel_title"), t("manage_rel_select_table"), [{ id: "table", label: t("main_data_table"), type: "select", options: tables }], t("next_btn"));
        if (res && res.table) manageRelations(res.table);
    };

const fSelect = document.getElementById("formula-select");
    if (fSelect) {
        fSelect.onchange = renderFormulaBuilder;
    }

    const insertFormulaBtn = document.getElementById("insert-built-formula-button");
    if (insertFormulaBtn) {
        insertFormulaBtn.onclick = insertBuiltFormula;
    }

    const clearFormulaBtn = document.getElementById("clear-formula-button");
    if (clearFormulaBtn) {
        clearFormulaBtn.onclick = clearFormulaForm;
    }

    // Setup Theme Toggle
    const themeToggle = document.getElementById("theme-toggle");
    if (themeToggle) themeToggle.onchange = toggleTheme;

    // Display App Version
    const versionDisplay = document.getElementById("app-version-display");
    if (versionDisplay) versionDisplay.innerText = APP_VERSION;

    const langSelect = document.getElementById("language-select") as HTMLSelectElement;
    if (langSelect) {
        langSelect.onchange = async (e) => {
            currentLang = (e.target as HTMLSelectElement).value;
            await idbSet(IDB_KEYS.LANGUAGE, currentLang);
            applyTranslations();
            renderDashboard(); // Re-render dynamic text on change
        };
    }

   // Only bind global document listeners once
    if (!globalEventsBound) {
        // Close Settings with Escape
        document.addEventListener("keydown", (e) => {
            if (e.key === "Escape") {
                const settingsView = document.getElementById('settings-view');
                if (settingsView && settingsView.style.display === 'block') {
                    const isModalOpen = Array.from(document.querySelectorAll('.modal-overlay')).some(m => (m as HTMLElement).style.display === 'flex');
                    if (!isModalOpen) {
                        toggleSettings();
                    }
                }
            }
        });

        // Handle automatic cell reference insertion for Formula Builder
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async (eventArgs) => {
            
            if (isCapturingForFormula && formulaCaptureTarget) {                
                executeCellCapture();
                return; // Important: prevent falling through to Live-Sync while capturing
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

                    if (newId && newId !== currentId && newId.trim() !== "") {
                        if (isSwitchingRecord) return;
                        isSwitchingRecord = true;
                        const editorCols = (editorColsProp.value || "").split(",").filter(c => c);
                        setTimeout(async () => {
                            try {
                                await switchGridEditorRecord(newId, relation.foreignKey, editorTargetTable, editorCols, subEditorSheet!.name, matchedMainTable);
                            } catch (err) {
                                console.error("Switch error:", err);
                            } finally {
                                isSwitchingRecord = false;
                            }
                        }, 10);
                    }
                });
            } catch (e) { console.error("Setup/Detection error:", e); }
        });

        globalEventsBound = true;
    }
    

    await migrateFromLocalStorage(); // Migrate old data if present
    loadSettings();
    renderDashboard();

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
            if (!purposeProp.isNullObject && purposeProp.value === "SubDataEditor") showEditorView(!tableProp.isNullObject ? tableProp.value : t("unknown_table_name"), isLive);
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

export function customConfirm(title: string, message: string, confirmText?: string): Promise<boolean> {
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
      btnYes.innerText = confirmText || t("yes_btn");
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

        if (addBtnText) {
            btnAdd.style.display = "";
            btnAdd.innerHTML = `<i class="ms-Icon ms-Icon--Add" style="margin-right:8px;"></i>${addBtnText}`;
        } else {
            btnAdd.style.display = "none";
        }

        const createLi = (original: string, isNew: boolean) => {
            const li = document.createElement("li");
            li.className = "sortable-item";
            li.draggable = true;
            li.dataset.original = original;
            li.dataset.isNew = String(isNew);
            li.dataset.isDeleted = "false";
            
            li.innerHTML = `
                <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: var(--input-bg);">
                    <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-inline-end: 8px; color: #888; cursor: grab;"></i>
                    <input type="text" class="ms-TextField-field list-name-input" value="${original}" placeholder="${t("name_placeholder")}" style="flex: 1; padding: 4px 8px; font-weight: 600;" ${(!allowRename && !isNew) ? 'disabled' : ''} />
                    <button type="button" class="icon-btn list-delete-btn" style="margin-left: 8px; color: #d13438;" title="${t("delete_tooltip")}">
                        <i class="ms-Icon ms-Icon--Delete"></i>
                    </button>
                </div>
            `;
            const inputEl = li.querySelector("input[type='text']") as HTMLInputElement;
            inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
            
            const deleteBtn = li.querySelector(".list-delete-btn") as HTMLButtonElement;
            deleteBtn.onclick = async () => {
                const itemName = inputEl.value || t("this_item");
            const confirmed = await customConfirm(t("mark_for_deletion"), t("mark_deletion_msg", itemName), t("del_version_confirm"));
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
                      <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: var(--input-bg);">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-inline-end: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-right: 8px;" ${isId ? `disabled title="${t("cannot_delete_id_col_tooltip")}"` : ''} />
                  <input type="text" class="ms-TextField-field col-name-input" value="${item}" placeholder="${t("col_name_placeholder")}" style="flex: 1; padding: 4px 8px; font-weight: 600;" />
                  ${isId ? `<span style="font-size:10px; color:#0078d4; margin-left:8px; font-weight:bold;" title="${t("primary_id_col_tooltip")}">(ID)</span>` : ''}
                  <input type="hidden" class="col-formula-hidden" value="${formula.replace(/"/g, '&quot;')}" />
                  <button type="button" class="icon-btn edit-formula-btn" style="margin-left: 8px; color: ${formula ? '#0078d4' : '#888'};" title="${formula ? t("edit_formula_tooltip", formula.replace(/"/g, '&quot;')) : t("add_formula_tooltip")}">
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
              const colName = nameInput.value.trim() || item || t("new_column_default");
              
              modal.style.display = "none";
              const res = await customFormPrompt(t("formula_for", colName), t("define_calc_formula"), [
                  { id: "vFormula", label: t("formula_def_label"), type: "formula", varsList: varNames, tablesWithFields: tablesWithFields, fieldsList: currentFields, value: formulaHidden.value }
              ]);
              modal.style.display = "flex";

              if (res !== null) {
                  const newFormula = res.vFormula || "";
                  formulaHidden.value = newFormula;
                  if (newFormula) {
                      formulaBtn.style.color = "#0078d4";
                      formulaBtn.title = t("edit_formula_tooltip", newFormula.replace(/"/g, '&quot;'));
                  } else {
                      formulaBtn.style.color = "#888";
                      formulaBtn.title = t("add_formula_tooltip");
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
              <div style="display:flex; align-items:center; width:100%; margin-bottom: 8px; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: var(--input-bg);">
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-inline-end: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-inline-end: 8px;" />
                  <input type="text" class="ms-TextField-field col-name-input" value="" placeholder="${t("new_col_name_placeholder")}" style="flex: 1; padding: 4px 8px; font-weight: 600;" />
                  <input type="hidden" class="col-formula-hidden" value="" />
                  <button type="button" class="icon-btn edit-formula-btn" style="margin-left: 8px; color: #888;" title="${t("add_formula_tooltip")}">
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
              const colName = nameInput.value.trim() || t("new_column_default");
              
              modal.style.display = "none";
              const res = await customFormPrompt(t("formula_for", colName), t("define_calc_formula"), [
                  { id: "vFormula", label: t("formula_def_label"), type: "formula", varsList: varNames, tablesWithFields: tablesWithFields, fieldsList: currentFields, value: formulaHidden.value }
              ]);
              modal.style.display = "flex";

              if (res !== null) {
                  const newFormula = res.vFormula || "";
                  formulaHidden.value = newFormula;
                  if (newFormula) {
                      formulaBtn.style.color = "#0078d4";
                      formulaBtn.title = t("edit_formula_tooltip", newFormula.replace(/"/g, '&quot;'));
                  } else {
                      formulaBtn.style.color = "#888";
                      formulaBtn.title = t("add_formula_tooltip");
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
      btnDeleteSelected.style.display = "inline-flex";
      btnNext.style.display = "inline-flex";
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
                      badge.style.cssText = "font-size:10px; color:#0078d4; margin-inline-start:8px; font-weight:bold;";
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
                  <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-inline-end: 8px; color: #888; cursor: grab;"></i>
                  <input type="checkbox" class="col-delete-checkbox" style="margin-inline-end: 8px;" />
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

      const handleKeyDown = (e: KeyboardEvent) => {
          if (e.key === "Escape") {
              e.preventDefault();
              btnCancel.click();
          } else if (e.key === "Enter") {
              const activeEl = document.activeElement as HTMLElement;
              // Allow Enter to work normally if the user has explicitly tabbed to another button
              if (activeEl && activeEl.tagName === "BUTTON" && activeEl !== btnNext && activeEl !== btnOk && activeEl !== btnCancel && activeEl !== btnBack && activeEl !== btnDeleteSelected) return;
              
              e.preventDefault();
              if (btnNext.style.display !== "none") btnNext.click();
              else if (btnOk.style.display !== "none") btnOk.click();
          }
      };
      document.addEventListener("keydown", handleKeyDown);

      const cleanup = () => {
          modal.style.display = "none";
          btnDeleteSelected.onclick = null;
          btnNext.onclick = null;
          btnBack.onclick = null;
          btnOk.onclick = null;
          btnCancel.onclick = null;
          idSelect.onchange = null;
          listEl.removeEventListener("dragover", handleDragOver);
          document.removeEventListener("keydown", handleKeyDown);
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
          btnBack.style.display = "inline-flex";
          btnOk.style.display = "inline-flex";
          
          if (recordsCountEl) recordsCountEl.innerText = String(finalRecords.length); 
          if (idColumnEl) idColumnEl.innerText = validationData.idField;
          if (finalColumnsEl) finalColumnsEl.innerText = newFieldsList.join(", ");
      };

      btnBack.onclick = () => {
          step1Div.style.display = "block";
          step2Div.style.display = "none";
          btnDeleteSelected.style.display = "inline-flex";
          btnNext.style.display = "inline-flex";
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
        if (visited.has(vName)) throw new Error(t("circular_reference_detected"));
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

export function customFormPrompt(title: string, message: string, fields: FormField[], confirmText?: string): Promise<Record<string, string> | null> {
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
        if (confirmText) {
            btnOk.innerText = confirmText;
        } else {
            btnOk.innerText = t("save_btn");
        }
        inputsContainer.innerHTML = "";

        const inputElements: { id: string, el: HTMLInputElement | HTMLSelectElement | HTMLDivElement, type: string }[] = [];

        fields.forEach(f => {
            const fieldDiv = document.createElement("div");
            fieldDiv.className = f.type === 'select' ? "ms-Dropdown" : "ms-TextField";
            fieldDiv.style.marginBottom = "12px";

            const label = document.createElement("label");
            label.className = "ms-Label";
            label.innerText = f.label;
            fieldDiv.appendChild(label);

            if (f.type === 'select') {
                const select = document.createElement("select");
                select.className = "ms-Dropdown-title w-100";
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
                wrap.style.backgroundColor = "var(--input-bg)";
                
                let fieldsHtml = "";
                if (f.fieldsList && f.fieldsList.length > 0) {
                    fieldsHtml = `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center;">
                        <span style="font-size:9px; font-weight:bold; color:#888;">${t("fields_label")}</span>
                        ${f.fieldsList.map(i => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; white-space:nowrap; color:var(--text-color);" data-op="[${i}]">${i}</button>`).join('')}
                    </div>`;
                }
                
                const tablesWithFields = f.tablesWithFields || {};

                const toolbarHtml = `
                <div style="background: var(--hover-bg); padding: 6px; border-bottom: 1px solid var(--border-color); display: flex; flex-direction: column; gap: 6px;">
                    <div style="display:flex; gap:4px;">
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" + ">+</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" - ">-</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" * ">*</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" / ">/</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" ( ">(</button>
                        <button type="button" class="col-insert-op" style="min-width:24px; padding:2px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; color:var(--text-color);" data-op=" ) ">)</button>
                    </div>
                    <div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;">
                        <span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">${t("funcs_label")}</span>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op="DC.SUM('Table', 'Col')">SUM()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op="DC.COUNT('Table')">COUNT()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op="DC.VAR('VarName')">VAR()</button>
                        <button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#dff6dd; border:1px solid #b7e0b5; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op=" ( condition ? true_val : false_val ) ">IF()</button>
                    </div>
                    ${(f.varsList && f.varsList.length > 0) ? `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;"><span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">${t("vars_label")}</span>${f.varsList.map(v => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#c7e0f4; border:1px solid #99c9ef; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op="DC.VAR('${v}')">${v}</button>`).join('')}</div>` : ''}
                    ${Object.keys(tablesWithFields).length > 0 ? `<div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;"><span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">${t("tables_label")}</span>${Object.keys(tablesWithFields).map(t => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:#fce1cb; border:1px solid #f9c79f; border-radius:2px; cursor:pointer; white-space:nowrap; color:#222;" data-op="'${t}'">${t}</button>`).join('')}</div>` : ''}
                    <div style="display:flex; gap:4px; overflow-x:auto; padding-bottom:2px; align-items:center; width:100%;">
                        <span style="font-size:9px; font-weight:bold; color:#888; flex-shrink:0;">${t("fields_label")}</span>
                        <select class="field-table-selector" style="font-size:10px; padding:2px; max-width:100px; border:1px solid var(--border-color); border-radius:2px; flex-shrink:0; background:var(--input-bg); color:var(--text-color);">
                            ${(f.fieldsList && f.fieldsList.length > 0) ? `<option value="_CURRENT_">${t("current_table_opt")}</option>` : `<option value="">${t("select_table_opt")}</option>`}
                            ${Object.keys(tablesWithFields).map(t => `<option value="${t}">${t}</option>`).join('')}
                        </select>
                        <span class="dynamic-fields-container" style="display:flex; gap:4px; flex-shrink:0;">
                            ${(f.fieldsList || []).map(i => `<button type="button" class="col-insert-op" style="font-size:10px; padding:2px 6px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; white-space:nowrap; color:var(--text-color);" data-op="[${i}]">${i}</button>`).join('')}
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
                input.placeholder = t("formula_builder_placeholder");
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
                            btn.style.cssText = "font-size:10px; padding:2px 6px; background:var(--card-bg); border:1px solid var(--border-color); border-radius:2px; cursor:pointer; white-space:nowrap; color:var(--text-color);";
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
                            emptyOpt.text = t("manage_rel_none");
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

const formulaDefinitions: Record<string, { id: string, labelKey: string, required: boolean }[]> = {
    "GET": [
        { id: "param-id", labelKey: "param_record_id", required: false },
        { id: "param-fieldName", labelKey: "param_field_name", required: false },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false }
    ],
    "SEARCH": [
        { id: "param-searchField", labelKey: "param_search_field", required: true },
        { id: "param-searchValue", labelKey: "param_search_value", required: true },
        { id: "param-returnField", labelKey: "param_return_field", required: true },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false }
    ],
    "FILTER": [
        { id: "param-searchField", labelKey: "param_search_field", required: true },
        { id: "param-searchValue", labelKey: "param_search_value", required: true },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false },
        { id: "param-exactMatch", labelKey: "param_exact_match", required: false }
    ],
    "SUM": [
        { id: "param-sumField", labelKey: "param_sum_field", required: true },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false }
    ],
    "SUMIFS": [
        { id: "param-sumField", labelKey: "param_sum_field", required: true },
        { id: "param-criteriaField", labelKey: "param_criteria_field", required: true },
        { id: "param-criteriaValue", labelKey: "param_criteria_value", required: true },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false }
    ],
    "JOIN": [
        { id: "param-baseTableName", labelKey: "param_base_table_name", required: true },
        { id: "param-foreignKeyField", labelKey: "param_link_column", required: true },
        { id: "param-foreignTableName", labelKey: "param_target_table_name", required: true },
        { id: "param-foreignReturnField", labelKey: "param_target_return_field", required: true }
    ],
    "SORT": [
        { id: "param-sortField", labelKey: "param_sort_field", required: true },
        { id: "param-ascending", labelKey: "param_ascending", required: false },
        { id: "param-dataTableName", labelKey: "param_data_table_name", required: false },
        { id: "param-rev", labelKey: "param_revision", required: false }
    ],
    "VAR": [
        { id: "param-varName", labelKey: "param_variable_name", required: true }
    ]
};

export async function renderFormulaBuilder() {
    const select = document.getElementById("formula-select") as HTMLSelectElement;
    const container = document.getElementById("formula-inputs-container");
    if (!select || !container) return;

    const formula = select.value;
    const fields = formulaDefinitions[formula] || [];

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
                const fInput = document.getElementById("param-foreignTableName") as HTMLSelectElement;
                return fInput && fInput.value.trim() !== "" ? fInput.value.trim() : "";
            } else {
                const bInput = document.getElementById("param-baseTableName") as HTMLSelectElement;
                return bInput && bInput.value.trim() !== "" ? bInput.value.trim() : defaultTable;
            }
        } else {
            const tInput = document.getElementById("param-dataTableName") as HTMLSelectElement;
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
                const defaultOpt = document.createElement("option");
                defaultOpt.value = "";
                defaultOpt.text = t("select_option");
                targetEl.appendChild(defaultOpt);
            }

            if (ref.id.toLowerCase().includes("tablename")) {
                tables.forEach(t => {
                    const opt = document.createElement("option");
                    opt.value = t;
                    if (ref.el.tagName === "SELECT") opt.text = t;
                    targetEl.appendChild(opt);
                });
                if (ref.el.tagName === "SELECT" && tables.includes(currentVal)) {
                    ref.el.value = currentVal;
                }
                return;
            }

            if (ref.id === "param-exactMatch" || ref.id === "param-ascending") {
                [t("true_text"), t("false_text")].forEach(tStr => {
                    const opt = document.createElement("option");
                    opt.value = tStr;
                    opt.text = tStr;
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
                } else if (ref.id === "param-rev") {
                    const maxRev = dataSet.revision || 1;
                    for (let i = maxRev; i >= 1; i--) {
                        if (i === maxRev || (dataSet.history && dataSet.history[i])) {
                            options.push(String(i));
                        }
                    }
                }
            }
            // Remove duplicates and empty strings
            options = [...new Set(options)].filter(o => o !== undefined && o !== null && String(o).trim() !== "");

            options.forEach(opt => {
                const optionElement = document.createElement("option");
                optionElement.value = opt;
                if (ref.el.tagName === "SELECT") {
                    if (ref.id === "param-rev") {
                        const tName = getTargetTableForInput(ref.id);
                        const dSet = store[tName];
                        const isMax = dSet && String(dSet.revision || 1) === opt;
                        optionElement.text = t("rev_latest_history", opt, isMax ? t("rev_latest_suffix") : "");
                    } else {
                        optionElement.text = opt;
                    }
                }
                targetEl.appendChild(optionElement);
            });

            if (ref.el.tagName === "SELECT" && options.includes(currentVal)) {
                ref.el.value = currentVal;
            }
        });
    };

    container.innerHTML = "";

    fields.forEach(f => {
        const isSelect = f.id.toLowerCase().includes("field") || f.id === "param-exactMatch" || f.id === "param-ascending" || f.id.toLowerCase().includes("tablename") || f.id === "param-rev";
        const fieldDiv = document.createElement("div");
        fieldDiv.className = isSelect ? "ms-Dropdown" : "ms-TextField";
        fieldDiv.style.marginBottom = "8px";

        const label = document.createElement("label");
        label.className = "ms-Label";
        label.innerHTML = `${t(f.labelKey)} ${f.required ? '<span style="color:#d13438;">*</span>' : `<span style="color:#888;font-weight:normal;">${t("optional_label")}</span>`}`;

        let inputEl: HTMLInputElement | HTMLSelectElement;
        let dataList: HTMLDataListElement | undefined;

        if (isSelect || f.id.toLowerCase().includes("tablename")) { // Data Table Names should also be a select for consistency
            inputEl = document.createElement("select");
            inputEl.id = f.id;
            inputEl.className = "ms-Dropdown-title w-100";
            inputRefs.push({ id: f.id, el: inputEl });
            
            if (f.id.toLowerCase().includes("tablename")) {
                inputEl.addEventListener("change", updateDataLists);
            }
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
        if (isSelect || f.id.toLowerCase().includes("tablename")) {
            fieldDiv.appendChild(inputEl);
            if (dataList) fieldDiv.appendChild(dataList);
        } else {
            const inputWrapper = document.createElement("div");
            inputWrapper.style.display = "flex";
            inputWrapper.style.gap = "6px";
            inputEl.style.flex = "1";
            inputWrapper.appendChild(inputEl);
            if (dataList) inputWrapper.appendChild(dataList);
            
            const captureBtn = document.createElement("button");
            captureBtn.type = "button";
            captureBtn.className = "icon-btn";
            captureBtn.title = "Capture Selected Cell";
            captureBtn.innerHTML = `<i class="ms-Icon ms-Icon--TouchPointer"></i>`;
            captureBtn.style.border = "1px solid var(--border-color)";
            captureBtn.style.background = "var(--card-bg)";
            captureBtn.style.borderRadius = "2px";
            captureBtn.style.padding = "0 8px";
            captureBtn.style.cursor = "pointer";
            captureBtn.style.color = "var(--primary-color)";
            captureBtn.style.display = "flex";
            captureBtn.style.alignItems = "center";

            captureBtn.onclick = (e) => {
                e.preventDefault();
                if (isCapturingForFormula) {
                    // Cancel capture mode if the button is clicked again
                    isCapturingForFormula = false;
                    formulaCaptureTarget = null;
                    originalFormulaCellAddress = null;
                    hideCaptureMessage();
                    showStatus(t("status_ready"), "info");
                    return;
                }
                
                // Instant UI state update
                isCapturingForFormula = true;
                formulaCaptureTarget = inputEl as HTMLInputElement;
                
                showCaptureMessage(t("app_title"), t("select_cell_to_capture"), () => {
                    executeCellCapture();
                });
                showStatus(t("select_cell_to_capture"), "info");

                // Fetch active cell in the background to avoid blocking the UI
                Excel.run(async (context) => {
                    const cell = context.workbook.getActiveCell();
                    cell.load("address");
                    await context.sync();
                    originalFormulaCellAddress = cell.address;
                }).catch(err => {
                });
            };

            inputWrapper.appendChild(captureBtn);
            fieldDiv.appendChild(inputWrapper);
        }
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
                throw new Error(t("field_is_required", t(f.labelKey)));
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
            status.innerText = t("inserted_formula", formulaStr);
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
}

export function toggleSettings() {
  const settingsView = document.getElementById('settings-view');
  const mainView = document.getElementById('main-view');
  const settingsBtn = document.getElementById('settings-button');
  const isSettingsOpen = settingsView?.style.display === 'block';

  settingsBtn?.classList.remove('active');

  if (isSettingsOpen) {
    if (settingsView) settingsView.style.display = 'none';
    if (mainView) mainView.style.display = 'block';
  } else {
    if (settingsView) settingsView.style.display = 'block';
    if (mainView) mainView.style.display = 'none';
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
 
export function customTableWizardPrompt(
    captureDataFn: () => Promise<{ values: any[][], formulas: any[][], colStartIndex: number }>,
    workspaces: string[],
    allTables: string[],
    defaultWorkspace: string = "Public",
    defaultTableName: string = ""
): Promise<{ finalData: { tableName: string, workspace: string, parentTable: string, foreignKey: string, idField: string, fields: string[], records: any[] }, captured: { formulas: any[][], colStartIndex: number, originalHeaders: string[] } } | null> {
    return new Promise((resolve) => {
        const modal = document.getElementById("custom-wizard-modal")!;
        const stepIndicator = document.getElementById("wizard-step-indicator")!;
        const errorEl = document.getElementById("wizard-error")!;
        
        const stepRange = document.getElementById("wizard-step-range")!;
        const step1 = document.getElementById("wizard-step-1")!;
        const step2 = document.getElementById("wizard-step-2")!;
        const step3 = document.getElementById("wizard-step-3")!;
        const step4 = document.getElementById("wizard-step-4")!;
        
        const btnCancel = document.getElementById("wizard-btn-cancel") as HTMLButtonElement;
        const btnBack = document.getElementById("wizard-btn-back") as HTMLButtonElement;
        const btnNext = document.getElementById("wizard-btn-next") as HTMLButtonElement;
        const btnSave = document.getElementById("wizard-btn-save") as HTMLButtonElement;
        const btnDeleteSelected = document.getElementById("wizard-delete-selected") as HTMLButtonElement;

        // Step 1 Inputs
        const inputTableName = document.getElementById("wizard-table-name") as HTMLInputElement;
        const inputWorkspace = document.getElementById("wizard-workspace") as HTMLInputElement;
        const inputWorkspaceList = document.getElementById("wizard-workspace-list")!;
        const selectParent = document.getElementById("wizard-parent-table") as HTMLSelectElement;

        // Step 2 Inputs
        const selectId = document.getElementById("wizard-id-select") as HTMLSelectElement;
        const columnsList = document.getElementById("wizard-columns-list")!;

        // Step 3 Inputs
        const selectFk = document.getElementById("wizard-fk-select") as HTMLSelectElement;
        const parentNameDisplay = document.getElementById("wizard-parent-name-display")!;

        // Step 4 Displays
        const sumName = document.getElementById("wizard-summary-name")!;
        const sumWs = document.getElementById("wizard-summary-ws")!;
        const sumRecords = document.getElementById("wizard-summary-records")!;
        const sumId = document.getElementById("wizard-summary-id")!;
        const sumCols = document.getElementById("wizard-summary-cols")!;
        const sumParentRow = document.getElementById("wizard-summary-parent-row")!;
        const sumParent = document.getElementById("wizard-summary-parent")!;

        let currentStepIndex = 0;
        let logicalSteps = [0, 1, 2, 4];
        let hasParent = false;

        let headers: string[] = [];
        let dataRows: any[][] = [];
        let capturedFormulas: any[][] = [];
        let colStartIndex: number = 0;

        let finalData = {
            tableName: "",
            workspace: "",
            parentTable: "",
            foreignKey: "",
            idField: "",
            fields: [] as string[],
            records: [] as any[]
        };

        const showError = (msg: string) => {
            errorEl.innerText = msg;
            errorEl.style.display = "block";
        };
        const hideError = () => { errorEl.style.display = "none"; };

        const showStep = () => {
            hideError();
            const physicalStep = logicalSteps[currentStepIndex];
            
            stepRange.style.display = physicalStep === 0 ? "block" : "none";
            step1.style.display = physicalStep === 1 ? "block" : "none";
            step2.style.display = physicalStep === 2 ? "block" : "none";
            step3.style.display = physicalStep === 3 ? "block" : "none";
            step4.style.display = physicalStep === 4 ? "block" : "none";

            btnBack.style.display = currentStepIndex > 0 ? "inline-flex" : "none";
            btnNext.style.display = currentStepIndex < logicalSteps.length - 1 ? "inline-flex" : "none";
            btnSave.style.display = currentStepIndex === logicalSteps.length - 1 ? "inline-flex" : "none";

            stepIndicator.innerText = t("wizard_step_indicator", currentStepIndex + 1, logicalSteps.length);
        };

        const updateStepTotals = () => {
            hasParent = selectParent.value !== "-- None --";
            logicalSteps = hasParent ? [0, 1, 2, 3, 4] : [0, 1, 2, 4];
            showStep();
        };

        // --- INIT STEP 1 ---
        inputTableName.value = defaultTableName;
        inputWorkspace.value = defaultWorkspace;
        
        inputWorkspaceList.innerHTML = "";
        workspaces.forEach(ws => {
            const opt = document.createElement("option");
            opt.value = ws;
            inputWorkspaceList.appendChild(opt);
        });

        selectParent.innerHTML = `<option value="-- None --">${t("manage_rel_none")}</option>`;
        allTables.forEach(t => {
            const opt = document.createElement("option");
            opt.value = t;
            selectParent.appendChild(opt);
        });
        
        selectParent.onchange = updateStepTotals;
        updateStepTotals();

        // --- INIT STEP 2 ---
        const initStep2 = () => {
            selectId.innerHTML = "";
            headers.forEach((h, i) => {
                const opt = document.createElement("option");
                opt.value = h;
                opt.text = h || t("column_generic", i + 1);
                selectId.appendChild(opt);
            });
            
            columnsList.innerHTML = "";
            headers.forEach((item) => {
                const li = document.createElement("li");
                li.className = "sortable-item";
                li.draggable = true;
                li.dataset.original = item;
                li.innerHTML = `
                    <div style="display:flex; align-items:center; width:100%; border: 1px solid var(--border-color); padding: 4px 8px; border-radius: 4px; background: var(--input-bg); margin-bottom: 4px;">
                        <i class="ms-Icon ms-Icon--GlobalNavButton" style="margin-inline-end: 8px; color: #888; cursor: grab;"></i>
                        <input type="checkbox" class="col-delete-checkbox" style="margin-inline-end: 8px;" />
                      <input type="text" class="ms-TextField-field list-name-input" value="${item}" style="flex: 1; padding: 2px 8px; font-weight: 600; border: none; outline: none; background: transparent; color: var(--text-color);" placeholder="${t("name_placeholder")}" />
                    </div>
                `;
                const inputEl = li.querySelector("input[type='text']");
                if (inputEl) inputEl.addEventListener("mousedown", (e) => { e.stopPropagation(); });
                li.addEventListener("dragstart", () => li.classList.add("dragging"));
                li.addEventListener("dragend", () => li.classList.remove("dragging"));
                columnsList.appendChild(li);
            });

            const renderBadges = () => {
                const idName = selectId.value;
                Array.from(columnsList.children).forEach(li => {
                    const orig = (li as HTMLElement).dataset.original;
                    let badge = li.querySelector('.id-badge');
                    if (orig === idName) {
                        if (!badge) {
                            badge = document.createElement("span");
                            badge.className = "id-badge";
                            badge.style.cssText = "font-size:10px; color:#0078d4; margin-inline-start:8px; font-weight:bold;";
                            badge.innerText = t("id_badge");
                            li.querySelector('div')?.appendChild(badge);
                        }
                    } else {
                        if (badge) badge.remove();
                    }
                });
            };
            selectId.onchange = renderBadges;
            renderBadges();
        };

        const handleDragOver = (e: DragEvent) => {
            e.preventDefault();
            const dragging = columnsList.querySelector('.dragging') as HTMLElement;
            if (!dragging) return;
            const siblings = [...columnsList.querySelectorAll('.sortable-item:not(.dragging)')] as HTMLElement[];
            const nextSibling = siblings.find(sibling => {
                const box = sibling.getBoundingClientRect();
                return (e.clientY - box.top - box.height / 2) < 0;
            });
            if (nextSibling) columnsList.insertBefore(dragging, nextSibling);
            else columnsList.appendChild(dragging);
        };
        columnsList.addEventListener("dragover", handleDragOver);

        btnDeleteSelected.onclick = () => {
            const checkboxes = columnsList.querySelectorAll('.col-delete-checkbox:checked');
            checkboxes.forEach(cb => {
                const li = cb.closest('.sortable-item');
                if (li && (li as HTMLElement).dataset.original !== selectId.value) li.remove(); // Allow deletion if not the ID field
                else if (li) { showError(t("cannot_delete_id_column_msg", selectId.value)); }
            });
        };

        // --- INIT STEP 3 ---
        const initStep3 = () => {
            parentNameDisplay.innerText = selectParent.value;
            const newFieldsList = Array.from(columnsList.children).map(li => ((li as HTMLElement).querySelector(".list-name-input") as HTMLInputElement).value.trim());
            selectFk.innerHTML = "";
            newFieldsList.forEach(f => {
                const opt = document.createElement("option");
                opt.value = f;
                opt.text = f;
                selectFk.appendChild(opt);
            });
        };

        // --- INIT STEP 4 ---
        const initStep4 = () => {
            sumName.innerText = finalData.tableName;
            sumWs.innerText = finalData.workspace;
            sumRecords.innerText = String(finalData.records.length);
            sumId.innerText = finalData.idField;
            sumCols.innerText = finalData.fields.join(", ");
            
            if (hasParent) {
                sumParentRow.style.display = "block";
                sumParent.innerText = `${finalData.parentTable} (via ${finalData.foreignKey})`;
            } else {
                sumParentRow.style.display = "none";
            }
        };

        // --- NAVIGATION LOGIC ---
        btnNext.onclick = async () => {
            const physicalStep = logicalSteps[currentStepIndex];
            if (physicalStep === 0) {
                btnNext.disabled = true;
                try {
                    const captured = await captureDataFn();
                    headers = captured.values[0];
                    dataRows = captured.values.slice(1);
                    capturedFormulas = captured.formulas;
                    colStartIndex = captured.colStartIndex;
                    initStep2(); // Initialize columns list now that we have headers
                    currentStepIndex++;
                    showStep();
                } catch (err: any) {
                    showError(err.message);
                } finally {
                    btnNext.disabled = false;
                }
            } 
            else if (physicalStep === 1) {
                finalData.tableName = inputTableName.value.trim();
                finalData.workspace = inputWorkspace.value.trim();
                finalData.parentTable = selectParent.value !== "-- None --" ? selectParent.value : "";
                
                if (!finalData.tableName) { showError(t("table_name_required_error")); return; }
                if (allTables.includes(finalData.tableName)) { showError(t("table_already_exists_error", finalData.tableName)); return; }
                if (!finalData.workspace) { showError(t("workspace_name_required_error")); return; }

                currentStepIndex++;
                showStep();
            } 
            else if (physicalStep === 2) {
                const selectedIdOrigName = selectId.value;
                const idLi = Array.from(columnsList.children).find(li => (li as HTMLElement).dataset.original === selectedIdOrigName);
                if (!idLi) { showError(t("primary_id_column_missing_error")); return; }

                const idIndexOrig = headers.indexOf(selectedIdOrigName);
                const idSet = new Set();
                for (let i = 0; i < dataRows.length; i++) {
                    const val = String(dataRows[i][idIndexOrig]);
                    if (!val || val.trim() === "") { showError(t("row_has_empty_id_error", i + 1)); return; }
                    if (idSet.has(val)) { showError(t("duplicate_id_found_error", val)); return; }
                    idSet.add(val);
                }
                
                const newFieldsList = Array.from(columnsList.children).map(li => ((li as HTMLElement).querySelector(".list-name-input") as HTMLInputElement).value.trim());
                const uniqueFields = new Set(newFieldsList);
                if (uniqueFields.size !== newFieldsList.length) { showError(t("all_column_names_must_be_unique_error")); return; }
                
                const columnsInfo = Array.from(columnsList.children).map(li => ({
                    oIdx: headers.indexOf((li as HTMLElement).dataset.original as string),
                    newName: ((li as HTMLElement).querySelector(".list-name-input") as HTMLInputElement).value.trim()
                }));

                finalData.records = dataRows.map(row => {
                    const rec: any = {};
                    columnsInfo.forEach(col => { rec[col.newName] = row[col.oIdx]; });
                    rec.__DC_ID__ = String(row[idIndexOrig]);
                    return rec;
                });

                finalData.idField = ((idLi as HTMLElement).querySelector(".list-name-input") as HTMLInputElement).value.trim();
                finalData.fields = newFieldsList;

                if (hasParent) {
                    initStep3();
                } else {
                    finalData.foreignKey = "";
                    initStep4();
                }
                currentStepIndex++;
                showStep();
            }
            else if (physicalStep === 3) {
                finalData.foreignKey = selectFk.value;
                if (!finalData.foreignKey) { showError(t("link_column_required_error")); return; }
                initStep4();
                currentStepIndex++;
                showStep();
            }
        };

        btnBack.onclick = () => {
            currentStepIndex--;
            showStep();
        };

        const handleKeyDown = (e: KeyboardEvent) => {
            if (e.key === "Escape") {
                e.preventDefault();
                btnCancel.click();
            } else if (e.key === "Enter") {
                const activeEl = document.activeElement as HTMLElement;
                // Allow Enter to work normally if the user has explicitly tabbed to another button
                if (activeEl && activeEl.tagName === "BUTTON" && activeEl !== btnNext && activeEl !== btnSave && activeEl !== btnCancel && activeEl !== btnBack) return;
                
                e.preventDefault();
                if (btnNext.style.display !== "none") btnNext.click();
                else if (btnSave.style.display !== "none") btnSave.click();
            }
        };
        document.addEventListener("keydown", handleKeyDown);

        const cleanup = () => {
            modal.style.display = "none";
            btnNext.onclick = null;
            btnBack.onclick = null;
            btnSave.onclick = null;
            btnCancel.onclick = null;
            selectId.onchange = null;
            selectParent.onchange = null;
            columnsList.removeEventListener("dragover", handleDragOver);
            document.removeEventListener("keydown", handleKeyDown);
        };

        btnSave.onclick = () => { cleanup(); resolve({ finalData, captured: { formulas: capturedFormulas, colStartIndex, originalHeaders: headers } }); };
        btnCancel.onclick = () => { cleanup(); resolve(null); };

        modal.style.display = "flex";
        showStep();
    });
}

export async function executeNewTableCapture(defaultWorkspace: string = "Public", defaultTableName: string = "") {
  const status = document.getElementById("status-text");
  try {
      const storedData = await idbGet(IDB_KEYS.STORE);
      const store: Store = storedData ? JSON.parse(storedData) : {};
      const allTableKeys = Object.keys(store);

      let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
      let wsOrder: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];
      const familiesSet = new Set<string>(wsOrder);
      allTableKeys.forEach(k => familiesSet.add(store[k].family || 'Public'));
      const workspaces = Array.from(familiesSet);
      if (!workspaces.includes("Public")) workspaces.push("Public");

      const captureDataFn = async () => {
          return await Excel.run(async (context) => {
              const range = context.workbook.getSelectedRange();
              range.load(["values", "formulas", "columnIndex"]);
              await context.sync();
              const values = range.values;
              const formulas = range.formulas;
              if (values.length < 2) {
                  throw new Error(t("select_range_for_table_creation"));
              }
              return { values, formulas, colStartIndex: range.columnIndex };
          });
      };

      const wizardResult = await customTableWizardPrompt(
          captureDataFn,
          workspaces,
          allTableKeys,
          defaultWorkspace,
          defaultTableName
      );

      if (!wizardResult) return; // User cancelled

      const { finalData, captured } = wizardResult;
      const dataName = finalData.tableName;
      
      let rev = 1;
      store[dataName] = {
          dataTableName: dataName,
          family: finalData.workspace,
          idField: finalData.idField,
          revision: rev,
          history: {},
          fields: finalData.fields,
          records: finalData.records
      };

      // Apply calculation logic to initialize them structurally if necessary
      await applyCalculatedFields(store[dataName].records, store[dataName].fields, store[dataName].calculatedFields, store, dataName);

      await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
      if (status) { status.innerText = t("saved_records_in_table", finalData.records.length, dataName); status.style.color = "green"; }

      // Add relationship if parent table was selected
      if (finalData.parentTable && store[finalData.parentTable]) {
          const pName = finalData.parentTable;
          store[pName].relations = store[pName].relations || [];
          store[pName].relations.push({ subTable: dataName, foreignKey: finalData.foreignKey });
          await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
      }

      await idbSet(IDB_KEYS.DEFAULT_TABLE, dataName);
      await idbSet(IDB_KEYS.DEFAULT_REVISION, String(rev));

      // Refresh UI
      await renderDashboard();
      await refreshFormulas(true);

      let detectedFormulas: { header: string, formula: string }[] = [];
      if (captured.formulas.length > 1) {
          captured.originalHeaders.forEach((h: string, i: number) => {
              if (typeof captured.formulas[1][i] === 'string' && captured.formulas[1][i].startsWith('=')) {
                  detectedFormulas.push({ header: String(h), formula: captured.formulas[1][i] });
              }
          });
      }

      // Post-Wizard Prompt for detected formulas
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
                        const relativeIdx = colIdx - captured.colStartIndex;
                        if (relativeIdx >= 0 && relativeIdx < finalData.fields.length) {
                            return `[${finalData.fields[relativeIdx]}]`;
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
                fieldsList: finalData.fields,
                tablesWithFields: tablesWithFields,
                value: mappedCalcs[df.header].dc
            }));
            
            const reviewRes = await customFormPrompt(t("review_formulas_title"), t("review_formulas_msg"), reviewFields, t("save_formulas"));
            
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
        throw new Error(t("select_range_with_headers_and_data"));
      }

      const headers = values[0];
      const dataRows = values.slice(1);

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(t("data_table_not_found_error", dataTableName));
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
          t("replace_ver_title"),
          t("replace_ver_msg", dataTableName, values.length - 1),
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
        status.innerText = isLatest ? t("replaced_table_with_records", dataTableName, summary.records.length, summary.idField) : t("replaced_rev_of_table_with_records", selectedRev, dataTableName, summary.records.length);
        status.style.color = "green";
      }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus(t("error_replacing_data") + error.message, "error");
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
        throw new Error(t("select_range_with_headers_and_data"));
      }

      const headers = values[0];
      const dataRows = values.slice(1);

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
      if (existingStore) {
        store = JSON.parse(existingStore);
      }

      if (!store[dataTableName]) {
        throw new Error(t("data_table_not_found_error", dataTableName));
      }

      const dataSet = store[dataTableName];
      const existingIdField = dataSet.idField || dataSet.fields[0];

      const summary = await customDataSummaryPrompt(
          t("capture_rev_title"),
          t("capture_rev_msg", dataTableName, values.length - 1),
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
      
      if (status) { status.innerText = t("captured_rev_for_table", dataSet.revision, dataTableName, dataRows.length); status.style.color = "green"; }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus(t("error_capturing_new_revision") + error.message, "error");
  }
}

export async function renderDashboard() {
  const list = document.getElementById("workspaces-container");
  if (!list) return;
  
  const storedData = await idbGet(IDB_KEYS.STORE);
  const defaultSelect = document.getElementById("default-data-table-select") as HTMLSelectElement;
  const defaultRevSelect = document.getElementById("default-revision-select") as HTMLSelectElement;

  if (!storedData) {
    list.innerHTML = `<div style='font-size: 12px; color: #666; padding: 8px;'>${t("no_tables")}</div>`;
    if (defaultSelect) defaultSelect.innerHTML = "";
    if (defaultRevSelect) defaultRevSelect.innerHTML = "";
    renderFormulaBuilder();
    await renderVariables();
    return;
  }

  let store: Store = JSON.parse(storedData);
  let needsSave = false;

  // Cleanup: migrate old single-data-table format if present
  if ((store as any).collectionName && typeof (store as any).collectionName === "string") {
    const oldName = (store as any).collectionName;
    if (!store[oldName]) {
      store[oldName] = {
        dataTableName: (store as any).collectionName,
        fields: (store as any).fields,
        records: (store as any).records
      } as any;
    }
    delete (store as any).collectionName;
    delete (store as any).entityName;
    delete (store as any).fields;
    delete (store as any).records;
    needsSave = true;
  }

  // Remove any remaining invalid keys
  for (const key of Object.keys(store)) {
    if (!store[key] || typeof (store as any)[key] !== "object" || Array.isArray(store[key])) {
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

  let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
  let wsOrder: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];

  if (keys.length === 0 && wsOrder.length === 0) {
    list.innerHTML = `<div style='font-size: 12px; color: #666; padding: 8px;'>${t("no_workspaces")}</div>`;
    if (defaultSelect) defaultSelect.innerHTML = "";
    if (defaultRevSelect) defaultRevSelect.innerHTML = "";
    renderFormulaBuilder();
    await renderVariables();
    return;
  }

  // Setup Default Data Table dropdown
  if (defaultSelect) {
    let currentDefault = await idbGet(IDB_KEYS.DEFAULT_TABLE);

    defaultSelect.innerHTML = "";

    keys.forEach(key => {
      const opt = document.createElement("option");
      opt.value = key;
      opt.text = key;
      defaultSelect.appendChild(opt);
    });
    
    if (!currentDefault || !keys.includes(currentDefault)) {
      currentDefault = keys.length > 0 ? keys[0] : "";
      if (currentDefault) {
          await idbSet(IDB_KEYS.DEFAULT_TABLE, currentDefault);
      } else {
          await idbSet(IDB_KEYS.DEFAULT_TABLE, "");
      }
    }
    if (currentDefault) defaultSelect.value = currentDefault;
    
    const updateDefaultRevOptions = async (tableName: string) => {
      if (!defaultRevSelect) return;
      let currentDefaultRev = await idbGet(IDB_KEYS.DEFAULT_REVISION);

      defaultRevSelect.innerHTML = "";
      if (!tableName) return;
      const tData = store[tableName];
      if (tData) {
        const maxRev = tData.revision || 1;
        for (let i = maxRev; i >= 1; i--) {
          if (i === maxRev || (tData.history && tData.history[i])) {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.text = t("rev_latest_history", i, i === maxRev ? t("rev_latest_suffix") : "");
            defaultRevSelect.appendChild(opt);
          }
        }
      }
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
      captureParentSelect.innerHTML = `<option value="">${t("manage_rel_none")}</option>`;
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
      orderedFamilies.forEach(fam => {
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
                  <div style="flex:1; min-width: 0; margin-inline-end: 8px;">
                      <div style="font-weight:bold; font-size:13px; color:var(--primary-color); margin-bottom: 2px;">
                          <i class="ms-Icon ms-Icon--Table"></i> ${rel.main} 
                          <i class="ms-Icon ms-Icon--Forward" style="font-size:10px; margin:0 4px;"></i> 
                          <i class="ms-Icon ms-Icon--Table"></i> ${rel.sub}
                      </div>
                      <div style="font-size:11px; color:var(--text-color); opacity:0.8;">${t("link_column")}: ${rel.fk}</div>
                  </div>
              `;
  
              const actionDiv = document.createElement("div");
              actionDiv.style.display = "flex";
              actionDiv.style.gap = "4px";
              actionDiv.style.flexShrink = "0";
  
              const openBtn = document.createElement("button");
              openBtn.className = "icon-btn";
              openBtn.style.color = "#0078d4";
              openBtn.title = t("open_split_editor_btn");
              openBtn.innerHTML = '<i class="ms-Icon ms-Icon--SplitObject"></i>';
              openBtn.onclick = () => openRelationEditor(rel.main, rel.sub, rel.fk);
  
              const deleteBtn = document.createElement("button");
              deleteBtn.className = "icon-btn";
              deleteBtn.style.color = "#d13438";
              deleteBtn.title = t("delete_link_btn");
              deleteBtn.innerHTML = '<i class="ms-Icon ms-Icon--Delete"></i>';
              deleteBtn.onclick = async () => {
                  const confirm = await customConfirm(t("delete_link_title"), t("delete_link_msg", rel.main, rel.sub), t("yes_delete_btn"));
                  if (!confirm) return;
                  const storedData = await idbGet(IDB_KEYS.STORE);
                  if (storedData) {
                      const store = JSON.parse(storedData);
                      if (store[rel.main] && store[rel.main].relations) {
                          store[rel.main].relations = store[rel.main].relations.filter((r: any) => r.subTable !== rel.sub || r.foreignKey !== rel.fk);
                          await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
                          renderDashboard();
                          showStatus(t("link_removed_success"), "success");
                      }
                  }
              };
  
              actionDiv.appendChild(openBtn);
              actionDiv.appendChild(deleteBtn);
              li.appendChild(actionDiv);
              relList.appendChild(li);
          });
      } else {
          relList.innerHTML = `<li style="font-size: 12px; color: #666;">${t("no_relations")}</li>`;
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
    titleSpan.innerHTML = `<i class="ms-Icon ms-Icon--FabricFolder" style="margin-inline-end: 8px; color: var(--primary-color);"></i> <span style="font-weight: 600; font-size: 14px;">${fam}</span>`;
    famHeaderContent.appendChild(titleSpan);
    
    const buttonsContainer = document.createElement("div");
    buttonsContainer.style.display = "flex";
    buttonsContainer.style.gap = "6px";

    const addTableBtn = document.createElement("button");
    addTableBtn.className = "ms-Button ms-Button--primary";
    addTableBtn.style.minWidth = "auto";
    addTableBtn.style.padding = "0 8px";
    addTableBtn.style.height = "24px";
    addTableBtn.style.lineHeight = "24px";
    addTableBtn.title = t("add_new_table");
    addTableBtn.innerHTML = `<span class="ms-Button-label" style="font-size: 11px;"><i class="ms-Icon ms-Icon--Add icon-no-margin-sm" style="margin-inline-end: 4px;"></i><span class="hide-sm">${t("add_new_table")}</span></span>`;
    addTableBtn.onclick = (e) => {
        e.stopPropagation();
        executeNewTableCapture(fam, "");
    };

    const manageTablesBtn = document.createElement("button");
    manageTablesBtn.className = "ms-Button ms-Button--default";
    manageTablesBtn.style.minWidth = "auto";
    manageTablesBtn.style.padding = "0 8px";
    manageTablesBtn.style.height = "24px";
    manageTablesBtn.style.lineHeight = "24px";
    manageTablesBtn.title = t("manage_btn");
    manageTablesBtn.innerHTML = `<span class="ms-Button-label" style="font-size: 11px;"><i class="ms-Icon ms-Icon--Settings icon-no-margin-sm" style="margin-inline-end: 4px;"></i><span class="hide-sm">${t("tables_btn")}</span></span>`;
    manageTablesBtn.onclick = (e) => { e.stopPropagation(); manageWorkspaceTables(fam); };
    
    buttonsContainer.appendChild(addTableBtn);
    buttonsContainer.appendChild(manageTablesBtn);
    famHeaderContent.appendChild(buttonsContainer);

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
          <div><i class="ms-Icon ms-Icon--Table" style="margin-inline-end: 8px;"></i> ${key}</div>
          <div style="font-size:12px; font-weight:normal; opacity:0.8;">${t("rev_count_rows", rev, count)}</div>
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
            btn.style.backgroundColor = "var(--card-bg)";
            btn.style.cursor = "pointer";
            btn.style.minHeight = "60px";
            btn.style.textAlign = "center";
            btn.style.boxShadow = "0 1px 2px rgba(0,0,0,0.05)";
            btn.style.transition = "all 0.15s ease-in-out";
            
            btn.onmouseover = () => { btn.style.backgroundColor = "var(--neutral-lighter, #f3f2f1)"; };
            btn.onmouseout = () => { btn.style.backgroundColor = "var(--card-bg)"; };
            
            let iconColor = "var(--text-color)";
            if (colorCls === 'primary') iconColor = "#0078d4";
            else if (colorCls === 'danger') iconColor = "#d13438";

            btn.innerHTML = `<i class="ms-Icon ms-Icon--${icon}" style="font-size: 16px; margin-bottom: 6px; color: ${iconColor};"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${text}</span>`;
            btn.onclick = onClick;
            return btn;
        };

        const insertTableBtn = createGridBtn("Table", t("insert_sheet"), "default", () => insertTable(key));
        const replaceBtn = createGridBtn("Sync", t("replace_version"), "default", () => replaceTableData(key));
        const appendBtn = createGridBtn("Add", t("append_data"), "default", () => appendTableData(key));
        const snapshotBtn = createGridBtn("Camera", t("snapshot"), "primary", () => createSnapshot(key));
        const deleteVersionBtn = createGridBtn("RemoveEvent", t("del_version"), "danger", () => deleteCurrentVersion(key));
        const deleteBtn = createGridBtn("Delete", t("del_table"), "danger", () => deleteDataTable(key));
        const exportCSVBtn = createGridBtn("Download", t("export_csv"), "default", () => exportCSV(key));
        const insertDropdownBtn = createGridBtn("Dropdown", t("headers_dropdown"), "default", () => insertDropdown(key));
        const editBtn = createGridBtn("Edit", t("form_editor"), "primary", () => loadRecordForEdit(key));
        const gridEditBtn = createGridBtn("GridViewSmall", t("grid_editor"), "primary", () => openGridEditor(key));
        const duplicateRecordBtn = createGridBtn("Copy", t("clone_record"), "default", () => duplicateRecordPrompt(key));
        const editColumnsBtn = createGridBtn("Sort", t("manage_columns"), "default", () => manageColumns(key));
        const moveWorkspaceBtn = createGridBtn("FabricFolder", t("move_workspace"), "default", () => moveTableWorkspace(key));
        const captureRevBtn = createGridBtn("Camera", t("new_revision"), "primary", () => captureNewRevision(key));
        const cloneSubBtn = createGridBtn("Copy", t("clone_sub_records"), "default", () => cloneSubRecordsPrompt(key));

        const revSelectContainer = document.createElement("div");
        revSelectContainer.style.marginBottom = "12px";
        revSelectContainer.innerHTML = `<label style="font-size:12px; font-weight:600; color:var(--text-color); margin-inline-end:8px;">${t("target_revision")}</label>`;

        const revSelect = document.createElement("select");
        revSelect.id = `fb-rev-${key}`;
        revSelect.className = "ms-Dropdown-title";
        revSelect.style.width = "auto";
        revSelect.style.display = "inline-block";
        for (let i = rev; i >= 1; i--) {
          if (i === rev || (dataTable.history && dataTable.history[i])) {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.text = t("revision_latest_history", i, i === rev ? t("rev_latest_suffix") : "");
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
          replaceBtn.innerHTML = `<i class="ms-Icon ms-Icon--Sync" style="font-size: 16px; margin-bottom: 6px; color: var(--text-color);"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${t("replace_version")}</span>`;
          snapshotBtn.innerHTML = `<i class="ms-Icon ms-Icon--${isLatest ? 'Camera' : 'Undo'}" style="font-size: 16px; margin-bottom: 6px; color: #0078d4;"></i><span style="font-size: 10px; line-height: 1.1; color: var(--text-color); font-weight: 600;">${isLatest ? t("snapshot") : t("snapshot")}</span>`;

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
                const isCurrentlyOpen = accHeader.classList.contains("open");
                
                // Collapse all others
                accordions.forEach(acc => {
                    acc.content.style.display = "none";
                    acc.header.classList.remove("open");
                    const closedIconColor = acc.theme === 'danger' ? '#d13438' : acc.theme === 'fast' ? '#0078d4' : 'var(--primary-color)';
                    const accIcon = `<span class="icon" style="display:inline-block; width:15px; color: ${closedIconColor};">&#9654;</span>`;
                    acc.header.innerHTML = `${accIcon}<span>${acc.title}</span>`;
                });

                // If it was closed, open it. (If it was open, it's now closed from the loop above).
                if (!isCurrentlyOpen) {
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

        const secEdit = buildAccordion(t("data_entry_views"), [
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

        const secSchema = buildAccordion(t("schema_operations"), schemaElements, 'default', false, true);

        const secVersion = buildAccordion(t("versioning_danger_zone"), [
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
        emptyMsg.innerText = t("no_tables_in_ws");
        famContent.appendChild(emptyMsg);
    }
  });
  
  renderFormulaBuilder();
  await renderVariables();
}

export async function deleteDataTable(dataTableName: string) {
  const confirmed = await customConfirm(t("del_table_title"), t("del_table_msg", dataTableName), t("del_table_confirm"));
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
        status.innerText = t("deleted_entire_data_table", dataTableName);
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

    const msg = isLatest ? t("del_version_latest_msg", dataTableName) : t("del_version_hist_msg", selectedRev, dataTableName);
    const confirmed = await customConfirm(t("del_version_title"), msg, t("del_version_confirm"));
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
          status.innerText = t("deleted_current_version_rolled_back", highestOldRev);
          status.style.color = "green";
        }
      } else {
        delete store[dataTableName];
        if (status) {
          status.innerText = t("deleted_only_version_table_removed", dataTableName);
          status.style.color = "blue";
        }
      }
    } else {
      if (dataSet.history && dataSet.history[selectedRev]) {
        delete dataSet.history[selectedRev];
        if (status) {
          status.innerText = t("deleted_historical_rev", selectedRev, dataTableName);
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
    showStatus(t("error_deleting_version") + error.message, "error");
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
          status.innerText = t("locked_rev_current_is", currentRev, dataSet.revision);
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
          status.innerText = t("restored_rev_as_new_active", selectedRev, dataSet.revision);
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
        showStatus(t("storage_limit_reached"), "error");
    } else {
        showStatus(t("error_creating_snapshot") + error.message, "error");
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
      if (status) { status.innerText = t("cannot_resort_historical_revision"); status.style.color = "red"; }
      return;
    }

    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
    const varNames = Object.keys(variables);
    const tablesWithFields: Record<string, string[]> = {};
    Object.keys(store).forEach(t => tablesWithFields[t] = store[t].fields || []);

    const idField = dataSet.idField || dataSet.fields[0];
    const promptResult = await customManageColumnsPrompt(t("edit_cols_title"), t("edit_cols_msg"), dataSet.fields, idField, dataSet.calculatedFields || {}, varNames, tablesWithFields);
    
    if (!promptResult) return; // User cancelled

    const columnChanges = promptResult.changes;
    const newFields = columnChanges.map(c => c.newName);
    
    // Check for duplicates
    const uniqueFields = new Set(newFields);
    if (uniqueFields.size !== newFields.length) {
        if (status) { status.innerText = t("column_names_must_be_unique"); status.style.color = "red"; }
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
      status.innerText = t("columns_updated_successfully", dataSet.revision);
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus(t("error_general") + error.message, "error");
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

    const id = await customPrompt(t("edit_record_title"), t("edit_record_msg", dataTableName), "", allIds);
    if (!id || id.trim() === "") return;

    const record = targetDataSet.records.find((r: any) => String(r.__DC_ID__) === String(id));

    if (!record) {
        if (status) { status.innerText = t("record_not_found_error", id); status.style.color = "red"; }
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

    const editResult = await customFormPrompt(t("edit_record_title_with_id", id), t("edit_record_update_msg"), formFields);
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
        status.innerText = t("record_updated_refreshing_excel");
        status.style.color = "green";
    }

    await refreshFormulas(true);
  } catch (error: any) {
    if (error.name === "QuotaExceededError" || (error.message && error.message.includes("exceeded the quota"))) {
        showStatus(t("storage_limit_reached_cannot_save"), "error");
    } else {
        showStatus(t("error_saving_changes") + error.message, "error");
    }
  }
}

export async function exportCSV(dataTableName: string) {
  const status = document.getElementById("status-text");
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) throw new Error(t("no_data_found_error"));

    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) throw new Error(t("data_table_not_found_error", dataTableName));

    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;

    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
    }

    if (!targetDataSet || !targetDataSet.fields || !targetDataSet.records) {
      throw new Error(t("no_data_found_for_revision_error"));
    }

    await exportCSVData(dataTableName, selectedRev, targetDataSet);

    if (status) {
      status.innerText = t("exported_table_to_csv", dataTableName, selectedRev);
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus(t("error_exporting_csv") + error.message, "error");
  }
}

export async function refreshFormulas(silent: boolean = false) {
  const status = document.getElementById("status-text");
  try {
    if (status && !silent) status.innerText = t("refreshing_dashboard_formulas");

    if (!silent) {
      await renderDashboard();
    }

    const count = await executeRefreshFormulas();
    if (status && !silent) {
      status.innerText = t("refreshed_dashboard_formulas", count);
      status.style.color = "green";
    }
  } catch (error: any) {
    if (!silent) showStatus(t("error_general") + error.message, "error");
  }
}

export async function convertToValues() {
  const status = document.getElementById("status-text");
  try {
    if (status) status.innerText = t("converting_please_wait");
    const count = await executeConvertToValues();
    if (status) {
      status.innerText = t("converted_formulas_to_values", count);
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus(t("error_general") + error.message, "error");
  }
}

export async function backupData() {
  try {
    await downloadBackup();

    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = t("backup_downloaded_successfully");
      status.style.color = "green";
    }
  } catch (error: any) {
    showStatus(t("backup_error") + error.message, "error");
  }
}

export function triggerRestore() {
  document.getElementById("restore-file-input").click();
}

export async function restoreData(event: any) {
  const status = document.getElementById("status-text");
  try {
    await processRestoreFile(event);
    if (status) { status.innerText = t("data_restored_successfully"); status.style.color = "green"; }
    await renderDashboard();
    await refreshFormulas(true);
  } catch (error: any) {
    showStatus(t("restore_error") + error.message, "error");
  }
}

export async function insertDropdown(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    if (!dataSet) throw new Error(t("no_data_found_error"));
    const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
    const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
    let targetDataSet = dataSet;
    if (selectedRev !== (dataSet.revision || 1) && dataSet.history && dataSet.history[selectedRev]) {
      targetDataSet = dataSet.history[selectedRev];
    }
    if (!targetDataSet || !targetDataSet.fields) throw new Error(t("no_headers_found_error"));

    await executeInsertDropdown(targetDataSet.fields);
    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = t("inserted_headers_dropdown", dataTableName);
      status.style.color = "green";
    }
  } catch (error: any) {
    console.error(error);
    showStatus(t("error_inserting_dropdown") + error.message, "error");
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
      throw new Error(t("no_data_found_for_table_error"));
    }

    const res = await customFormPrompt(t("insert_table_title"), t("select_columns_to_insert"), [
        { id: "cols", label: t("columns_label"), type: "checkboxes", options: targetDataSet.fields }
    ], t("insert_table_confirm"));
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
        table.showBandedRows = false;
        await context.sync();
        
        if (rows.length > 0) {
            table.getDataBodyRange().formulas = rows;
            try {
                await context.sync();
            } catch (e: any) {
                table.getDataBodyRange().values = rows.map(r => r.map((c: any) => typeof c === 'string' && c.startsWith('=') ? `'${c}` : c));
                await context.sync();
                throw new Error(t("excel_formula_js_syntax_error"));
            }
        }
    });

    const status = document.getElementById("status-text");
    if (status) {
       status.innerText = t("inserted_table_for", dataTableName);
       status.style.color = "green";
    }
  } catch (error: any) {
    showStatus(t("error_inserting_table") + error.message, "error");
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
                showEditorView(!tableProp.isNullObject ? tableProp.value : t("unknown_table_name"), isLive || purposeProp.value === "MainDataEditor");
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
    if (indicator) { indicator.style.display = isLiveSync ? "block" : "none"; }
    
    applyTranslations(); // Trigger translations to format Editor mode correctly upon rendering
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
        if (status) { status.innerText = t("loading_split_editor"); status.style.color = "blue"; }
        const saved = await saveGridEditor(false); // ensure any open editor saves
        if (!saved) {
             showStatus(t("fix_errors_before_new_editor"), "error");
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
        const firstId = targetMainData.records && targetMainData.records.length > 0 ? targetMainData.records[0][idField] : t("new_record_id_placeholder");

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
        
        showStatus(t("opened_split_editor_arrange_windows", mainTableName, subTableName), "success");
    } catch (e: any) {
        showStatus(t("error_opening_split_editor") + e.message, "error");
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
        
        const msg = filterField ? t("opened_sub_table_in_editor", filterField, filterValue) : t("opened_table_in_grid_editor", dataTableName);
        if (status) { status.innerText = msg; status.style.color = "blue"; }
    });
  } catch (error: any) {
      showStatus(t("error_opening_grid_editor") + error.message, "error");
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
                if (!val || val.trim() === "") throw new Error(t("row_has_empty_id_in_table_error", i + 1, dataTableName));
                if (idSet.has(val)) throw new Error(t("duplicate_id_found_in_table_error", val, dataTableName));
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
                status.innerText = t("saved_records", savedCount);
                status.style.color = "green"; 
            } else {
                if (!shouldClose) status.innerText = t("no_changes_detected");
            }
        }
        
        await refreshFormulas(true);
    });
    return true;
  } catch (error: any) { 
      showStatus(t("error_saving_grid_editor") + error.message, "error");
      return false;
  }
}

export async function switchGridEditorRecord(newId: string, foreignKey: string, targetTable: string, selectedColumns: string[], subSheetName: string = "DC_Grid_Editor", mainTableName: string = "") {
    const status = document.getElementById("status-text");
    try {
        if (status) { status.innerText = t("auto_syncing_switching_to_record", newId); status.style.color = "blue"; }
        const saved = await saveGridEditor(false, subSheetName); // Save ONLY the sub-sheet without closing
        if (!saved) {
            return;
        }

        // Recreate the sheet cleanly in the background without stealing focus
        await openGridEditor(targetTable, foreignKey, newId, mainTableName, selectedColumns, subSheetName, "SubDataEditor", false);

        if (status) { status.innerText = t("switched_sub_table_to_record", newId); status.style.color = "green"; }
    } catch (error: any) { 
        console.error("Error switching record:", error);
        showStatus(t("error_switching_record") + error.message, "error");
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
  } catch (error: any) { showStatus(t("error_canceling_grid_editor") + error.message, "error"); }
}

export async function manageRelations(dataTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const dataSet = store[dataTableName];
    const existingRelations = dataSet.relations || [];
    
    const allTables = Object.keys(store).filter(t => t !== dataTableName);
    const relationStrings = existingRelations.map((r: any) => t("relation_link_suffix", r.subTable, r.foreignKey));

    const tablesWithFields: Record<string, string[]> = {};
    allTables.forEach(t => tablesWithFields[t] = store[t].fields || []);

    const fields: FormField[] = [];
    if (relationStrings.length > 0) {
        fields.push({ id: "keepRelations", label: t("manage_rel_keep"), type: "checkboxes", options: relationStrings });
    }
    
    fields.push({ id: "newSubTable", label: t("manage_rel_add_target"), type: "select", options: [t("manage_rel_none"), ...allTables] });
    fields.push({ id: "newForeignKey", label: t("manage_rel_add_link"), type: "select", options: [t("manage_rel_none")], dependsOn: "newSubTable", optionsMap: tablesWithFields });

    const res = await customFormPrompt(t("manage_rel_title"), t("manage_rel_manage_msg", dataTableName), fields);
    
    if (!res) return;

    let finalRelations: any[] = [];
    if (relationStrings.length > 0 && res.keepRelations !== undefined) {
        const kept = res.keepRelations.split(",").map((s: string) => s.trim());
        finalRelations = existingRelations.filter((r: any) => kept.includes(t("relation_link_suffix", r.subTable, r.foreignKey)));
    }

    if (res.newSubTable && res.newSubTable !== t("manage_rel_none") && res.newForeignKey && res.newForeignKey.trim() !== "" && res.newForeignKey !== t("manage_rel_none")) {
        if (!finalRelations.find(r => r.subTable === res.newSubTable)) {
            finalRelations.push({ subTable: res.newSubTable, foreignKey: res.newForeignKey.trim() });
        }
    }

    dataSet.relations = finalRelations;
    
    await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
    renderDashboard();
    showStatus(t("relations_updated_for", dataTableName), "success");
    if (res.newSubTable && res.newSubTable !== t("manage_rel_none") && res.newForeignKey && res.newForeignKey.trim() !== "" && res.newForeignKey !== t("manage_rel_none")) {
            const addRollup = await customConfirm(t("add_rollup_title"), t("add_rollup_msg", dataTableName, res.newSubTable), t("yes_btn"));
            if (addRollup) {
                const subFields = store[res.newSubTable]?.fields || [];
                const rollupRes = await customFormPrompt(t("rollup_field_title"), t("define_the_rollup"), [
                    { id: "fieldName", label: t("rollup_field_name"), type: "text" },
                    { id: "type", label: t("rollup_type"), type: "select", options: ["SUM", "COUNT"] },
                    { id: "col", label: t("rollup_col"), type: "select", options: [t("manage_rel_none"), ...subFields] }
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
                        showStatus(t("added_rollup_field", rollupRes.fieldName), "success");
                    }
                }
            }
        }

  } catch (error: any) { showStatus(t("error_managing_relations") + error.message, "error"); }
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
        throw new Error(t("select_range_with_headers_and_data"));
      }

      let store: Store = {};
      const existingStore = await idbGet(IDB_KEYS.STORE);
      if (existingStore) store = JSON.parse(existingStore);

      const dataSet = store[dataTableName];
      if (!dataSet) throw new Error(`Data table '${dataTableName}' not found.`);

      const revSelect = document.getElementById(`fb-rev-${dataTableName}`) as HTMLSelectElement;
      const selectedRev = revSelect ? parseInt(revSelect.value) : dataSet.revision || 1;
      if (selectedRev !== (dataSet.revision || 1)) {
          throw new Error(t("cannot_append_to_historical_revision"));
      }

      const summary = await customDataSummaryPrompt(
          t("append_data_title"),
          t("review_data_to_append", dataTableName, values.length - 1),
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
              throw new Error(t("duplicate_id_already_exists_error", id, dataTableName));
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
        status.innerText = t("appended_records_current_is", newRecords.length, dataSet.revision);
        status.style.color = "green";
      }
      await renderDashboard();
      await refreshFormulas(true);
    });
  } catch (error: any) {
    showStatus(t("error_appending_data") + error.message, "error");
  }
}

export async function cloneSubRecordsPrompt(mainTableName: string) {
  try {
    const storedData = await idbGet(IDB_KEYS.STORE);
    if (!storedData) return;
    const store = JSON.parse(storedData);
    const mainDataSet = store[mainTableName];
    if (!mainDataSet || !mainDataSet.relations || mainDataSet.relations.length === 0) {
        showStatus(t("no_relations_defined_for_table"), "error");
        return;
    }

    const allMainIds = mainDataSet.records.map((r: any) => String(r.__DC_ID__));
    const relationsOptions = mainDataSet.relations.map((r: any) => t("relation_fk_suffix", r.subTable, r.foreignKey));

    const res1 = await customFormPrompt(t("clone_sub_title"), t("clone_sub_select_target"), [
        { id: "relation", label: t("clone_sub_relation"), type: "select", options: relationsOptions }
    ]);
    if (!res1 || !res1.relation) return;

    const selectedRel = mainDataSet.relations.find((r: any) => res1.relation.includes(r.subTable));
    if (!selectedRel) return;

    const res2 = await customFormPrompt(t("clone_from", selectedRel.subTable), t("clone_select_ids"), [
        { id: "sourceId", label: t("clone_source_id"), type: "autocomplete", options: allMainIds },
        { id: "targetIds", label: t("clone_target_ids"), type: "checkboxes", options: allMainIds }
    ]);
    if (!res2 || !res2.sourceId || !res2.targetIds) return;

    const targetIds = res2.targetIds.split(",");
    if (targetIds.length === 0 || targetIds.includes("")) return;

    const subDataSet = store[selectedRel.subTable];
    if (!subDataSet) return;

    const sourceRecords = subDataSet.records.filter((r: any) => String(r[selectedRel.foreignKey]) === String(res2.sourceId));
    if (sourceRecords.length === 0) {
        showStatus(t("no_sub_records_found", res2.sourceId), "error");
        return;
    }

    const newRecords: any[] = [];
    targetIds.forEach(tId => {
        if (tId === res2.sourceId) return; // skip self
        sourceRecords.forEach((sr: any) => {
            const cloned = { ...sr };
            cloned.__DC_ID__ = `${t("cloned_id_prefix")}${Date.now()}_${Math.floor(Math.random() * 1000)}`;
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

    showStatus(t("successfully_cloned_sub_records", sourceRecords.length, targetIds.length), "success");
  } catch (error: any) {
    showStatus(t("error_cloning_sub_records") + error.message, "error");
  }
}

export async function moveTableWorkspace(tableName: string) {
    try {
        const storedData = await idbGet(IDB_KEYS.STORE);
        if (!storedData) return;
        const store = JSON.parse(storedData);
        const dataSet = store[tableName];
        if (!dataSet) return;

        let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
        let wsOrder: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];

        const familiesSet = new Set<string>(wsOrder);
        Object.keys(store).forEach(k => familiesSet.add(store[k].family || 'Public'));
        const familiesList = Array.from(familiesSet);

        const res = await customFormPrompt(t("move_ws_title"), t("select_or_type_new_workspace", tableName), [
            { id: "newWorkspace", label: t("workspace_label"), type: "autocomplete", options: familiesList, value: dataSet.family || t("public_workspace") }
        ]);

        if (!res || !res.newWorkspace || res.newWorkspace.trim() === "") return;
        const newWorkspace = res.newWorkspace.trim();
        if (newWorkspace === (dataSet.family || 'Public')) return; // No change

        dataSet.family = newWorkspace;
        await idbSet(IDB_KEYS.STORE, JSON.stringify(store));
        renderDashboard();
        showStatus(t("moved_table_to_workspace", tableName, newWorkspace), "success");

    } catch (error: any) {
        showStatus(t("error_moving_table") + error.message, "error");
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
    
    const res = await customFormPrompt(t("add_var_title"), t("add_var_msg"), [
        { id: "vName", label: t("add_var_name"), type: "text" },
        { id: "vFormula", label: t("add_var_formula"), type: "formula", varsList: varNames, tablesWithFields: tablesWithFields }
    ]);

    if (res) {
        if (!res.vName || res.vName.trim() === "") {
            showStatus(t("variable_name_required"), "error");
            return;
        }
        if (!res.vFormula || res.vFormula.trim() === "") {
            showStatus(t("formula_required"), "error");
            return;
        }
        try {
            // Validation Layer: Test compiling and executing the formula
            const evaluateVar = (vName: string, visited: Set<string>): any => {
                if (visited.has(vName)) throw new Error(t("loop_detected"));
                visited.add(vName);
                if (vName === res.vName) throw new Error(t("variable_cannot_reference_itself"));
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
            showStatus(t("variable_saved", res.vName), "success");
        } catch (error: any) {
            showStatus(t("invalid_formula_for_variable", res.vName, error.message), "error");
        }
    }
}

export async function addWorkspace() {
    try {
        const newName = await customPrompt(t("add_ws_title"), t("add_ws_msg"));
        if (!newName || newName.trim() === "") return;

        let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
        let wsOrder: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];
        if (!wsOrder.includes(newName)) {
            wsOrder.push(newName);
            await idbSet(IDB_KEYS.WORKSPACES_ORDER, JSON.stringify(wsOrder));
            renderDashboard();
        }

        const res = await customConfirm(t("add_ws_added_title"), t("add_ws_added_msg", newName), t("add_ws_capture_btn"));
        if (res) {
            await executeNewTableCapture(newName);
        }
    } catch (error: any) {
        showStatus(t("error_adding_workspace") + error.message, "error");
    }
}

export async function renderVariables() {
    const vStoreRaw = await idbGet(IDB_KEYS.VARIABLES);
    const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
    const list = document.getElementById("variables-list");
    if (!list) return;

    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};

    const evaluateVar = (vName: string, visited: Set<string>): any => {
        if (visited.has(vName)) throw new Error(t("loop_detected"));
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

    list.innerHTML = "";

    for (const [vName, vFormula] of Object.entries(variables)) {
        let result: any = t("error_text");
        try {
            result = evaluateVar(vName, new Set());
        } catch(e: any) { result = e.message === t("loop_detected") ? t("loop_error") : t("error_text"); }

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
            <div style="flex:1; min-width: 0; margin-inline-end: 8px;">
                <div style="font-weight:bold; font-size:13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${vName}">${vName}</div>
                <div style="font-size:11px; color:var(--text-color); opacity:0.8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title='${vFormula.replace(/'/g, "&#39;")}'>${vFormula}</div>
            </div>
            <div style="color:green; font-weight:bold; margin-inline-end: 8px; font-size:14px; max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0;" title="${result}">${result}</div>
        `;

        const actionDiv = document.createElement("div");
        actionDiv.style.display = "flex";
        actionDiv.style.gap = "4px";
        actionDiv.style.flexShrink = "0";

        const insertBtn = document.createElement("button");
        insertBtn.className = "icon-btn";
        insertBtn.style.color = "#0078d4";
        insertBtn.innerHTML = `<i class="ms-Icon ms-Icon--Insert"></i>`;
        insertBtn.title = t("insert_to_sheet");
        insertBtn.onclick = async () => {
            const formulaStr = `=DC.VAR("${vName}")`;
            try {
                await Excel.run(async (context) => {
                    const cell = context.workbook.getActiveCell();
                    cell.formulas = [[formulaStr]];
                    await context.sync();
                });
                showStatus(t("inserted_variable_to_sheet", vName), "success");
            } catch (error: any) {
                showStatus(t("error_inserting_variable") + error.message, "error");
            }
        };

        const delBtn = document.createElement("button");
        delBtn.className = "icon-btn";
        delBtn.style.color = "#d13438";
        delBtn.innerHTML = `<i class="ms-Icon ms-Icon--Delete"></i>`;
        delBtn.title = t("delete_variable");
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

    const res = await customFormPrompt(t("dup_record_title"), t("dup_record_msg"), [
        { id: "sourceId", label: t("source_record_id_label"), type: "autocomplete", options: allMainIds },
        { id: "newId", label: t("dup_record_new_id"), type: "text" }
    ]);

    if (!res || !res.sourceId || !res.newId) return;

    const sourceId = res.sourceId.trim();
    const newId = res.newId.trim();

    if (sourceId === "" || newId === "") return;
    if (allMainIds.includes(newId)) {
        showStatus(t("record_id_already_exists", newId, dataTableName), "error");
        return;
    }

    const sourceRecord = mainDataSet.records.find((r: any) => String(r.__DC_ID__) === sourceId);
    if (!sourceRecord) {
        showStatus(t("source_record_not_found", sourceId), "error");
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

    showStatus(t("successfully_duplicated_record", sourceId, newId, String(clonedSubRecordsCount)), "success");
  } catch (error: any) {
    showStatus(t("error_duplicating_record") + error.message, "error");
  }
}

export async function manageWorkspaces() {
    const storedData = await idbGet(IDB_KEYS.STORE);
    const store = storedData ? JSON.parse(storedData) : {};
    
    let wsOrderRaw = await idbGet(IDB_KEYS.WORKSPACES_ORDER);
    let familiesList: string[] = wsOrderRaw ? JSON.parse(wsOrderRaw) : [];
    
    const currentFamilies = new Set<string>();
    Object.keys(store).forEach(k => currentFamilies.add(store[k].family || 'Public'));

    currentFamilies.forEach(f => { if (!familiesList.includes(f)) familiesList.push(f); });

    const res = await customManageListPrompt(t("manage_ws_title"), t("manage_ws_desc"), familiesList, "", true);

    if (!res) return;

    let hasChanges = false;
    let newOrder: string[] = [];

    for (const r of res) {
        if (r.isDeleted) {
            if (!r.isNew) {
                const ws = r.original;
                const tablesInWs = Object.keys(store).filter(k => (store[k].family || 'Public') === ws);
                if (tablesInWs.length > 0) {
                    const confirm = await customConfirm(t("delete_ws_title"), t("delete_ws_msg", ws, tablesInWs.length), t("del_table_confirm"));
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

    const res = await customManageListPrompt(t("manage_tb_title", fam), t("manage_tb_desc"), tablesInWs, "", false);
    if (!res) return;

    let hasChanges = false;
    let newOrder: string[] = tableOrder.filter(t => !tablesInWs.includes(t));

    for (const r of res) {
        if (r.isDeleted) {
            if (!r.isNew) {
                const tbName = r.original;
                const confirm = await customConfirm(t("delete_tb_title"), t("delete_tb_msg", tbName), t("del_version_confirm"));
                if (confirm) {
                    delete store[tbName];
                    for (const key of Object.keys(store)) {
                        if (store[key].relations) {
                            store[key].relations = store[key].relations.filter((rel: any) => rel.subTable !== tbName);
                        }
                    }
                    hasChanges = true;
                } else {
                    newOrder.push(tbName);
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
    await executeNewTableCapture(fam, defaultName);
}