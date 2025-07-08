# Python MOVE Tracker Automation Script Specification as at 2025-07-08 12:11

## 1\. Script Name & Purpose

  * **Name:** `move_tracker_report.py`
  * **Purpose:** To automate the process of generating MOVE (Minimal Outcome-Value Effort) project progress reports. It will read project data and configurations from a structured Excel workbook, compute the `Progress_Log` for a specified snapshot date, generate Work Execution and Fever Charts (as SVG images for testability), and insert these updated data and charts back into the Excel workbook. It also provides functionality to generate a template Excel file for new users.

-----

## 2\. Dependencies

The script will be a single-file script adhering to **PEP 723**. It will include a shebang for `uv` and specify the minimum Python version.

```python
#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.13"
# dependencies = [
#     "pandas>=2.0",
#     "numpy>=1.20",
#     "matplotlib>=3.5",
#     "openpyxl>=3.0",
#     "rich>=12.0",
#     "typer>=0.9.0"
# ]
# ///
```

-----

## 3\. Command-Line Interface (CLI)

The script should be executable from the command line and support the following arguments using `typer`:

  * `--excel-path <path>`:
      * **Required (unless `--create-template` is used).** Path to the Excel workbook (`.xlsx`) containing the MOVE project data.
      * Example: `--excel-path "MOVE_Project_Data.xlsx"`
  * `--create-template`:
      * **Optional.** A boolean flag. If present, the script will **create a new Excel file** at the location specified by `--excel-path` (or a default name if `--excel-path` is omitted and this flag is used) populated with the required sheet names and column headers. It will **not** process data or generate charts. This flag makes `--excel-path` optional for template creation, where it dictates the output file path.
      * Example: `--create-template --excel-path "New_MOVE_Template.xlsx"`
  * `--snapshot-date <date>`:
      * **Required (unless `--create-template` is used).** Specifies the single date (`YYYY-MM-DD`) for which the `Progress_Log` should be processed and charts generated. This allows for focused testing and analysis at specific points in time.
      * Example: `--snapshot-date "2025-08-15"`
  * `--log-level <level>`:
      * **Optional.** Sets the logging verbosity.
      * **Choices:** `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`.
      * **Default:** `INFO`
      * Example: `--log-level DEBUG`
  * `--no-chart-insertion`:
      * **Optional.** A boolean flag. If present, the script will process data but skip inserting charts into the Excel file. Useful for debugging data generation without chart overhead.
      * Example: `--no-chart-insertion`
  * `--no-data-update`:
      * **Optional.** A boolean flag. If present, the script will generate charts but skip updating the `Progress_Log` sheet in the Excel file.

**Example Usage:**

  * To create a template: `uv run move_tracker_report.py --create-template --excel-path "MyProjectTemplate.xlsx"`
  * To process data and generate reports for a specific date: `uv run move_tracker_report.py --excel-path "MyProject.xlsx" --snapshot-date "2025-08-15" --log-level INFO`

-----

## 4\. Input Excel File Structure

The script expects a single Excel workbook (`.xlsx`) with the following sheets and data structures:

### Sheet: `Instructions`

  * **Purpose:** Placeholder for user-facing instructions on how to use the template and manage data. The script will ignore this sheet during data processing but will create it for the template.
  * **Content (for template creation):** Basic guidance on populating `Historic_Work_Items`, `Current_Work_Items`, and `MOVE_Configuration`.

### Sheet: `Historic_Work_Items`

  * **Purpose:** Contains past project work item data for calculating historical flow time.
  * **Columns:**
      * `Historical_WI_ID` (String): Unique ID for the historical work item.
      * `Actual_Start_Date` (Datetime): Date and time the work item actually started.
      * `Actual_Completion_Date` (Datetime): Date and time the work item actually completed.
      * `Flow_Time_Days` (Number): Calculated column (Excel formula or Python calculated) representing `INT(Actual_Completion_Date) - INT(Actual_Start_Date) + 1`. This column should be present, but the script can optionally re-calculate if missing or incorrect.
  * **Data Range:** Assumed to start from A1 with headers.

### Sheet: `Current_Work_Items`

  * **Purpose:** Contains the specific work items for the current MOVE project.
  * **Columns:**
      * `Work_Item_ID` (String): Unique ID for the current work item.
      * `Description` (String): Description of the work item.
      * `Status` (String): Current status (e.g., "Not Started", "In Progress", "Completed"). Only "Completed" items will count towards `Actual_Work_Completed`.
      * `Actual_Start_Date` (Datetime, Optional): Date and time the work item actually became part of the scope / started.
      * `Actual_Completion_Date` (Datetime, Optional): Date and time the work item actually completed. This is the crucial field for `Actual_Work_Completed` count.
      * `Date_Withdrawn` (Datetime, Optional): Date and time the work item was formally removed from the project scope. If this column has a date, the item is considered withdrawn from that date onwards. Users should be instructed *not* to delete work items but to mark them as withdrawn.
  * **Data Range:** Assumed to start from A1 with headers.

### Sheet: `MOVE_Configuration`

  * **Purpose:** Contains key parameters for the current MOVE project.
  * **Structure:** A simple key-value table.
  * **Columns:**
      * `Parameter` (String): Name of the configuration parameter.
      * `Value` (Dynamic): The value for the parameter.
  * **Expected Parameters:**
      * `Planned_Start_Date`: (Date, `YYYY-MM-DD`): The official start date of the current MOVE.
      * `Planned_Delivery_Date`: (Date, `YYYY-MM-DD`): The initial planned delivery date for the entire MOVE.
      * `Historic_50th_Percentile_Flow_Time_Override`: (Number, Optional): An optional override for the 50th percentile historical flow time. If provided, use this. If not, calculate from `Historic_Work_Items`.
      * `Buffer_Green_Date`: (Date, `YYYY-MM-DD`): Date defining the boundary for the "Green" buffer zone.
      * `Buffer_Yellow_Date`: (Date, `YYYY-MM-DD`): Date defining the boundary for the "Yellow" buffer zone.
      * `Buffer_Red_Date`: (Date, `YYYY-MM-DD`): Date defining the boundary for the "Red" buffer zone.
      * `Fever_Green_Yellow_Left_Y` (Float, 0.0-1.0): Y-axis percentage (Buffer Consumed) where the Green-Yellow zone boundary line intersects the left (Work Done 0%) axis of the Fever Chart.
      * `Fever_Green_Yellow_Right_Y` (Float, 0.0-1.0): Y-axis percentage (Buffer Consumed) where the Green-Yellow zone boundary line intersects the right (Work Done 100%) axis of the Fever Chart.
      * `Fever_Yellow_Red_Left_Y` (Float, 0.0-1.0): Y-axis percentage (Buffer Consumed) where the Yellow-Red zone boundary line intersects the left (Work Done 0%) axis of the Fever Chart.
      * `Fever_Yellow_Red_Right_Y` (Float, 0.0-1.0): Y-axis percentage (Buffer Consumed) where the Yellow-Red zone boundary line intersects the right (Work Done 100%) axis of the Fever Chart.
  * **Data Range:** Assumed to start from A1 with headers.

-----

## 5\. Output Excel File Changes

The script will modify the *same* Excel workbook provided via `--excel-path`.

### Sheet: `Progress_Log`

  * **Purpose:** To store the calculated time-series data of project progress, reflecting changes in scope and completion over time. This sheet will serve as a comprehensive historical record of the project's measured progress.
  * **Action:** The script will **recalculate the entire `Progress_Log` from the `Planned_Start_Date` up to the specified `--snapshot-date`**. This recalculated historical data will then **overwrite the entire `Progress_Log` sheet** (excluding headers) in the Excel workbook. This ensures the log is always consistent with the latest `Current_Work_Items` and `MOVE_Configuration` up to the current processing date.
  * **Columns:**
      * `Snapshot_Date` (Datetime): The date the snapshot was taken.
      * `Scope_At_Snapshot` (Number): The dynamic scope at the `Snapshot_Date`. This is calculated by counting `Current_Work_Items` where `Actual_Start_Date` \<= `Snapshot_Date` AND (`Date_Withdrawn` is NULL or `Date_Withdrawn` \> `Snapshot_Date`).
      * `Actual_Work_Completed` (Number): Count of work items completed where `Status` is "Completed" AND `Actual_Completion_Date` \<= `Snapshot_Date`.
      * `Elapsed_Time_Days` (Number): Days since `Planned_Start_Date` to `Snapshot_Date`.
      * `Actual_Operational_Throughput` (Number): `Actual_Work_Completed` / `Elapsed_Time_Days`.
      * `Current_50th_Percentile_Flow_Time` (Number): Calculated 50th percentile flow time based on **completed work items up to `Snapshot_Date`** from `Current_Work_Items`. If insufficient completed data, use `Historic_50th_Percentile_Flow_Time` or a default.
      * `Forecasted_Delivery_Date` (Datetime): `Snapshot_Date` + (`Remaining_Work` / `Actual_Operational_Throughput`). `Remaining_Work` is `Scope_At_Snapshot` - `Actual_Work_Completed`.
      * `Buffer_Consumption_Percentage` (Float, 0.0 to 1.0+): Percentage of buffer consumed. Calculated as `(Forecasted_Delivery_Date - Planned_Delivery_Date) / (Buffer_Red_Date - Planned_Delivery_Date)`. Capped at 0 if ahead of schedule.
      * `Work_Done_Percentage` (Float, 0.0 to 1.0): `Actual_Work_Completed` / `Scope_At_Snapshot`.
      * `Current_Buffer_Signal` (String): "Green", "Yellow", "Red", "Beyond Red" based on `Forecasted_Delivery_Date` relative to buffer dates.
  * **Data Range:** Will start from A1 with headers.

### Sheets: `Work_Execution_Chart` and `Fever_Chart`

  * **Purpose:** To display the generated charts as images.
  * **Action:** The script will **create these sheets if they don't exist**. It will **clear any existing images** in these sheets and **insert the newly generated SVG images** of the charts.
  * **Placement:** Charts should be placed starting from cell A1 or a reasonable offset (e.g., A2, B2) to allow for potential titles or notes above them.

-----

## 6\. Core Logic / Workflow

**Test-Driven Development (TDD) Recommendation:** The script should be developed following a TDD approach. Tests for core calculations, data parsing, and chart element presence (e.g., specific text in SVG output) should be written *before* implementing the features. This ensures robustness and maintainability.

1.  **Initialization:**

      * Set up logging based on `--log-level`.
      * Parse CLI arguments using `typer`.
      * Initialize `rich` console for output.

2.  **Template Creation Logic:**

      * If `--create-template` is present:
          * Implement `_create_excel_template` function.
          * Create `openpyxl.Workbook`.
          * Add/rename sheets: `Instructions`, `MOVE_Configuration`, `Historic_Work_Items`, `Current_Work_Items`, `Progress_Log`, `Work_Execution_Chart`, `Fever_Chart`.
          * Populate headers for each data sheet as per Section 4.
          * Add descriptive text to `Instructions` sheet.
          * Add sample configuration parameters and values to `MOVE_Configuration`.
          * Save the workbook to the specified `excel-path` (or default if not provided).
          * Exit successfully.

3.  **Read Input Data & Configuration:**

      * Implement `_read_excel_data` function.
      * Load `MOVE_Configuration` into a `MOVEConfiguration` dataclass instance (parsing dates and numbers).
      * Load `Historic_Work_Items` into a Pandas DataFrame.
      * Load `Current_Work_Items` into a Pandas DataFrame, ensuring date columns are `datetime` objects.
      * **Validation:** Check for missing sheets, critical configuration parameters, and required columns in data sheets. Log warnings/errors and exit on critical failures.

4.  **Calculate Derived Configuration & Flow Times:**

      * Determine `Historic_50th_Percentile_Flow_Time`:
          * If `Historic_50th_Percentile_Flow_Time_Override` is provided, use it.
          * Otherwise, calculate the 50th percentile (median) of `Flow_Time_Days` from `Historic_Work_Items`. Log the calculated value.
      * **Calculate buffer dates/boundaries for Work Execution Chart:**
          * `Initial_Scope` (at `Planned_Start_Date`) should be derived from `Current_Work_Items` (count items where `Actual_Start_Date` \<= `Planned_Start_Date` AND (`Date_Withdrawn` is NULL or `Date_Withdrawn` \> `Planned_Start_Date`)).
          * `Initial_Ideal_Completion_Flow_Time = Initial_Scope * Historic_50th_Percentile_Flow_Time`
          * `Green_Buffer_Start_Date = Planned_Start_Date + Initial_Ideal_Completion_Flow_Time`
          * `Yellow_Buffer_Start_Date = Green_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)`
          * `Red_Buffer_Start_Date = Yellow_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)`
          * `Beyond_Red_Date = Red_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)`

5.  **Generate `Progress_Log` Data:**

      * Implement `_generate_full_progress_log` function. This function will generate the historical `Progress_Log` from `Planned_Start_Date` up to the CLI `--snapshot-date`.
      * It will iterate through dates from `Planned_Start_Date` up to the `--snapshot-date`, typically in increments of 1 day, or a frequency that captures all relevant changes (e.g., dates of `Actual_Start_Date`, `Actual_Completion_Date`, `Date_Withdrawn` from `Current_Work_Items`, plus weekly intervals). For each `current_date_in_loop` in this range:
          * **`Scope_At_Snapshot` calculation:** Count `Current_Work_Items` where `Actual_Start_Date` \<= `current_date_in_loop` AND (`Date_Withdrawn` is NULL or `Date_Withdrawn` \> `current_date_in_loop`).
          * `Actual_Work_Completed` calculation: Count `Current_Work_Items` where `Status` is "Completed" AND `Actual_Completion_Date` \<= `current_date_in_loop`.
          * `Current_50th_Percentile_Flow_Time` calculation:
              * Filter `Current_Work_Items` for items `Completed` by `current_date_in_loop` that started on or before `current_date_in_loop`.
              * Calculate flow time for these items.
              * Compute the 50th percentile. If less than 2 completed items, fall back to `Historic_50th_Percentile_Flow_Time` or a reasonable default (e.g., 5 days).
          * All other calculations (Elapsed\_Time\_Days, Actual\_Operational\_Throughput, Forecasted\_Delivery\_Date, Buffer\_Consumption\_Percentage, Work\_Done\_Percentage, Current\_Buffer\_Signal) will use `current_date_in_loop` as the reference point and the dynamically calculated `Scope_At_Snapshot`.
      * The function will return a DataFrame containing all generated log entries, sorted by `Snapshot_Date`.

6.  **Update Excel with `Progress_Log` Data:**

      * If `--no-data-update` is not set:
          * Implement `_update_progress_log_sheet` function.
          * Open the Excel workbook using `openpyxl`.
          * Select/create the `Progress_Log` sheet.
          * **Clear the entire `Progress_Log` sheet's data** (excluding the header row).
          * Write the **full, newly generated `Progress_Log` DataFrame** (from `_generate_full_progress_log`) back to the sheet.
          * Save the workbook.

7.  **Generate Charts (Matplotlib, SVG Output):**

      * If `--no-chart-insertion` is not set:
          * Implement `_generate_charts` function.
          * The **entire `Progress_Log` DataFrame** generated in step 5 will be used for plotting, ensuring the charts accurately display the historical progression up to the `--snapshot-date`.
          * **Work Execution Chart:**
              * **Title:** "Work Execution Signal Chart as at {snapshot\_date}"
              * Plot all `Actual_Work_Completed` vs. `Snapshot_Date` from the `Progress_Log` DataFrame (line with markers).
              * Plot all `Scope_At_Snapshot` vs. `Snapshot_Date` from the `Progress_Log` DataFrame (line, as it can change over time).
              * Plot a **linear trendline** for `Actual_Work_Completed` from the full `Progress_Log` data, projecting forward from the *last `Snapshot_Date` in the log* (which is the `--snapshot-date` from CLI).
              * **Extend the `Scope_At_Snapshot` line:** The `Scope_At_Snapshot` value corresponding to the *last `Snapshot_Date` in the `Progress_Log`* (i.e., the `--snapshot-date` from CLI) should be extended as a horizontal dashed line from that date to the end of the chart's X-axis range.
              * Add **vertical lines** at `Planned_Delivery_Date`, `Green_Buffer_Start_Date`, `Yellow_Buffer_Start_Date`, `Red_Buffer_Start_Date`, using the dynamically calculated buffer dates.
              * **Visual Intersection Indicator:** Clearly mark the intersection point of the **projected linear trendline** and the **extended last known `Scope_At_Snapshot` line**. Annotate this point with the `Forecasted_Delivery_Date` from the `--snapshot-date`'s `Progress_Log` entry.
              * Set X-axis range from `Planned_Start_Date` to `Beyond_Red_Date` (from calculated buffer dates).
              * Set Y-axis range from 0 to slightly above the maximum `Scope_At_Snapshot` **observed within the entire `Progress_Log` data** (i.e., up to `snapshot_date`).
              * Save the chart as `YYYY-MM-DD_work_execution_chart.svg` (where `YYYY-MM-DD` is the `--snapshot-date`).
          * **Fever Chart:**
              * **Title:** "Fever Chart as at {snapshot\_date}"
              * Plot all `Work_Done_Percentage` vs. `Buffer_Consumption_Percentage` from the `Progress_Log` DataFrame (line with markers). This shows the historical path on the fever chart.
              * **Implement background colored zones using `ax.fill_between` or `ax.fill` with polygons.** These zones will be defined by two sloping lines, derived from `Fever_Green_Yellow_Left_Y`, `Fever_Green_Yellow_Right_Y`, `Fever_Yellow_Red_Left_Y`, `Fever_Yellow_Red_Right_Y` configuration parameters. The Green zone will be below the first line, Yellow between the two, and Red above the second.
              * **Highlight Current Point:** Add a distinct, larger marker or annotation for the data point corresponding to the `--snapshot-date` on the Fever Chart.
              * Set X-axis range from 0% to 100% (`Work_Done_Percentage`).
              * Set Y-axis range from 0% to a reasonable maximum (e.g., `max(1.0, progress_log_df['Buffer_Consumption_Percentage'].max() * 1.2)`).
              * Save the chart as `YYYY-MM-DD_fever_chart.svg` (where `YYYY-MM-DD` is the `--snapshot-date`).

8.  **Insert Charts into Excel:**

      * If `--no-chart-insertion` is not set:
          * Implement `_insert_charts_into_excel` function.
          * Open the Excel workbook.
          * For `Work_Execution_Chart` sheet, clear existing images and insert the generated SVG.
          * For `Fever_Chart` sheet, clear existing images and insert the generated SVG.
          * Save the workbook.

9.  **Cleanup:** Remove temporary chart SVG files (`.svg`).

-----

## 7\. Data Structures

  * **`enum` for `BufferSignal`:**
    ```python
    from enum import Enum
    class BufferSignal(Enum):
        GREEN = "Green"
        YELLOW = "Yellow"
        RED = "Red"
        BEYOND_RED = "Beyond Red"
    ```
  * **`dataclass` for `MOVEConfiguration`:**
    ```python
    from dataclasses import dataclass
    from datetime import date

    @dataclass
    class MOVEConfiguration:
        planned_start_date: date
        planned_delivery_date: date
        # Target_Scope now derived from Current_Work_Items and dynamic
        historic_50th_percentile_flow_time_override: float = None
        buffer_green_date: date
        buffer_yellow_date: date
        buffer_red_date: date
        # Snapshot_Frequency_Days removed from config
        fever_green_yellow_left_y: float
        fever_green_yellow_right_y: float
        fever_yellow_red_left_y: float
        fever_yellow_red_right_y: float
    ```
    (Dates will be parsed from strings in Excel to `datetime.date` objects, percentages for fever chart boundaries are floats 0.0-1.0).

-----

## 8\. Logging

  * Utilize Python's built-in `logging` module.
  * Messages should be logged to the console (standard output).
  * Log levels will correspond to the `--log-level` argument.
      * `INFO`: General progress messages (e.g., "Reading data...", "Charts generated...").
      * `DEBUG`: Detailed information (e.g., values of calculated parameters, DataFrame heads).
      * `WARNING`: Non-critical issues (e.g., optional config missing).
      * `ERROR`: Critical failures (e.g., file not found, sheet missing, invalid data).

-----

## 9\. Rich Output

  * Use the `rich` library for aesthetically pleasing console output.
  * Display status messages, progress bars (if processing large data), and formatted tables for key information (e.g., configuration summary, final `Progress_Log` entry) using `rich.console.Console` and `rich.table.Table`.
  * Error messages should also be formatted clearly with `rich`.

-----

## 10\. Error Handling

  * Implement `try-except` blocks for file operations, data parsing, and critical calculations.
  * Provide informative error messages to the user using `rich.console.Console.log` or `rich.console.Console.print` (with color/style) and `logging.error`).
  * Graceful exit on critical errors (e.g., missing Excel file, invalid configuration values, `snapshot-date` before `Planned_Start_Date`).

-----