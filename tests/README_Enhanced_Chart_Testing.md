# Enhanced Chart Visual Testing for MOVE Tracker

This document explains the enhanced chart visual testing functionality that automatically generates reference charts for all snapshot dates in the Progress_Log.

## Overview

The enhanced testing system provides comprehensive visual regression testing by:

1. **Automatically discovering all snapshot dates** from the Progress_Log sheet
2. **Generating reference charts** for each snapshot date
3. **Comparing generated charts** against reference images
4. **Detecting visual regressions** across the entire project timeline

## Files

- `test_chart_visuals_enhanced.py` - Enhanced test suite with comprehensive functionality
- `test_chart_visuals.py` - Original test suite (still functional)
- `reference_images/` - Directory containing reference chart images

## Key Functions

### `get_snapshot_dates_from_progress_log(excel_path)`
Extracts all unique snapshot dates from the Progress_Log sheet after running the MOVE tracker script.

**Returns:** List of date strings in YYYY-MM-DD format

### `generate_all_reference_charts(test_scenario_name="basic_scenario")`
Generates reference charts for all snapshot dates found in the Progress_Log.

**Process:**
1. Creates test Excel file with sample data
2. Runs MOVE tracker to generate Progress_Log
3. Extracts all snapshot dates
4. Generates Work Execution and Fever charts for each date
5. Saves charts as reference images

**Returns:** Boolean indicating success/failure

### `test_comprehensive_chart_visual_consistency()`
Comprehensive test that validates charts for ALL snapshot dates automatically.

**Process:**
1. Discovers all snapshot dates from Progress_Log
2. Generates charts for each date
3. Compares against reference images
4. Auto-generates missing references
5. Reports any visual mismatches

## Usage Examples

### 1. Generate All Reference Charts

```bash
# Using pytest (recommended)
pytest tests/test_chart_visuals_enhanced.py::test_generate_all_reference_charts -s

# Using the convenience script
python tmp_rovodev_generate_references.py

# Using Python directly
python -c "
import sys; sys.path.append('tests')
from test_chart_visuals_enhanced import generate_all_reference_charts
generate_all_reference_charts('basic_scenario')
"
```

### 2. Run Comprehensive Visual Tests

```bash
# Test all snapshot dates automatically
pytest tests/test_chart_visuals_enhanced.py::test_comprehensive_chart_visual_consistency -s

# Run with verbose output
pytest tests/test_chart_visuals_enhanced.py::test_comprehensive_chart_visual_consistency -s -v
```

### 3. Run Original Single-Date Tests

```bash
# Test specific dates (original functionality)
pytest tests/test_chart_visuals.py::test_chart_visual_consistency -s
```

## Test Scenarios

### Basic Scenario
The default test scenario includes:
- **Historic Work Items:** 3 completed items with flow times of 2, 4, and 7 days
- **Current Work Items:** 6 work items with various statuses and dates
- **Project Timeline:** January 1-21, 2025
- **Expected Snapshot Dates:** Multiple dates based on work item events

### Expected Snapshot Dates (Basic Scenario)
Based on the test fixture data, you should see charts for dates like:
- 2025-01-01 (Project start)
- 2025-01-05 (First completion)
- 2025-01-10 (Second completion)
- 2025-01-12 (Scope change)
- 2025-01-13 (Withdrawal)
- 2025-01-18 (Third completion)
- 2025-01-21 (Final snapshot)

## Reference Image Naming Convention

Reference images are saved with the following naming pattern:
```
{test_scenario_name}_{snapshot_date}_{chart_type}.png
```

Examples:
- `basic_scenario_2025-01-01_work_execution_chart.png`
- `basic_scenario_2025-01-01_fever_chart.png`
- `basic_scenario_2025-01-15_work_execution_chart.png`
- `basic_scenario_2025-01-15_fever_chart.png`

## Benefits

### 1. Comprehensive Coverage
- Tests charts for ALL dates in the project timeline
- Catches regressions that might only appear on specific dates
- Validates the entire progression of the project

### 2. Automated Discovery
- No need to manually specify test dates
- Automatically adapts to changes in test data
- Discovers edge cases and boundary conditions

### 3. Visual Regression Detection
- Pixel-perfect comparison of chart images
- Generates diff images for failed comparisons
- Configurable similarity thresholds

### 4. Easy Maintenance
- Auto-generates missing reference images
- Simple regeneration of all references
- Clear reporting of failures and successes

## Workflow

### Initial Setup
1. Run `test_generate_all_reference_charts` to create initial references
2. Review generated charts manually to ensure they look correct
3. Commit reference images to version control

### Development Workflow
1. Make changes to chart generation code
2. Run `test_comprehensive_chart_visual_consistency` 
3. Review any failures and diff images
4. Update references if changes are intentional
5. Commit updated references

### Continuous Integration
```bash
# In CI pipeline
pytest tests/test_chart_visuals_enhanced.py::test_comprehensive_chart_visual_consistency
```

## Troubleshooting

### No Snapshot Dates Found
- Ensure the Excel file has proper data in Current_Work_Items
- Check that the MOVE tracker script runs successfully
- Verify Progress_Log sheet is generated

### Chart Generation Failures
- Check that all required dependencies are installed (matplotlib, pandas, etc.)
- Verify the --save-charts-only option is working
- Review error messages in test output

### Visual Comparison Failures
- Review diff images in the test output directory
- Adjust comparison threshold if needed
- Regenerate references if changes are intentional

## Configuration

### Comparison Threshold
Adjust the similarity threshold in `compare_images()`:
```python
# More strict (default)
threshold=0.01  # 1% difference allowed

# More lenient
threshold=0.05  # 5% difference allowed
```

### Test Scenarios
Add new test scenarios by:
1. Creating new test fixture data in `conftest.py`
2. Adding new scenario names to the parametrized tests
3. Generating references for the new scenario

## Integration with Existing Tests

The enhanced testing system is designed to complement, not replace, the existing test infrastructure:

- **Original tests** (`test_chart_visuals.py`) still work for specific date testing
- **Enhanced tests** (`test_chart_visuals_enhanced.py`) provide comprehensive coverage
- **Both can be run together** for maximum confidence

This provides a robust, automated visual regression testing system that ensures chart quality across the entire project timeline.