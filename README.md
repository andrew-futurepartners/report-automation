# PowerPoint Report Automation

A Streamlit application that exports new PowerPoint reports from crosstab Excel files and updates existing decks with new data. It creates charts, optional data tables, slide titles, question text, and base text, with an updater that refreshes those objects without breaking formatting.

## Features

- **Excel Parsing**: Intelligently parses Q-style crosstab Excel files
- **PowerPoint Export**: Creates new presentations with branded charts and tables
- **Deck Updates**: Updates existing presentations while preserving formatting
- **Flexible Mapping**: Both automatic and manual mapping systems
- **Multiple Chart Types**: Bar, Donut, Line, and Chart+Table options

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

1. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

2. Upload a crosstab Excel file
3. Choose between exporting a new PowerPoint or updating an existing one
4. Configure chart types, titles, and base text for each table
5. Export or update your presentation

### Enhanced Mapping System

The application now supports both automatic and manual mapping, making it easy to retroactively map existing PowerPoint files.

#### Automatic Mapping (New Reports)

When you create new reports through the tool, shapes are automatically tagged with:
- **Alt Text**: Detailed mapping information stored in PowerPoint's alternative text field
- **XML Integration**: Uses PowerPoint's native XML structure for reliable alt text storage

#### Manual Mapping (Existing Reports)

For existing PowerPoint files, you can manually create mappings using the `mapping_helper.py` script:

##### Step 1: Generate a Mapping Template

```bash
python mapping_helper.py template your_presentation.pptx your_crosstab.xlsx
```

This creates a `mapping_template.py` file that looks like:

```python
# PowerPoint Mapping Template

# Instructions:
# 1. Edit the mapping values below
# 2. Save this file
# 3. Use the apply_mapping function to update your PowerPoint

# Available crosstab tables:
# - Q Age (columns: 18-24, 25-34, 35-44, 45-54, 55+, Total)
# - Q Gender (columns: Male, Female, Total)

MAPPINGS = {
    # Slide 1, Shape 1: Chart1
    'Chart1': {
        'type': 'chart',
        'table_title': 'Q Age',
        'column': 'Total',  # or specific column name
        'exclude_rows': 'base, mean, average, avg',
        'auto_update': 'yes'
    },

    # Slide 1, Shape 2: Table1
    'Table1': {
        'type': 'table', 
        'table_title': 'Q Age',
        'columns': '*',
        'exclude_rows': 'base, mean, average, avg',
        'auto_update': 'yes'
    },
}
```

##### Step 2: Edit the Mapping

Modify the `table_title` and `column` values to match your crosstab data. The `table_title` should exactly match one of the table titles from your Excel file.

##### Step 3: Apply the Mapping

```bash
python mapping_helper.py apply your_presentation.pptx mapping_template.py
```

This updates your PowerPoint file with the new mappings (saved as `your_presentation_mapped.pptx`).

##### Step 4: Validate Mappings

```bash
python mapping_helper.py validate your_presentation_mapped.pptx your_crosstab.xlsx
```

This checks that all your mappings are valid and points out any issues.

#### Manual Mapping in PowerPoint

You can also create mappings directly in PowerPoint:

1. **Right-click** on a chart or table shape
2. **Select "Format Shape"**
3. **Go to "Alt Text" tab**
4. **In the Description field, add your mapping:**

```
type: chart
table_title: Q Age
column: Total
exclude_rows: base, mean, average, avg
auto_update: yes
```

**Note**: The tool automatically sets alt text for new reports. For existing reports, you can manually add these mappings in PowerPoint, and the tool will read them during updates.

### Mapping Helper Commands

```bash
# List all shapes in a PowerPoint file
python mapping_helper.py list presentation.pptx

# Generate mapping template
python mapping_helper.py template presentation.pptx crosstab.xlsx

# Apply mappings from file
python mapping_helper.py apply presentation.pptx mapping_template.py

# Validate existing mappings
python mapping_helper.py validate presentation.pptx crosstab.xlsx
```

## Mapping Configuration Options

### For Charts
- `type`: Must be "chart"
- `table_title`: The exact title of the crosstab table
- `column`: Which column to chart (e.g., "Total", "Male", "18-24")
- `exclude_rows`: Rows to exclude (default: "base, mean, average, avg")
- `auto_update`: Set to "no" to skip automatic updates

### For Tables
- `type`: Must be "table"
- `table_title`: The exact title of the crosstab table
- `columns`: Which columns to include (use "*" for all)
- `exclude_rows`: Rows to exclude
- `auto_update`: Set to "no" to skip automatic updates

### For Text Objects
- `type`: "question_text" or "base_text"
- `auto_update`: Set to "no" to skip automatic updates

## File Structure

- `app.py` - Main Streamlit application
- `crosstab_parser.py` - Excel file parser
- `pptx_exporter.py` - PowerPoint creation and export
- `deck_update.py` - PowerPoint update functionality
- `mapping_helper.py` - Manual mapping assistance
- `requirements.txt` - Python dependencies

## Troubleshooting

### Common Issues

1. **Font Not Found**: The app now uses Arial as a fallback font
2. **Mapping Not Working**: Use the validation command to check for issues
3. **Shape Names**: Ensure shapes have descriptive names for easier mapping

### Debugging

- Use `python mapping_helper.py list` to see all shapes and their mapping status
- Check the console output during updates for detailed logging
- Validate mappings before running updates

## Advanced Features

### Custom Row Selection

You can specify which rows to include in charts by modifying the `exclude_rows` field:

```
exclude_rows: base, mean, average, avg, other_unwanted_rows
```

### Binding Question and Base Text

Charts can automatically update question and base text by setting:

```
bind_question: TEXT_QUESTION
bind_base: TEXT_BASE
```

### Skipping Updates

Set `auto_update: no` to prevent a shape from being automatically updated.

## Contributing

Feel free to submit issues and enhancement requests. The mapping system is designed to be extensible for additional mapping types and configurations.
