# PowerPoint Report Automation

A Streamlit application that creates and updates PowerPoint reports from Q-style crosstab Excel files. It generates branded charts, data tables, slide titles, question/base text, row-based callouts, and AI-generated insights — with an updater that refreshes those objects from new data without breaking formatting.

## Features

- **Excel Parsing**: Parses Q-style crosstab Excel workbooks, detecting table blocks, banners, metrics, and footnotes
- **PowerPoint Export**: Creates new presentations with branded charts, tables, titles, and annotations
- **Deck Updates**: Updates existing presentations with new data while preserving formatting
- **AI Insights**: Generates per-table summaries via the OpenAI API and compiles an Executive Summary slide
- **Brand System**: Centralized brand configuration (colors, fonts, palettes) with automatic palette rotation across slides
- **Column Selection**: Choose which data column (e.g., "Total", "2024", a specific banner segment) to visualize per table
- **Row-Based Callouts**: Create callouts tied to specific rows with customizable text and a `[Value]` placeholder
- **Toggle-Based Callout UI**: Clean toggle controls for managing callouts per table, with persistence and "Previously:" labels for existing callouts
- **Flexible Mapping**: Alt-text-based mapping system with both automatic (new reports) and manual (existing reports) workflows
- **Multiple Chart Types**: Horizontal Bar, Vertical Bar, Stacked Bar, Donut, Line, Pie, and Chart+Table combos

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Create a `.env` file in the project root with your OpenAI API key:
   ```
   OPENAI_API_KEY=sk-your-key-here
   ```

## Usage

### Running the App

```bash
streamlit run app.py
```

### Workflow 1: Create a New Report

1. Upload a crosstab Excel file
2. Configure each table: chart type, data column, metric, title, base text, question, row sorting, and callouts
3. Export a new `.pptx` — all shapes are automatically tagged with mapping metadata
4. If AI insights are enabled, an Executive Summary slide is generated and per-table insight text boxes are added

### Workflow 2: Update an Existing Report

1. Upload an existing `.pptx` (with mapping metadata in alt text)
2. Review detected connections (charts, tables, text, callouts)
3. Upload a new crosstab Excel file
4. The tool refreshes all mapped shapes with fresh data, preserving formatting
5. Unmapped tables are listed on a summary slide

### AI Insights

When enabled, the app calls the OpenAI API (`gpt-4o-mini`) to generate a short, factual 2-3 sentence summary for each crosstab table. These summaries are:

- Displayed as a text strip at the bottom of each chart slide
- Aggregated into a branded Executive Summary slide at the front of the deck

Insights require a valid `OPENAI_API_KEY` in your `.env` file. If the key is missing or the API call fails, the app continues without insights.

```python
from ai_insights import generate_all_summaries

summaries = generate_all_summaries(tables, column_key="Total", selections=selections)
```

### Column Selection

Each table has a "Data column" dropdown. The selected column drives:
- Chart data
- Base N values
- Callout values
- Table data display

### Row-Based Callouts

1. **Enable**: Check "Enable callouts for this table"
2. **Select row and column**: Pick which data point the callout references
3. **Customize text**: Edit the text box; use `[Value]` as a placeholder for the actual data value
4. **Examples**:
   - `Gen Z: [Value]` renders as `Gen Z: 20.2%`
   - `Millennials represent [Value] of respondents` renders as `Millennials represent 15.2% of respondents`

When updating an existing deck, previously set callouts display a "Previously:" label showing the old text for context.

## Brand Configuration

All colors, fonts, and sizing are defined in `brand_config.py` — derived from the Future Partners 2023 brand guidelines. No branding values are hardcoded elsewhere.

- **4 infographic palettes** (Blue, Green, Purple, Red) with 5 tones each
- **Automatic palette rotation** across slides for visual variety
- **Typography**: Headline (Arial Black), Body (Arial), Narrative (Times New Roman) as system fallbacks for GT America
- **Slide styling**: Background color, gridlines, text colors, chart defaults, executive summary layout, and AI insight text box positioning

To customize branding, edit `brand_config.py`. All modules import from it.

## Mapping System

Shapes are connected to crosstab data via alt-text metadata stored on each PowerPoint shape. When a report is created, `pptx_exporter.py` tags each shape automatically. When updating, `deck_update.py` reads these tags to match shapes to the correct table and column.

### Automatic Mapping (New Reports)

Shapes are tagged automatically during export with alt text like:

```
type=chart; table_title=Q Age; column=Total; exclude_rows=NET
```

### Manual Mapping (Existing Reports)

For existing PowerPoint files, use `mapping_helper.py`:

```bash
# List all shapes and their mapping status
python mapping_helper.py list presentation.pptx

# Generate an editable mapping template
python mapping_helper.py template presentation.pptx crosstab.xlsx

# Apply mappings from a template file
python mapping_helper.py apply presentation.pptx mapping_template.py

# Validate mappings against crosstab data
python mapping_helper.py validate presentation.pptx crosstab.xlsx
```

You can also map shapes directly in PowerPoint: right-click a shape, open **Format Shape > Alt Text**, and add mapping fields in the Description box:

```
type: chart
table_title: Q Age
column: Total
exclude_rows: base, mean, average, avg
auto_update: yes
```

### Mapping Options

| Shape Type | Fields |
|---|---|
| **Chart** | `type: chart`, `table_title`, `column`, `exclude_rows`, `auto_update` |
| **Table** | `type: table`, `table_title`, `columns` (`*` for all), `exclude_rows`, `auto_update` |
| **Question text** | `type: text_question`, `auto_update` |
| **Base text** | `type: text_base`, `auto_update` |
| **Title text** | `type: text_title`, `auto_update` |
| **Callout** | `type: text_callout`, `table_title`, `row_label`, `column_key`, `auto_update` |
| **AI Insight** | `type: ai_insight`, `table_title`, `auto_update` |

Set `auto_update: no` on any shape to skip it during updates.

## File Structure

```
app.py                 — Streamlit UI (main entry point)
crosstab_parser.py     — Excel crosstab parser
pptx_exporter.py       — PowerPoint creation and export
deck_update.py         — PowerPoint update logic
mapping_helper.py      — CLI for mapping management
ai_insights.py         — AI-powered table summaries (OpenAI API)
brand_config.py        — Centralized brand colors, fonts, and layout
requirements.txt       — Python dependencies
.env                   — API keys (not tracked in git)
README.md              — This file

templates/
  Template_ReportSlides.pptx   — Master slide template
  chart_templates/             — Brand chart style templates (.crtx)

sample_data/
  Test Crosstab - Smaller - V1.xlsx  — Sample crosstab for testing
  Test Crosstab - Smaller - V2.xlsx  — Alternate sample for testing updates
  SATS/                              — State of the American Traveler data

docs/
  FuturePartners_BrandGuidelines2023_ReducedSize.pdf — Brand reference
```

## Troubleshooting

- **Font Not Found**: The app uses Arial/Arial Black as system fallbacks for the brand fonts
- **AI Insights Not Generating**: Ensure `OPENAI_API_KEY` is set in `.env` and the `openai` package is installed
- **Mapping Not Working**: Run `python mapping_helper.py validate` to diagnose issues
- **Shape Names**: Use `python mapping_helper.py list` to see all shapes and their mapping status
- **Console Logging**: Check terminal output during updates for detailed diagnostics
