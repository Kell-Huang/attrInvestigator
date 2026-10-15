# Attribute Investigator

`attInvestigator.py` is a local command-line tool that analyzes a Productsup SlimEdit page saved as a `.txt` file and exports the attribute mapping to an Excel workbook.

It reads Import, Intermediate, and Export attributes, traces their relationships, identifies the associated data source, marks paused rules, and highlights standalone or unused attributes in the generated report.

## Features

- Parses Import, Intermediate, and optional Export sections from SlimEdit HTML.
- Keeps isolated Import, Intermediate, and Export attributes in the report.
- Maps Import → Intermediate → Export attribute dependencies.
- Identifies `Main Source`, `Additional` sources, and `Missing Source`.
- Appends `🚫Paused` to paused rules.
- Cleans HTML entities from attribute and source values.
- Generates a formatted `.xlsx` report.
- Uses the selected TXT filename as the report name.
- Adds a timestamp to reduce the risk of overwriting an existing report.

## Requirements

- Python 3.9 or later
- `openpyxl`

All other imported modules are part of the Python standard library.

## Project Structure

```text
attribute-investigator/
├── attInvestigator.py
├── requirements.txt
└── README.md
```

Place the SlimEdit `.txt` files that you want to analyze in the same working directory where you run the script.

## Installation

### 1. Open a terminal in the project directory

```bash
cd path/to/attribute-investigator
```

### 2. Create a virtual environment

```bash
python -m venv .venv
```

On Windows, if `python` is unavailable, use:

```powershell
py -m venv .venv
```

### 3. Activate the virtual environment

#### Windows Command Prompt

```bat
.venv\Scripts\activate
```

#### Windows PowerShell

```powershell
.\.venv\Scripts\Activate.ps1
```

#### macOS or Linux

```bash
source .venv/bin/activate
```

### 4. Install the dependency

```bash
python -m pip install -r requirements.txt
```

On Windows, you can also use:

```powershell
py -m pip install -r requirements.txt
```

## Preparing the Input File

1. Save the Productsup SlimEdit page HTML as a `.txt` file.
2. Put the `.txt` file in the directory from which the script will be executed.
3. You may place multiple `.txt` files in the directory. The program will list them and ask you to select one.

The input must contain the SlimEdit Import and Intermediate sections. The Export section is optional.

## Running the Tool

Run:

```bash
python attInvestigator.py
```

Or on Windows:

```powershell
py attInvestigator.py
```

The script displays the available `.txt` files:

```text
1 => Otto Product ChannelEngine Shared.txt
2 => Another Export.txt
Type the number of the slimedit page you want to process> 
```

Enter the number of the file you want to analyze.

## Output

The generated Excel file is saved in the current working directory.

The normal filename format is:

```text
<selected_txt_filename>_<YYYYMMDDHHMM>_mapping.xlsx
```

Example:

```text
Otto Product ChannelEngine Shared_202610151731_mapping.xlsx
```

If a usable report name cannot be created, the script falls back to a filename beginning with:

```text
import_intermediate_attribute_
```

## Excel Report Columns

When the SlimEdit page contains an Export section, the workbook contains:

```text
Import | Rule | Intermediate | Export Rule | Export | Data Source
```

When no Export section exists, it contains:

```text
Import | Rule | Intermediate | Data Source
```

The generated worksheet is named:

```text
all_attrs_inter_export
```

## Report Formatting

- Import and Export headers use green formatting.
- The Intermediate header uses blue formatting.
- Standalone Intermediate attributes are highlighted in purple.
- Import attributes without a matching Intermediate attribute are highlighted in purple.
- Attribute, source, and rule columns are formatted for readability.

## Troubleshooting

### `No .txt file found in the current directory.`

Make sure at least one `.txt` file is present in the directory where you run the script.

You can verify the current directory with:

```bash
pwd
```

On Windows Command Prompt:

```bat
cd
```

### `ModuleNotFoundError: No module named 'openpyxl'`

Activate the virtual environment and reinstall the dependency:

```bash
python -m pip install -r requirements.txt
```

### `HTML does not contain the import / intermediate sections`

The selected file does not contain the expected SlimEdit section markers:

```text
source-columns
intermediate-columns
```

Save the complete SlimEdit page content and try again.

### The wrong TXT file was selected

Run the program again and enter the number corresponding to the correct file.

### Existing report was overwritten

The standard timestamp has minute-level precision. Running the tool more than once within the same minute for the same input filename can create the same output filename and overwrite the previous report. Rename or move the first report before rerunning within the same minute.

### PowerShell blocks activation

Run this command for the current PowerShell session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate the environment again:

```powershell
.\.venv\Scripts\Activate.ps1
```
