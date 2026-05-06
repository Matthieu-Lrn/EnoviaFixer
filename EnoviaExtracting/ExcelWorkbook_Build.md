# Excel Workbook Setup

This is the recommended hardcoded workbook route for Langflow.

## Workbook path

Create this workbook on the CATIA/Langflow machine:

```text
\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\LangflowEnoviaExtraction.xlsm
```

The Python component already opens that exact path:

- [lf0_excel_vba_extraction_runner_component.py](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/lf0_excel_vba_extraction_runner_component.py)

## Recommended setup

Keep one shared workbook on the server at that hardcoded path.

Langflow does not build or rebuild the workbook at runtime. It simply:

1. opens `LangflowEnoviaExtraction.xlsm`
2. writes the current inputs into the `Inputs` sheet
3. runs the `RunFullExtraction` macro

That means the workbook should already exist and already contain the VBA modules listed below.
The `Inputs` sheet does not need to be created manually. Langflow will create or refresh it at runtime if needed.

## What to put in the workbook

For the Excel-only route, generate the merged import bundle:

```powershell
python .\EnoviaExtracting\build_excel_import_bundle.py
```

Then import everything from:

```text
EnoviaExtracting\excel_vba_merged\
```

At minimum, the core workbook modules are:

1. [ExcelCatiaBootstrap.bas](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/ExcelCatiaBootstrap.bas)
2. [ExcelCatiaBridge.bas](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/ExcelCatiaBridge.bas)
3. [EnoviaSearching.bas](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/EnoviaSearching.bas)
4. [ExcelLangflowSteps.bas](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/ExcelLangflowSteps.bas)
5. [ExcelOrchestrator.bas](/c:/Users/mlaurin/Desktop/EnoviaFixer/EnoviaExtracting/ExcelOrchestrator.bas)

## Inputs sheet

Langflow uses a worksheet named `Inputs`. If it is missing, the runtime creates it automatically.
These cells are used:

- `B1` top assembly
- `B2` revision
- `B3` project number
- `B4` export path
- `B5` sync from BSF (`TRUE` or `FALSE`)
- `B6` CI option
- `B7` Non-CI option

The orchestrator entrypoint is:

```vb
RunFullExtraction
```

Column `A` labels are written automatically for readability. The runtime writes the current run values into column `B`.

## What the workbook actually runs

The workbook contains the ENOVIA search code directly and can now host the PVRSync and SaveFileBase logic directly too.

The generated `excel_vba_merged` bundle:

- keeps the PVRSync side with its original names
- adds the SaveFileBase side with `SFB_` prefixes so it can coexist in one Excel VBA project
- includes a few helper classes/forms sourced from [extracted_vba](/c:/Users/mlaurin/Desktop/EnoviaFixer/extracted_vba) that were missing from the trimmed `vba_source`

## Runtime flow

1. Langflow runs `LF0 Excel VBA Extraction Runner`
2. Python opens the hardcoded `.xlsm`
3. Python writes the Inputs sheet values
4. Excel runs `RunFullExtraction`
5. Excel connects to CATIA
6. Excel opens the top assembly through ENOVIA
7. Excel runs the in-workbook PVRSync code
8. Excel runs the in-workbook SaveFileBase export code

## Current assumptions

- CATIA is already open
- ENOVIA login is already active in that CATIA session
- `LangflowEnoviaExtraction.xlsm` already exists at the hardcoded server path
- that workbook already contains the four Excel VBA modules listed above
- the WebService / BDI environment that the original macros relied on is still available on that machine
- Excel is installed on the machine running the Langflow component
