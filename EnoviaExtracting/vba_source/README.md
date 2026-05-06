# Langflow CATIA Extraction VBA Source

These folders are the source packages behind the Excel-driven Langflow extraction flow.

Target server folder:

```text
\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\
```

The repo now includes an Excel import bundle generator:

```powershell
python .\EnoviaExtracting\build_excel_import_bundle.py
```

That creates:

```text
EnoviaExtracting\excel_vba_merged\
```

The generated folder contains:

- core Excel modules
- the PVRSync VBA source
- a prefixed SaveFileBase VBA copy that can coexist in the same Excel project
- a few helper classes/forms pulled from `extracted_vba` where the trimmed `vba_source` was missing them

Excel workbook route:

- The hardcoded Langflow runner opens `LangflowEnoviaExtraction.xlsm`.
- That workbook lives on the shared server path and is expected to already exist before runtime.
- That workbook should import the generated `excel_vba_merged` bundle.
- Langflow creates the `Inputs` sheet if needed, writes the current inputs into it, then runs `RunFullExtraction`.
- The workbook then runs `RunFullExtraction`, which:
  1. connects to the running CATIA session,
  2. opens the top assembly through ENOVIA search,
  3. runs the in-workbook PVRSync VBA,
  4. runs the in-workbook SaveFileBase VBA.

Notes:

- The source was trimmed from the original CATVBA exports but still keeps the needed shared modules for BDI/ENOVIA metadata, PVR sync, and SaveFileBase export.
- `LangflowPVRSync` bypasses the first user form by using parameters from Langflow.
- To keep the run fully unattended, use `ciOption="DisplayNever"` or `ciOption="DisplayNonRel"`. `DisplayAlways` can still trigger the old effective-document selection form deeper in the original PVR sync logic.
- `LangflowSaveFileBase` bypasses the folder picker by setting `sLangflowDestinationFolder` before calling `mdlSave3DStructure.StartProcess`.
- Do not import the raw `LangflowPVRSync\` and `LangflowSaveFileBase\` folders side by side into one Excel VBA project. Use the generated `excel_vba_merged` bundle instead.
