# Langflow CATIA Extraction VBA Source

These folders are source packages for the CATVBA files launched by the Langflow components.

Target server folder:

```text
\\VADER\Apps\m170 - wp4\WP 4.2.1 Cabinet\09.   Monuments\36. MSB monument\14.Data Transfer\DATA TRANS 3.0\Temp-Matthieu\
```

Build these CATVBA files in CATIA V5:

```text
LangflowEnoviaSearch.catvba
  Import: LangflowEnoviaSearch\EnoviaSearching.bas
  Entry:  EnoviaSearching.OpenTopAssemblyFromEnovia(topAssy, expectedRevision)

LangflowPVRSync.catvba
  Import all modules/classes from: LangflowPVRSync\
  Entry: BA_KBE_GCC_DDP.SyncActivePVR(projectNumber, syncFromBSF, ciOption, nonCIOption, kbePathFile, toolbarPath)

LangflowSaveFileBase.catvba
  Import all modules/classes from: LangflowSaveFileBase\
  Entry: BA_KBE_GCC_SAVEFILEBASE.ExportActivePVR(exportFolder, kbePathFile, toolbarPath)
```

Notes:

- The source was trimmed from the original CATVBA exports but still keeps the needed shared modules for BDI/ENOVIA metadata, PVR sync, and SaveFileBase export.
- `LangflowPVRSync` bypasses the first user form by using parameters from Langflow.
- To keep the run fully unattended, use `ciOption="DisplayNever"` or `ciOption="DisplayNonRel"`. `DisplayAlways` can still trigger the old effective-document selection form deeper in the original PVR sync logic.
- `LangflowSaveFileBase` bypasses the folder picker by setting `sLangflowDestinationFolder` before calling `mdlSave3DStructure.StartProcess`.
