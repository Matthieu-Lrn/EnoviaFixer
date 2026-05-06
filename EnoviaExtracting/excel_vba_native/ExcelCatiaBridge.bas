Attribute VB_Name = "ExcelCatiaBridge"
Option Explicit

Public Const DEFAULT_KBE_PATH_FILE As String = _
    "\\aero\mtlplm\catia\V5_KBE_Tools\Production\00_KBE_Env\06_KBE_CATScript\01_MACROS\02_BA_GCC\08_PATH_FILES\BA_COMMON_KBE_PATH.txt"

Public Sub SyncActivePVR(ByVal projectNumber As String, ByVal syncFromBSF As String, ByVal ciOption As String, ByVal nonCiOption As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Call EnsureCatiaSession

    If Len(Trim$(kbePathFile)) = 0 Then
        kbePathFile = DEFAULT_KBE_PATH_FILE
    End If

    Call BA_KBE_GCC_DDP.SyncActivePVR( _
        projectNumber, _
        syncFromBSF, _
        ciOption, _
        nonCiOption, _
        kbePathFile, _
        toolbarPath _
    )
End Sub

Public Sub ExportActivePVR(ByVal exportFolder As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Call EnsureCatiaSession

    If Len(Trim$(kbePathFile)) = 0 Then
        kbePathFile = DEFAULT_KBE_PATH_FILE
    End If

    Call ExcelNativeExtraction.ExportActivePVRNative(exportFolder)
End Sub
