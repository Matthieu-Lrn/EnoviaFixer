Attribute VB_Name = "ExcelLangflowSteps"
Option Explicit

Public Sub RunTopAssySearch(ByVal topAssy As String, ByVal expectedRevision As String)
    Call EnsureCatiaSession
    Call OpenTopAssemblyFromEnovia(topAssy, expectedRevision)
End Sub

Public Sub RunPVRSyncStep(ByVal projectNumber As String, ByVal syncFromBSF As String, ByVal ciOption As String, ByVal nonCiOption As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Call SyncActivePVR(projectNumber, syncFromBSF, ciOption, nonCiOption, kbePathFile, toolbarPath)
End Sub

Public Sub RunPVRExportStep(ByVal exportFolder As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Call ExportActivePVR(exportFolder, kbePathFile, toolbarPath)
End Sub
