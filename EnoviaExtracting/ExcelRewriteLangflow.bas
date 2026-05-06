Attribute VB_Name = "ExcelRewriteLangflow"
Option Explicit

Public Sub RunTopAssySearch(ByVal topAssy As String, ByVal expectedRevision As String)
    Call EnsureCatiaSession
    Call OpenTopAssemblyFromEnovia(topAssy, expectedRevision)
End Sub

Public Sub RunPVRSyncStep(ByVal projectNumber As String, ByVal syncFromBSF As String, ByVal ciOption As String, ByVal nonCiOption As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Err.Raise vbObjectError + 6300, , "PVRSync rewrite is not implemented yet. Extraction rewrite is available through RunPVRExportStep."
End Sub

Public Sub RunPVRExportStep(ByVal exportFolder As String, Optional ByVal kbePathFile As String = "", Optional ByVal toolbarPath As String = "")
    Call RewriteExportActivePVR(exportFolder)
End Sub

Public Sub RunFullExtraction()
    Err.Raise vbObjectError + 6301, , "Full extraction rewrite is not available yet because the PVRSync rewrite is not implemented."
End Sub
