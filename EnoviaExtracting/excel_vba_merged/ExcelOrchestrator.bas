Attribute VB_Name = "ExcelOrchestrator"
Option Explicit

Public Sub RunFullExtraction()
    Dim ws As Worksheet
    Dim topAssy As String
    Dim revision As String
    Dim projectNumber As String
    Dim exportPath As String
    Dim syncFromBSF As String
    Dim ciOption As String
    Dim nonCiOption As String

    Set ws = ThisWorkbook.Worksheets("Inputs")

    topAssy = Trim$(CStr(ws.Range("B1").Value))
    revision = Trim$(CStr(ws.Range("B2").Value))
    projectNumber = Trim$(CStr(ws.Range("B3").Value))
    exportPath = Trim$(CStr(ws.Range("B4").Value))
    syncFromBSF = UCase$(Trim$(CStr(ws.Range("B5").Value)))
    ciOption = Trim$(CStr(ws.Range("B6").Value))
    nonCiOption = Trim$(CStr(ws.Range("B7").Value))

    If topAssy = "" Then Err.Raise vbObjectError + 3000, , "Top assembly is required."
    If revision = "" Then Err.Raise vbObjectError + 3001, , "Revision is required."
    If exportPath = "" Then Err.Raise vbObjectError + 3002, , "Export path is required."
    If syncFromBSF <> "TRUE" And projectNumber = "" Then Err.Raise vbObjectError + 3003, , "Project number is required when Sync From BSF is FALSE."

    Call RunTopAssySearch(topAssy, revision)
    Call RunPVRSyncStep(projectNumber, syncFromBSF, ciOption, nonCiOption)
    Call RunPVRExportStep(exportPath)
End Sub
