Attribute VB_Name = "ImportTableFromExcel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'***    Update Sept 15 2016 By Abhishek Kamboj
'***    Update Sept28 2016 Bug Correction
Public Sub LaunchImport(sArray As Variant, sExcelFilePath As String, Optional ByVal sWorksheetName As String = "", Optional ByVal StartRow As Integer = 2, Optional ByVal StartCol As Integer = 2)
'*** imports values from excel file & dumps it in an array
'*** if aWorksheet name is give then it will import from a particular sheet
'*** otherwise it will import from active sheet

    Dim sAnswer As String
    Dim oExcelAp As New clsExcel
    '***** Launch Excel
    
    Set oExcelAp = New clsExcel
    Call oExcelAp.GetNewExcel
    
    'oExcelAp.App.Visible = False
    oExcelAp.App.DisplayAlerts = False
    oExcelAp.App.ScreenUpdating = False
    
    Call oExcelAp.CloseAlreadyOpenWorkbook(sExcelFilePath)
    'Call oExcelAp.OpenExcelFile(sExcelFilePath)
    
    Dim oXLWorkbook 'As Workbook
    Set oXLWorkbook = oExcelAp.App.Workbooks.Add(sExcelFilePath)
    Dim oXLSheet 'As Worksheet
    
    If sWorksheetName = "" Then sWorksheetName = oXLWorkbook.Sheets.Item(1).Name
    Set oXLSheet = oXLWorkbook.Sheets(sWorksheetName)
    
    If oExcelAp.IsOneWkbOpen = False Then
        sAnswer = MsgBox("Cannot Open Excel workbook" & vbCrLf & "Macro aborted.", vbExclamation)
        GoTo endsub
    End If
    
    iNbRow = StartRow
    Do While oXLSheet.Cells(iNbRow, StartCol).Value <> ""
        iNbRow = iNbRow + 1
    Loop
    iNbRow = iNbRow - 1
    
    iNbColumn = StartCol
    Do While oXLSheet.Cells(StartRow, iNbColumn).Value <> ""
        iNbColumn = iNbColumn + 1
    Loop
    iNbColumn = iNbColumn - 1
    
    '***** Transfer values from Excel file to sArray
    ReDim sArray(1 To iNbRow, 1 To iNbColumn)
    sArray = oXLSheet.Range(oXLSheet.Cells(1, 1), oXLSheet.Cells(iNbRow, iNbColumn)).Value

    oExcelAp.App.DisplayAlerts = False
    oXLWorkbook.Close
    DoEvents
    oExcelAp.App.ScreenUpdating = True

endsub:
    oExcelAp.App.Quit
    DoEvents
    
    Set oExcelAp = Nothing
End Sub


