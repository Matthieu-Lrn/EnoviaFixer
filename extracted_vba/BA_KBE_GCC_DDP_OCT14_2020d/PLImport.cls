Attribute VB_Name = "PLImport"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub Launch()

Dim oDrawingDocument As DrawingDocument
Dim oSheet As DrawingSheet
Dim sSheetName As String
Dim oBkgView As DrawingView
Dim oDwgText As DrawingText
Dim iNbOfSheets As Integer
Dim oDwgTable As DrawingTable
Dim dTablePositionX As Double
Dim dTablePositionY As Double
Dim dColumnSize(1 To 12) As Double
Dim iExcelLastRow As Integer
Dim oPLValues() As Variant
Dim iSheetNb As Integer
Dim iPLLastNonEmptyRow As Integer
Dim iDwgTableRow As Integer
Dim iDwgTableColumn As Integer
Dim dFontsize As Double
Dim sFontColor() As String
Dim iBkgColor As Long
Dim iFontColor As Long
Dim oVisProperty As VisPropertySet
Dim sAnswer As String
Dim sCellValue As String
Dim oTextFct As New clsDwgTextFunctions
Dim j As Integer
Dim oExcel As clsExcel

'*** Initial set
Set oDrawingDocument = CATIA.ActiveDocument
Set oSelection = oDrawingDocument.Selection

'*** Drawing Table Columns Size
dColumnSize(1) = 0.3568 * 25.4
dColumnSize(2) = 0.3568 * 25.4
dColumnSize(3) = 0.3568 * 25.4
dColumnSize(4) = 0.446 * 25.4
dColumnSize(5) = 0.5352 * 25.4
dColumnSize(6) = 2.3192 * 25.4
dColumnSize(7) = 2.676 * 25.4
dColumnSize(8) = 0.6244 * 25.4
dColumnSize(9) = 1.8732 * 25.4
dColumnSize(10) = 0.446 * 25.4
dColumnSize(11) = 0.1784 * 25.4
dColumnSize(12) = 0.2676 * 25.4

'Log File
Call AddToLogFile("Import Part List", oDrawingDocument.Name)


'*** Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetNewExcel
Call oExcel.OpenExcelFile("C:\Temp\DDPToolbar_PL.xls")

'*** Find last row
oExcel.RunMacro ("FindLastRow")
iExcelLastRow = oExcel.App.Cells(1, 17).Value
oExcel.App.Cells(1, 17).Value = ""

'*** Transfer all the values in the Excel file to oPLValues
oPLValues = oExcel.App.Range(oExcel.App.Cells(2, 2), oExcel.App.Cells(iExcelLastRow, 13)).Value

'*** Get Font Color
ReDim sFontColor(1 To UBound(oPLValues, 1), 1 To 12)

For iRow = 2 To iExcelLastRow

    For iColumn = 2 To 13
    
        'Font color
        iFontColor = oExcel.App.Cells(iRow, iColumn).Font.color
        Select Case iFontColor
            Case 16711935
            sFontColor(iRow - 1, iColumn - 1) = "Magenta"
            Case Else
            sFontColor(iRow - 1, iColumn - 1) = "Standard"
        End Select

    Next
Next


'Close Excel workbook and kill process
Call oExcel.CloseAlreadyOpenWorkbook("C:\Temp\DDPToolbar_PL.xls")
Set oExcel = Nothing


'*** How many PL sheets required
If (iExcelLastRow - 1) <= 23 Then
    iNbOfSheets = 1
Else
    iNbOfSheets = Int((iExcelLastRow - 1 - 23 - 1) / 26) + 2
End If


Call frmProgress.progressBarInitialize("Importing Part List")

'*** Scan all PL sheets
For iSheetNb = 1 To iNbOfSheets
    
    '*** Sheet Name
    If iSheetNb <= 9 Then
        sSheetName = "PL0" & iSheetNb
    Else
        sSheetName = "PL" & iSheetNb
    End If
    
    Call frmProgress.progressBarRepaint("Importing Part List: " & sSheetName)

    '*** Retrieve PL sheet
    On Error Resume Next
    Set oSheet = Nothing
    Set oSheet = oDrawingDocument.Sheets.Item(sSheetName)
    If oSheet Is Nothing Then
        sAnswer = MsgBox(sSheetName & " can't be found.", vbExclamation)
        GoTo Error_Was_Found
    End If
    On Error GoTo 0

    '*** First and last row (in oPLValues) for the active sheet
    If iSheetNb = 1 Then
        iPLFirstRow = 1
        iPLLastRow = 23
        dTablePositionX = 0.2818 * 25.4
        dTablePositionY = 6.29881 * 25.4
    Else
        iPLFirstRow = 24 + (iSheetNb - 2) * 26
        iPLLastRow = 49 + (iSheetNb - 2) * 26
        dTablePositionX = 0.2818 * 25.4
        dTablePositionY = 6.69881 * 25.4
    End If
    If iPLLastRow > UBound(oPLValues, 1) Then
        iPLLastRow = UBound(oPLValues, 1)
    End If
    

    '*** Retrieve the last non-empty row of the table in the active sheet
    iPLLastNonEmptyRow = PLLastNonEmptyRow(oPLValues, iPLFirstRow, iPLLastRow)
    
    
    '*** Background view
    Set oBkgView = oSheet.Views.Item(2)
    
    '*** Delete useless text and "PL_Table"
    Call oTextFct.DeleteUselessText(oDrawingDocument, oSheet)
    If oBkgView.Tables.Count > 0 Then
        For i = 1 To oBkgView.Tables.Count
            Set oDwgTable = oBkgView.Tables.Item(i)
            If oDwgTable.Name = "PL_Table" Then
                oBkgView.Tables.Remove (i)
                Exit For
            End If
        Next
    End If
    

    '*** Show Header
    oSelection.Clear
    oSelection.Add oBkgView
    oSelection.Search "Name=PL'.'Header*,sel"
    If oSelection.Count >= 1 Then
        oSelection.VisProperties.SetShow catVisPropertyShowAttr
    End If
    
    oSelection.Clear

    '*** Create Table
    Set oDwgTable = oBkgView.Tables.Add(dTablePositionX, dTablePositionY, iPLLastNonEmptyRow - iPLFirstRow + 1, 12, 0.24 * 25.4, 1)
    oDwgTable.Name = "PL_Table"

    oDwgTable.ComputeMode = CatTableComputeOFF
    
    '*** Fill Table
    For iDwgTableRow = 1 To oDwgTable.NumberOfRows

        For iDwgTableColumn = 1 To 12
              
            sCellValue = Trim(oPLValues(iPLFirstRow + iDwgTableRow - 1, iDwgTableColumn))
            Call oTextFct.CalculateFontSize(sCellValue, dFontsize, iDwgTableColumn)

            Call oDwgTable.SetCellString(iDwgTableRow, iDwgTableColumn, sCellValue)
            Set oDwgText = oDwgTable.GetCellObject(iDwgTableRow, iDwgTableColumn)

            oDwgText.SetFontSize 0, 0, dFontsize
            oDwgText.SetFontName 0, 0, "SSS1"

            'Cell alignment
            If iDwgTableColumn <= 5 Then
                Call oDwgTable.SetCellAlignment(iDwgTableRow, iDwgTableColumn, CatTableMiddleCenter)
            End If
        Next
    Next
    
    '*** Set column width
    For i = 1 To 12
        Call oDwgTable.SetColumnSize(i, dColumnSize(i))
    Next
    
    '*** Set table color and line width
    oSelection.Clear
    oSelection.Add oDwgTable
    Set oVisProperty = oSelection.VisProperties
    Call oVisProperty.SetRealColor(0, 0, 0, 0)
    Call oVisProperty.SetRealWidth(2, 0)
    
    
    '*** Set Text Color
    For iDwgTableRow = 1 To oDwgTable.NumberOfRows

        For iDwgTableColumn = 1 To 12
            Set oDwgText = oDwgTable.GetCellObject(iDwgTableRow, iDwgTableColumn)
            If sFontColor(iPLFirstRow + iDwgTableRow - 1, iDwgTableColumn) = "Magenta" Then
                oDwgText.SetParameterOnSubString catColor, 0, 0, -16711681
            Else
                oDwgText.SetParameterOnSubString catColor, 0, 0, 8388863
            End If
        Next
    Next

    
    oDwgTable.ComputeMode = CatTableComputeON
Next
    
    
    
'Reset unused PL sheets
Call ResetPLText(oDrawingDocument, CInt(Right(sSheetName, 2)) + 1)

Error_Was_Found:

Call frmProgress.Hide
Call frmKBEMain.resetToolbar

End Sub

Private Sub ResetPLText(ByVal oDrawingDocument As DrawingDocument, iFirstPLSheet As Integer)

Dim oSheet As DrawingSheet
Dim oDwgText As DrawingText
Dim oBkgView As DrawingView
Dim sAnswer As String
Dim sColumnNameList As Variant
Dim oTextFct As New clsDwgTextFunctions

'*** Initial set
Set oSelection = oDrawingDocument.Selection


'*** Scan all sheets
For Each oSheet In oDrawingDocument.Sheets
    If Left(oSheet.Name, 2) = "PL" Then
        If CInt(Right(oSheet.Name, 2)) >= iFirstPLSheet Then
    
            '*** Background view
            Set oBkgView = oSheet.Views.Item(2)

            '*** Ask user to reset PL
            sAnswer = MsgBox("Imported Part List is using " & iFirstPLSheet - 1 & " pages." & vbCrLf & "Do you want to reset PL text on " & oSheet.Name, vbYesNo)
            If sAnswer <> vbYes Then GoTo NextSheet

            '*** Delete useless text
            Call oTextFct.DeleteUselessText(oDrawingDocument, oSheet)
            
            '*** Delete "PL_Table"
            If oBkgView.Tables.Count > 0 Then
                For i = 1 To oBkgView.Tables.Count
                    Set oDwgTable = oBkgView.Tables.Item(i)
                    If oDwgTable.Name = "PL_Table" Then
                        oBkgView.Tables.Remove (i)
                        Exit For
                    End If
                Next
            End If
        End If
    End If

NextSheet:
Next
End Sub

Private Function PLLastNonEmptyRow(ByVal oPLValues As Variant, ByVal iPLFirstRow As Integer, ByVal iPLLastRow As Integer) As Integer

Dim iRow As Integer
Dim iColumn As Integer
Dim iPLLastNonEmptyRow As Integer
Dim sLine As String

'*** Scan row starting from iPLLastRow and retrieve the first non-empty row
For iRow = iPLLastRow To iPLFirstRow Step -1
    
    '*** Concatenate all the values on the same row
    sLine = ""
    For iColumn = LBound(oPLValues, 2) To UBound(oPLValues, 2)
        sLine = sLine & "|" & Trim(oPLValues(iRow, iColumn))
    Next
    
    '*** If the row is not empty we exit the for
    If sLine <> "||||||||||||" Then
        iPLLastNonEmptyRow = iRow
        Exit For
    End If
Next

PLLastNonEmptyRow = iPLLastNonEmptyRow
End Function















