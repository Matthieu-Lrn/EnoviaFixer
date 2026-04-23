Attribute VB_Name = "PLTransferToExcel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oExcel As clsExcel
Public Sub Launch()

Dim oDrawingDocument As DrawingDocument
Dim oSheet As DrawingSheet
Dim oBkgView As DrawingView
Dim oDwgText As DrawingText
Dim iColumn As Integer
Dim oDwgTable As DrawingTable
Dim sString As String
Dim sAllWarning As String
Dim sOneWarning As String
Dim oColumnNameList As New clsCollection
Dim oDwgTextList As clsCollection


'Log File
Call AddToLogFile("Transfer Part List")


'*** Initialize
Set oDrawingDocument = CATIA.ActiveDocument
Call frmProgress.progressBarInitialize("Extracting, Please Wait...")

'*** Get Excel file
Set oExcel = New clsExcel
Call oExcel.GetNewExcel

'*** Close already opened file
Call oExcel.CloseAlreadyOpenWorkbook("C:\Temp\DDPToolbar_PL.xls")

'***Open Excel file
Call oExcel.OpenExcelFile(sPLTemplateFile)
oExcel.ShowExcelWindow
oExcel.RunMacro ("DeleteTypeColumn")

'*** Initialize oColumnNameList
' Initialize Column Name
sString = "PL_QtyRH|PL_QtyLH|PL_FindNo|PL_ZoneSheet|PL_CageCode|PL_PartNumber|PL_Description|PL_Size|"
sString = sString & "PL_MatSpec|PL_FinishCode|PL_PCCN"

' Transfer value in array
sTempArray = Split(sString, "|")

'Add elements to oColumnNameList
For i = LBound(sTempArray) To UBound(sTempArray)
    oColumnNameList.Add sTempArray(i), sTempArray(i)
Next


'*** Scan all sheets
For Each oSheet In oDrawingDocument.Sheets

    'Only PL sheets
    If Left(oSheet.Name, 2) = "PL" Then

        Call frmProgress.progressBarRepaint("Extracting " & oSheet.Name)
        
        '*** Background view
        Set oBkgView = oSheet.Views.Item(2)
        
        '*** We check if we have a drawing table named "PL_Table"
        Set oDwgTable = Nothing
        If oBkgView.Tables.Count > 0 Then
            For i = 1 To oBkgView.Tables.Count
                Set oDwgTable = oBkgView.Tables.Item(i)
                If oDwgTable.Name = "PL_Table" Then
                    Exit For
                End If
                Set oDwgTable = Nothing
            Next
        End If

        '*** If we dont have a drawing table we extract the value from the "PL_xxx" text
        If oDwgTable Is Nothing And oBkgView.Texts.Count > 0 Then
        

            sAnswer = MsgBox("Drawing table named 'PL_Table' can't be found in sheet " & oSheet.Name & "." & vbCrLf & _
                             "Macro will extract value from 'PL_xxx' text object shown in the background view.", vbInformation)
            
            
            sAllWarning = ""
            
            'Check for visible text object when name <> "PL_*"
            Set oSelection = oDrawingDocument.Selection
            oSelection.Clear
            oSelection.Add oBkgView
            oSelection.Search "(CATDrwSearch.DrwText.Name_CAP!=PL_* & CATDrwSearch.DrwText.Visibility=Visible),sel"
           
            'Remove
            '- 'NOTES' text
            '- Text in the Header section
            '- Text in the Note column
            If oSelection.Count > 0 Then
                For j = oSelection.Count To 1 Step -1
                
                    If TypeName(oSelection.Item(j).Value) = "DrawingText" Then
                        Set oDwgText = oSelection.Item(j).Value
                        If Not (oDwgText.Y < 6.71 * 25.4 And oDwgText.X < 10.45 * 25.4 And oDwgText.Text <> "NOTES") Then
                            oSelection.Remove (j)
                        End If
                    Else
                    End If
                Next
            End If

            'If we still have text in the selection
            If oSelection.Count > 0 Then
                sOneWarning = " - Text objects were found but they are not named 'PL_xxx'." & vbCrLf & _
                              "   They won't be extracted in the Excel file."
                If sAllWarning = "" Then
                    sAllWarning = sOneWarning
                Else
                    sAllWarning = sAllWarning & vbCrLf & sOneWarning
                End If
                
            End If
            
            
            'Initial
            Set oDwgTextList = New clsCollection
    
            'Check for visible text object when name start with "PL_*"
            oSelection.Clear
            oSelection.Add oBkgView
            oSelection.Search "(CATDrwSearch.DrwText.Name_CAP=PL_* & CATDrwSearch.DrwText.Visibility=Visible),sel"
    
            'No text were found
            If oSelection.Count = 0 Then

                sOneWarning = " - No text object named 'PL_xxx' were found." & vbCrLf & _
                              "   Nothing will be extracted in the Excel file."
            
                If sAllWarning = "" Then
                    sAllWarning = sOneWarning
                Else
                    sAllWarning = sAllWarning & vbCrLf & sOneWarning
                End If

            ''PL_xx' text were found
            Else
            
                For i = 1 To oSelection.Count
                    Set oDwgText = oSelection.Item(i).Value
                
                
                    '* text name was not already found in oDwgTextList
                    If Not oDwgTextList.Exists(oDwgText.Name) Then
                        
                        'Add to oDwgTextList
                        oDwgTextList.Add oDwgText.Name, oDwgText
                        
                        'Transfer text to Excel
                        iColumn = oColumnNameList.GetIndex(oDwgText.Name) + 1 '+1 because because first text will be in column B in Excel
                        Call TransferValueToExcel(oSheet.Name, iColumn, oDwgText.Text)
                    
                    '* two text with the same name in the same sheet
                    Else
                        sOneWarning = " - There is more than one text named " & oDwgText.Name & "." & vbCrLf & _
                                      "   Only the first one will be extracted to the Excel file."
                    
                        If sAllWarning = "" Then
                            sAllWarning = sOneWarning
                        Else
                            sAllWarning = sAllWarning & vbCrLf & sOneWarning
                        End If
                    End If
                Next
                              
                If sAllWarning <> "" Then
                    sAllWarning = "Following warnings were found on " & oSheet.Name & " :" & vbCrLf & sAllWarning
                    sAnswer = MsgBox(sAllWarning, vbExclamation)
                End If
            End If
            
        '*** If we have a drawing table we extract the values
        ElseIf Not oDwgTable Is Nothing Then
            Call TransferDwgTableToExcel(oSheet.Name, oDwgTable)
        End If
    End If
    

Next



oExcel.RunMacro ("ColumnsAutoFit")
oExcel.SaveWorkbook ("C:\Temp\DDPToolbar_PL.xls")

Set oExcel = Nothing
frmProgress.Hide
Call frmKBEMain.resetToolbar


End Sub


Private Sub TransferValueToExcel(ByVal sSheetName As String, ByVal iColumn As Integer, ByVal sText As String)

Dim iSheetNb As Integer
Dim iStartRow As Integer
Dim sTextArray
Dim i As Integer


'Convert Sheet Name in number
iSheetNb = CInt(Right(sSheetName, 2))

'Calculate first row in Excel file in reference to the sheet number
If iSheetNb = 1 Then
    iStartRow = 2
Else
    iStartRow = 25 + (iSheetNb - 2) * 26
End If

'Convert sText into an array
sTextArray = Split(sText, vbLf)
If (sTextArray(0) = "X" Or sTextArray(0) = "XX" Or sTextArray(0) Like "XXX*") Then Exit Sub

'*** Transfer all values in the array to Excel
oExcel.App.Range(oExcel.App.Cells(iStartRow, iColumn), oExcel.App.Cells(iStartRow + UBound(sTextArray), iColumn)).Value = oExcel.App.Transpose(sTextArray)

End Sub

Private Sub TransferDwgTableToExcel(ByVal sSheetName As String, ByVal oDwgTable As DrawingTable)

Dim iSheetNb As Integer
Dim iExcelStartRow As Integer
Dim oText As DrawingText
Dim lColor As Long

Dim oPLValues() As Variant
Dim oFontSize() As Variant
Dim iDwgTableRow As Integer
Dim iDwgTableColumn As Integer
Dim iExcelRow As Integer
Dim iExcelColumn As Integer


'Convert Sheet Name in number
iSheetNb = CInt(Right(sSheetName, 2))

'Calculate first row in Excel file in reference to the sheet number
If iSheetNb = 1 Then
    iStartRow = 2
Else
    iStartRow = 25 + (iSheetNb - 2) * 26
End If

'Convert drawing table in array
ReDim oPLValues(1 To oDwgTable.NumberOfRows, 1 To 12)
ReDim oFontSize(1 To oDwgTable.NumberOfRows, 1 To 12)

For iDwgTableRow = 1 To oDwgTable.NumberOfRows
    For iDwgTableColumn = 1 To 12
        oPLValues(iDwgTableRow, iDwgTableColumn) = Trim(oDwgTable.GetCellString(iDwgTableRow, iDwgTableColumn))
        Set oText = oDwgTable.GetCellObject(iDwgTableRow, iDwgTableColumn)
        lColor = oText.GetParameterOnSubString(catColor, 0, 0)
        If lColor = -16711681 Then
            oFontSize(iDwgTableRow, iDwgTableColumn) = "Magenta"
        Else
            oFontSize(iDwgTableRow, iDwgTableColumn) = "Standard"
        End If
    Next
Next


' Transfer all values in the array to Excel
oExcel.App.Range(oExcel.App.Cells(iStartRow, 2), oExcel.App.Cells(iStartRow + oDwgTable.NumberOfRows - 1, 13)).Value = oPLValues


' Apply Font color
For i = 1 To UBound(oFontSize, 1)
    For j = 1 To 12
        
        If oFontSize(i, j) = "Magenta" Then
            oExcel.App.Cells(iStartRow + i - 1, j + 1).Font.color = RGB(255, 0, 255)
        End If
        
    Next
Next

End Sub
















