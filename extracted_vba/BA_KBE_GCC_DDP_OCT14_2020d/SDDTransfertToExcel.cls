Attribute VB_Name = "SDDTransfertToExcel"
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
Dim oView As DrawingView
Dim oDwgText As DrawingText
Dim iColumn As Integer
Dim sTempString As String
Dim sString As String

Dim oColumnNameList As New clsCollection
Dim oDwgTextList As clsCollection

'Log File
Call AddToLogFile("SDD Transfer")

'*** Initialize
Set oDrawingDocument = CATIA.ActiveDocument
Call frmProgress.progressBarInitialize("Extracting, Please Wait...")

'*** Get Excel file
Set oExcel = New clsExcel
Call oExcel.GetExcel

'*** Open Excel file
Call oExcel.OpenExcelFile(sSDDTemplateFile)
oExcel.ShowExcelWindow


'*** Initialize oColumnNameList
' Initialize Column Name
sString = "SDD_QtyRH|SDD_QtyLH|SDD_FindNo|SDD_ZoneSheet|SDD_CageCode|SDD_PartNumber|SDD_Description|SDD_ConfigCode|"
sString = sString & "SDD_Zone1|SDD_Zone2|SDD_Zone3|SDD_Zone4|SDD_Zone5|SDD_Zone6|SDD_Zone7|SDD_Zone8|SDD_Zone9|"
sString = sString & "SDD_Zone10|SDD_Zone11BP|SDD_Zone11DB|SDD_Zone11BI|SDD_Zone11BR|SDD_ATANumber|SDD_PCCN"

' Transfer value in array
sTempArray = Split(sString, "|")

'Add elements to oColumnNameList
For i = LBound(sTempArray) To UBound(sTempArray)
    oColumnNameList.Add sTempArray(i), sTempArray(i)
Next


'*** Scan all sheets
For Each oSheet In oDrawingDocument.Sheets

    '*** Only SH sheets
    If Left(oSheet.Name, 2) = "SH" Then

        Call frmProgress.progressBarRepaint("Extracting " & oSheet.Name)
        
        '*** Background view
        Set oView = oSheet.Views.Item(2)

        '*** If there is no text object on the sheet we go to the next one
        If oView.Texts.Count = 0 Then GoTo NextSheet

        'Initial
        Set oDwgTextList = New clsCollection

        '*** Scan all text object in Background view, bypass the ones not in oColumnNameList
        For Each oDwgText In oView.Texts
            If oColumnNameList.Exists(oDwgText.Name) Then

                '* text name was not already found in oDwgTextList
                If Not oDwgTextList.Exists(oDwgText.Name) Then
                    
                    'Add to oDwgTextList
                    oDwgTextList.Add oDwgText.Name, oDwgText
                    
                    'Transfer text to Excel
                    iColumn = oColumnNameList.GetIndex(oDwgText.Name) + 1
                    Call TransferValueToExcel(oSheet.Name, iColumn, oDwgText.Text)
                
                '* two text with the same name in the same sheet
                Else
                    sAnswer = MsgBox("There is more than one text object named " & oDwgText.Name & " in Sheet Background of " & oSheet.Name & "." & vbCrLf & "Please delete useless text or rename and run macro again.", vbExclamation)
                    GoTo Error_Was_Found
                End If
            End If
        Next
        
        '*** Check if all first row text objects were found
        For i = 1 To oColumnNameList.Count
            If Not oDwgTextList.Exists(oColumnNameList.GetKey(i)) Then
                sAnswer = MsgBox("Text object named " & oColumnNameList.GetKey(i) & " can't be found in Sheet Background of " & oSheet.Name & ".", vbExclamation)
            End If
        Next

    End If
    
NextSheet:
Next

Error_Was_Found:
oExcel.RunMacro ("MoveShapes")
Set oExcel = Nothing

frmProgress.Hide
Call frmKBEMain.resetToolbar

End Sub

Private Sub TransferValueToExcel(ByVal sSheetName As String, ByVal iColumn As Integer, ByVal sText As String)

Dim iSheetNb As Integer
Dim iStartRow As Integer
Dim sTextArray
Dim i As Integer


'*** Convert Sheet Name in number
iSheetNb = CInt(Right(sSheetName, 2))

'*** Calculate first row in Excel file in reference to the sheet number
If iSheetNb = 1 Then
    iStartRow = 3
ElseIf iSheetNb = 2 Then
    iStartRow = 22
Else
    iStartRow = 46
End If

'*** Convert sText into an array
sTextArray = Split(sText, vbLf)
If sTextArray(0) = "X" Then Exit Sub

'*** Transfer all values in the array to Excel
oExcel.App.Range(oExcel.App.Cells(iStartRow, iColumn), oExcel.App.Cells(iStartRow + UBound(sTextArray), iColumn)).Value = oExcel.App.Transpose(sTextArray)

End Sub



