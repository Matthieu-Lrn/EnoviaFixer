Attribute VB_Name = "SDDImport"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim oColumnNameList As clsCollection

'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
' SHOULD BE UPDATED AS PER PLImport
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************


Public Sub Launch()

Dim oDrawingDocument As DrawingDocument
Dim oSheet As DrawingSheet
Dim sPreviousSheetName As String
Dim sSheetName As String
Dim oView As DrawingView

Dim oDwgText As DrawingText
Dim sDwgTextShowState() As Boolean
Dim sTextNewLine As String
Dim sTempArray() As String

Dim iRow As Integer
Dim iColumn As Integer
Dim iLastRow As Integer
Dim iItem As Integer
Dim iSheetNb As Integer
Dim oComponent As DrawingComponent
Dim bFirstLine As Boolean
Dim oExcel As clsExcel

Dim oDwgTextList As clsCollection

Set oDrawingDocument = CATIA.ActiveDocument
Set oSelection = oDrawingDocument.Selection

'Log File
Call AddToLogFile("SDD Import")


'*** Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetExcel
Call oExcel.OpenExcelFile(, "Select the Excel file to be imported")


'*** Find last row
oExcel.RunMacro ("FindLastRow")
iLastRow = oExcel.App.Cells(1, 26).Value
oExcel.App.Cells(1, 26).Value = ""

'*** Initialize oColumnNameList
' Initialize Column Name
sString = "SDD_QtyRH|SDD_QtyLH|SDD_FindNo|SDD_ZoneSheet|SDD_CageCode|SDD_PartNumber|SDD_Description|SDD_ConfigCode|"
sString = sString & "SDD_Zone1|SDD_Zone2|SDD_Zone3|SDD_Zone4|SDD_Zone5|SDD_Zone6|SDD_Zone7|SDD_Zone8|SDD_Zone9|"
sString = sString & "SDD_Zone10|SDD_Zone11BP|SDD_Zone11DB|SDD_Zone11BI|SDD_Zone11BR|SDD_ATANumber|SDD_PCCN"

' Transfer value in array
sTempArray = Split(sString, "|")

'Add elements to oColumnNameList
Set oColumnNameList = New clsCollection
For i = LBound(sTempArray) To UBound(sTempArray)
    oColumnNameList.Add sTempArray(i), sTempArray(i)
Next

'*** Scan Excel file
sPreviousSheetName = ""
For iRow = 3 To iLastRow
    
    iItem = iRow - 2
    sSheetName = ""
    
    '*** Associate item to a sheet
    If iItem <= 19 Then
        sSheetName = "SH01"
    ElseIf iItem <= 43 Then
        sSheetName = "SH02"
    Else
        sSheetName = "SH03"
    End If
    
    '*** If you change sheet
    If sPreviousSheetName <> sSheetName Then
    
        bFirstLine = True
        
        '*** Retrieve SDD sheet
        On Error Resume Next
        Set oSheet = Nothing
        Set oSheet = oDrawingDocument.Sheets.Item(sSheetName)
        If oSheet Is Nothing Then
            sAnswer = MsgBox(sSheetName & " can't be found", vbExclamation)
            Exit For
        End If
        On Error GoTo 0
        
        '*** Background view
        Set oView = oSheet.Views.Item(2)
       
        'Initial
        Set oDwgTextList = New clsCollection
        
        '*** Scan all text object in Background view, bypass the ones not in oColumnNameList
        For Each oDwgText In oView.Texts
            If oColumnNameList.Exists(oDwgText.Name) Then

                If Not oDwgTextList.Exists(oDwgText.Name) Then
                    oDwgTextList.Add oDwgText.Name, oDwgText
                Else
                    sAnswer = MsgBox("There is more than one text object named " & oDwgText.Name & " in Sheet Background of " & oSheet.Name & "." & vbCrLf & "Please delete useless text object and run macro again.", vbExclamation)
                    GoTo Error_Was_Found
                End If
            End If
        Next


        '*** Check if all first row text objects were found
        For i = 1 To oColumnNameList.Count
            If Not oDwgTextList.Exists(oColumnNameList.GetKey(i)) Then
                sAnswer = MsgBox("Text object named " & oColumnNameList.GetKey(i) & " can't be found in Sheet Background of " & oSheet.Name & "." & vbCrLf & "Please rename manually and run macro again.", vbExclamation)
                GoTo Error_Was_Found
            End If
        Next
        
        '*** Hide all text ojects found in oDwgTextList
        ReDim sDwgTextShowState(1 To oDwgTextList.Count)
        oSelection.Clear
        For i = 1 To oDwgTextList.Count
            oSelection.Add oDwgTextList.GetItemByIndex(i)
            sDwgTextShowState(i) = False
        Next
        oSelection.VisProperties.SetShow catVisPropertyNoShowAttr

        '*** Show Header
        oSelection.Clear

        Set oComponent = Nothing
        For Each oComponent In oView.Components
            If oComponent.Name Like "PL.Header*" Then Exit For
        Next
        If Not oComponent Is Nothing Then
            oSelection.Add oComponent
            oSelection.VisProperties.SetShow catVisPropertyShowAttr
        End If

        oSelection.Clear

    Else
        bFirstLine = False
    End If
    
    sPreviousSheetName = sSheetName
    
    
    '*** Scan columns in Excel file and populate values in the different text objects.
    For iColumn = 2 To 25
    
        '*** Retrieve text object: PL_FindNo, PL_Description......
        Set oDwgText = oDwgTextList.GetItemByKey(oColumnNameList.GetKey(iColumn - 1))
        
        '*** Retrieve and format value from Excel file
        If oExcel.App.Cells(iRow, iColumn).Value = "" Then
            sTextNewLine = ""
        
        '*** Qty RH, Qty LH or Find No
        ElseIf iColumn <= 3 And IsNumeric(oExcel.App.Cells(iRow, iColumn).Value) Then
            
            '*** Value < 10 we leave a blank space in front of the text to center it in the column
            If CInt(oExcel.App.Cells(iRow, iColumn).Value) < 10 Then
                sTextNewLine = " " & oExcel.App.Cells(iRow, iColumn).Value
            Else
                sTextNewLine = oExcel.App.Cells(iRow, iColumn).Value
            End If
        Else
            sTextNewLine = oExcel.App.Cells(iRow, iColumn).Value
        End If
        
        '*** Show text object if value <> "" and hidden
        If sTextNewLine <> "" And sDwgTextShowState(iColumn - 1) = False Then
            oSelection.Clear
            oSelection.Add oDwgText
            oSelection.VisProperties.SetShow catVisPropertyShowAttr
            sDwgTextShowState(iColumn - 1) = True
        End If
        
        
        '*** First line on a new PL sheet
        If bFirstLine = True And sTextNewLine <> "" Then
            oDwgText.Text = sTextNewLine
        ElseIf bFirstLine = True And sTextNewLine = "" Then
            oDwgText.Text = "X"
        
        '*** Other Line
        ElseIf bFirstLine = False Then
            oDwgText.Text = oDwgText.Text & vbCrLf & sTextNewLine
            
            '*** First line = "X" and other line is not empty we need to remove the "X" on the first line
            If Split(oDwgText.Text, vbLf)(0) = "X" And sTextNewLine <> "" Then
                oDwgText.Text = "" & vbCrLf & Split(oDwgText.Text, vbLf, 2)(1)
            End If
        End If
    Next
    
Next


Error_Was_Found:

'Show Excel
Call oExcel.ShowExcelWindow
Set oExcel = Nothing

Call frmKBEMain.resetToolbar

End Sub
