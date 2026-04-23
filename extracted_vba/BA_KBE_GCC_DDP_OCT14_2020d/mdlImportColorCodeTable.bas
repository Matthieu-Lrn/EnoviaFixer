Attribute VB_Name = "mdlImportColorCodeTable"
Option Explicit

Private Type ExcelAddress
    HeaderRow As Integer
    FMLColumn As Integer
    TitleColumn As Integer
    PNColumn As Integer
    DefPartColumn As Integer
    NomenclatureColumn As Integer
    DatasetTypeColumn As Integer
End Type

Private Type DwgInfo
    DwgNumber As String
    DwgRev As String
    DwgIteration As String
End Type

Private Type DataTableRowInfo
    RowNb As Integer
    RowHeight As Double
    Text As String
    Frame As String
    Alignment As String
    FontSize As String
    FontName As String
End Type

Private oExcel As clsExcel


Sub ColorCodeMain()

Dim oExcelData As clsXML
Dim vExcelData As Variant
Dim sReturnMsg As String, sMsg As String
Dim oTableColl As Collection
Dim oColorCodeList As clsCollection
Dim iActionNb As Integer
Dim oDwgInfo As DwgInfo
Dim oWks As Object

'Clear cache
Call StartWebServiceTool
WebServiceAccessTool.ClearCache

'Select the action to be done
iActionNb = SelectAction
If iActionNb = 0 Then GoTo ResetTheToolbar

'Go fetch data from Excel file and drawing color coding table if required
If iActionNb <> 4 Then
    
    'Get "COMPLETE STRUCTURE" worksheet from IDL document
    sReturnMsg = ""
    Call GetIDLWorkbook(oWks, sReturnMsg)
    If sReturnMsg = "ExitSub" Then GoTo ResetTheToolbar
    
    'Get Excel data
    sReturnMsg = ""
    Call GetExcelData(oExcelData, sReturnMsg, oWks)
    If sReturnMsg = "ExitSub" Then GoTo ResetTheToolbar
    
    'Retrieve the Color Code list
    Call PopulateColorCodeList(oColorCodeList)
    
    'Retrieve all the color coding tables in the drawing
    Set oTableColl = RetrieveColorTables

End If

'Get the drawing info
Call GetDrawingInfo(oDwgInfo)

'Perform Action
Call PerformAction(oDwgInfo, iActionNb, oTableColl, oExcelData, oColorCodeList)

ResetTheToolbar:
Call frmKBEMain.resetToolbar

End Sub

Private Function TransformDataToArray(ByVal oExcelData As clsXML) As Variant

Dim vArray As Variant
Dim oNode As IXMLDOMNode
Dim oNodeList As IXMLDOMNodeList
Dim i As Integer

Set oNodeList = oExcelData.SelectNodes("./Part", oExcelData.RootNode)
ReDim vArray(1 To oNodeList.Length, 1 To 4)

For i = 1 To oNodeList.Length
    Set oNode = oNodeList.Item(i - 1)
    vArray(i, 1) = oNode.Attributes.getNamedItem("ColorCode").nodeValue
    vArray(i, 2) = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    vArray(i, 3) = oNode.Attributes.getNamedItem("Title").nodeValue
    vArray(i, 4) = oNode.Attributes.getNamedItem("Qty").nodeValue
Next

'Return
TransformDataToArray = vArray

End Function

Private Sub GetExcelData(ByRef oExcelData As clsXML, ByRef sReturnMsg As String, ByVal oWks As Object)

Dim oExcelAddress As ExcelAddress
Dim vArray As Variant
Dim i As Integer
Dim sPartNumber As String, sColorCode As String, sSortingValue As String, sAttNames(1 To 5) As String, sAttValues(1 To 5) As String
Dim oExistingNode As IXMLDOMNode
Dim oNewElem As IXMLDOMElement


'Get header row
oExcelAddress.HeaderRow = GetHeaderRow(oWks)
If oExcelAddress.HeaderRow = 0 Then
    Call MsgBox("Header row can't be retrieved in 'COMPLETE STRUCTURE'. Process aborted.", vbCritical, "Import Color Coding Table")
    sReturnMsg = "ExitSub"
    Exit Sub
End If

'Get columns
oExcelAddress.FMLColumn = GetColumn("FML MATERIAL", oExcelAddress.HeaderRow, oWks)
oExcelAddress.TitleColumn = GetColumn("TITLE", oExcelAddress.HeaderRow, oWks)
oExcelAddress.PNColumn = GetColumn("PART NUMBER", oExcelAddress.HeaderRow, oWks)
oExcelAddress.DefPartColumn = GetColumn("DEFINING PART", oExcelAddress.HeaderRow, oWks)
oExcelAddress.NomenclatureColumn = GetColumn("NOMENCLATURE", oExcelAddress.HeaderRow, oWks)
oExcelAddress.DatasetTypeColumn = GetColumn("DATASET TYPE", oExcelAddress.HeaderRow, oWks)

'Transfer Excel content in vArray
vArray = oWks.Cells(oExcelAddress.HeaderRow, 1).CurrentRegion.Value
Call oExcel.CloseExcelFile
Set oExcel = Nothing

'Initialize oExcelData
Set oExcelData = New clsXML
Call oExcelData.AddElement("Root")

'Initialize the attributes names to be added to oExcelData
sAttNames(1) = "SortString": sAttNames(2) = "ColorCode": sAttNames(3) = "PartNumber": sAttNames(4) = "Qty": sAttNames(5) = "Title"

'Scan all the rows
For i = 2 To UBound(vArray, 1)

    'Get the color code
    sColorCode = UCase(Trim(vArray(i, oExcelAddress.FMLColumn)))
    
    'Only consider rows with a color code
    If sColorCode <> "" Then
    
        'Get the part number
        sPartNumber = GetPartNumber(vArray(i, oExcelAddress.PNColumn), vArray(i, oExcelAddress.NomenclatureColumn), vArray(i, oExcelAddress.DefPartColumn), vArray(i, oExcelAddress.DatasetTypeColumn))
    
        'Define the sorting value
        sSortingValue = IIf(sColorCode = "SYSTEM", "ZZZZZZSYSTEM_" & sPartNumber, sColorCode & "_" & sPartNumber)
    
        'Part already in oExcelData, just increase the qty
        If oExcelData.Exists("./Part[@SortString='" & sSortingValue & "']", oExcelData.RootNode) Then
            Set oExistingNode = oExcelData.SelectSingleNode("./Part[@SortString='" & sSortingValue & "']", oExcelData.RootNode)
            Call oExcelData.EditAttributeValue(oExistingNode, "Qty", CInt(oExcelData.GetAttributeValue(oExistingNode, "Qty") + 1))
        
        'Add part to oExcelData
        Else
            Set oNewElem = oExcelData.AddElement("Part", oExcelData.RootNode)
            sAttValues(1) = sSortingValue: sAttValues(2) = sColorCode: sAttValues(3) = sPartNumber: sAttValues(4) = 1: sAttValues(5) = UCase(Trim(vArray(i, oExcelAddress.TitleColumn)))
            Call oExcelData.AddMultipleAttributes(oNewElem, sAttNames, sAttValues)
        End If
        
    End If
Next

'Sort oExcelData
Call oExcelData.SortNodes(oExcelData.RootNode, "SortString")

End Sub

Private Sub GetIDLWorkbook(ByRef oWks As Object, ByRef sReturnMsg As String)

Dim sAnswer As String, sMsg As String

'Get Excel
Set oExcel = New clsExcel
Call oExcel.GetNewExcel

'Ask user to select the IDL Extract file
Call oExcel.OpenExcelFile(, "Select the IDL extract file with the color codes")

'Define messages
sMsg = "No 'COMPLETE STRUCTURE' sheet were found in the selected workbook. Process aborted."

'No COMPLETE STRUCTURE" sheet found
If oExcel.WorksheetExists("COMPLETE STRUCTURE") = False Then
    Call MsgBox(sMsg, vbCritical, "Import Color Coding Table")
    Call oExcel.CloseExcelFile
    Set oExcel = Nothing
    sReturnMsg = "ExitSub"
    Exit Sub
End If

'Get "COMPLETE STRUCTURE" sheet
Set oWks = oExcel.App.Worksheets("COMPLETE STRUCTURE")

End Sub

Private Function GetHeaderRow(ByVal oWks As Object) As Integer

Dim iRow As Integer

For iRow = 1 To 50
    If Trim(oWks.Cells(iRow, 1).Value) = "ITEM #" Then
        Exit For
    End If
Next

If iRow < 50 Then
    GetHeaderRow = iRow
Else
    GetHeaderRow = 0
End If

End Function

Private Function GetColumn(ByVal sName As String, ByVal iHeaderRow As Integer, ByVal oWks As Object) As Integer

Dim iColumn As Integer

For iColumn = 1 To 100
    If Trim(oWks.Cells(iHeaderRow, iColumn).Value) Like sName & "*" Then
        Exit For
    End If
Next

If iColumn < 100 Then
    GetColumn = iColumn
Else
    GetColumn = 0
End If

End Function

Private Function GetPartNumber(ByVal sPartNumber As String, ByVal sNomenclature As String, ByVal sDefiningPart As String, ByVal sDatasetType As String) As String

'Define Part Number
If (sDatasetType = "FLEXIBLE REPRESENTATION" Or sDatasetType = "CATALOG LIGHT VERSION") And sDefiningPart <> "" Then
    sPartNumber = sDefiningPart
ElseIf sNomenclature <> "" Then
    sPartNumber = sNomenclature
End If

'Ucase
sPartNumber = UCase(sPartNumber)

'Clean
sPartNumber = CleanString(sPartNumber)

'Return
GetPartNumber = sPartNumber

End Function

Private Function CleanString(ByVal sString As String) As String

Dim sCleanString As String

sCleanString = sString
sCleanString = Trim(Replace(sCleanString, "(DON'T USE THIS PART)", ""))
sCleanString = Trim(Replace(sCleanString, "(CANCELLED)", ""))

CleanString = sCleanString
End Function

Private Sub PopulateColorCodeList(ByRef oColorCodeList As clsCollection)

    Dim sTextRow As String, iFileNo As Integer
    Dim iLine As Integer
    
    'Open file
    iFileNo = FreeFile
    Open sColorCodeList For Input As iFileNo

    'Create clsCollection
    Set oColorCodeList = New clsCollection

    'Scan file and transfer part number to oColorCodePartExclusionList
    iLine = 0
    Do While Not EOF(iFileNo)
        
        iLine = iLine + 1
        Line Input #iFileNo, sTextRow
        
        'The first two lines in the file should be skipped
        If iLine >= 3 Then
            Call oColorCodeList.Add(Split(sTextRow, ",")(1), Split(sTextRow, ",")(1))
        End If
    Loop
    
    'Add "System" to the collection
    Call oColorCodeList.Add("System", "System")
    
    Close #iFileNo
End Sub

Private Function RetrieveColorTables() As Collection

Dim oSelection
Dim oSheet As DrawingSheet
Dim i As Integer
Dim oTable As DrawingTable
Dim oBkgView As DrawingView
Dim oTableColl As Collection

'Initialize
Set oSelection = CATIA.ActiveDocument.Selection
Set oTableColl = New Collection

'Scan each drawing sheet and look for tables named "ColorCode_Table"
For Each oSheet In CATIA.ActiveDocument.Sheets

    'Only check in drawing sheet
    If oSheet.Name Like "SH##" Then
    
        'Get Background view of the sheet
        Set oBkgView = oSheet.Views.Item(2)
        
        'Select the color code tables
        oSelection.Clear
        oSelection.Add oBkgView
        oSelection.Search "(CATDrwSearch.DrwTable.Name=ColorCode_Table & CATDrwSearch.DrwTable.Visibility=Visible),sel"
    
        'Transfer the tables in the collection
        If oSelection.Count > 0 Then
            For i = 1 To oSelection.Count
                Set oTable = oSelection.Item(i).Value
    
                'Add Table to collection.
                Call oTableColl.Add(oTable)
    
            Next
        End If
    End If
Next

'Return
Set RetrieveColorTables = oTableColl
End Function

Private Function SelectAction() As Integer

Dim iActionNb As Integer
Dim sAnswer As String

'Ask user what he wants to do
sAnswer = InputBox("Select the action to be performed:" & vbCrLf _
        & " 1) Create new or overwrite existing Color Code Table" & vbCrLf _
        & " 2) Add a new Color Code Table" & vbCrLf _
        & " 3) Compare Color Code Table" & vbCrLf _
        & " 4) Validate compare info in the LOG database", "Import Color Coding Table")

'Trim
sAnswer = Trim(sAnswer)

'Check response
If IsNumeric(sAnswer) Then
    iActionNb = CInt(sAnswer)
    If iActionNb < 1 And iActionNb > 4 Then iActionNb = 0
Else
    iActionNb = 0
End If

'Return
SelectAction = iActionNb

End Function

Private Sub GetDrawingInfo(ByRef oDwgInfo As DwgInfo)

'Drawing is saved in Enovia
If CATIA.ActiveDocument.FullName Like "ENOVIA5*" Then
    oDwgInfo.DwgNumber = Split(CATIA.ActiveDocument.Name, "-")(0)
    oDwgInfo.DwgRev = Right(Split(CATIA.ActiveDocument.Name, ".")(0), 2)
    oDwgInfo.DwgIteration = GetDocumentIteration(oDwgInfo.DwgNumber, oDwgInfo.DwgRev)
'Drawing not saved in Enovia
Else
    oDwgInfo.DwgNumber = Split(CATIA.ActiveDocument.Name, ".")(0)
    oDwgInfo.DwgRev = "N/A"
    oDwgInfo.DwgIteration = "0"
End If
End Sub

Private Function GetDocumentIteration(ByVal sPartNumber As String, ByVal sRevision As String) As String

Dim oAttList As New clsAttributesList
Dim sIteration As String

'Get attributes
sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRevision, False, "DOCUMENT_ITERATION")

'Return value
If sIteration <> "Part or Attribute doesn't exist" Then
    GetDocumentIteration = sIteration
Else
    GetDocumentIteration = ""
End If

End Function

Private Sub PerformAction(ByRef oDwgInfo As DwgInfo, ByVal iActionNb As Integer, ByVal oTableColl As Collection, ByVal oExcelData As clsXML, ByVal oColorCodeList As clsCollection)

Dim oBkgView As DrawingView
Dim oDwgDoc As DrawingDocument
Dim oDwgData As clsXML
Dim sErrorMsg As String, sCompare As String, sMsg As String
Dim vExcelData As Variant

'Get background view of active drawing sheet
Set oDwgDoc = CATIA.ActiveDocument
Set oBkgView = oDwgDoc.Sheets.ActiveSheet.Views.Item(2)

'Transform oExcelData in an array
If Not oExcelData Is Nothing Then
    vExcelData = TransformDataToArray(oExcelData)
End If

Select Case iActionNb
    
    Case 1
        Call DeleteTable(oTableColl)
        Call CreateTable(oBkgView, vExcelData, "ColorCode_Table", "COLOR CODING TABLE", oColorCodeList)
        Call AddToLogFile("Import Color Code Table", , oDwgInfo.DwgNumber, oDwgInfo.DwgRev, CInt(oDwgInfo.DwgIteration), , , "Overwrite Existing Table")

    Case 2
        Call RenameTable(oTableColl, "OLD_ColorCode_Table", "*** OLD COLOR CODING TABLE - TO DELETE ***")
        Call CreateTable(oBkgView, vExcelData, "ColorCode_Table", "COLOR CODING TABLE", oColorCodeList)
        Call AddToLogFile("Import Color Code Table", , oDwgInfo.DwgNumber, oDwgInfo.DwgRev, CInt(oDwgInfo.DwgIteration), , , "Rename Existing Table and create new one")

    Case 3
        sErrorMsg = ""
        Call GetTableFromDrawing(oDwgData, sErrorMsg, oTableColl)
        If sErrorMsg <> "" Then
            Call MsgBox(sErrorMsg, vbCritical, "Import Color Coding Table")
            Exit Sub
        End If

        'Compare the color table from Excel vs the one found in the CATDrawing
        sCompare = CompareData(oDwgData, oExcelData)
        
        'Table in Excel is the same as the one in the CATDrawing
        If sCompare = "OK" Then
            Call AddToLogFile("Import Color Code Table", , oDwgInfo.DwgNumber, oDwgInfo.DwgRev, CInt(oDwgInfo.DwgIteration), sCompare, , "Compare Table", oDwgData.XMLDoc.xml)
            Sleep 500
            sLog = ConfirmLogWasCreated(oDwgInfo, sCompare)
        'Table in Excel is firrefernt from the one in the CATDrawing
        Else
            Call AddToLogFile("Import Color Code Table", , oDwgInfo.DwgNumber, oDwgInfo.DwgRev, CInt(oDwgInfo.DwgIteration), sCompare, , "Compare Table")
        End If
        
        If sCompare = "OK" And sLog = "OK" Then
            sMsg = "Data in Excel is the same as existing Color Code Table in CATDrawing. "
            sMsg = sMsg & vbCrLf & "Compare table info was correctly populated in LOG database for iteration " & oDwgInfo.DwgIteration & ". "
            sMsg = sMsg & vbCrLf & "eChecker will be able to retrieve this information to validate DWG659. "
            Call MsgBox(sMsg, 64, "Import Color Coding Table")
        ElseIf sCompare = "OK" And sLog = "KO" Then
            sMsg = "Data in Excel is the same as existing Color Code Table in CATDrawing. "
            sMsg = sMsg & "However the compare table info was not correctly populated in LOG database for iteration " & oDwgInfo.DwgIteration & ". "
            sMsg = sMsg & vbCrLf & "This info is required by eChecker to validate DWG659."
            sMsg = sMsg & vbCrLf
            sMsg = sMsg & vbCrLf & "Please wait 10 minutes for database to refresh and run tool again using option 4."
            Call MsgBox(sMsg, 48, "Import Color Coding Table")

        Else
            Call CreateTable(oBkgView, vExcelData, "CHECK_ColorCode_Table", "*** FOR CHECKING PURPOSE ONLY - TO DELETE ***", oColorCodeList)
            Call MsgBox("Selected data in Excel is different from existing Color Code Table.", vbCritical, "Import Color Coding Table")
        End If

    Case 4
        sLog = ConfirmLogWasCreated(oDwgInfo, "OK")
       
        If sLog = "OK" Then
            sMsg = "Compare table info was correctly populated in LOG database for iteration " & oDwgInfo.DwgIteration & ". "
            sMsg = sMsg & vbCrLf & "eChecker will be able to retrieve this information to validate DWG659."
            Call MsgBox(sMsg, 64, "Import Color Coding Table")
        Else
            sMsg = "Compare table info was not correctly populated in LOG database for iteration " & oDwgInfo.DwgIteration & ". "
            sMsg = sMsg & vbCrLf & "This info is required by eChecker to validate DWG659."
            sMsg = sMsg & vbCrLf
            sMsg = sMsg & vbCrLf & "Please run tool again using option 3."
            Call MsgBox(sMsg, 48, "Import Color Coding Table")
       End If
    Case Else
        Exit Sub
End Select

End Sub

Private Function CompareData(ByVal oDwgData As clsXML, ByVal oExcelData As clsXML) As String
    
Dim i As Integer
Dim oDwgNode As IXMLDOMNode, oExcelNode As IXMLDOMNode
Dim oCopyDwgData As New clsXML, oCopyExcelData As New clsXML
Dim sQuery As String

'Copy data from original
Call oCopyDwgData.XMLDoc.loadXML(oDwgData.XMLDoc.xml)
Call oCopyExcelData.XMLDoc.loadXML(oExcelData.XMLDoc.xml)

'Scan all the nodes in oCopyDwgData and check if the same note exist in oExcelNode.
Do
    If oCopyDwgData.RootNode.childNodes.Length = 0 Then Exit Do
    
    For i = 1 To oDwgData.RootNode.childNodes.Length
        Set oDwgNode = oCopyDwgData.RootNode.childNodes.Item(0)
    
        sQuery = "./Part[@ColorCode='" & oDwgNode.Attributes.getNamedItem("ColorCode").nodeValue & "' and "
        sQuery = sQuery & " @PartNumber='" & oDwgNode.Attributes.getNamedItem("PartNumber").nodeValue & "' and "
        sQuery = sQuery & " @Qty='" & oDwgNode.Attributes.getNamedItem("Qty").nodeValue & "']"
        
        Set oExcelNode = oCopyExcelData.SelectSingleNode(sQuery, oCopyExcelData.RootNode)
        
        If Not oExcelNode Is Nothing Then
            Call oCopyExcelData.DeleteNode(oExcelNode)
            Call oCopyDwgData.DeleteNode(oDwgNode)
            Exit For
        Else
            Exit Do
        End If
    Next
Loop

'Check if there is remaining data in any of the two collections
If oCopyExcelData.RootNode.childNodes.Length = 0 And oCopyDwgData.RootNode.childNodes.Length = 0 Then
    CompareData = "OK"
Else
    CompareData = "KO"
End If

End Function

Private Sub GetTableFromDrawing(ByRef oDwgData As clsXML, ByRef sErrorMsg As String, ByVal oTableColl As Collection)

Dim i As Integer, j As Integer
Dim oTable As DrawingTable
Dim sColorCode As String, sRowType As String
Dim oItem As clsCollection
Dim sAttNames(1 To 4) As String, sAttValues(1 To 4) As String
Dim oNewElem As IXMLDOMElement

'Initial set
Set oDwgData = New clsXML
Call oDwgData.AddElement("Root")

'Initialize the attributes names to be added to oDwgData
sAttNames(1) = "ColorCode": sAttNames(2) = "PartNumber": sAttNames(3) = "Qty": sAttNames(4) = "Title"

'Scan all tables
For i = 1 To oTableColl.Count

    'Get table from collection
    Set oTable = oTableColl.Item(i)
    
    'Scan all rows in oTable
    sColorCode = ""
    For j = 1 To oTable.NumberOfRows
        
        If Trim(UCase(oTable.GetCellString(j, 1))) = "COLOR CODING TABLE" Then
            sRowType = "Header"
        ElseIf UCase(oTable.GetCellString(j, 1)) Like "SYSTEM*" And oTable.GetCellString(j, 2) = "" Then
            sColorCode = "SYSTEM"
            sRowType = "Color Code"
        ElseIf UCase(oTable.GetCellString(j, 1)) Like "*PART NUMBER*" And UCase(oTable.GetCellString(j, 2)) Like "*DESCRIPTION*" Then
            sRowType = "Sub Header"
        ElseIf oTable.GetCellString(j, 1) <> "" And oTable.GetCellString(j, 2) = "" Then
            sColorCode = UCase(oTable.GetCellString(j, 1))
            sRowType = "Color Code"
        Else
            sRowType = "Part"
        End If
        
        If sRowType = "Part" Then
        
            'Paste values in array
            sAttValues(1) = sColorCode: sAttValues(2) = oTable.GetCellString(j, 1): sAttValues(3) = CStr(oTable.GetCellString(j, 3)): sAttValues(4) = CStr(oTable.GetCellString(j, 2))
            
            'Add new element in oDwgData
            If Not oDwgData.Exists("./Part[@ColorCode='" & sAttValues(1) & "' and @PartNumber='" & sAttValues(2) & "']", oDwgData.RootNode) Then
                Set oNewElem = oDwgData.AddElement("Part", oDwgData.RootNode)
                Call oDwgData.AddMultipleAttributes(oNewElem, sAttNames, sAttValues)
            Else
                sErrorMsg = oTable.GetCellString(j, 1) & " with color code " & sColorCode & " is found more than one in the table. Please clean up and run the tool again."
                Exit Sub
            End If
        End If
    Next
Next

End Sub

Private Sub RenameTable(ByVal oTableColl As Collection, ByVal sNewTableName As String, ByVal sHeader As String)

Dim oTable As DrawingTable
Dim i As Integer

For i = 1 To oTableColl.Count

    Set oTable = oTableColl.Item(i)

    oTable.Name = sNewTableName
    Call oTable.SetCellString(1, 1, sHeader)
Next
    
End Sub

Private Sub DeleteTable(ByVal oTableColl As Collection)

Dim i As Integer
Dim oSelection

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

For i = 1 To oTableColl.Count
    oSelection.Add oTableColl.Item(i)
Next

oSelection.Delete
End Sub

Private Sub CreateTable(ByVal oView As DrawingView, ByVal oData As Variant, ByVal sTableName As String, ByVal sHeader As String, ByVal oColorCodeList As clsCollection)

Dim oDwgTable As DrawingTable
Dim sLogNumeric As String, sLogMissing As String, sMessage As String
Dim bNewColorCodeSection As Boolean
Dim oTableRow As DataTableRowInfo
Dim i As Integer

'Create Initial Table
Set oDwgTable = oView.Tables.Add(0, 0, 1, 4, 0.41 * 25.4, 1)
oDwgTable.Name = sTableName
oDwgTable.ComputeMode = CatTableComputeOFF

'Set colum width
Call oDwgTable.SetColumnSize(1, 3.5 * 25.4)
Call oDwgTable.SetColumnSize(2, 5.5 * 25.4)
Call oDwgTable.SetColumnSize(3, 0.5 * 25.4)
Call oDwgTable.SetColumnSize(4, 1.4 * 25.4)

'Table Header
oTableRow.RowNb = 1
oTableRow.RowHeight = 0.41
oTableRow.Text = sHeader
oTableRow.Frame = "0"
oTableRow.Alignment = "4" 'Center
oTableRow.FontSize = "0.16"
oTableRow.FontName = "SSS1"

Call AddRowToTable(oTableRow, oDwgTable)

sLogNumeric = ""
sLogMissing = ""
For i = LBound(oData, 1) To UBound(oData, 1)

    'Check if we start a new color code section
    If i = LBound(oData, 1) Then
        bNewColorCodeSection = True
    Else
        If oData(i, 1) = oData(i - 1, 1) Then
            bNewColorCodeSection = False
        Else
            bNewColorCodeSection = True
        End If
    End If
    
    'New Color Code Section
    If bNewColorCodeSection = True Then
        oTableRow.RowNb = oTableRow.RowNb + 1
        oTableRow.RowHeight = 0.75
        oTableRow.Alignment = "4" 'Center
        oTableRow.FontSize = "0.16"
        oTableRow.FontName = "SSS1"
        
        If oData(i, 1) = "SYSTEM" Then
            oTableRow.Text = "SYSTEM COLOR CODED PARTS (FROM SDD)"
            oTableRow.Frame = "0"
        Else
            oTableRow.Text = oData(i, 1)
            oTableRow.Frame = "9"
            
            'Check Color Code: Last two digits must be numeric
            If Not oTableRow.Text Like "*##" Then
                sLogNumeric = sLogNumeric & oTableRow.Text & "|"
            
            'Check Color Code: Other digits must be in the oColorCodeList
            Else
                If Not oColorCodeList.Exists(Left(oTableRow.Text, Len(oTableRow.Text) - 2)) Then
                    sLogMissing = sLogMissing & oTableRow.Text & "|"
                End If
            End If
        End If

        Call AddRowToTable(oTableRow, oDwgTable)

        oTableRow.RowNb = oTableRow.RowNb + 1
        oTableRow.RowHeight = 0.41
        oTableRow.Text = "PART NUMBER|DESCRIPTION|QTY|NOTES"
        oTableRow.Frame = "0|0|0|0"
        oTableRow.Alignment = "4|4|4|4" 'Center
        oTableRow.FontSize = "0.16|0.16|0.16|0.16"
        oTableRow.FontName = "SSS1|SSS1|SSS1|SSS1"
        
        Call AddRowToTable(oTableRow, oDwgTable)
    End If

    'Add row with part number, title and qty
    oTableRow.RowNb = oTableRow.RowNb + 1
    oTableRow.RowHeight = 0.41
    oTableRow.Text = oData(i, 2) & "|" & oData(i, 3) & "|" & oData(i, 4) & "| "
    oTableRow.Frame = "0|0|0|0"
    oTableRow.Alignment = "1|1|4|4"
    oTableRow.FontSize = "0.16|0.16|0.16|0.16"
    oTableRow.FontName = "SSS1|SSS1|SSS1|SSS1"
    
    Call AddRowToTable(oTableRow, oDwgTable)

Next

'Delete last row
oDwgTable.RemoveRow (oTableRow.RowNb + 1)
oDwgTable.ComputeMode = CatTableComputeON


'Display sLog
If sLogNumeric <> "" Then
    sMessage = "The last two digits of the following Color Code must be numeric:"
    
    For i = LBound(Split(sLogNumeric, "|")) To UBound(Split(sLogNumeric, "|")) - 1
        sMessage = sMessage & vbCrLf & " - " & Split(sLogNumeric, "|")(i)
    Next
    
    Call MsgBox(sMessage, vbCritical, "Import Color Coding Table")
    
End If

'Display sLog
If sLogMissing <> "" Then
    sMessage = "The following Color Code were not found in the list:"
    
    For i = LBound(Split(sLogMissing, "|")) To UBound(Split(sLogMissing, "|")) - 1
        sMessage = sMessage & vbCrLf & " - " & Split(sLogMissing, "|")(i)
    Next
    
    sMessage = sMessage & vbCrLf & vbCrLf & "Please contact KBE group to add a code in the list."
    Call MsgBox(sMessage, vbCritical, "Import Color Coding Table")
    
End If

End Sub

Private Sub AddRowToTable(ByRef oTableRow As DataTableRowInfo, ByVal oDwgTable As DrawingTable)
   
Dim i As Integer
Dim sString As String
Dim oDwgText As DrawingText

'Add row
Call oDwgTable.AddRow(oTableRow.RowNb)

'Merge cells if required
If LBound(Split(oTableRow.Text, "|")) = UBound(Split(oTableRow.Text, "|")) Then
    Call oDwgTable.MergeCells(oTableRow.RowNb, 1, 1, 4)
End If

'Set row height
oDwgTable.SetRowSize oTableRow.RowNb, oTableRow.RowHeight * 25.4

'Scan all columns
For i = LBound(Split(oTableRow.Text, "|")) To UBound(Split(oTableRow.Text, "|"))
    
    'Add Text in cell
    sString = Split(oTableRow.Text, "|")(i)
    Call oDwgTable.SetCellString(oTableRow.RowNb, i + 1, sString)
    
    'Get Text object
    Set oDwgText = oDwgTable.GetCellObject(oTableRow.RowNb, i + 1)
    
    'Add Frame
    sString = Split(oTableRow.Frame, "|")(i)
    oDwgText.SetParameterOnSubString catBorder, 0, 0, sString
    
    'Alignment
    sString = Split(oTableRow.Alignment, "|")(i)
    Call oDwgTable.SetCellAlignment(oTableRow.RowNb, i + 1, CInt(sString))
    
    'Font Size
    sString = Split(oTableRow.FontSize, "|")(i)
    oDwgText.SetFontSize 0, 0, CDbl(sString) * 25.4
    
    'Font Name
    sString = Split(oTableRow.FontName, "|")(i)
    oDwgText.SetFontName 0, 0, sString
Next

End Sub

Private Function ConfirmLogWasCreated(ByRef oDwgInfo As DwgInfo, ByVal sCompare As String) As String

Dim oColl 'As clsCollection : The clsCollection in WebService Access tool and DDP toolbar are different
Dim sReturn As String, sSQLIteration As String, sFctName As String, sLogStatus As String, sValue As String, sUserID As String
Dim i As Integer
Dim sUserCode As String

'Initialize
sReturn = "KO"

'Get user ID
If sUserCode = "" Then

    On Error Resume Next
    sUserCode = UCase(CATIA.SystemService.Environ("V5START_USERID"))
    On Error GoTo 0
    
    If sUserCode = "" Then
        Dim WSHnet
        Set WSHnet = CreateObject("WScript.Network")
        sUserCode = UCase(WSHnet.UserName)
        Set WSHnet = Nothing
    End If
End If

'Retrieve the log
Set oColl = Nothing
Set oColl = WebServiceAccessTool.GetToolLogUsage("LOG", _
                                            False, _
                                            Format(DateAdd("d", -90, Now), "YYYY-MM-dd"), _
                                            Format(DateAdd("d", 1, Now), "YYYY-MM-dd"), _
                                            oDwgInfo.DwgNumber, _
                                            oDwgInfo.DwgRev, _
                                            "BA_KBE_GCC")

'Scan all log items return from SQL search
If Not oColl Is Nothing Then
    For i = oColl.Count To 1 Step -1
    
        'Get info from SQL log item
        sSQLIteration = oColl.GetItem(i).GetItem("part_iter")
        sFctName = oColl.GetItem(i).GetItem("function_name")
        sValue = oColl.GetItem(i).GetItem("value")
        sLogStatus = oColl.GetItem(i).GetItem("log_status")
        sUserID = oColl.GetItem(i).GetItem("usr_bn")
    
        'Compare
        If sSQLIteration = oDwgInfo.DwgIteration And _
           sFctName = "Import Color Code Table" And _
           sValue = "Compare Table" And _
           sLogStatus = sCompare And _
           UCase(sUserID) = UCase(sUserCode) Then
           
           ConfirmLogWasCreated = "OK"
           Exit Function
        End If
    Next
End If

'Return
ConfirmLogWasCreated = sReturn

End Function





