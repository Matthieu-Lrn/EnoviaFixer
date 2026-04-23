Attribute VB_Name = "mdlLoadSaveDrawings"
Option Explicit
Private oDwgList As clsCollection
Private oAttList As clsAttributesList
Private sDestinationFolder As String
Private oLoadError As clsCollection
Private oSaveError As clsCollection

Private Sub CleanDrawingList()

Dim i As Integer
Dim sDwgNb As String
Dim sRev As String

'Remove from the list all the documents that matches "G########-###". I know these are not drawing.
For i = oDwgList.Count To 1 Step -1

    sDwgNb = oDwgList.GetItemByIndex(i).GetItemByKey("Document Number")
        
    If sDwgNb Like "G########-###" Then
        Call oDwgList.RemoveByIndex(i)
    End If
Next

'List is empty
If oDwgList.Count = 0 Then
    Call MsgBox("No drawing were found in the list. Process aborted", vbInformation)
    bCancelAction = True
End If
End Sub

Private Sub DisplayErrors()

Dim sMessage As String
Dim i As Integer

If oLoadError.Count > 0 Then

    sMessage = "The following documents could not be loaded:"
    
    For i = 1 To oLoadError.Count
        sMessage = sMessage & vbCrLf & " - " & Left(oLoadError.GetKey(i), Len(oLoadError.GetKey(i)) - 2) & " " & Right(oLoadError.GetKey(i), 2)
    Next

    Call MsgBox(sMessage, vbCritical)

End If



If oSaveError.Count > 0 Then

    sMessage = "The following documents could not be save:"
    
    For i = 1 To oLoadError.Count
        sMessage = sMessage & vbCrLf & " - " & Left(oSaveError.GetKey(i), Len(oSaveError.GetKey(i)) - 2) & " " & Right(oSaveError.GetKey(i), 2)
    Next

    Call MsgBox(sMessage, vbCritical)

End If


End Sub

Private Sub GetDrawingList()

Dim sPath As String
Dim sFileType As String

'Select file with drawing list
sPath = OpenFileDialog("Drawing List (*.xls;*.xlsx;*.xlsm;*.xml)|*.xls;*.xlsx;*.xlsm;*.xml", False, , sDestinationFolder, , "Select the file with the drawing list")
If sPath = "" Then
    bCancelAction = True
    Exit Sub
End If

'Get the drawing list
Select Case UCase(Right(sPath, 3))
    Case "XML"
        Call GetDrawingListFromXML(sPath)
    Case Else
        Call GetDrawingListFromExcel(sPath)
End Select

End Sub

Sub StartProcess()

Dim dSectionTimer As Double
Dim sAnswer As String
Dim sMessage As String

dSectionTimer = Timer

'Initialize
Set oDwgList = New clsCollection
Set oAttList = New clsAttributesList
Set oLoadError = New clsCollection
Set oSaveError = New clsCollection
bCancelAction = False

'Check Tools + Options
sMessage = "To speed up the process make sure Tools + Options + General + Load referenced documents is NOT selected."
sMessage = sMessage & vbCrLf & "Do you want to continue ?"
If MsgBox(sMessage, 36) = vbNo Then Exit Sub

'Get the destination folder
Call SelectDestinationFolder
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Get the list of CATDrawings
Call GetDrawingList
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Clean the drawing list. Remove everything that is not a drawing
Call CleanDrawingList
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Get the document attributes
Call GetDocumentAttributes
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Load / Save and Close CATDrawings
Call LoadAndSave
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Update Metadata
Call GenerateMetadataHTML
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Update DTExport
Call GenerateDTExportReport
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Display load error
Call DisplayErrors

'The end
endsub:
Unload frmProgress

Call MsgBox("Duration time is: " & Round(Timer - dSectionTimer, 5))

End Sub


Private Sub GetDrawingListFromExcel(ByVal sPath As String)

Dim oExcel As clsExcel
Dim iRow As Integer
Dim sDwgNb As String
Dim sRev As String
Dim oItem As clsCollection


'Open the Excel file
Set oExcel = New clsExcel
Call oExcel.GetExcel
Call oExcel.OpenExcelFile(sPath)

'No file was open
If oExcel.IsOpen = False Then
    bCancelAction = True
    Exit Sub
End If

'Get data from file
If oExcel.App.cells(2, 1).Value <> "" Then
    
    iRow = 2
    Do While oExcel.App.cells(iRow, 1).Value <> ""
    
        sDwgNb = oExcel.App.cells(iRow, 1).Value
        sRev = oExcel.App.cells(iRow, 2).Value
        
        'Add to collection
        If Not oDwgList.Exists(sDwgNb & sRev) Then
            Set oItem = New clsCollection
            Call oItem.Add("Document Number", sDwgNb)
            Call oItem.Add("Document Revision", sRev)
            Call oDwgList.Add(sDwgNb & sRev, oItem)
        End If
        
        iRow = iRow + 1
    Loop
Else
    Call oExcel.CloseExcelFile
    Call MsgBox("Excel file has the wrong format. Process aborted", vbCritical)
    bCancelAction = True
    Exit Sub
End If

'Close Excel file
Call oExcel.CloseExcelFile

Set oExcel = Nothing
End Sub


Private Sub GetDrawingListFromXML(ByVal sPath As String)

Dim sDwgNb As String
Dim sRev As String
Dim sString As String
Dim oItem As clsCollection
Dim oXMLDoc As DOMDocument60
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim dTimer As Double
Dim bTimeout As Boolean


'Open the xml file
dTimer = Timer
Set oXMLDoc = New DOMDocument60
Call oXMLDoc.Load(sPath)
bTimeout = False
Do
    DoEvents
    Sleep 150
    If Timer - dTimer > 30 Then bTimeout = True
Loop Until oXMLDoc.readyState = 4 Or bTimeout = True

'Check that the file was open
If bTimeout = True Then
    Call MsgBox("XML file could not be open in the 30sec limit. Macro aborted", vbCritical)
    bCancelAction = True
    Exit Sub
End If

'Search all drawing linked to a CI
Set oNodes = oXMLDoc.selectNodes(".//Part[@IsCI='True' and @SelectedDwg!='']")
For Each oNode In oNodes

    'Get info from XML
    sString = oNode.Attributes.getNamedItem("SelectedDwg").nodeValue
    
    If sString <> "" Then
        sDwgNb = Left(sString, Len(sString) - 2)
        sRev = Right(sString, 2)
    
        'Add to collection
        If Not oDwgList.Exists(sDwgNb & sRev) Then
            Set oItem = New clsCollection
            Call oItem.Add("Document Number", sDwgNb)
            Call oItem.Add("Document Revision", sRev)
            Call oDwgList.Add(sDwgNb & sRev, oItem)
        End If
    End If
Next

'Search all primary doc linked to a non-CI
Set oNodes = oXMLDoc.selectNodes(".//Part[@IsCI='False' and @Type='NONE']")
For Each oNode In oNodes

    'Get info from XML
    sString = oNode.Attributes.getNamedItem("PrimaryDocument").nodeValue
    
    If sString <> "" Then
        sDwgNb = Left(sString, Len(sString) - 2)
        sRev = Right(sString, 2)
    
        'Add to collection
        If Not oDwgList.Exists(sDwgNb & sRev) Then
            Set oItem = New clsCollection
            Call oItem.Add("Document Number", sDwgNb)
            Call oItem.Add("Document Revision", sRev)
            Call oDwgList.Add(sDwgNb & sRev, oItem)
        End If
    End If
Next

'Exit
Set oXMLDoc = Nothing

End Sub


Private Sub GetDocumentAttributes()

Dim sDwgNb As String
Dim sRev As String
Dim sExtension As String
Dim sOID As String
Dim oItem As clsCollection
Dim i As Integer
Dim iMax As Integer

Call frmProgress.progressBarInitialize("Saving drawings file base")

iMax = oDwgList.Count
For i = oDwgList.Count To 1 Step -1

    If bCancelAction = True Then Exit Sub

    'Get info from collection
    sDwgNb = oDwgList.GetItemByIndex(i).GetItemByKey("Document Number")
    sRev = oDwgList.GetItemByIndex(i).GetItemByKey("Document Revision")
    
    Call frmProgress.progressBarRepaint("Step 1 of 4 - Retrieving attributes", 4, 1, "Retrieving attributes " & sDwgNb & " " & sRev & " / " & (iMax - i + 1) & " of " & iMax, iMax, (iMax - i + 1))
    
    'Get the extension
    sExtension = oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "EXTENSION")
    
    'Get the OID
    sOID = oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "OID")
    
    'Add to collection
    If sExtension = "CATDrawing" Then
        Call oDwgList.GetItemByIndex(i).Add("Extension", sExtension)
        Call oDwgList.GetItemByIndex(i).Add("OID", sOID)
    
    'Remove non-CATDrawing from collection
    Else
        Call oDwgList.RemoveByIndex(i)
    End If
Next

End Sub

Private Sub LoadAndSave()

Dim i As Integer
Dim sDwgNb As String
Dim sRev As String
Dim sExtension As String
Dim sOID As String
Dim oDocument As DrawingDocument
Dim bAlreadyOpen As Boolean
Dim sIteration As String
Dim sMessage As String
Dim sAnswer As String

Call frmProgress.progressBarInitialize("Saving drawings file base")

For i = 1 To oDwgList.Count

    If bCancelAction = True Then Exit Sub

    'Get info from collection
    sDwgNb = oDwgList.GetItemByIndex(i).GetItemByKey("Document Number")
    sRev = oDwgList.GetItemByIndex(i).GetItemByKey("Document Revision")
    sExtension = oDwgList.GetItemByIndex(i).GetItemByKey("Extension")
    sOID = oDwgList.GetItemByIndex(i).GetItemByKey("OID")
    sIteration = oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "DOCUMENT_ITERATION")
    
    Call frmProgress.progressBarRepaint("Step 2 of 4 - Saving drawing", 4, 2, "Saving " & sDwgNb & " " & sRev & " / " & i & " of " & oDwgList.Count, oDwgList.Count, i)

    'Load document
    Call LoadDocFromEnovia(oDocument, bAlreadyOpen, sDwgNb, sRev, sExtension, sOID)
    
    'We have a document
    If Not oDocument Is Nothing Then
        On Error Resume Next
        Err.Clear
        Call oDocument.SaveAs(sDestinationFolder & sDwgNb & " " & sRev & "_" & sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            sMessage = "The tool can't save " & sDwgNb & " " & sRev & "_" & sIteration & "."
            sMessage = sMessage & vbCrLf & "Do you want to stop the process ?"
            sAnswer = MsgBox(sMessage, 20)
            
            'Exit process
            If sAnswer = vbYes Then
                bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call oSaveError.Add(sDwgNb & sRev, sDwgNb & " " & sRev & "_" & sIteration)
                Call oDocument.Close
            End If
        Else
            Call oDocument.Close
        End If
        On Error GoTo 0
    Else
        Call oLoadError.Add(sDwgNb & sRev, sDwgNb & sRev)
    End If
        
Next

End Sub

Private Sub GenerateMetadataHTML()

Dim sFilePath As String
Dim sEntireFile As String
Dim sTextPart1 As String
Dim sTextPart2 As String
Dim sDwgNb As String
Dim sRev As String
Dim i As Integer
Dim sAnswer As String
Dim oFSO
Dim oTextStream

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")

Do
    'User to select HTML template
    sFilePath = OpenFileDialog("HTML file (*.html)|*.html", , , sDestinationFolder, , "Select Metadata file to update")
    
    'No file selected
    If sFilePath = "" Then
        sAnswer = MsgBox("No file was selected. Do want to select again ?", 20)
        If sAnswer = vbNo Then
            bCancelAction = True
            Exit Sub
        Else
            GoTo SelectAgain
        End If
    End If
    
    'Read Metadata file
    Set oTextStream = oFSO.OpenTextFile(sFilePath, 1, False, 0)
    sEntireFile = oTextStream.ReadAll
    oTextStream.Close

    'Check the we have the right template
    If UBound(Split(sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")) > 0 Then Exit Do

    'Wrong template selected
    sAnswer = MsgBox("Looks like the wrong file was selected. Do want to select again ?", 20)
    
    'Exit
    If sAnswer = vbNo Then
        bCancelAction = True
        Exit Sub
    End If
    
SelectAgain:
Loop

'Split text
sTextPart1 = Split(sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(0) & "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>"
sTextPart2 = Split(sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(1)

'Full text is the first part
sEntireFile = sTextPart1

'Add all rows to full text
For i = 1 To oDwgList.Count

    If bCancelAction = True Then Exit Sub

    Call frmProgress.progressBarRepaint("Step 3 of 4 - Updating Metadata file", 4, 1, "Updating Metadata " & sDwgNb & " " & sRev & " / " & i & " of " & oDwgList.Count, oDwgList.Count, i)

    'The drawing info
    sDwgNb = oDwgList.GetItemByIndex(i).GetItemByKey("Document Number")
    sRev = oDwgList.GetItemByIndex(i).GetItemByKey("Document Revision")

    'Add drawing to HTML file
    Call AddNewLineToMetadata(sEntireFile, sDwgNb, sRev)

Next

'Add second part of the text
sEntireFile = sEntireFile & sTextPart2

'Save file
Set oTextStream = oFSO.OpenTextFile(sFilePath, 2, False, 0)
oTextStream.WriteLine sEntireFile
oTextStream.Close
Set oTextStream = Nothing
Set oFSO = Nothing


End Sub


Private Sub AddNewLineToMetadata(ByRef sFullText As String, ByVal sDwgNb As String, ByVal sRev As String)

Dim sReleaseDate, sFormatReleaseDate As String

'Get and format release date
sReleaseDate = ""
sReleaseDate = oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "REV_LAST_MOD_DATE")

If sReleaseDate <> "" Then
    sFormatReleaseDate = Left(sReleaseDate, 4)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(sReleaseDate, 5, 2)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(sReleaseDate, 7)
Else
    sFormatReleaseDate = ""
End If


'Add all attributes of a drawing
sFullText = sFullText & "<TR>"
sFullText = sFullText & "<A NAME=" & Chr(34) & sDwgNb & Chr(34) & ">"
sFullText = sFullText & "<TH>"
sFullText = sFullText & "<A HREF=" & Chr(34) & "#TOP" & Chr(34) & ">" & sDwgNb & "</A>"
sFullText = sFullText & "</TH>"
sFullText = sFullText & "</A>"
sFullText = sFullText & "<TD>"
sFullText = sFullText & "<PRE>Base Number          :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Base Number") & "</STRONG>"
sFullText = sFullText & "<BR>Dash Number          :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Dash Number") & "</STRONG>"
sFullText = sFullText & "<BR>Document Revision    :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "BA Document Revision") & "</STRONG>"
sFullText = sFullText & "<BR>Revision Status      :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Revision Status") & "</STRONG>"
sFullText = sFullText & "<BR>Title                :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Title") & "</STRONG>"
sFullText = sFullText & "<BR>Major Supplier Code  :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Major Supplier Code") & "</STRONG>"
sFullText = sFullText & "<BR>Release Date         :<STRONG> " & sFormatReleaseDate & "</STRONG>"
sFullText = sFullText & "<BR>Dataset Type         :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Dataset Type") & "</STRONG>"
sFullText = sFullText & "<BR>File Name            :<STRONG> " & sDwgNb & " " & sRev & ".CATDrawing</STRONG>"
sFullText = sFullText & "<BR>Organization(Part)   :<STRONG> " & "" & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Doc)    :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Revision Organization") & "</STRONG>"
sFullText = sFullText & "<BR>Shareable            :<STRONG> " & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Shareable") & "</STRONG>"

sFullText = sFullText & "</PRE>"
sFullText = sFullText & "</TD>"
sFullText = sFullText & "</TR>"



End Sub



Private Sub LoadDocFromEnovia(ByRef oDocument As DrawingDocument, ByRef bAlreadyOpen As Boolean, ByVal sPartNumber As String, ByVal sRevision As String, ByVal sExtension As String, ByVal sOID As String)

Dim dLoadTimer As Double
Dim EnoviaDoc As EnoviaDocument
Dim EV5product As Product

'Initialize
Set oDocument = Nothing
bAlreadyOpen = True
Set EnoviaDoc = CATIA.Application

'Check if document already loaded in session
On Error Resume Next
Set oDocument = CATIA.Documents.Item(sPartNumber & sRevision & "." & sExtension)
On Error GoTo 0

'Load document from ENOVIA
If oDocument Is Nothing Then
    
    'Initialize
    bAlreadyOpen = False

    'Load
    On Error Resume Next
    Err.Clear
    Set oDocument = EnoviaDoc.LoadWithOIDObjectType(sOID, "BA_CAD_DocRevision")

    'Make sure the document is loaded
    dLoadTimer = Timer
    Do

        'We have a document, everything is OK
        If Not oDocument Is Nothing Then Exit Do

        DoEvents
        Sleep 150
        If Abs(Timer - dLoadTimer) > 10 Then
            Exit Do
        End If
    Loop
    On Error GoTo 0
    
End If

End Sub

Private Sub SelectDestinationFolder()

Dim sAnswer As String
Do
    'Select destination folder
    sDestinationFolder = OpenDirectoryDialog("Select Destination Folder")
    
    'Check sDestinationFolder <> ""
    If sDestinationFolder = "" Then
        bCancelAction = True
        Exit Sub
    End If
    
    'Check destination folder doesn't contain a white space
    If sDestinationFolder Like "* *" Then
        sAnswer = MsgBox(sDestinationFolder & " is invalid. Please select a path without any white spaces.", 16, "Destination Folder")
        GoTo TryAgain
    End If

    'Destination folder is ok, exit
    sDestinationFolder = sDestinationFolder & "\"
    Exit Do
    
TryAgain:
Loop

End Sub

Private Sub GenerateDTExportReport()

Dim sPath As String
Dim sEntireFile As String
Dim sTextPart1 As String
Dim sTextPart2 As String
Dim sDwgNb As String
Dim sRev As String
Dim i As Integer
Dim sIteration As String
Dim sAnswer As String
Dim oFSO
Dim oTextStream

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")

Do
    'User to select DTExport file
    sPath = OpenFileDialog("HTML file (*.html)|*.html", , , sDestinationFolder, , "Select DTExport file to update")

    'No file selected
    If sPath = "" Then
        sAnswer = MsgBox("No file was selected. Do want to select again ?", 20)
        If sAnswer = vbNo Then
            bCancelAction = True
            Exit Sub
        Else
            GoTo SelectAgain
        End If
    End If
    
    'Read DTExport file
    Set oTextStream = oFSO.OpenTextFile(sPath, 1, False, 0)
    sEntireFile = oTextStream.ReadAll
    oTextStream.Close


    'Check the we have the right template
    If UBound(Split(sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")) > 0 Then Exit Do

    'Wrong template selected
    sAnswer = MsgBox("Looks like the wrong file was selected. Do want to select again ?", 20)
    
    'Exit
    If sAnswer = vbNo Then
        bCancelAction = True
        Exit Sub
    End If

SelectAgain:

Loop

'Split text
sTextPart1 = Split(sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(0) & "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>" & vbLf
sTextPart2 = Split(sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(1)

'Full text is the first part
sEntireFile = sTextPart1

'Add all rows to full text
For i = 1 To oDwgList.Count
        
    If bCancelAction = True Then Exit Sub
    
    Call frmProgress.progressBarRepaint("Step 3 of 4 - Updating Metadata file", 4, 1, "Updating Metadata " & sDwgNb & " " & sRev & " / " & i & " of " & oDwgList.Count, oDwgList.Count, i)
    
    'The drawing info
    sDwgNb = oDwgList.GetItemByIndex(i).GetItemByKey("Document Number")
    sRev = oDwgList.GetItemByIndex(i).GetItemByKey("Document Revision")
    sIteration = oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "DOCUMENT_ITERATION")
    
    'Add drawing to HTML file
    Call AddNewLineToDTExport(sEntireFile, sDwgNb, sRev, sIteration)
            
Next

'Add second part of the text
sEntireFile = sEntireFile & sTextPart2

'Save file
Set oTextStream = oFSO.OpenTextFile(sPath, 2, False, 0)
oTextStream.WriteLine sEntireFile
oTextStream.Close
Set oTextStream = Nothing
Set oFSO = Nothing

End Sub


Private Sub AddNewLineToDTExport(ByRef sFullText As String, ByVal sDwgNb As String, ByVal sRev As String, ByVal sIteration As String)

Dim sSecurityCheck As String

'Get the security check
sSecurityCheck = Right(oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Security Check"), 2)


'Add all attributes of a drawing
sFullText = sFullText & "<tr>"
sFullText = sFullText & "<td>" & sDwgNb & "</td>"
sFullText = sFullText & "<td>" & sRev & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "DOCUMENT_ITERATION") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Revision Status") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Revision Organization") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Revision Project") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Shareable") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sDwgNb, sRev, False, "Title") & "</td>"
sFullText = sFullText & "<td>" & sDwgNb & " " & sRev & "_" & sIteration & ".CATDrawing</td>"
sFullText = sFullText & "<td>" & Left(sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "<td>" & Right(sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "</tr>"
sFullText = sFullText & vbLf
End Sub

