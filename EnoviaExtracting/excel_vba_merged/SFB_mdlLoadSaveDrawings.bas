Attribute VB_Name = "SFB_mdlLoadSaveDrawings"
Option Explicit
Private SFB_oDwgList As SFB_clsCollection
Private SFB_oAttList As SFB_clsAttributesList
Private SFB_sDestinationFolder As String
Private SFB_oLoadError As SFB_clsCollection
Private SFB_oSaveError As SFB_clsCollection

Private Sub SFB_CleanDrawingList()

Dim SFB_i As Integer
Dim SFB_sDwgNb As String
Dim SFB_sRev As String

'Remove from the list all the documents that matches "G########-###". I know these are not drawing.
For SFB_i = SFB_oDwgList.Count To 1 Step -1

    SFB_sDwgNb = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Number")
        
    If SFB_sDwgNb Like "G########-###" Then
        Call SFB_oDwgList.RemoveByIndex(SFB_i)
    End If
Next

'List is empty
If SFB_oDwgList.Count = 0 Then
    Call MsgBox("No drawing were found in the list. Process aborted", vbInformation)
    SFB_bCancelAction = True
End If
End Sub

Private Sub SFB_DisplayErrors()

Dim SFB_sMessage As String
Dim SFB_i As Integer

If SFB_oLoadError.Count > 0 Then

    SFB_sMessage = "The following documents could not be loaded:"
    
    For SFB_i = 1 To SFB_oLoadError.Count
        SFB_sMessage = SFB_sMessage & vbCrLf & " - " & Left(SFB_oLoadError.GetKey(SFB_i), Len(SFB_oLoadError.GetKey(SFB_i)) - 2) & " " & Right(SFB_oLoadError.GetKey(SFB_i), 2)
    Next

    Call MsgBox(SFB_sMessage, vbCritical)

End If



If SFB_oSaveError.Count > 0 Then

    SFB_sMessage = "The following documents could not be save:"
    
    For SFB_i = 1 To SFB_oLoadError.Count
        SFB_sMessage = SFB_sMessage & vbCrLf & " - " & Left(SFB_oSaveError.GetKey(SFB_i), Len(SFB_oSaveError.GetKey(SFB_i)) - 2) & " " & Right(SFB_oSaveError.GetKey(SFB_i), 2)
    Next

    Call MsgBox(SFB_sMessage, vbCritical)

End If


End Sub

Private Sub SFB_GetDrawingList()

Dim SFB_sPath As String
Dim SFB_sFileType As String

'Select file with drawing list
SFB_sPath = SFB_OpenFileDialog("Drawing List (*.xls;*.xlsx;*.xlsm;*.xml)|*.xls;*.xlsx;*.xlsm;*.xml", False, , SFB_sDestinationFolder, , "Select the file with the drawing list")
If SFB_sPath = "" Then
    SFB_bCancelAction = True
    Exit Sub
End If

'Get the drawing list
Select Case UCase(Right(SFB_sPath, 3))
    Case "XML"
        Call SFB_GetDrawingListFromXML(SFB_sPath)
    Case Else
        Call SFB_GetDrawingListFromExcel(SFB_sPath)
End Select

End Sub

Sub SFB_StartProcess()

Dim SFB_dSectionTimer As Double
Dim SFB_sAnswer As String
Dim SFB_sMessage As String

SFB_dSectionTimer = Timer

'Initialize
Set SFB_oDwgList = New SFB_clsCollection
Set SFB_oAttList = New SFB_clsAttributesList
Set SFB_oLoadError = New SFB_clsCollection
Set SFB_oSaveError = New SFB_clsCollection
SFB_bCancelAction = False

'Check Tools + Options
SFB_sMessage = "To speed up the process make sure Tools + Options + General + Load referenced documents is NOT selected."
SFB_sMessage = SFB_sMessage & vbCrLf & "Do you want to continue ?"
If MsgBox(SFB_sMessage, 36) = vbNo Then Exit Sub

'Get the destination folder
Call SFB_SelectDestinationFolder
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Get the list of CATDrawings
Call SFB_GetDrawingList
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Clean the drawing list. Remove everything that is not SFB_a drawing
Call SFB_CleanDrawingList
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Get the document attributes
Call SFB_GetDocumentAttributes
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Load / Save and Close CATDrawings
Call SFB_LoadAndSave
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Update Metadata
Call SFB_GenerateMetadataHTML
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Update DTExport
Call SFB_GenerateDTExportReport
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If

'Display load error
Call SFB_DisplayErrors

'The end
endsub:
Unload SFB_frmProgress

Call MsgBox("Duration time is: " & Round(Timer - SFB_dSectionTimer, 5))

End Sub


Private Sub SFB_GetDrawingListFromExcel(ByVal SFB_sPath As String)

Dim SFB_oExcel As SFB_clsExcel
Dim SFB_iRow As Integer
Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_oItem As SFB_clsCollection


'Open the Excel file
Set SFB_oExcel = New SFB_clsExcel
Call SFB_oExcel.GetExcel
Call SFB_oExcel.OpenExcelFile(SFB_sPath)

'No file was open
If SFB_oExcel.IsOpen = False Then
    SFB_bCancelAction = True
    Exit Sub
End If

'Get data from file
If SFB_oExcel.App.cells(2, 1).Value <> "" Then
    
    SFB_iRow = 2
    Do While SFB_oExcel.App.cells(SFB_iRow, 1).Value <> ""
    
        SFB_sDwgNb = SFB_oExcel.App.cells(SFB_iRow, 1).Value
        SFB_sRev = SFB_oExcel.App.cells(SFB_iRow, 2).Value
        
        'Add to collection
        If Not SFB_oDwgList.Exists(SFB_sDwgNb & SFB_sRev) Then
            Set SFB_oItem = New SFB_clsCollection
            Call SFB_oItem.Add("Document Number", SFB_sDwgNb)
            Call SFB_oItem.Add("Document Revision", SFB_sRev)
            Call SFB_oDwgList.Add(SFB_sDwgNb & SFB_sRev, SFB_oItem)
        End If
        
        SFB_iRow = SFB_iRow + 1
    Loop
Else
    Call SFB_oExcel.CloseExcelFile
    Call MsgBox("Excel file has the wrong format. Process aborted", vbCritical)
    SFB_bCancelAction = True
    Exit Sub
End If

'Close Excel file
Call SFB_oExcel.CloseExcelFile

Set SFB_oExcel = Nothing
End Sub


Private Sub SFB_GetDrawingListFromXML(ByVal SFB_sPath As String)

Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_sString As String
Dim SFB_oItem As SFB_clsCollection
Dim SFB_oXMLDoc As DOMDocument60
Dim SFB_oNodes As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode
Dim SFB_dTimer As Double
Dim SFB_bTimeout As Boolean


'Open the xml file
SFB_dTimer = Timer
Set SFB_oXMLDoc = New DOMDocument60
Call SFB_oXMLDoc.Load(SFB_sPath)
SFB_bTimeout = False
Do
    DoEvents
    SFB_Sleep 150
    If Timer - SFB_dTimer > 30 Then SFB_bTimeout = True
Loop Until SFB_oXMLDoc.readyState = 4 Or SFB_bTimeout = True

'Check that the file was open
If SFB_bTimeout = True Then
    Call MsgBox("XML file could not be open in the 30sec limit. Macro aborted", vbCritical)
    SFB_bCancelAction = True
    Exit Sub
End If

'Search all drawing linked to SFB_a CI
Set SFB_oNodes = SFB_oXMLDoc.selectNodes(".//Part[@IsCI='True' and @SelectedDwg!='']")
For Each SFB_oNode In SFB_oNodes

    'Get info from XML
    SFB_sString = SFB_oNode.Attributes.getNamedItem("SelectedDwg").nodeValue
    
    If SFB_sString <> "" Then
        SFB_sDwgNb = Left(SFB_sString, Len(SFB_sString) - 2)
        SFB_sRev = Right(SFB_sString, 2)
    
        'Add to collection
        If Not SFB_oDwgList.Exists(SFB_sDwgNb & SFB_sRev) Then
            Set SFB_oItem = New SFB_clsCollection
            Call SFB_oItem.Add("Document Number", SFB_sDwgNb)
            Call SFB_oItem.Add("Document Revision", SFB_sRev)
            Call SFB_oDwgList.Add(SFB_sDwgNb & SFB_sRev, SFB_oItem)
        End If
    End If
Next

'Search all primary doc linked to SFB_a non-CI
Set SFB_oNodes = SFB_oXMLDoc.selectNodes(".//Part[@IsCI='False' and @Type='NONE']")
For Each SFB_oNode In SFB_oNodes

    'Get info from XML
    SFB_sString = SFB_oNode.Attributes.getNamedItem("PrimaryDocument").nodeValue
    
    If SFB_sString <> "" Then
        SFB_sDwgNb = Left(SFB_sString, Len(SFB_sString) - 2)
        SFB_sRev = Right(SFB_sString, 2)
    
        'Add to collection
        If Not SFB_oDwgList.Exists(SFB_sDwgNb & SFB_sRev) Then
            Set SFB_oItem = New SFB_clsCollection
            Call SFB_oItem.Add("Document Number", SFB_sDwgNb)
            Call SFB_oItem.Add("Document Revision", SFB_sRev)
            Call SFB_oDwgList.Add(SFB_sDwgNb & SFB_sRev, SFB_oItem)
        End If
    End If
Next

'Exit
Set SFB_oXMLDoc = Nothing

End Sub


Private Sub SFB_GetDocumentAttributes()

Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_sExtension As String
Dim SFB_sOID As String
Dim SFB_oItem As SFB_clsCollection
Dim SFB_i As Integer
Dim SFB_iMax As Integer

Call SFB_frmProgress.SFB_progressBarInitialize("Saving drawings file base")

SFB_iMax = SFB_oDwgList.Count
For SFB_i = SFB_oDwgList.Count To 1 Step -1

    If SFB_bCancelAction = True Then Exit Sub

    'Get info from collection
    SFB_sDwgNb = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Number")
    SFB_sRev = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Revision")
    
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 1 of 4 - Retrieving attributes", 4, 1, "Retrieving attributes " & SFB_sDwgNb & " " & SFB_sRev & " / " & (SFB_iMax - SFB_i + 1) & " of " & SFB_iMax, SFB_iMax, (SFB_iMax - SFB_i + 1))
    
    'Get the extension
    SFB_sExtension = SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "EXTENSION")
    
    'Get the OID
    SFB_sOID = SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "OID")
    
    'Add to collection
    If SFB_sExtension = "CATDrawing" Then
        Call SFB_oDwgList.GetItemByIndex(SFB_i).Add("Extension", SFB_sExtension)
        Call SFB_oDwgList.GetItemByIndex(SFB_i).Add("OID", SFB_sOID)
    
    'Remove non-CATDrawing from collection
    Else
        Call SFB_oDwgList.RemoveByIndex(SFB_i)
    End If
Next

End Sub

Private Sub SFB_LoadAndSave()

Dim SFB_i As Integer
Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_sExtension As String
Dim SFB_sOID As String
Dim SFB_oDocument As DrawingDocument
Dim SFB_bAlreadyOpen As Boolean
Dim SFB_sIteration As String
Dim SFB_sMessage As String
Dim SFB_sAnswer As String

Call SFB_frmProgress.SFB_progressBarInitialize("Saving drawings file base")

For SFB_i = 1 To SFB_oDwgList.Count

    If SFB_bCancelAction = True Then Exit Sub

    'Get info from collection
    SFB_sDwgNb = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Number")
    SFB_sRev = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Revision")
    SFB_sExtension = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Extension")
    SFB_sOID = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("OID")
    SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "DOCUMENT_ITERATION")
    
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 2 of 4 - Saving drawing", 4, 2, "Saving " & SFB_sDwgNb & " " & SFB_sRev & " / " & SFB_i & " of " & SFB_oDwgList.Count, SFB_oDwgList.Count, SFB_i)

    'Load document
    Call SFB_LoadDocFromEnovia(SFB_oDocument, SFB_bAlreadyOpen, SFB_sDwgNb, SFB_sRev, SFB_sExtension, SFB_sOID)
    
    'We have SFB_a document
    If Not SFB_oDocument Is Nothing Then
        On Error Resume Next
        Err.Clear
        Call SFB_oDocument.SaveAs(SFB_sDestinationFolder & SFB_sDwgNb & " " & SFB_sRev & "_" & SFB_sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            SFB_sMessage = "The tool can't save " & SFB_sDwgNb & " " & SFB_sRev & "_" & SFB_sIteration & "."
            SFB_sMessage = SFB_sMessage & vbCrLf & "Do you want to stop the process ?"
            SFB_sAnswer = MsgBox(SFB_sMessage, 20)
            
            'Exit process
            If SFB_sAnswer = vbYes Then
                SFB_bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call SFB_oSaveError.Add(SFB_sDwgNb & SFB_sRev, SFB_sDwgNb & " " & SFB_sRev & "_" & SFB_sIteration)
                Call SFB_oDocument.Close
            End If
        Else
            Call SFB_oDocument.Close
        End If
        On Error GoTo 0
    Else
        Call SFB_oLoadError.Add(SFB_sDwgNb & SFB_sRev, SFB_sDwgNb & SFB_sRev)
    End If
        
Next

End Sub

Private Sub SFB_GenerateMetadataHTML()

Dim SFB_sFilePath As String
Dim SFB_sEntireFile As String
Dim SFB_sTextPart1 As String
Dim SFB_sTextPart2 As String
Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_i As Integer
Dim SFB_sAnswer As String
Dim SFB_oFSO
Dim SFB_oTextStream

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")

Do
    'User to select HTML template
    SFB_sFilePath = SFB_OpenFileDialog("HTML file (*.html)|*.html", , , SFB_sDestinationFolder, , "Select Metadata file to update")
    
    'No file selected
    If SFB_sFilePath = "" Then
        SFB_sAnswer = MsgBox("No file was selected. Do want to select again ?", 20)
        If SFB_sAnswer = vbNo Then
            SFB_bCancelAction = True
            Exit Sub
        Else
            GoTo SelectAgain
        End If
    End If
    
    'Read Metadata file
    Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sFilePath, 1, False, 0)
    SFB_sEntireFile = SFB_oTextStream.ReadAll
    SFB_oTextStream.Close

    'Check the we have the right template
    If UBound(Split(SFB_sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")) > 0 Then Exit Do

    'Wrong template selected
    SFB_sAnswer = MsgBox("Looks like the wrong file was selected. Do want to select again ?", 20)
    
    'Exit
    If SFB_sAnswer = vbNo Then
        SFB_bCancelAction = True
        Exit Sub
    End If
    
SelectAgain:
Loop

'Split text
SFB_sTextPart1 = Split(SFB_sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(0) & "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>"
SFB_sTextPart2 = Split(SFB_sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(1)

'Full text is the first part
SFB_sEntireFile = SFB_sTextPart1

'Add all rows to full text
For SFB_i = 1 To SFB_oDwgList.Count

    If SFB_bCancelAction = True Then Exit Sub

    Call SFB_frmProgress.SFB_progressBarRepaint("Step 3 of 4 - Updating Metadata file", 4, 1, "Updating Metadata " & SFB_sDwgNb & " " & SFB_sRev & " / " & SFB_i & " of " & SFB_oDwgList.Count, SFB_oDwgList.Count, SFB_i)

    'The drawing info
    SFB_sDwgNb = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Number")
    SFB_sRev = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Revision")

    'Add drawing to HTML file
    Call SFB_AddNewLineToMetadata(SFB_sEntireFile, SFB_sDwgNb, SFB_sRev)

Next

'Add second part of the text
SFB_sEntireFile = SFB_sEntireFile & SFB_sTextPart2

'Save file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sFilePath, 2, False, 0)
SFB_oTextStream.WriteLine SFB_sEntireFile
SFB_oTextStream.Close
Set SFB_oTextStream = Nothing
Set SFB_oFSO = Nothing


End Sub


Private Sub SFB_AddNewLineToMetadata(ByRef sFullText As String, ByVal SFB_sDwgNb As String, ByVal SFB_sRev As String)

Dim SFB_sReleaseDate, sFormatReleaseDate As String

'Get and format release date
SFB_sReleaseDate = ""
SFB_sReleaseDate = SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "REV_LAST_MOD_DATE")

If SFB_sReleaseDate <> "" Then
    sFormatReleaseDate = Left(SFB_sReleaseDate, 4)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(SFB_sReleaseDate, 5, 2)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(SFB_sReleaseDate, 7)
Else
    sFormatReleaseDate = ""
End If


'Add all attributes of SFB_a drawing
sFullText = sFullText & "<TR>"
sFullText = sFullText & "<A NAME=" & Chr(34) & SFB_sDwgNb & Chr(34) & ">"
sFullText = sFullText & "<TH>"
sFullText = sFullText & "<A HREF=" & Chr(34) & "#TOP" & Chr(34) & ">" & SFB_sDwgNb & "</A>"
sFullText = sFullText & "</TH>"
sFullText = sFullText & "</A>"
sFullText = sFullText & "<TD>"
sFullText = sFullText & "<PRE>Base Number          :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Base Number") & "</STRONG>"
sFullText = sFullText & "<BR>Dash Number          :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Dash Number") & "</STRONG>"
sFullText = sFullText & "<BR>Document Revision    :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "BA Document Revision") & "</STRONG>"
sFullText = sFullText & "<BR>Revision Status      :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Revision Status") & "</STRONG>"
sFullText = sFullText & "<BR>Title                :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Title") & "</STRONG>"
sFullText = sFullText & "<BR>Major Supplier Code  :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Major Supplier Code") & "</STRONG>"
sFullText = sFullText & "<BR>Release Date         :<STRONG> " & sFormatReleaseDate & "</STRONG>"
sFullText = sFullText & "<BR>Dataset Type         :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Dataset Type") & "</STRONG>"
sFullText = sFullText & "<BR>File Name            :<STRONG> " & SFB_sDwgNb & " " & SFB_sRev & ".CATDrawing</STRONG>"
sFullText = sFullText & "<BR>Organization(Part)   :<STRONG> " & "" & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Doc)    :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Revision Organization") & "</STRONG>"
sFullText = sFullText & "<BR>Shareable            :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Shareable") & "</STRONG>"

sFullText = sFullText & "</PRE>"
sFullText = sFullText & "</TD>"
sFullText = sFullText & "</TR>"



End Sub



Private Sub SFB_LoadDocFromEnovia(ByRef SFB_oDocument As DrawingDocument, ByRef SFB_bAlreadyOpen As Boolean, ByVal SFB_sPartNumber As String, ByVal SFB_sRevision As String, ByVal SFB_sExtension As String, ByVal SFB_sOID As String)

Dim SFB_dLoadTimer As Double
Dim SFB_EnoviaDoc As EnoviaDocument
Dim SFB_EV5product As Product

'Initialize
Set SFB_oDocument = Nothing
SFB_bAlreadyOpen = True
Set SFB_EnoviaDoc = CATIA.Application

'Check if document already loaded in session
On Error Resume Next
Set SFB_oDocument = CATIA.Documents.Item(SFB_sPartNumber & SFB_sRevision & "." & SFB_sExtension)
On Error GoTo 0

'Load document from ENOVIA
If SFB_oDocument Is Nothing Then
    
    'Initialize
    SFB_bAlreadyOpen = False

    'Load
    On Error Resume Next
    Err.Clear
    Set SFB_oDocument = SFB_EnoviaDoc.LoadWithOIDObjectType(SFB_sOID, "BA_CAD_DocRevision")

    'Make sure the document is loaded
    SFB_dLoadTimer = Timer
    Do

        'We have SFB_a document, everything is OK
        If Not SFB_oDocument Is Nothing Then Exit Do

        DoEvents
        SFB_Sleep 150
        If Abs(Timer - SFB_dLoadTimer) > 10 Then
            Exit Do
        End If
    Loop
    On Error GoTo 0
    
End If

End Sub

Private Sub SFB_SelectDestinationFolder()

Dim SFB_sAnswer As String
Do
    'Select destination folder
    SFB_sDestinationFolder = SFB_OpenDirectoryDialog("Select Destination Folder")
    
    'Check SFB_sDestinationFolder <> ""
    If SFB_sDestinationFolder = "" Then
        SFB_bCancelAction = True
        Exit Sub
    End If
    
    'Check destination folder doesn't contain SFB_a white space
    If SFB_sDestinationFolder Like "* *" Then
        SFB_sAnswer = MsgBox(SFB_sDestinationFolder & " is invalid. Please select SFB_a path without any white spaces.", 16, "Destination Folder")
        GoTo TryAgain
    End If

    'Destination folder is ok, exit
    SFB_sDestinationFolder = SFB_sDestinationFolder & "\"
    Exit Do
    
TryAgain:
Loop

End Sub

Private Sub SFB_GenerateDTExportReport()

Dim SFB_sPath As String
Dim SFB_sEntireFile As String
Dim SFB_sTextPart1 As String
Dim SFB_sTextPart2 As String
Dim SFB_sDwgNb As String
Dim SFB_sRev As String
Dim SFB_i As Integer
Dim SFB_sIteration As String
Dim SFB_sAnswer As String
Dim SFB_oFSO
Dim SFB_oTextStream

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")

Do
    'User to select DTExport file
    SFB_sPath = SFB_OpenFileDialog("HTML file (*.html)|*.html", , , SFB_sDestinationFolder, , "Select DTExport file to update")

    'No file selected
    If SFB_sPath = "" Then
        SFB_sAnswer = MsgBox("No file was selected. Do want to select again ?", 20)
        If SFB_sAnswer = vbNo Then
            SFB_bCancelAction = True
            Exit Sub
        Else
            GoTo SelectAgain
        End If
    End If
    
    'Read DTExport file
    Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sPath, 1, False, 0)
    SFB_sEntireFile = SFB_oTextStream.ReadAll
    SFB_oTextStream.Close


    'Check the we have the right template
    If UBound(Split(SFB_sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")) > 0 Then Exit Do

    'Wrong template selected
    SFB_sAnswer = MsgBox("Looks like the wrong file was selected. Do want to select again ?", 20)
    
    'Exit
    If SFB_sAnswer = vbNo Then
        SFB_bCancelAction = True
        Exit Sub
    End If

SelectAgain:

Loop

'Split text
SFB_sTextPart1 = Split(SFB_sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(0) & "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>" & vbLf
SFB_sTextPart2 = Split(SFB_sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(1)

'Full text is the first part
SFB_sEntireFile = SFB_sTextPart1

'Add all rows to full text
For SFB_i = 1 To SFB_oDwgList.Count
        
    If SFB_bCancelAction = True Then Exit Sub
    
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 3 of 4 - Updating Metadata file", 4, 1, "Updating Metadata " & SFB_sDwgNb & " " & SFB_sRev & " / " & SFB_i & " of " & SFB_oDwgList.Count, SFB_oDwgList.Count, SFB_i)
    
    'The drawing info
    SFB_sDwgNb = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Number")
    SFB_sRev = SFB_oDwgList.GetItemByIndex(SFB_i).GetItemByKey("Document Revision")
    SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "DOCUMENT_ITERATION")
    
    'Add drawing to HTML file
    Call SFB_AddNewLineToDTExport(SFB_sEntireFile, SFB_sDwgNb, SFB_sRev, SFB_sIteration)
            
Next

'Add second part of the text
SFB_sEntireFile = SFB_sEntireFile & SFB_sTextPart2

'Save file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sPath, 2, False, 0)
SFB_oTextStream.WriteLine SFB_sEntireFile
SFB_oTextStream.Close
Set SFB_oTextStream = Nothing
Set SFB_oFSO = Nothing

End Sub


Private Sub SFB_AddNewLineToDTExport(ByRef sFullText As String, ByVal SFB_sDwgNb As String, ByVal SFB_sRev As String, ByVal SFB_sIteration As String)

Dim SFB_sSecurityCheck As String

'Get the security check
SFB_sSecurityCheck = Right(SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Security Check"), 2)


'Add all attributes of SFB_a drawing
sFullText = sFullText & "<tr>"
sFullText = sFullText & "<td>" & SFB_sDwgNb & "</td>"
sFullText = sFullText & "<td>" & SFB_sRev & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "DOCUMENT_ITERATION") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Revision Status") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Revision Organization") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Revision Project") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Shareable") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sDwgNb, SFB_sRev, False, "Title") & "</td>"
sFullText = sFullText & "<td>" & SFB_sDwgNb & " " & SFB_sRev & "_" & SFB_sIteration & ".CATDrawing</td>"
sFullText = sFullText & "<td>" & Left(SFB_sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "<td>" & Right(SFB_sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "</tr>"
sFullText = sFullText & vbLf
End Sub

