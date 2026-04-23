Attribute VB_Name = "BA_KBE_GCC_COMMON"
'*****************************************************************************************************
'                     SCRIPT VERSION
'*****************************************************************************************************
Private Const sCommonModuleVersion As String = "KBE_GCC_COMMON_MAY18_2016"
'*****************************************************************************************************

'*****************************************************************************************************
'                     PATHS FILE
'*****************************************************************************************************
Public sKBEPathFile As String
'*****************************************************************************************************

#If Vba7 Then
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare PtrSafe Function GetCursor Lib "user32" () As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long '***** To verify if a window is minimized
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare PtrSafe Function FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpclassname As Any, ByVal lpCaption As Any) '***** To give title bar a toolbar look, used also to call OpenFileDialog to give window handler
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '***** To give title bar a toolbar look
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long '***** To give title bar a toolbar look
    Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long '***** To find screen size, and if user has 1 or 2 screens (to position toolbar inside screen)
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer '***** To detect ESC key press
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long) '***** To make a pause when running script
    Public Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long  'To browse a directory
    Public Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'To browse a directory
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'To browse a directory
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long 'To browse for a file
    Public Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long 'To browse for a file
    Public Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long 'To get user ID
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '***** Used in Manage Data (XML) script
#Else
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare Function GetCursor Lib "user32" () As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long '***** To set cursor shape as a "Hand" when a selection is needed
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long '***** To verify if a window is minimized
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare Function FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpclassname As Any, ByVal lpCaption As Any) '***** To give title bar a toolbar look, used also to call OpenFileDialog to give window handler
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '***** To give title bar a toolbar look
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long '***** To give title bar a toolbar look
    Public Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long '***** To find screen size, and if user has 1 or 2 screens (to position toolbar inside screen)
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer '***** To detect ESC key press
    Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long) '***** To make a pause when running script
    Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long  'To browse a directory
    Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'To browse a directory
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'To browse a directory
    Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long 'To browse for a file
    Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long 'To browse for a file
    Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long 'To get user ID
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '***** Used in Manage Data (XML) script
#End If

Public Const SW_SHOWNORMAL = 1  '***** Works with function "ShowWindow" > "half" maximize window
Public Const SW_MAXIMIZE = 3    '***** Works with function & "ShowWindow" > maximize window
Public Const SW_MINIMIZE = 6    '***** Works with function & "ShowWindow" > minimize window (could be 2)
Public Const SW_RESTORE = 9     '***** Works with function "ShowWindow" > restore window

Public Const PI As Double = 3.14159265358979 '*** value of Pi
Public Const SMALLNUMBER As Double = 0.00000001 '*** to avoid division by zero

Private Type OpenFilename 'To browse for a file or save a file
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BROWSEINFO 'To browse a directory
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public dPointsToPixelRatioH As Double
Public dPointsToPixelRatioV As Double
Public dTitleBarHeight As Double        '***** Used to determine if Windows classic or XP Style
Public sTempDirectory As String         '***** Temp Directory
Public sSettingsPath As String          '***** Saved Settings
Public bToolbarCommandSelected As Boolean   '***** Used to determine if user selected clicks on toolbar
Public bStopMultiSelection As Boolean       '***** Used to determine if user cancelled multi selection

Public Function setCommonVariables()
    
    CATIA.Caption = "CATIA V5 R" & CATIA.SystemConfiguration.Release & " SP" & CATIA.SystemConfiguration.ServicePack
    CATIA.RefreshDisplay = True
    CATIA.Interactive = True
        
    '*****Restore Catia window if minimized
    If IsIconic(FindWindow(vbNullString, CATIA.Caption)) Then ShowWindow FindWindow(vbNullString, CATIA.Caption), SW_RESTORE
    
    '*****SystemMetrics API is in pixels
    '*****Application (CATIA) left and top is in pixels
    '*****Userform left and top is in points
    '****************************************************
    '*****Get the conversion of points/pixel for H & V
    dPointsToPixelRatioH = 72 / GetDeviceCaps(GetDC(0), 88) '***** 72 points per inch / ??? pixels per inch
    dPointsToPixelRatioV = 72 / GetDeviceCaps(GetDC(0), 90) '***** 72 points per inch / ??? pixels per inch
    
    '*****Determine if user has Windows Classic or XP Style
    dTitleBarHeight = GetSystemMetrics32(31) 'Returns 18 Windows Classic Mode / 25 Windows XP Mode
    
    '***** Temp Directory (%tmp% or %temp%), saved settings path and Common Module name
    sTempDirectory = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\"
    sSettingsPath = sTempDirectory & GetTagValueFromFile(sKBEPathFile, "Saved Settings")

    
End Function


'********************************************************************************
'* Name: ReadTextFile
'* Purpose: Read a text file, then return the value found on line "iLineIndex"
'*          If "iLineIndex = 0" (Default value), function returns the value of last line
'*          If "iLineIndex" is greater than the number of lines, function returns the value of last line
'*          If "bCompleteText" is = true, function returns the complete text in the file, all lines separated by a return (vbcrlf)
'*          NOTE: Two strings separated by a coma "," is considered on two different lines!!!
'*
'* Assumption:
'*
'* Author: http://www.vbforums.com/showthread.php?342619-Classic-VB-How-can-I-read-write-a-text-file
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function ReadTextFile(ByVal sFilePath As String, Optional iLineIndex As Long = 0, Optional bCompleteText As Boolean = False) As String

    Dim FileNumber As Integer
    Dim iLineCount As Long
    Dim iLineText As String

    ' ensure that the file exists
    ReadTextFile = ""
    If Len(Dir$(sFilePath)) = 0 Then Exit Function
    
    'Open file
    FileNumber = FreeFile
    Open sFilePath$ For Input As #FileNumber

    'Read lines
    iLineCount = 0
    Do While Not EOF(FileNumber)

        iLineCount = iLineCount + 1
        Input #FileNumber, iLineText

        If bCompleteText Then
            If Trim(iLineText) <> "" Then
                ReadTextFile = IIf(ReadTextFile = "", iLineText, ReadTextFile & vbCrLf & iLineText)
            End If
        Else
            ReadTextFile = Trim(iLineText)
        End If

        If iLineCount = iLineIndex Then Exit Do

    Loop

    'Close the file
    Close #FileNumber

End Function


'********************************************************************************
'* Name: WriteTextFile
'* Purpose: Add a line at the end of a text file
'*
'* Assumption:
'*
'* Author: http://www.devhut.net/2011/06/06/vba-append-text-to-a-text-file/
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Sub WriteTextFile(ByVal sFilePath As String, Optional ByVal sLineText As String = "", Optional ByVal bResetFile = True)

    Dim FileNumber As Integer

    'Exit if folder does not exist
    If Len(Dir$(Left(sFilePath, InStrRev(sFilePath, "\")))) = 0 Then Exit Sub
    
    On Error Resume Next
    FileNumber = FreeFile                                               ' Get unused file number
    If Not (bResetFile) Then Open sFilePath For Append As #FileNumber   ' Connect to the file
    If bResetFile Then Open sFilePath For Output As #FileNumber         ' Connect to the file
    If sLineText <> "" Then Print #FileNumber, sLineText                ' Write string
    Close #FileNumber                                                   ' Close the file
    On Error GoTo 0
End Sub


'********************************************************************************
'* Name: TrimLine
'* Purpose: In a string, keep only what's found after the first ">", remove all
'*          spaces and all tabs. Replace all "\" by "/"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function TrimLine(ByVal sLineText As String) As String

    TrimLine = sLineText
    TrimLine = Mid(TrimLine, InStr(sLineText, ">") + 1) '***** Getting everything after the first ">"
    TrimLine = Replace(TrimLine, vbTab, "") '***** Remove all tabs
    TrimLine = Replace(TrimLine, "/", "\") '***** Replace "/" by "\" in string
    TrimLine = Trim(TrimLine) '***** Remove spaces before and after string

End Function


'********************************************************************************
'* Name: Get Tag Value From File
'* Purpose: In a file, returns Value associated to a tag , that is, value after
'*          identifier between <""> is returned. sFilepath is the path of text file to read
'*          if no tag value is found then "" is returned.
'*          <  XX XX  > are equal to <XX XX>
'*          tags are not case sensitive.
'*
'* Assumption: if more than one tag exist in a file then first tag value is returned
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function GetTagValueFromFile(ByVal sFilePath As String, ByVal sTag As String) As String

Dim sTagtosearch As String
Dim sEntireFile As String
Dim aEntireFile() As String

GetTagValueFromFile = ""
sEntireFile = ReadTextFile(sFilePath, 0, True)      '**** read entire file
aEntireFile() = Split(sEntireFile, vbNewLine)

On Error Resume Next '**** to handle lines where no tags are found
For i = 0 To UBound(aEntireFile)
    sTagtosearch = aEntireFile(i)
    If InStr(sTagtosearch, "<") > 0 And InStr(sTagtosearch, ">") > 0 Then
        sTagtosearch = Mid(sTagtosearch, InStr(sTagtosearch, "<") + 1) '***** Getting everything after the first "<"
        sTagtosearch = Mid(sTagtosearch, 1, InStr(sTagtosearch, ">") - 1)  '***** Getting everything before the first ">"
        sTagtosearch = Trim(sTagtosearch)
        If sTagtosearch = sTag Then
            GetTagValueFromFile = TrimLine(aEntireFile(i))
            On Error GoTo 0
            Exit Function
        End If
    End If
Next
On Error GoTo 0
End Function


'********************************************************************************
'* Name: Set Tag Value From File
'* Purpose: In a file, replace Value associated to a tag , that is, value after
'*          identifier between <""> is returned. sFilepath is the path of text file to read
'*          if no tag value is found then "" is returned.
'*          <  XX XX  > are equal to <XX XX>
'*          tags are not case sensitive.
'*
'* Assumption: if more than one tag exist in a file then first tag value is returned
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SetTagValueToFile(ByVal sFilePath As String, ByVal sTag As String, ByVal sNewTagValue As String)

Dim bTagFound As Boolean
Dim sTagtosearch As String
Dim sEntireFile As String
Dim aEntireFile() As String

bTagFound = False
sEntireFile = ReadTextFile(sFilePath, 0, True)      '**** read entire file
aEntireFile() = Split(sEntireFile, vbNewLine)
Call WriteTextFile(sFilePath) '**** Empty file content

On Error Resume Next '**** to handle lines where no tags are found
For i = LBound(aEntireFile) To UBound(aEntireFile)
    sTagtosearch = aEntireFile(i)
    sTagtosearch = Mid(sTagtosearch, InStr(sTagtosearch, "<") + 1) '***** Getting everything after the first "<"
    sTagtosearch = Mid(sTagtosearch, 1, InStr(sTagtosearch, ">") - 1)  '***** Getting everything before the first ">"
    sTagtosearch = Trim(sTagtosearch)
    If UCase(sTagtosearch) Like "*" & UCase(sTag) & "*" Then
        bTagFound = True
        aEntireFile(i) = "<" & sTag & ">" & vbTab & vbTab & TrimLine(sNewTagValue)
    End If
    Call WriteTextFile(sFilePath, aEntireFile(i), False)
Next

If Not (bTagFound) Then Call WriteTextFile(sFilePath, "<" & sTag & ">" & vbTab & vbTab & TrimLine(sNewTagValue), False)
On Error GoTo 0

End Sub


'********************************************************************************
'* Name: Open With Notepad
'* Purpose: Open a file using Notepad
'*
'* Assumption:
'*
'* Author: François Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function OpenWithNotepad() As Long
    
    sWIN_DIR = CATIA.SystemService.Environ("windir")
    
    On Error Resume Next
    Err.Clear
    Shell sWIN_DIR & "\NOTEPAD.EXE " & sLogFileName, SW_SHOWNORMAL
    OpenWithNotepad = Err.Number
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: KBE Search
'* Purpose: Query (Catia Search) accordind to criteria, using HSO Synchronize property to increase performance
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function KBESearch(ByVal StringToSearch As String, Optional ByVal iObject As AnyObject = Nothing, Optional oCurrentSelection As Selection = Nothing)
    
    If oCurrentSelection Is Nothing Then Set oCurrentSelection = CATIA.ActiveDocument.Selection
    
    If Not iObject Is Nothing Then
        oCurrentSelection.Clear
        oCurrentSelection.Add iObject
    End If
    
    Err.Clear
    On Error Resume Next
    CATIA.HSOSynchronized = False
    oCurrentSelection.Search StringToSearch
    CATIA.HSOSynchronized = True
    On Error GoTo 0
    
    Set KBESearch = oCurrentSelection
    
End Function


'********************************************************************************
'* Name: Multi Window Select Element
'* Purpose: Loop until user does a selection or cancel action
'*
'* Assumption: another sub can set bStopMultiSelection to false during loop, so the returned selection is empty
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function MultiWindowSelectElement(ByVal selectionMessage As String)
    
    On Error Resume Next '***** Error if user select something else than a feature (e.g. clicking on toolbar)
    CATIA.ActiveDocument.Selection.Clear
    
    Set MultiWindowSelectElement = Nothing
    bStopMultiSelection = False
    bToolbarCommandSelected = False
    
    Do
        DoEvents
        SetHandCursor
        CATIA.StatusBar = selectionMessage
        If CATIA.ActiveDocument.Selection.Count2 > 0 Then
            Set MultiWindowSelectElement = CATIA.ActiveDocument.Selection.Item2(1)
            Exit Do
        End If
        If CATIA.ActiveDocument Is Nothing Then bStopMultiSelection = True
        If GetAsyncKeyState(vbKeyEscape) <> 0 Then bStopMultiSelection = True
    Loop Until bStopMultiSelection Or bToolbarCommandSelected
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Set Hand Cursor
'* Purpose: Modify shape of cursor to a hand
'*          This is used with multi selection (previous sub) so user knows that he needs to select something
'*
'* Assumption:
'*
'* Author:
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SetHandCursor()
    
    Dim HandCursor As Long
    On Error Resume Next
        If HandCursor = 0 Then
            HandCursor = LoadCursor(0, 32649&)
        ElseIf GetCursor() = HandCursor Then
            Exit Function
        End If
        SetCursor HandCursor
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Disable Option Name Check
'* Purpose: If name check is enabled, this sub disable name check in options (CATSettings)
'*          See "Tools / Options / Part Infrastructure / Display / Checking Operation When Renaming"
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function disableOptionNameCheck()
    
    Dim eNamingCheck As PartInfrastructureSettingAtt
    Set eNamingCheck = CATIA.SettingControllers.Item("CATMmuPartInfrastructureSettingCtrl")
    
    If eNamingCheck.NamingMode <> catNoNamingCheck Then
        eNamingCheck.NamingMode = catNoNamingCheck
        MsgBox "Name check was disabled. To change this setting again:  " & vbCrLf & _
                "Tools / Options / Part Infrastructure / Display / Checking Operation When Renaming        ", _
                vbExclamation, "Name check disabled"
    End If
    
    eNamingCheck.SaveRepository
    eNamingCheck.Commit
    
End Function


'********************************************************************************
'* Name: Enable Relation update
'* Purpose: If name formula update/synchronous are disabled, this sub enable both options (CATSettings)
'*          See "Tools / Options / Parameters and Measure / Knowledge / Relations update in part context"
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function enableRelationUpdate()
    
    Dim eRelationUpdate As KnowledgeSheetSettingAtt
    Set eRelationUpdate = CATIA.SettingControllers.Item("CATLieKnowledgeSheetSettingCtrl")
    
    If eRelationUpdate.RelationsUpdateInPartContextEvaluateDuringUpdate <> 1 Then eRelationUpdate.RelationsUpdateInPartContextEvaluateDuringUpdate = 1
    If eRelationUpdate.RelationsUpdateInPartContextSynchronousRelations <> 1 Then eRelationUpdate.RelationsUpdateInPartContextSynchronousRelations = 1
    
    eRelationUpdate.SaveRepository
    eRelationUpdate.Commit
    
End Function


'********************************************************************************
'* Name: Clean empty publication
'* Purpose: Delete a publication that is not linked to any element
'*          e.g. A published element was deleted, but not it's publication
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub deleteBrokenPublications(ByVal oDocumentToClean As Document, Optional ByVal sProgressBarCaption As String = "Clean Publications", Optional ByVal bProgressMain As Boolean = False)
    
    Dim sPublishedName As String
    
    On Error Resume Next
    Set oPublications = oDocumentToClean.Product.Publications
    If frmProgress.Visible = False Then frmProgress.progressBarInitialize (sProgressBarCaption)
    
    iProgressCount2 = 0
    iProgressCount2Max = oDocumentToClean.Product.Publications.Count
    
    For i = oPublications.Count To 1 Step -1
        
        iProgressCount2 = iProgressCount2 + 1
        sProgBarComment = "Cleaning publication (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        If bProgressMain = False Then Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount2Max, iProgressCount2 + 1 - i, , , , frmProgress.lblTimer.Caption)
        If bProgressMain = True Then Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, frmProgress.pbProgressMain.Max, frmProgress.pbProgressMain.Value, sProgBarComment, iProgressCount2Max, iProgressCount2 + 1 - i, frmProgress.lblTimer.Caption)
        
        sPublishedName = ""
        sPublishedName = oPublications.Item(i).Valuation.Name 'Not valuated for an empty published parameter
        sPublishedName = oPublications.Item(i).Valuation.DisplayName 'DisplayName = "" for an empty published feature or body
        
        If sPublishedName = "" Then 'If Valuation.Name was not valuated (empty parameter) or Valuation.DisplayName = "" (empty feature)
            oPublications.Remove (oPublications.Item(i).Name)
        End If
    Next
    
    '***** Tree may not refresh, and publication could stay visible in tree
    '***** even if they are deleted (bug in Part Mode only)
    '***** Trying to publish again the MainBody refreshes the tree
    If TypeName(oDocumentToClean) = "PartDocument" Then
        oPublications.Add (oDocumentToClean.Part.MainBody.Name)
        oPublications.SetDirect oDocumentToClean.Part.MainBody.Name, _
        oDocumentToClean.Product.CreateReferenceFromName(oDocumentToClean.Part.Name & "/!" & oDocumentToClean.Part.MainBody.Name)
    End If
    
    On Error GoTo 0
    
End Sub


'********************************************************************************
'* Name: Select Assembly
'* Purpose: Ask user to select an assembly, and verifies if its name contains a string "productName" (arguments)
'*          If the assy has the wrong name, it returns Nothing
'*          If the assy has the good name, the active workbench is set to "Assembly Design"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub selectAssy(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal productName As String = "*", Optional ByVal userMessage As String = "")
      
    Dim sDocumentType As String
    Dim oSelectedItem As AnyObject
    On Error Resume Next
    
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    If TypeName(oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    '***** if we cannot access PartNumber property, put oSelectedItem.LeafProduct in default mode
    If oSelectedItem.LeafProduct.PartNumber = "" Then
        oSelectedItem.LeafProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    sDocumentType = ""
    sDocumentType = TypeName(oSelectedItem.Value.ReferenceProduct.Parent)
    If sDocumentType = "ProductDocument" Then
        Set selectedProduct = oSelectedItem.Value
    ElseIf sDocumentType = "PartDocument" Then
        If TypeName(oSelectedItem.Value.Parent.Parent.ReferenceProduct.Parent) = "ProductDocument" Then
            Set selectedProduct = oSelectedItem.Value.Parent.Parent
        End If
    Else
        sDocumentType = TypeName(oSelectedItem.LeafProduct.ReferenceProduct.Parent)
        If sDocumentType = "PartDocument" Then
            sDocumentType = TypeName(oSelectedItem.LeafProduct.Parent.Parent.ReferenceProduct.Parent)
            If sDocumentType = "ProductDocument" Then
                Set selectedProduct = oSelectedItem.LeafProduct.Parent.Parent
            End If
        End If
    End If
    
    If Not selectedProduct Is Nothing And selectedProduct.PartNumber Like productName Then
        '****If name contains "productName" and selection is not a part instance
        Set selectedRootDocument = oSelectedItem.Document
        CATIA.ActiveDocument.Selection.Clear
        
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
        If CATIA.GetWorkbenchId <> "Assembly" Then
            CATIA.StartWorkbench "Assembly"
        End If
    Else
        Set selectedProduct = Nothing
        Set selectedRootDocument = Nothing
        CATIA.ActiveDocument.Selection.Clear
    End If
    
End Sub


'********************************************************************************
'* Name: Select Root Assembly
'* Purpose: Ask user to select any element. If the selected element is a product
'*          with the name "productName", this product is returned. Otherwise, the
'*          root product is found, and script verifies if its name contains a string
'*          "productName" (arguments)
'*          If the assy has the wrong name, it returns Nothing
'*          If the assy has the good name, the active workbench is set to "Assembly Design"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub selectRootAssy(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal productName As String = "*", Optional ByVal userMessage As String = "")
    
    Dim oSelectedItem As AnyObject
    On Error Resume Next
    
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    If TypeName(oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    Set selectedRootDocument = Nothing
    Set selectedProduct = Nothing
    Set selectedProduct = oSelectedItem.LeafProduct.ReferenceProduct 'Returns nothing if part is not in design mode, Returns nothing if we are in a part (not in product mode)
    
    'If a product "productName" was directly selected by user, return it (don't return Root product, but selected product)
    If Not selectedProduct Is Nothing Then
        If Not oSelectedItem.LeafProduct.PartNumber Like productName Or TypeName(selectedProduct.Parent) <> "ProductDocument" Then
            Set selectedProduct = Nothing
        End If
    End If
    
    'Else, return Root product
    If selectedProduct Is Nothing Then
        If oSelectedItem.Document.Product.Name Like productName And TypeName(oSelectedItem.Document) = "ProductDocument" Then
            Set selectedProduct = oSelectedItem.Document.Product
        Else
            Set selectedProduct = Nothing
        End If
    End If
    
    CATIA.ActiveDocument.Selection.Clear
    
    If Not selectedProduct Is Nothing Then
        Set selectedRootDocument = oSelectedItem.Document
        
        '****If name contains SEED_ASSY and selection is not a part instance
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
        If CATIA.GetWorkbenchId <> "Assembly" Then
            CATIA.StartWorkbench "Assembly"
        End If
    End If
    
End Sub


'********************************************************************************
'* Name: Select Part
'* Purpose: Ask user to select a part, and verifies if its name contains a string "partName" (arguments)
'*          If the part has the wrong name, it returns Nothing
'*          Note: user can select any element inside the part, or its instance, this will find the part
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub selectPart(ByRef selectedPart As Part, ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal PartName As String = "*", Optional ByVal userMessage As String = "")
    
    Dim oSelectedItem As AnyObject
    On Error Resume Next
    
    ' Initialize variables
    Set selectedPart = Nothing
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    
    ' Valuate selectedPart, selectedProduct and selectedRootDocument (depending if we are in product or part mode)
    If TypeName(oSelectedItem.Document) = "ProductDocument" Then
        ' Put part in design mode if needed (works only in product mode)
        If TypeName(oSelectedItem.LeafProduct.ReferenceProduct.Parent) = "PartDocument" Then
            oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        End If
        
        Set selectedPart = oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part 'Works only in product mode.
        Set selectedProduct = oSelectedItem.LeafProduct
    ElseIf TypeName(oSelectedItem.Document) = "PartDocument" Then
        Set selectedPart = oSelectedItem.Document.Part 'Works only in part mode.
        Set selectedProduct = oSelectedItem.Document.Product
    End If
    Set selectedRootDocument = oSelectedItem.Document
    CATIA.ActiveDocument.Selection.Clear
    
    If Not selectedPart.Name Like PartName Then
        '****If name does not contain partName
        Set selectedPart = Nothing
        Set selectedProduct = Nothing
        Set selectedRootDocument = Nothing
    End If
    
End Sub


'********************************************************************************
'* Name: Select Instance
'* Purpose: Ask user to select an instance of a CATPart or from a CATProduct
'*          If the assy has the wrong name, it returns Nothing
'*          If the assy has the good name, the active workbench is set to "Assembly Design"
'*
'* Assumption:
'*
'* Author: Francois Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SelectInstance(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")
      
    Dim sDocumentType As String
    On Error Resume Next
    
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    If TypeName(oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    '***** if we cannot access PartNumber property, put oSelectedItem.LeafProduct in default mode
    If oSelectedItem.LeafProduct.PartNumber = "" Then
        oSelectedItem.LeafProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    Set selectedProduct = oSelectedItem.LeafProduct
    
    If Not selectedProduct Is Nothing Then
        Set selectedRootDocument = oSelectedItem.Document
        CATIA.ActiveDocument.Selection.Clear
        
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
        If CATIA.GetWorkbenchId <> "Assembly" Then
            CATIA.StartWorkbench "Assembly"
        End If
    Else
        Set selectedProduct = Nothing
        Set selectedRootDocument = Nothing
        CATIA.ActiveDocument.Selection.Clear
    End If
    
End Sub


'********************************************************************************
'* Name: Select Planar Face
'* Purpose:
'*
'* Assumption:
'*
'* Author: Francois Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub selectPlanarFace(ByRef selectedFace As Face, ByRef parentPart As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")

    ' Initialize variables
    Set selectedFace = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    
    '*** putting part in design mode if needed
    On Error Resume Next
        Set xxx = oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
    If Err.Number <> 0 Then
        On Error GoTo 0
        oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        sAnswer = MsgBox("Part has been put in design mode" + vbNewLine + "Please reselect the face.", vbExclamation)
        Set oSelectedItem = MultiWindowSelectElement(userMessage)
    End If
    On Error GoTo 0

    'The selection must be done in a CATProduct window
    'The selected face must be a "PlanarFace"
    If TypeName(oSelectedItem.Value) = "PlanarFace" And TypeName(oSelectedItem.Document) = "ProductDocument" Then
        Set selectedFace = oSelectedItem.Value
        Set parentPart = oSelectedItem.LeafProduct
        Set selectedRootDocument = oSelectedItem.Document
    End If

End Sub


'********************************************************************************
'* Name: Select Face
'* Purpose:
'*
'* Assumption:
'*
'* Author: Francois Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub selectFace(ByRef selectedFace As Face, ByRef parentPart As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")

    ' Initialize variables
    Set selectedFace = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    
    '*** putting part in design mode if needed
    On Error Resume Next
        Set xxx = oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
    If Err.Number <> 0 Then
        On Error GoTo 0
        oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        sAnswer = MsgBox("Part has been put in design mode" + vbNewLine + "Please reselect the face.", vbExclamation)
        Set oSelectedItem = MultiWindowSelectElement(userMessage)
    End If
    On Error GoTo 0

    'The selection must be done in a CATProduct window
    If TypeName(oSelectedItem.Value) Like "*Face" And TypeName(oSelectedItem.Document) = "ProductDocument" Then
        Set selectedFace = oSelectedItem.Value
        Set parentPart = oSelectedItem.LeafProduct
        Set selectedRootDocument = oSelectedItem.Document
    End If

End Sub


'********************************************************************************
'* Name: Get body from object
'* Purpose: Finding the body of any selected object. This function calls itself
'*          recursively to access the parent of the object passed in until
'*          a Body object is reached.
'*
'* Assumption:
'*
'* Author: Julien Bigaouette (inspired from http://v5vb.wordpress.com/2009/12/15/get-part/ )
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function GetBodyFromObject(ByRef iObject As Variant) As Body
    
    Dim oChildBody As Body
    Dim oParentBody As Body
    Dim oSelection As Selection
    
    On Error Resume Next '****Possible error if iObject.Parent does not exist
    Set oSelection = CATIA.ActiveDocument.Selection '***** Adding object to selection, then setting object = selection, makes us sure to get the good type on the object. Otherwise, a body could be typed as "Hybridbodies" or "Shapes" for example
    oSelection.Clear
    oSelection.Add iObject
    If oSelection.Count2 > 0 Then Set iObject = oSelection.Item2(1).Value
    oSelection.Clear
    
    If TypeName(iObject) = "Body" Then
        'If the object passed in is a body under a boolean operation, find its parent body, then loop again on that body (parent body could be under another boolean operation)
        If iObject.InBooleanOperation Then
            For Each oBody In iObject.Parent
                For Each oShape In oBody.Shapes
                    Set oChildBody = Nothing
                    Set oChildBody = oShape.Body
                    If Not oChildBody Is Nothing Then
                        If oChildBody.Name = iObject.Name Then
                            oSelection.Clear
                            oSelection.Add iObject
                            oSelection.Add oChildBody
                            If oSelection.Count2 = 1 Then 'If only 1 item is selected, it means iObject and oChildBody is the exact same body > We found the good oParentBody
                                Set oParentBody = oBody
                                Set GetBodyFromObject = GetBodyFromObject(oParentBody)
                                oSelection.Clear
                                Exit Function
                            End If
                        End If
                    End If
                Next
            Next
        'If the object passed in is a body that is not under a boolean operation, return it and exit
        Else
             Set GetBodyFromObject = iObject
             Exit Function
        End If
        
    ElseIf TypeName(iObject) = "Part" Or TypeName(iObject) = "Product" Then
        'If iObject is a part or a product, iObject is too high in the tree
        Set GetBodyFromObject = Nothing
        Exit Function
        
    ElseIf TypeName(iObject) = TypeName(iObject.Parent) Then
        'If the type of this object is the same as the parent object then return nothing.
        'The reason is the Parent property of some objects simply returns the same object.
        'This will result in an infinite loop (e.g. An UDF could return itself as parent)
    
        Set GetBodyFromObject = Nothing
        Exit Function
    
    Else
        'Call the function again and pass it the objects parent
        Set GetBodyFromObject = GetBodyFromObject(iObject.Parent)
        
    End If
    
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Get parent product description from product
'* Purpose: Returns the description of the given product, or its closest parent
'*
'* Assumption:
'*
'* Author: Julien Bigaouette (inspired from http://v5vb.wordpress.com/2009/12/15/get-part/ )
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function GetDescriptionFromProduct(ByRef iObject As Variant, Optional ByVal sDescription As String = "*") As String
    
    On Error Resume Next '****Possible error if iObject.Parent does not exist
    GetDescriptionFromProduct = ""
    
    If TypeName(iObject) = "Product" Then
        'If iObject is a part or a product, iObject is too high in the tree
        
        If iObject.DescriptionInst <> "" And iObject.DescriptionInst Like sDescription Then
            GetDescriptionFromProduct = iObject.DescriptionInst
            Exit Function
        Else
            GetDescriptionFromProduct = GetDescriptionFromProduct(iObject.Parent)
        End If
        
    ElseIf TypeName(iObject) = TypeName(iObject.Parent) And iObject.Name = iObject.Parent.Name Then
        'If the type of this object is the same as the parent object then return nothing.
        'The reason is the Parent property of some objects simply returns the same object.
        'This will result in an infinite loop (e.g. An UDF returns itself as parent)
        
        GetDescriptionFromProduct = ""
        Exit Function
        
    Else
        'Call the function again and pass it the objects parent
        GetDescriptionFromProduct = GetDescriptionFromProduct(iObject.Parent)
    End If
    
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Scan Seed Part, find Subset parameters, create a collection of objects "SSParameter"
'*
'* Purpose:
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub ScanSubsets(ByVal oProductCollection As Collection)
    
    Dim iObject As Variant
    Set SSParameters = New Collection 'Public Variable to be declared in another module
    
    On Error Resume Next
    For Each iObject In oProductCollection
        '**** Two names of parameter set to cover: "SubSetParameters" & "ParametersSubSet"
        Set oParameterSet = Nothing
        Set oParameterSet = iObject.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
        Set oParameterSet = iObject.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
        If Not oParameterSet Is Nothing Then
            'Create a collection of objects "SSParameter" for each SS parameter found
            For Each oParameter In oParameterSet.DirectParameters
                sCurrentSSParameterName = Mid(oParameter.Name, InStrRev(oParameter.Name, "\SS_") + Len("\SS_"))
                SSParameters.Add New SSParameter, sCurrentSSParameterName
                SSParameters.Item(sCurrentSSParameterName).ParameterObject = oParameter
                SSParameters.Item(sCurrentSSParameterName).ParameterName = sCurrentSSParameterName
                
                For Each sSubSetItem In ConvertStringToCollection(oParameter.ValueAsString, "|")
                    SSParameters.Item(sCurrentSSParameterName).AddToSubsetValues sSubSetItem
                Next
            Next
        End If
    Next
    On Error GoTo 0
    
End Sub


'********************************************************************************
'* Name: Convert String to Collection
'* Purpose: From a string containing separators, create a collection of strings
'*          (e.g. sStringValue = "AAA;BBB;CCC" and sSeparator = ";" >> Collection.Item(2) = "BBB"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function ConvertStringToCollection(ByRef sStringValue As String, ByRef sSeparator As String, Optional sTag As Boolean = False) As Collection
    
    Dim iObject As Variant
    Set ConvertStringToCollection = New Collection
    
    On Error Resume Next
    For Each iObject In Split(sStringValue, sSeparator)
        If Not iObject = "" And sTag = False Then
            ConvertStringToCollection.Add iObject
        ElseIf Not iObject = "" And sTag = True Then
            ConvertStringToCollection.Add iObject, iObject
        End If
    Next
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Convert Collection to String
'* Purpose: From a collection of strings, create one string separating collection values with a specific character
'*          (e.g. Collection.Item(1) = "AAA", Collection.Item(2) = "BBB", Separator = ":" >> String = "AAA:BBB"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function ConvertCollectionToString(ByRef oCollection As Collection, ByRef sSeparator As String) As String
    
    Dim iObject As Variant
    
    On Error Resume Next
    If oCollection.Item(1) = Empty Then Exit Function
    On Error GoTo 0
    
    ConvertCollectionToString = ""
    For Each iObject In oCollection
        If ConvertCollectionToString = "" Then
            ConvertCollectionToString = iObject
        Else
            ConvertCollectionToString = ConvertCollectionToString & sSeparator & iObject
        End If
    Next
    
End Function


'********************************************************************************
'* Name: Listbox move down
'* Purpose: Change selection > Select element under selected one (bMoveElementDown = false)
'*          Move selected elements down inside a listbox (bMoveElementDown = true)
'*
'* Assumption:
'*
'* Author: http://www.xtremevbtalk.com/archive/index.php/t-180834.html
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function ListBoxMoveDown(ByRef oListBox As ListBox, Optional bMoveElementDown As Boolean = False)
    
    Dim iPosition As Long
    Dim sNextItem As String
    
    iPosition = oListBox.ListCount - 1
    For i = oListBox.ListCount - 1 To 0 Step -1
        If oListBox.Selected(i) Then
            If Not i = iPosition Then
                With oListBox
                    If bMoveElementDown Then
                        sNextItem = .List(i + 1)
                        .List(i + 1) = .List(i)
                        .List(i) = sNextItem
                        .ListIndex = i + 1
                    End If
                    .Selected(i) = False
                    .Selected(i + 1) = True
                End With
            End If
            iPosition = iPosition - 1
        End If
    Next i
    
End Function


'********************************************************************************
'* Name: Listbox move up
'* Purpose: Change selection > Select element above selected one (bMoveElementUp = false)
'*          Move selected elements up inside a listbox (bMoveElementDown = true)
'*
'* Assumption:
'*
'* Author: http://www.xtremevbtalk.com/archive/index.php/t-180834.html
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function ListBoxMoveUp(ByRef oListBox As ListBox, Optional bMoveElementUp As Boolean = False)
    
    Dim iPosition As Long
    Dim sNextItem As String
    
    iPosition = 0
    For i = 0 To oListBox.ListCount - 1
        If oListBox.Selected(i) Then
            If Not i = iPosition Then
                With oListBox
                    If bMoveElementUp Then
                        sNextItem = .List(i - 1)
                        .List(i - 1) = .List(i)
                        .List(i) = sNextItem
                        .ListIndex = i - 1
                    End If
                    .Selected(i) = False
                    .Selected(i - 1) = True
                End With
            End If
            iPosition = iPosition + 1
        End If
    Next i
    
End Function


'********************************************************************************
'* Name: Listbox reorder
'* Purpose: Reorder elements in alphabetical order inside a listbox
'*
'* Assumption:
'*
'* Author: http://www.vbaexpress.com/forum/showthread.php?t=26064
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function ListBoxReorder(ByRef oListBox As ListBox)
    
    Dim sNextItem As String
    
    For i = 0 To oListBox.ListCount - 1
        For j = i + 1 To oListBox.ListCount - 1
            If oListBox.List(i) > oListBox.List(j) Then
                sNextItem = oListBox.List(j)
                oListBox.List(j) = oListBox.List(i)
                oListBox.List(i) = sNextItem
            End If
        Next j
    Next i
    
End Function


'********************************************************************************
'* Name: Show Open File Dialog
'* Purpose: Promt user to select a file / to save a file as...
'*
'* Assumption: sFilter = "Filter1 Name|Filter1 Value|Filter2 Name|Filter2 Value ..."
'*              Example: "All Files(*.*)|*.*|JPG Image(*.jpg)|*.jpg"
'*
'* Information: http://docvb.free.fr/apidetail.php?idapi=136
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function OpenFileDialog(ByVal sFilter As String, _
    Optional ByVal bSaveAsDialog As Boolean = False, _
    Optional ByVal sDefaultExtension As String, _
    Optional ByVal sInitialDirectory As String, _
    Optional ByVal sInitialFileName As String, _
    Optional ByVal sWindowTitle As String, _
    Optional ByVal iParentWindowHandler As Long) As String
    
    Dim OFN As OpenFilename
    Dim bFileSelected As Boolean
    
    Const OFN_FILEMUSTEXIST = &H1000 'The user can type only names of existing files in the File Name entry field
    Const OFN_HIDEREADONLY = &H4 'Hides Read Only check box
    Const OFN_OVERWRITEPROMPT = &H2 'Prompt user when overwriting a file
    
    On Error Resume Next
    ' set the values for the OpenFileName struct
    With OFN
        .hwndOwner = iParentWindowHandler
        .lStructSize = Len(OFN)
        .lpstrFilter = Replace(sFilter, "|", vbNullChar) & vbNullChar
        .lpstrFile = Left$(sInitialFileName & String$(1024, vbNullChar), 1024)
        .nMaxFile = Len(.lpstrFile)
        .flags = OFN_FILEMUSTEXIST + OFN_OVERWRITEPROMPT + OFN_HIDEREADONLY
        .lpstrInitialDir = sInitialDirectory
        .lpstrDefExt = sDefaultExtension
        .lpstrTitle = sWindowTitle
    End With
    
    ' show the dialog (Save As or Open)
    If Not (bSaveAsDialog) Then bFileSelected = GetOpenFileName(OFN)
    If bSaveAsDialog Then bFileSelected = GetSaveFileName(OFN)
    If bFileSelected Then
        ' extract the selected file (including the path)
        OpenFileDialog = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, vbNullChar) - 1)
    End If
End Function


'********************************************************************************
'* Name: Show Open Directory Dialog
'* Purpose: Promt user to select a directory
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette (Default folder)
'* Language: VBA
'********************************************************************************
Public Function OpenDirectoryDialog(Optional ByVal userMessage As String = "Select a folder", Optional ByRef sDefaultDirectory As String = vbNullString) As String
    
    Dim bInfo As BROWSEINFO
    Dim sPath As String
    
    With bInfo
        .lpszTitle = userMessage                         '   Prompt message
        .pidlRoot = 0&                                   '   Root folder = Desktop
        .ulFlags = &H1                                   '   Type of directory to return
        .lpfn = GetAddress(AddressOf BrowseCallbackProc) '   Default folder
        .lParam = StrPtr(sDefaultDirectory)              '   Default folder
    End With
    
    sPath = Space$(512)
    If SHGetPathFromIDList(ByVal SHBrowseForFolder(bInfo), ByVal sPath) Then
        OpenDirectoryDialog = Left(sPath, InStr(sPath, Chr$(0)) - 1)
        sDefaultDirectory = OpenDirectoryDialog
    Else
        OpenDirectoryDialog = ""
    End If
    
End Function

Private Function BrowseCallbackProc(ByVal hWnd&, ByVal msg&, ByVal lp&, ByVal InitDir$) As Long
   
   If (msg = 1) And (InitDir <> "") Then
      Call SendMessage(hWnd, &H466, 1, ByVal InitDir$)
   End If
   BrowseCallbackProc = 0
   
End Function

Private Function GetAddress(ByVal Addr As Long) As Long
   GetAddress = Addr
End Function


'********************************************************************************
'* Name: Open Web Page
'* Purpose: This sub receive a URL (strPartURL As String) as input
'*          It verifies if the page is already open
'*          If it is open, it maximize it and bring it front
'*          If the page is not already open, it opens a new one and maximize it
'
'*          Optional argument: length corresponding to the beginning of the string.
'*          Example:
'*          URLLength = 15 (corresponds to I:/V5_KBE_Tools > 15 characters)
'*          I want to open a new page I:/V5_KBE_Tools/Production/07_KBE_Broadcast/KBE_Info.htm
'*          ONLY if I don't find an existing page with the address starting with "I:/V5_KBE_Tools"
'*          Otherwise, I maximize the existing page (e.g. "I:\V5_KBE_Tools\Production\07_KBE_Broadcast\Global 6000\G6000ChangeManagementInfo.htm")
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub openPage(ByVal strPartURL As String, Optional ByVal URLLength As Integer)
    
    Dim objShellWindows As Variant
    Dim objShell As Variant
    Dim objIE As Variant
    
    If Len(strPartURL) = 0 Then Exit Sub
    
    On Error Resume Next
    Err.Clear
    Set objShell = CreateObject("Shell.Application")
    Set objShellWindows = objShell.Windows
    If Err <> 0 Then Exit Sub
    
    '*****Verify if there's a window open*****
    If objShellWindows.Count = 0 Then
        GoTo OpenNewWindow
    End If
    
    '*****Verify all windows, try to find the page with given URL and bring it to top*****
    '*****LIMIT CASE: if there's many tabs opened (IE Window), it does not activate it, it shows the actual active tab*****
    If URLLength = 0 Then URLLength = Len(strPartURL)
    
    For i = 0 To objShellWindows.Count - 1
        Set objIE = objShellWindows.Item(CLng(i))
        
        If Not (objIE Is Nothing) Then
            If InStr(Replace(objIE.LocationURL, "/", "\"), Left(strPartURL, URLLength)) Then
                
                '*****Maximize the page if it is minimized
                If IsIconic(objIE.hWnd) Then ShowWindow objIE.hWnd, SW_RESTORE
                
                '*****Bring the page to top in its actual state*****
                '*****Does not maximize a minimized window!! Idem as AppActivate method*****
                BringWindowToTop objIE.hWnd
                
                Set objShellWindows = Nothing
                Set objShell = Nothing
                Set objIE = Nothing
                
                Exit Sub
            End If
        End If
    Next i
    
    '*****Open URL in a new window*****
OpenNewWindow:
    
    Set objIE = CreateObject("InternetExplorer.Application")
    
    '*****Maximize the page
    ShowWindow objIE.hWnd, SW_MAXIMIZE
    objIE.Navigate (strPartURL)
    objIE.Visible = True
    
    '*****Reset variables*****
    Set objShellWindows = Nothing
    Set objShell = Nothing
    Set objIE = Nothing
    
End Sub


'********************************************************************************
'* Name: Product Name Check
'* Purpose: Inform user that the selected assembly does not follow the standards
'*          Used in: 3D Collector, 3D Collector Template, Cabinet Project, Reuse
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette (Default folder)
'* Language: VBA
'********************************************************************************
Public Function ProductNameCheck(ByVal NameToCheck As String, Optional ByVal bEngFileRequested As Boolean = True, Optional ByVal bDesignFileRequested As Boolean = True) As Boolean
    
    Dim bNHAPartNumberTest As Boolean
    Dim WarningMessage As String
    Dim FileTypeGuess As String
    
    bNHAPartNumberTest = False
    WarningMessage = ""
    FileTypeGuess = "Unknown Format:  "
    
    '********* Engineering File - Global 6000 nomenclature validation **********
    If Len(NameToCheck) = 14 And bEngFileRequested And NameToCheck Like "7*" Then
        
        WarningMessage = "Engineering file (Global 6000) detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
        
        If NameToCheck Like "7K*" Or NameToCheck Like "7G*" Then

            If Mid(NameToCheck, 11, 1) Like "-" Then
                If NameToCheck Like "??##???###-###*" And Not Mid(NameToCheck, 5, 3) Like "*#*" Then
                    bNHAPartNumberTest = True
                Else
                    WarningMessage = WarningMessage & """MM"", ""XXX"" and ""DDD"" should be numeric characters only " & vbCrLf & _
                                    """III"" should be alphabetic characters only "
                End If
            Else
                WarningMessage = WarningMessage & "Character 11 should be ""-"" "
            End If
        
        Else
            WarningMessage = "The first two char must be either '7K' or '7D'    "
        End If
    
    '********* Engineering File - Global 7000/8000 nomenclature validation **********
    ElseIf Len(NameToCheck) >= 17 And Len(NameToCheck) <= 20 And bEngFileRequested And NameToCheck Like "G025*" Then
        
        WarningMessage = "Engineering file (Global 7000/8000) detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
        
        If Mid(NameToCheck, 10, 1) Like "-" Then
            
            If NameToCheck Like "G025#####-###?###*" Then
                
                If NameToCheck Like "G025#####-###E###*" Then
                    
                    If NameToCheck Like "G025#####-###E###" Or NameToCheck Like "G025#####-###E###?##" Then
                        bNHAPartNumberTest = True
                    Else
                        WarningMessage = WarningMessage & """rr"" should be numeric characters only "
                    End If
                Else
                    WarningMessage = WarningMessage & "Character 14 should be ""E"" "
                End If
            
            Else
                WarningMessage = WarningMessage & """#####"", ""DDD"" and ""xxx"" should be numeric characters only    "
            End If
            
        Else
            WarningMessage = WarningMessage & "Character 10 should be ""-"" "
        End If
        

    '********* Design File - Global 6000 nomenclature validation **********
    ElseIf Len(NameToCheck) = 20 And bDesignFileRequested Then
        WarningMessage = "Design file detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
            
                 If NameToCheck Like "GXRS*" Or _
                NameToCheck Like "G500*" Or _
                NameToCheck Like "G600*" Or _
                NameToCheck Like "G700*" Or _
                NameToCheck Like "G800*" Or _
                NameToCheck Like "5000*" Or _
                NameToCheck Like "6000*" Or _
                NameToCheck Like "7000*" Or _
                NameToCheck Like "8000*" Then
           
                If NameToCheck Like "????????-????-??-???*" Then
                    If NameToCheck Like "????####-####-##-###*" Then
                        bNHAPartNumberTest = True
                    Else
                        WarningMessage = WarningMessage & """####"", ""MM"", ""SS"", ""XX"" and ""DDD"" should be numeric characters    "
                    End If
                Else
                    WarningMessage = WarningMessage & "Characters 9, 14 and 17 should be ""-"" "
                End If
                
            Else
                WarningMessage = WarningMessage & "Selected file name does not contain any of the following prefix:    " & vbCrLf & _
                                 "GXRS, G500, G600, G700, G800, 5000, 6000, 7000, 8000    "
            End If
    
    Else
        WarningMessage = "Unknown Format:  " & vbCrLf & "Selected NHA name is not valid    "
    End If
    
    
    '********* Warning message **********
    If bNHAPartNumberTest = False Then
        
        If bEngFileRequested Then WarningMessage = WarningMessage & vbCrLf & vbCrLf & _
                                                "      Eng. nomenclature (Global 6000):  7KMMIIIXXX-DDD  " & vbCrLf & vbCrLf & _
                                                "             MM = KBE monument number" & vbCrLf & _
                                                "             III = 3 digit iteration, alphabetic only (i.e. AAA, AAB...)" & vbCrLf & _
                                                "             XXX = KBE part index number" & vbCrLf & _
                                                "             DDD = 3 digit dash number" & vbCrLf & vbCrLf & _
                                                "      Eng. nomenclature (Global 7000/8000):  G025#####-DDDEEEErrr  " & vbCrLf & vbCrLf & _
                                                "             ##### = Base number" & vbCrLf & _
                                                "             DDD = 3 digit dash number" & vbCrLf & _
                                                "             EEEE = 4 digit envelope number (i.e. E001, E002...)" & vbCrLf & _
                                                "             rrr = Optional 3 digit representation (i.e. A01, A02...)"
                                                
        If bDesignFileRequested Then WarningMessage = WarningMessage & vbCrLf & vbCrLf & _
                                                "      Design nomenclature:  GXRS####-MMSS-XX-DDD        " & vbCrLf & vbCrLf & _
                                                "             #### = Project number" & vbCrLf & _
                                                "             MM = KBE monument number" & vbCrLf & _
                                                "             SS = sub-monument number" & vbCrLf & _
                                                "             XX = Additional information number" & vbCrLf & _
                                                "             DDD = 3 digit dash number"
        
        WarningMessage = vbCrLf & WarningMessage & vbCrLf & vbCrLf & "Do you want to continue?          " & vbCrLf & vbCrLf
        
        If MsgBox(WarningMessage, vbExclamation + vbYesNo, "Warning") = vbNo Then
            ProductNameCheck = False
        Else
            ProductNameCheck = True
        End If
    Else
        ProductNameCheck = True
    End If
    
End Function


'********************************************************************************
'* Name: Positioning functions
'* Purpose: All following functions are used to position a part in the space
'*
'* Assumption:
'*
'* Author: François Charette
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Sub SnapPartToPart(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceProduct As Product)
    
    'Purpose: Move child product instance (oProductToPosition) at inside its NHA (oProductToPositionNHA) at
    'same global position than reference product instance (oReferenceProduct)
    
    Dim dReferenceProductMatrix(11) As Variant
    Dim dTargetNHAMatrix(11) As Variant
    Dim dTargetMatrix(11) As Variant
    
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = BuildParentList(oReferenceProduct)
    Call CalculateInstanceGlobalPositionMatrix(sParentList, dReferenceProductMatrix)
    
    sParentList = BuildParentList(oProductToPositionNHA)
    Call CalculateInstanceGlobalPositionMatrix(sParentList, dTargetNHAMatrix)
    
    Call FindChildMatrix(dReferenceProductMatrix, dTargetNHAMatrix, dTargetMatrix)
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents dTargetMatrix
    
End Sub

Public Sub SnapPartToReferenceAxis(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceProduct As Product, ByVal oReferenceAxis As AxisSystem)
    
    'Purpose: Move child product instance (oProductToPosition) inside its NHA (oProductToPositionNHA) at
    'same global position than reference axis given (oReferenceAxis)
    
    Dim dReferenceProductMatrix(11) As Variant
    Dim dAxisMatrix(11) As Variant
    Dim dTargetNHAMatrix(11) As Variant
    Dim dTargetMatrix(11) As Variant
    
    'Retrieve Global Matrix of reference product
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = BuildParentList(oReferenceProduct)
    Call CalculateInstanceGlobalPositionMatrix(sParentList, dReferenceProductMatrix)
    
    'Retrieve Axis position (global matrix)
    Call GetAxisPosition(oReferenceAxis, dAxisMatrix)
    Call MultiplyVector(dAxisMatrix, dReferenceProductMatrix, dReferenceProductMatrix, True)
    Call MultiplyMatrix(dAxisMatrix, dReferenceProductMatrix, dReferenceProductMatrix)
    
    'Retrieve Global Matrix of target NHA
    sParentList = BuildParentList(oProductToPositionNHA)
    Call CalculateInstanceGlobalPositionMatrix(sParentList, dTargetNHAMatrix)
    
    'Retrieve target position of the part
    Call FindChildMatrix(dReferenceProductMatrix, dTargetNHAMatrix, dTargetMatrix)
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents dTargetMatrix
    
End Sub

Public Sub SnapAxisToOrigin(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceAxis As AxisSystem)
    
    'Purpose: Move product instance (oProductToPosition) at inside its NHA  (oProductToPositionNHA)
    'so the reference axis is at (0,0,0)
    
    Dim dAxisMatrix(11) As Variant
    Dim dNHAMatrix(11) As Variant
    Dim dReferenceProductMatrix(11) As Variant
    
    'Retrieve Global Matrix of Collector's NHA
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = BuildParentList(oProductToPositionNHA)
    Call CalculateInstanceGlobalPositionMatrix(sParentList, dNHAMatrix)
    
    'Retrieve Axis position matrix in the part
    Call GetAxisPosition(oReferenceAxis, dAxisMatrix)
    
    'Find dReferenceProductMatrix to have Axis position matrix at 0,0,0
    Call FindPartAxisMatrix(dAxisMatrix, dNHAMatrix, dReferenceProductMatrix)
    
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents dReferenceProductMatrix
    
End Sub

Public Sub FindPartAxisMatrix(ByVal dAxisMatrix As Variant, ByVal dNHAMatrix As Variant, ByRef dTargetPartMatrix As Variant)
    
    'Purpose: finding the position dTargetPartMatrix (relative to NHA) that correspond to a dAxisMatrix at (0,0,0)(Global position)
    
    Dim dTempMatrix As Variant
    dTempMatrix = dAxisMatrix
    
    'Invert Axis Matrix
    Call InvertMatrix(dAxisMatrix)
    
    'Set dTempMatrix Vector
    dTempMatrix(9) = -dAxisMatrix(9)
    dTempMatrix(10) = -dAxisMatrix(10)
    dTempMatrix(11) = -dAxisMatrix(11)
    
    'Caculate Vector Position > Return result in dAxisMatrix
    Call MultiplyVector(dTempMatrix, dAxisMatrix, dAxisMatrix)
    
    'Find part matrix (under NHA) > Return result in d3DCollMatrix
    Call FindChildMatrix(dAxisMatrix, dNHAMatrix, dTargetPartMatrix)
    
End Sub

Public Sub FindChildMatrix(ByRef dRefMatrix As Variant, ByRef dTargetNHAMatrix As Variant, ByRef dChildMatrix)
    
    'Purpose: Find the dChildMatrix (relative to NHA) to have same global position as dRefMatrix
    
    'Invert dRefMatrix
    Call InvertMatrix(dTargetNHAMatrix)
    
    'Remove Vector Position dTargetNHAMatrix from dRefMatrix, Return result in dRefMatrix Matrix, put dTargetNHAMatrix position to 0
    Call SubstractVector(dRefMatrix, dTargetNHAMatrix)
    
    'Caculate Vector Position > Return result in dChildMatrix
    Call MultiplyVector(dRefMatrix, dTargetNHAMatrix, dChildMatrix)
    
    'Caculate Rotation Matrix (dRefMatrix X dTargetMatrix) > Return result in dChildMatrix
    Call MultiplyMatrix(dRefMatrix, dTargetNHAMatrix, dChildMatrix)
    
End Sub

Public Function BuildParentList(ByVal iObject As Variant) As String
    
    'The parent list is a list of all the instances starting at the selection all the way up to the top product in the active window
    BuildParentList = iObject.Name & ";"
    
    Do
        Set iObject = iObject.Parent
        
        'We don't care about the "Products" or "ProductDocuments" objects
        If TypeName(iObject) = "Product" Then
            BuildParentList = BuildParentList & iObject.Name & ";"
        End If
    Loop Until TypeName(iObject) = "Application"

End Function

Public Sub CalculateInstanceGlobalPositionMatrix(ByVal sParentList As String, ByRef dGlobalMatrix As Variant)

    Dim oParent As Product
    Dim iParentList As Long
    Dim dLocalMatrix(11) As Variant
    
    'Initialize first parent and dGlobalMatrix (should be I Matrix)
    Set oParent = CATIA.ActiveDocument.Product
    Set oProdPosition = oParent.Position
    oProdPosition.GetComponents dGlobalMatrix
    
    'Loop sParentList
    For iParentList = UBound(Split(sParentList, ";")) - 2 To 0 Step -1
    
        'Retrieve oChild position matrix
        Set oParent = oParent.Products.Item(Split(sParentList, ";")(iParentList))
        Set oProdPosition = oParent.Position
        oProdPosition.GetComponents dLocalMatrix
        
        'Caculate Vector Position > Return result in Global Matrix
        Call MultiplyVector(dLocalMatrix, dGlobalMatrix, dGlobalMatrix, True)
        
        'Caculate Rotation Matrix (dLocalMatrix X dGlobalMatrix) > Return result in Global Matrix
        Call MultiplyMatrix(dLocalMatrix, dGlobalMatrix, dGlobalMatrix)
        
    Next iParentList
    
End Sub

Public Sub GetAxisPosition(ByVal oAxis As Object, ByRef dAxisMatrix As Variant)
    
    Dim dVectorPosition(2) As Variant
    
    'Retrieve axis position matrix
    oAxis.GetXAxis dVectorPosition
    dAxisMatrix(0) = dVectorPosition(0)
    dAxisMatrix(1) = dVectorPosition(1)
    dAxisMatrix(2) = dVectorPosition(2)
    
    oAxis.GetYAxis dVectorPosition
    dAxisMatrix(3) = dVectorPosition(0)
    dAxisMatrix(4) = dVectorPosition(1)
    dAxisMatrix(5) = dVectorPosition(2)
    
    oAxis.GetZAxis dVectorPosition
    dAxisMatrix(6) = dVectorPosition(0)
    dAxisMatrix(7) = dVectorPosition(1)
    dAxisMatrix(8) = dVectorPosition(2)
    
    oAxis.GetOrigin dVectorPosition
    dAxisMatrix(9) = dVectorPosition(0)
    dAxisMatrix(10) = dVectorPosition(1)
    dAxisMatrix(11) = dVectorPosition(2)
    
End Sub

Public Sub InvertMatrix(ByRef dMatrixToInvert As Variant)
    
    Dim dTempMatrix As Variant
    dTempMatrix = dMatrixToInvert
    
'Transpose Global Matrix
    dMatrixToInvert(0) = dTempMatrix(0)
    dMatrixToInvert(1) = dTempMatrix(3)
    dMatrixToInvert(2) = dTempMatrix(6)
    dMatrixToInvert(3) = dTempMatrix(1)
    dMatrixToInvert(4) = dTempMatrix(4)
    dMatrixToInvert(5) = dTempMatrix(7)
    dMatrixToInvert(6) = dTempMatrix(2)
    dMatrixToInvert(7) = dTempMatrix(5)
    dMatrixToInvert(8) = dTempMatrix(8)
    
End Sub

Public Sub MultiplyMatrix(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, ByRef dResultMatrix As Variant)
'Multiply two 3x3 matrix together (dMatrix1 * dMatrix2)
    
    dResultMatrix(0) = dMatrix1(0) * dMatrix2(0) + dMatrix1(1) * dMatrix2(3) + dMatrix1(2) * dMatrix2(6)
    dResultMatrix(1) = dMatrix1(0) * dMatrix2(1) + dMatrix1(1) * dMatrix2(4) + dMatrix1(2) * dMatrix2(7)
    dResultMatrix(2) = dMatrix1(0) * dMatrix2(2) + dMatrix1(1) * dMatrix2(5) + dMatrix1(2) * dMatrix2(8)

    dResultMatrix(3) = dMatrix1(3) * dMatrix2(0) + dMatrix1(4) * dMatrix2(3) + dMatrix1(5) * dMatrix2(6)
    dResultMatrix(4) = dMatrix1(3) * dMatrix2(1) + dMatrix1(4) * dMatrix2(4) + dMatrix1(5) * dMatrix2(7)
    dResultMatrix(5) = dMatrix1(3) * dMatrix2(2) + dMatrix1(4) * dMatrix2(5) + dMatrix1(5) * dMatrix2(8)
    
    dResultMatrix(6) = dMatrix1(6) * dMatrix2(0) + dMatrix1(7) * dMatrix2(3) + dMatrix1(8) * dMatrix2(6)
    dResultMatrix(7) = dMatrix1(6) * dMatrix2(1) + dMatrix1(7) * dMatrix2(4) + dMatrix1(8) * dMatrix2(7)
    dResultMatrix(8) = dMatrix1(6) * dMatrix2(2) + dMatrix1(7) * dMatrix2(5) + dMatrix1(8) * dMatrix2(8)
    
End Sub

Public Sub MultiplyVector(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, ByRef dResultMatrix As Variant, Optional ByVal bAddInitialValue As Boolean = False)
'Multiply 1x3 vector by 3x3 matrix ('dMatrix1(Vector) * dMatrix2)
    
    dResultMatrix(9) = dMatrix1(9) * dMatrix2(0) + dMatrix1(10) * dMatrix2(3) + dMatrix1(11) * dMatrix2(6)
    dResultMatrix(10) = dMatrix1(9) * dMatrix2(1) + dMatrix1(10) * dMatrix2(4) + dMatrix1(11) * dMatrix2(7)
    dResultMatrix(11) = dMatrix1(9) * dMatrix2(2) + dMatrix1(10) * dMatrix2(5) + dMatrix1(11) * dMatrix2(8)
    
    If bAddInitialValue Then
        dResultMatrix(9) = dResultMatrix(9) + dMatrix2(9)
        dResultMatrix(10) = dResultMatrix(10) + dMatrix2(10)
        dResultMatrix(11) = dResultMatrix(11) + dMatrix2(11)
    End If
    
End Sub

Public Sub SubstractVector(ByRef dMatrix1 As Variant, ByRef dMatrix2 As Variant)
'Remove 1x3 vector2 from 1x3 vector1

    dMatrix1(9) = dMatrix1(9) - dMatrix2(9)
    dMatrix1(10) = dMatrix1(10) - dMatrix2(10)
    dMatrix1(11) = dMatrix1(11) - dMatrix2(11)
    
End Sub

Public Function GetInstanceActivity(ByVal oInstance As Product, Optional ByVal bSetDefaultMode As Boolean = False) As Boolean
'********************************************************************************
'* Name: bInstanceActivity
'* Purpose: If Component activity state is found the it returns the value else it returns false.
'*          Hence false could either mean no parameter found or instance is not active.
'*          if bSetDefaultMode is set to true then it tries to set instance in default mode if no parameter is found.
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
    Dim oProdparams As Parameters
    
    On Error Resume Next
    Set oProdparams = oInstance.Parameters.SubList(oInstance, False)
   
    If oProdparams.Count = 0 And bSetDefaultMode = True Then
        oInstance.ApplyWorkMode DEFAULT_MODE
        Set oProdparams = oInstance.Parameters.SubList(oInstance, False)
    End If
    
    If oProdparams.Count > 0 Then GetInstanceActivity = oProdparams.GetItem("Component Activation State").Value
    
    On Error GoTo 0
    
End Function

'****--------------------------------------------------------------------------------------------------------------------------------
'****
'****
'****                           Following are Linear algebra & Geometry algorithms functions
'****
'**** Author: Abhishek Kamboj
'**** Updated by:
'**** Language: VBA
'****--------------------------------------------------------------------------------------------------------------------------------
Public Function DotProd(v1 As Variant, v2 As Variant) As Double
'*** dot product of two vectors of 3 elements
    DotProd = v1(0) * v2(0) + v1(1) * v2(1) + v1(2) * v2(2)
End Function
Public Sub CrossProduct(v1() As Double, v2() As Double, ByRef dCrossProduct As Variant)
    Dim V(2) As Double
    
    V(0) = v1(1) * v2(2) - v1(2) * v2(1)
    V(1) = v1(2) * v2(0) - v1(0) * v2(2)
    V(2) = v1(0) * v2(1) - v1(1) * v2(0)

    dCrossProduct = V
End Sub
Public Function LengthVector(v1 As Variant) As Double
'*** Magnitude of vector of 3 elements
    LengthVector = Sqr(v1(0) ^ 2 + v1(1) ^ 2 + v1(2) ^ 2)
End Function
Public Function ArcCos(oValue As Double) As Double
'*** Inverse cosine of value
    If Round(oValue, 8) = 1 Then ArcCos = 0: Exit Function
    If Round(oValue, 8) = -1 Then ArcCos = PI: Exit Function
    ArcCos = Atn(-oValue / Sqr(1 - oValue ^ 2)) + 2 * Atn(1)
End Function

Public Function AngleInRad(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, Optional ByRef sVector As String = "X") As Double
'*** Find angle between two vectors in 3 dimensions.Input is position matrix of two instance
'*** By assigning optional value to SVector , one can find angle between either X ,Y or Z axis of two Position matrix.
    Dim Vctr1(2)
    Dim Vctr2(2)
    
    Vctr1(0) = dMatrix1(0)
    Vctr1(1) = dMatrix1(1)
    Vctr1(2) = dMatrix1(2)
    
    Vctr2(0) = dMatrix2(0)
    Vctr2(1) = dMatrix2(1)
    Vctr2(2) = dMatrix2(2)
    
    If sVector = "Y" Then
        Vctr1(0) = dMatrix1(3)
        Vctr1(1) = dMatrix1(4)
        Vctr1(2) = dMatrix1(5)
        
        Vctr2(0) = dMatrix2(3)
        Vctr2(1) = dMatrix2(4)
        Vctr2(2) = dMatrix2(5)
    ElseIf sVector = "Z" Then
        Vctr1(0) = dMatrix1(6)
        Vctr1(1) = dMatrix1(7)
        Vctr1(2) = dMatrix1(8)
        
        Vctr2(0) = dMatrix2(6)
        Vctr2(1) = dMatrix2(7)
        Vctr2(2) = dMatrix2(8)
    End If
    
    AngleInRad = ArcCos(DotProd(Vctr1, Vctr2) / (LengthVector(Vctr1) * LengthVector(Vctr2)))
    
End Function
Public Function dDistanceLineSegToLineSeg(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double

'********--------------------
'***
'*** returns shortest distance between two line segments Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Adapted from http://geomalgorithms.com/a07-_distance.html#dist3D_Segment_to_Segment()
'***
'********--------------------
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim Result(2) As Double
    Dim PR1(2) As Double
    Dim PR2(2) As Double
    Dim a As Double, b As Double, c As Double, d As Double, DD As Double, sc As Double, sn As Double, sd As Double
    Dim tc As Double, tN As Double, tD As Double
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)
    V(0) = p4(0) - p3(0): V(1) = p4(1) - p3(1): V(2) = p4(2) - p3(2)
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)

    a = DotProd(U, U)   ' should be > = zero
    b = DotProd(U, V)
    c = DotProd(V, V)   ' should be > = zero
    d = DotProd(U, W)
    e = DotProd(V, W)
    
    DD = a * c - b * b
    sd = DD ' default
    tD = DD ' default
    
    If DD < SMALLNUMBER Then     ' lines are parallel
            sn = 0               ' forcing use of point P1 of Line 1
            sd = 1               ' to prevent possible division by 0.0 later
            tN = e
            tD = c
    Else                         ' if lines are not parallel then get closest point on infinite lines
        sn = (b * e - c * d)
        tN = (a * e - b * d)
            If sn < 0 Then       ' sc < 0 => the s=0 edge is visible
                sn = 0
                tN = e
                tD = c
            ElseIf sn > sd Then  ' sc > 1  => the s=1 edge is visible
                sn = sd
                tN = e + b
                tD = c
            End If
    End If
    If (tN < 0) Then             'tc < 0 => the t=0 edge is visible
        tN = 0
        ' recompute sc for this edge
            If (0 - d) < 0 Then
                sn = 0
            ElseIf (0 - d) > a Then
                sn = sd
            Else
                sn = 0 - d
                sd = a
            End If
    ElseIf tN > tD Then         ' tc > 1  => the t=1 edge is visible
        tN = tD
        'recompute sc for this edge
            If (0 - d + b) < 0 Then
                sn = 0
            ElseIf (0 - d + b) > a Then
                sn = sd
            Else
                sn = (0 - d + b)
                sd = a
            End If
    End If
        
    If Abs(sn) < SMALLNUMBER Then
        sc = 0
    Else
        sc = sn / sd
    End If
    
    If Abs(tN) < SMALLNUMBER Then
        tc = 0
    Else
        tc = tN / tD
    End If
    Result(0) = W(0) + (sc * U(0)) - (tc * V(0))
    Result(1) = W(1) + (sc * U(1)) - (tc * V(1))
    Result(2) = W(2) + (sc * U(2)) - (tc * V(2))
    dDistanceLineSegToLineSeg = LengthVector(Result) ' Length of Vector
    PR1(0) = p1(0) + sc * U(0)
    PR1(1) = p1(1) + sc * U(1)
    PR1(2) = p1(2) + sc * U(2)
    
    PR2(0) = p3(0) + tc * V(0)
    PR2(1) = p3(1) + tc * V(1)
    PR2(2) = p3(2) + tc * V(2)
End Function
Public Function dDistancePointToLineSeg(p1() As Double, p2() As Double, p3() As Double, Optional LineIsInfinite As Boolean = False) As Double
'*** calculate shortest distance between a point & a line segment
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim a As Double, b As Double, c As Double
    Dim ClosestPoint(2) As Double
    Dim Result(2) As Double
    
    U(0) = p1(0) - p3(0): U(1) = p1(1) - p3(1): U(2) = p1(2) - p3(2)     ' Vector of End point of line to the Point
    V(0) = p3(0) - p2(0): V(1) = p3(1) - p2(1): V(2) = p3(2) - p2(2)     ' Vector of Line
    W(0) = p1(0) - p2(0): W(1) = p1(1) - p2(1): W(2) = p1(2) - p2(2)     ' Vector of Startpoint of Line to the point
        
    a = DotProd(V, W)
    b = DotProd(V, V)
    
    If Not LineIsInfinite Then  ' conditions for line of finite length
        If a < 0 Then
            dDistancePointToLineSeg = LengthVector(W)
            Exit Function
        End If
        If b < a Then
            dDistancePointToLineSeg = LengthVector(U)
            Exit Function
        End If
    End If
    ' if line is infinte
    c = a / b
    ClosestPoint(0) = p2(0) + c * V(0)
    ClosestPoint(1) = p2(1) + c * V(1)
    ClosestPoint(2) = p2(2) + c * V(2)
    
    Result(0) = p1(0) - ClosestPoint(0)
    Result(1) = p1(1) - ClosestPoint(1)
    Result(2) = p1(2) - ClosestPoint(2)
    
    dDistancePointToLineSeg = Round(LengthVector(Result), 12)
End Function
Public Function dDistanceLineToLine(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double
'*** calculates shortest distance between lines of infinite length
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim PR1(2) As Double
    Dim PR2(2) As Double
    
    Dim Result(2) As Double
    Dim a As Double, b As Double, c As Double, d As Double, DD As Double, sc As Double, tc As Double
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)    ' vector of line 1
    V(0) = p4(0) - p3(0): V(1) = p4(1) - p3(1): V(2) = p4(2) - p3(2)    ' vector of line 2
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)

    a = DotProd(U, U)   ' should be > = zero
    b = DotProd(U, V)
    c = DotProd(V, V)   ' should be > = zero
    d = DotProd(U, W)
    e = DotProd(V, W)
    
    DD = a * c - b * b
     If DD < SMALLNUMBER Then     ' lines are parallel
        sc = 0
        If b > c Then
            tc = d / b
        Else
            tc = e / c
        End If
    Else
        sc = ((b * e) - (c * d)) / DD
        tc = ((a * e) - (b * d)) / DD
    End If

    Result(0) = W(0) + (sc * U(0)) - (tc * V(0))
    Result(1) = W(1) + (sc * U(1)) - (tc * V(1))
    Result(2) = W(2) + (sc * U(2)) - (tc * V(2))

    dDistanceLineToLine = Round(LengthVector(Result), 12)
    
    PR1(0) = p1(0) + sc * U(0)
    PR1(1) = p1(1) + sc * U(1)
    PR1(2) = p1(2) + sc * U(2)
    
    PR2(0) = p3(0) + tc * V(0)
    PR2(1) = p3(1) + tc * V(1)
    PR2(2) = p3(2) + tc * V(2)
    
End Function
Public Function dDistancePointToPoint(p1() As Double, p2() As Double) As Double
    Dim ResultVector(2) As Double
    
    ResultVector(0) = p2(0) - p1(0): ResultVector(1) = p2(1) - p1(1): ResultVector(2) = p2(2) - p1(2)
    dDistancePointToPoint = Round(LengthVector(ResultVector), 12)
End Function
Public Function dDistancePointToPlane(p1() As Double, p2() As Double, p3() As Double) As Double
    '*** p1 is the point from which we want to measure the distance to plane
    '*** Plane vector is created from p2 & p3 , p2 is also a point lying on a plane
    '*** distance is +ve when point is on the side of vector of plane
    '*** and -ve when it is on the other side
    Dim PlaneVector(2) As Double
    Dim Result(2) As Double
    Dim sb As Double
    Dim sn As Double
    Dim sd As Double
    Dim U(2) As Double
    
    PlaneVector(0) = p3(0) - p2(0): PlaneVector(1) = p3(1) - p2(1): PlaneVector(2) = p3(2) - p2(2)
    U(0) = p1(0) - p2(0): U(1) = p1(1) - p2(1): U(2) = p1(2) - p2(2)
    
    sn = DotProd(PlaneVector, U)
    sd = LengthVector(PlaneVector)
    dDistancePointToPlane = sn / sd
End Function

Public Function iFindLineSegmentIntersectionToPlane(p1() As Double, p2() As Double, p3() As Double, p4(), Optional ByRef dIntsctPT As Variant) As Integer
'***-------------
'*** p1 & p2 belongs to start & end of line segment
'*** Plane vector is created from p3 & p4 , p3 being the point lying on a plane
'*** Result = 0 No intersection
'*** Result = 1 Intersection exist
'*** Result =2 segement lies on plane
'*** if Intersection exist then dIntsctPT will have coordinates of intersection
'*** Adapted from --- http://geomalgorithms.com/a05-_intersect-1.html
'***-------------
    Dim PlaneVector(2) As Double
    Dim U(2) As Double
    Dim W(2) As Double
    Dim d As Double
    Dim N As Double
    Dim SI As Double
    Dim Result As Integer
    
    PlaneVector(0) = p4(0) - p3(0): PlaneVector(1) = p4(1) - p3(1): PlaneVector(2) = p4(2) - p3(2)
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)
    
    d = DotProd(PlaneVector, U)
    N = -DotProd(PlaneVector, W)
    
    If Abs(d) < SMALLNUMBER Then
        If N = 0 Then
            Result = 2 ' segment lies in plane
        Else
            Result = 0 ' no intersection
        End If
    End If
    SI = N / d
    If SI < 0 Or SI > 0 Then
        Result = 0  ' no intersection
    Else
        Result = 1 ' Intersection exist
        dIntsctPT(0) = p1(0) + SI * U(0)
        dIntsctPT(1) = p1(1) + SI * U(1)
        dIntsctPT(2) = p1(2) + SI * U(2)
    End If
    iFindLineSegmentIntersectionToPlane = Result
End Function



