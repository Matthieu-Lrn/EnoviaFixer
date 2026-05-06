Attribute VB_Name = "SFB_BA_KBE_GCC_COMMON"
'*****************************************************************************************************
'                     SCRIPT VERSION
'*****************************************************************************************************
Private Const SFB_sCommonModuleVersion As String = "KBE_GCC_COMMON_MAR14_2017"
'*****************************************************************************************************

'*****************************************************************************************************
'                     PATHS FILE
'*****************************************************************************************************
Public SFB_sKBEPathFile As String
'*****************************************************************************************************


#If VBA7 Then
    Public Declare PtrSafe Function SFB_SetCursor Lib "user32" (ByVal hCursor As Long) As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare PtrSafe Function SFB_GetCursor Lib "user32" () As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare PtrSafe Function SFB_LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare PtrSafe Function SFB_ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare PtrSafe Function SFB_IsIconic Lib "user32" (ByVal hwnd As Long) As Long '***** To verify if SFB_a window is minimized
    Public Declare PtrSafe Function SFB_BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare PtrSafe Function SFB_FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpclassname As Any, ByVal lpCaption As Any) '***** To give title bar SFB_a toolbar look, used also to call SFB_OpenFileDialog to give window handler
    Public Declare PtrSafe Function SFB_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '***** To give title bar SFB_a toolbar look
    Public Declare PtrSafe Function SFB_DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long '***** To give title bar SFB_a toolbar look
    Public Declare PtrSafe Function SFB_GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long '***** To find screen size, and if user has 1 or 2 screens (to position toolbar inside screen)
    Public Declare PtrSafe Function SFB_GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare PtrSafe Function SFB_GetDC Lib "user32" (ByVal hwnd As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare PtrSafe Function SFB_GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer '***** To detect ESC key press
    Public Declare PtrSafe Sub SFB_Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long) '***** To make SFB_a pause when running script
    Public Declare PtrSafe Function SFB_SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long  'To browse SFB_a directory
    Public Declare PtrSafe Function SFB_SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As SFB_BROWSEINFO) As Long 'To browse SFB_a directory
    Public Declare PtrSafe Function SFB_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'To browse SFB_a directory
    Public Declare PtrSafe Function SFB_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As SFB_OpenFilename) As Long 'To browse for SFB_a file
    Public Declare PtrSafe Function SFB_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As SFB_OpenFilename) As Long 'To browse for SFB_a file
    Public Declare PtrSafe Function SFB_GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long 'To get user ID
    Public Declare PtrSafe Function SFB_ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '***** Used in Manage Data (XML) script
#Else
    Public Declare Function SFB_SetCursor Lib "user32" (ByVal hCursor As Long) As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare Function SFB_GetCursor Lib "user32" () As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare Function SFB_LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long '***** To set cursor shape as SFB_a "Hand" when SFB_a selection is needed
    Public Declare Function SFB_ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare Function SFB_IsIconic Lib "user32" (ByVal hWnd As Long) As Long '***** To verify if SFB_a window is minimized
    Public Declare Function SFB_BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long '***** To show and maximize IE window when opening KBE info page
    Public Declare Function SFB_FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpCaption As Any) '***** To give title bar SFB_a toolbar look, used also to call SFB_OpenFileDialog to give window handler
    Public Declare Function SFB_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '***** To give title bar SFB_a toolbar look
    Public Declare Function SFB_DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long '***** To give title bar SFB_a toolbar look
    Public Declare Function SFB_GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long '***** To find screen size, and if user has 1 or 2 screens (to position toolbar inside screen)
    Public Declare Function SFB_GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare Function SFB_GetDC Lib "user32" (ByVal hWnd As Long) As Long '***** To convert Points to Pixel (to position toolbar inside screen)
    Public Declare Function SFB_GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer '***** To detect ESC key press
    Public Declare Sub SFB_Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long) '***** To make SFB_a pause when running script
    Public Declare Function SFB_SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long  'To browse SFB_a directory
    Public Declare Function SFB_SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As SFB_BROWSEINFO) As Long 'To browse SFB_a directory
    Public Declare Function SFB_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'To browse SFB_a directory
    Public Declare Function SFB_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As SFB_OpenFilename) As Long 'To browse for SFB_a file
    Public Declare Function SFB_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As SFB_OpenFilename) As Long 'To browse for SFB_a file
    Public Declare Function SFB_GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long 'To get user ID
    Public Declare Function SFB_ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, ByVal Filename As String, Optional ByVal Parameters As String, Optional ByVal Directory As String, Optional ByVal WindowStyle As Long = vbMinimizedFocus) As Long
    
#End If

Public Const SFB_SW_SHOWNORMAL = 1  '***** Works with function "SFB_ShowWindow" > "half" maximize window
Public Const SFB_SW_MAXIMIZE = 3    '***** Works with function & "SFB_ShowWindow" > maximize window
Public Const SFB_SW_MINIMIZE = 6    '***** Works with function & "SFB_ShowWindow" > minimize window (could be 2)
Public Const SFB_SW_RESTORE = 9     '***** Works with function "SFB_ShowWindow" > restore window

Public Const SFB_PI As Double = 3.14159265358979 '*** value of Pi
Public Const SFB_SMALLNUMBER As Double = 0.00000001 '*** to avoid division by zero

Private Type SFB_OpenFilename 'To browse for SFB_a file or save SFB_a file
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

Private Type SFB_BROWSEINFO 'To browse SFB_a directory
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public SFB_dPointsToPixelRatioH As Double
Public SFB_dPointsToPixelRatioV As Double
Public SFB_dTitleBarHeight As Double        '***** Used to determine if Windows classic or XP Style
Public SFB_sTempDirectory As String         '***** Temp Directory
Public SFB_sSettingsPath As String          '***** Saved Settings
Public SFB_sCommonModule As String          '***** Common Module name
Public SFB_bToolbarCommandSelected As Boolean   '***** Used to determine if user selected clicks on toolbar
Public SFB_bStopMultiSelection As Boolean       '***** Used to determine if user cancelled multi selection

Public Function SFB_setCommonVariables()
    
'    CATIA.Caption = "CATIA V5 R" & CATIA.SystemConfiguration.Release & " SP" & CATIA.SystemConfiguration.ServicePack
'    CATIA.RefreshDisplay = True
'    CATIA.Interactive = True
'
'    '*****Restore Catia window if minimized
'    If SFB_IsIconic(SFB_FindWindow(vbNullString, CATIA.Caption)) Then SFB_ShowWindow SFB_FindWindow(vbNullString, CATIA.Caption), SFB_SW_RESTORE
    
    '*****SystemMetrics API is in pixels
    '*****Application (CATIA) left and top is in pixels
    '*****Userform left and top is in points
    '****************************************************
    '*****Get the conversion of points/pixel for H & SFB_V
    SFB_dPointsToPixelRatioH = 72 / SFB_GetDeviceCaps(SFB_GetDC(0), 88) '***** 72 points per inch / ??? pixels per inch
    SFB_dPointsToPixelRatioV = 72 / SFB_GetDeviceCaps(SFB_GetDC(0), 90) '***** 72 points per inch / ??? pixels per inch
    
    '*****Determine if user has Windows Classic or XP Style
    SFB_dTitleBarHeight = SFB_GetSystemMetrics32(31) 'Returns 18 Windows Classic Mode / 25 Windows XP Mode
    
    '***** Temp Directory (%tmp% or %temp%), saved settings path and Common Module name
'    SFB_sTempDirectory = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\"
'    SFB_sSettingsPath = SFB_sTempDirectory & SFB_GetTagValueFromFile(SFB_sKBEPathFile, "Saved Settings")
'    SFB_sCommonModule = SFB_GetTagValueFromFile(SFB_sKBEPathFile, "Common Module")
'
End Function


'********************************************************************************
'* Name: SFB_ReadTextFile
'* Purpose: Read SFB_a text file, then return the value found on line "iLineIndex"
'*          If "iLineIndex = 0" (Default value), function returns the value of last line
'*          If "iLineIndex" is greater than the number of lines, function returns the value of last line
'*          If "bCompleteText" is = true, function returns the complete text in the file, all lines separated by SFB_a return (vbcrlf)
'*          NOTE: Two strings separated by SFB_a coma "," is considered on two different lines!!!
'*
'* Assumption:
'*
'* Author: http://www.vbforums.com/showthread.php?342619-Classic-VB-How-can-I-read-write-SFB_a-text-file
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_ReadTextFile(ByVal SFB_sFilePath As String, Optional iLineIndex As Long = 0, Optional bCompleteText As Boolean = False) As String

    Dim SFB_FileNumber As Integer
    Dim SFB_iLineCount As Long
    Dim SFB_iLineText As String

    ' ensure that the file exists
    SFB_ReadTextFile = ""
    If Len(Dir$(SFB_sFilePath)) = 0 Then Exit Function
    
    'Open file
    SFB_FileNumber = FreeFile
    Open SFB_sFilePath$ For Input As #SFB_FileNumber

    'Read lines
    SFB_iLineCount = 0
    Do While Not EOF(SFB_FileNumber)

        SFB_iLineCount = SFB_iLineCount + 1
        Input #SFB_FileNumber, SFB_iLineText

        If bCompleteText Then
            If Trim(SFB_iLineText) <> "" Then
                SFB_ReadTextFile = IIf(SFB_ReadTextFile = "", SFB_iLineText, SFB_ReadTextFile & vbCrLf & SFB_iLineText)
            End If
        Else
            SFB_ReadTextFile = Trim(SFB_iLineText)
        End If

        If SFB_iLineCount = iLineIndex Then Exit Do

    Loop

    'Close the file
    Close #SFB_FileNumber

End Function


'********************************************************************************
'* Name: SFB_WriteTextFile
'* Purpose: Add SFB_a line at the end of SFB_a text file
'*
'* Assumption:
'*
'* Author: http://www.devhut.net/2011/06/06/vba-append-text-to-SFB_a-text-file/
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Sub SFB_WriteTextFile(ByVal SFB_sFilePath As String, Optional ByVal sLineText As String = "", Optional ByVal bResetFile = True)

    Dim SFB_FileNumber As Integer

    'Exit if folder does not exist
    If Len(Dir$(Left(SFB_sFilePath, InStrRev(SFB_sFilePath, "\")))) = 0 Then Exit Sub
    
    On Error Resume Next
    SFB_FileNumber = FreeFile                                               ' Get unused file number
    If Not (bResetFile) Then Open SFB_sFilePath For Append As #SFB_FileNumber   ' Connect to the file
    If bResetFile Then Open SFB_sFilePath For Output As #SFB_FileNumber         ' Connect to the file
    If sLineText <> "" Then Print #SFB_FileNumber, sLineText                ' Write string
    Close #SFB_FileNumber                                                   ' Close the file
    On Error GoTo 0
End Sub


'********************************************************************************
'* Name: SFB_TrimLine
'* Purpose: In SFB_a string, keep only what's found after the first ">", remove all
'*          spaces and all tabs. Replace all "\" by "/"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_TrimLine(ByVal sLineText As String) As String

    SFB_TrimLine = sLineText
    SFB_TrimLine = Mid(SFB_TrimLine, InStr(sLineText, ">") + 1) '***** Getting everything after the first ">"
    SFB_TrimLine = Replace(SFB_TrimLine, vbTab, "") '***** Remove all tabs
    SFB_TrimLine = Replace(SFB_TrimLine, "/", "\") '***** Replace "/" by "\" in string
    SFB_TrimLine = Trim(SFB_TrimLine) '***** Remove spaces before and after string

End Function


'********************************************************************************
'* Name: Get Tag Value From File
'* Purpose: In SFB_a file, returns Value associated to SFB_a tag , that is, value after
'*          identifier between <""> is returned. sFilepath is the path of text file to read
'*          if no tag value is found then "" is returned.
'*          <  XX XX  > are equal to <XX XX>
'*          tags are not case sensitive.
'*
'* Assumption: if more than one tag exist in SFB_a file then first tag value is returned
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_GetTagValueFromFile(ByVal SFB_sFilePath As String, ByVal sTag As String, Optional bReset = False) As String

Static SFB_oDict As Object
Dim SFB_sTagtosearch As String
Dim SFB_sEntireFile As String
Dim SFB_aEntireFile() As String

SFB_GetTagValueFromFile = ""

'Reset SFB_oDict
If bReset = True Then
    Set SFB_oDict = Nothing
    Exit Function
End If

'Create the dictionary if not already created
If SFB_oDict Is Nothing Then
    Set SFB_oDict = CreateObject("Scripting.Dictionary")
End If

'Read file content or retrieve in SFB_oDict
If Not SFB_oDict.Exists(SFB_sFilePath) Then
    SFB_sEntireFile = SFB_ReadTextFile(SFB_sFilePath, 0, True)
    Call SFB_oDict.Add(SFB_sFilePath, SFB_sEntireFile)
Else
    SFB_sEntireFile = SFB_oDict(SFB_sFilePath)
End If

'Split with vbNewLine
SFB_aEntireFile() = Split(SFB_sEntireFile, vbNewLine)

On Error Resume Next '**** to handle lines where no tags are found
For SFB_i = 0 To UBound(SFB_aEntireFile)
    SFB_sTagtosearch = SFB_aEntireFile(SFB_i)
    If InStr(SFB_sTagtosearch, "<") > 0 And InStr(SFB_sTagtosearch, ">") > 0 Then
        SFB_sTagtosearch = Mid(SFB_sTagtosearch, InStr(SFB_sTagtosearch, "<") + 1) '***** Getting everything after the first "<"
        SFB_sTagtosearch = Mid(SFB_sTagtosearch, 1, InStr(SFB_sTagtosearch, ">") - 1)  '***** Getting everything before the first ">"
        SFB_sTagtosearch = Trim(SFB_sTagtosearch)
        If SFB_sTagtosearch = sTag Then
            SFB_GetTagValueFromFile = SFB_TrimLine(SFB_aEntireFile(SFB_i))
            On Error GoTo 0
            Exit Function
        End If
    End If
Next
On Error GoTo 0
End Function



'********************************************************************************
'* Name: Set Tag Value From File
'* Purpose: In SFB_a file, replace Value associated to SFB_a tag , that is, value after
'*          identifier between <""> is returned. sFilepath is the path of text file to read
'*          if no tag value is found then "" is returned.
'*          <  XX XX  > are equal to <XX XX>
'*          tags are not case sensitive.
'*
'* Assumption: if more than one tag exist in SFB_a file then first tag value is returned
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_SetTagValueToFile(ByVal SFB_sFilePath As String, ByVal sTag As String, ByVal sNewTagValue As String)

Dim SFB_bTagFound As Boolean
Dim SFB_sTagtosearch As String
Dim SFB_sEntireFile As String
Dim SFB_aEntireFile() As String

SFB_bTagFound = False
SFB_sEntireFile = SFB_ReadTextFile(SFB_sFilePath, 0, True)      '**** read entire file
SFB_aEntireFile() = Split(SFB_sEntireFile, vbNewLine)
Call SFB_WriteTextFile(SFB_sFilePath) '**** Empty file content

On Error Resume Next '**** to handle lines where no tags are found
For SFB_i = LBound(SFB_aEntireFile) To UBound(SFB_aEntireFile)
    SFB_sTagtosearch = SFB_aEntireFile(SFB_i)
    SFB_sTagtosearch = Mid(SFB_sTagtosearch, InStr(SFB_sTagtosearch, "<") + 1) '***** Getting everything after the first "<"
    SFB_sTagtosearch = Mid(SFB_sTagtosearch, 1, InStr(SFB_sTagtosearch, ">") - 1)  '***** Getting everything before the first ">"
    SFB_sTagtosearch = Trim(SFB_sTagtosearch)
    If UCase(SFB_sTagtosearch) Like "*" & UCase(sTag) & "*" Then
        SFB_bTagFound = True
        SFB_aEntireFile(SFB_i) = "<" & sTag & ">" & vbTab & vbTab & SFB_TrimLine(sNewTagValue)
    End If
    Call SFB_WriteTextFile(SFB_sFilePath, SFB_aEntireFile(SFB_i), False)
Next

If Not (SFB_bTagFound) Then Call SFB_WriteTextFile(SFB_sFilePath, "<" & sTag & ">" & vbTab & vbTab & SFB_TrimLine(sNewTagValue), False)
On Error GoTo 0

End Sub


'********************************************************************************
'* Name: Open With Notepad
'* Purpose: Open SFB_a file using Notepad
'*
'* Assumption:
'*
'* Author: François Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_OpenWithNotepad() As Long
    
    sWIN_DIR = CATIA.SystemService.Environ("windir")
    
    On Error Resume Next
    Err.Clear
    Shell sWIN_DIR & "\NOTEPAD.EXE " & sLogFileName, SFB_SW_SHOWNORMAL
    SFB_OpenWithNotepad = Err.Number
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
Public Function SFB_KBESearch(ByVal StringToSearch As String, Optional ByVal SFB_iObject As AnyObject = Nothing, Optional oCurrentSelection As Selection = Nothing)
    
    If oCurrentSelection Is Nothing Then Set oCurrentSelection = CATIA.ActiveDocument.Selection
    
    If Not SFB_iObject Is Nothing Then
        oCurrentSelection.Clear
        oCurrentSelection.Add SFB_iObject
    End If
    
    Err.Clear
    On Error Resume Next
    CATIA.HSOSynchronized = False
    oCurrentSelection.Search StringToSearch
    CATIA.HSOSynchronized = True
    On Error GoTo 0
    
    Set SFB_KBESearch = oCurrentSelection
    
End Function


'********************************************************************************
'* Name: Multi Window Select Element
'* Purpose: Loop until user does SFB_a selection or cancel action
'*
'* Assumption: another sub can set SFB_bStopMultiSelection to false during loop, so the returned selection is empty
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_MultiWindowSelectElement(ByVal selectionMessage As String)
    
    On Error Resume Next '***** Error if user select something else than SFB_a feature (e.g. clicking on toolbar)
    CATIA.ActiveDocument.Selection.Clear
    
    Set SFB_MultiWindowSelectElement = Nothing
    SFB_bStopMultiSelection = False
    SFB_bToolbarCommandSelected = False
    
    Do
        DoEvents
        SFB_SetHandCursor
        CATIA.StatusBar = selectionMessage
        If CATIA.ActiveDocument.Selection.Count2 > 0 Then
            Set SFB_MultiWindowSelectElement = CATIA.ActiveDocument.Selection.Item2(1)
            Exit Do
        End If
        If CATIA.ActiveDocument Is Nothing Then SFB_bStopMultiSelection = True
        If SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then SFB_bStopMultiSelection = True
    Loop Until SFB_bStopMultiSelection Or SFB_bToolbarCommandSelected
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Set Hand Cursor
'* Purpose: Modify shape of cursor to SFB_a hand
'*          This is used with multi selection (previous sub) so user knows that he needs to select something
'*
'* Assumption:
'*
'* Author:
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_SetHandCursor()
    
    Dim SFB_HandCursor As Long
    On Error Resume Next
        If SFB_HandCursor = 0 Then
            SFB_HandCursor = SFB_LoadCursor(0, 32649&)
        ElseIf SFB_GetCursor() = SFB_HandCursor Then
            Exit Function
        End If
        SFB_SetCursor SFB_HandCursor
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
Public Function SFB_disableOptionNameCheck()
    
    Dim SFB_eNamingCheck As PartInfrastructureSettingAtt
    Set SFB_eNamingCheck = CATIA.SettingControllers.Item("CATMmuPartInfrastructureSettingCtrl")
    
    If SFB_eNamingCheck.NamingMode <> catNoNamingCheck Then
        SFB_eNamingCheck.NamingMode = catNoNamingCheck
        MsgBox "Name check was disabled. To change this setting again:  " & vbCrLf & _
                "Tools / Options / Part Infrastructure / Display / Checking Operation When Renaming        ", _
                vbExclamation, "Name check disabled"
    End If
    
    SFB_eNamingCheck.SaveRepository
    SFB_eNamingCheck.Commit
    
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
Public Function SFB_enableRelationUpdate()
    
    Dim SFB_eRelationUpdate As KnowledgeSheetSettingAtt
    Set SFB_eRelationUpdate = CATIA.SettingControllers.Item("CATLieKnowledgeSheetSettingCtrl")
    
    If SFB_eRelationUpdate.RelationsUpdateInPartContextEvaluateDuringUpdate <> 1 Then SFB_eRelationUpdate.RelationsUpdateInPartContextEvaluateDuringUpdate = 1
    If SFB_eRelationUpdate.RelationsUpdateInPartContextSynchronousRelations <> 1 Then SFB_eRelationUpdate.RelationsUpdateInPartContextSynchronousRelations = 1
    
    SFB_eRelationUpdate.SaveRepository
    SFB_eRelationUpdate.Commit
    
End Function


'********************************************************************************
'* Name: Clean empty publication
'* Purpose: Delete SFB_a publication that is not linked to any element
'*          e.g. A published element was deleted, but not it's publication
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_deleteBrokenPublications(ByVal oDocumentToClean As Document, Optional ByVal sProgressBarCaption As String = "Clean Publications", Optional ByVal bProgressMain As Boolean = False)
    
    Dim SFB_sPublishedName As String
    
    On Error Resume Next
    Set oPublications = oDocumentToClean.Product.Publications
    If SFB_frmProgress.Visible = False Then SFB_frmProgress.SFB_progressBarInitialize (sProgressBarCaption)
    
    iProgressCount2 = 0
    iProgressCount2Max = oDocumentToClean.Product.Publications.Count
    
    For SFB_i = oPublications.Count To 1 Step -1
        
        iProgressCount2 = iProgressCount2 + 1
        sProgBarComment = "Cleaning publication (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        If bProgressMain = False Then Call SFB_frmProgress.SFB_progressBarRepaint(sProgBarComment, iProgressCount2Max, iProgressCount2 + 1 - SFB_i, , , , SFB_frmProgress.lblTimer.Caption)
        If bProgressMain = True Then Call SFB_frmProgress.SFB_progressBarRepaint(SFB_frmProgress.lblMessageMain.Caption, SFB_frmProgress.pbProgressMain.Max, SFB_frmProgress.pbProgressMain.Value, sProgBarComment, iProgressCount2Max, iProgressCount2 + 1 - SFB_i, SFB_frmProgress.lblTimer.Caption)
        
        SFB_sPublishedName = ""
        SFB_sPublishedName = oPublications.Item(SFB_i).Valuation.Name 'Not valuated for an empty published parameter
        SFB_sPublishedName = oPublications.Item(SFB_i).Valuation.DisplayName 'DisplayName = "" for an empty published feature or body
        
        If SFB_sPublishedName = "" Then 'If Valuation.Name was not valuated (empty parameter) or Valuation.DisplayName = "" (empty feature)
            oPublications.Remove (oPublications.Item(SFB_i).Name)
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
'* Purpose: Ask user to select an assembly, and verifies if its name contains SFB_a string "productName" (arguments)
'*          If the assy has the wrong name, it returns Nothing
'*          If the assy has the good name, the active workbench is set to "Assembly Design"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_selectAssy(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal productName As String = "*", Optional ByVal userMessage As String = "")
      
    Dim SFB_sDocumentType As String
    Dim SFB_oSelectedItem As AnyObject
    On Error Resume Next
    
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    If TypeName(SFB_oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    '***** if we cannot access PartNumber property, put SFB_oSelectedItem.LeafProduct in default mode
    If SFB_oSelectedItem.LeafProduct.PartNumber = "" Then
        SFB_oSelectedItem.LeafProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    SFB_sDocumentType = ""
    SFB_sDocumentType = TypeName(SFB_oSelectedItem.Value.ReferenceProduct.Parent)
    If SFB_sDocumentType = "ProductDocument" Then
        Set selectedProduct = SFB_oSelectedItem.Value
    ElseIf SFB_sDocumentType = "PartDocument" Then
        If TypeName(SFB_oSelectedItem.Value.Parent.Parent.ReferenceProduct.Parent) = "ProductDocument" Then
            Set selectedProduct = SFB_oSelectedItem.Value.Parent.Parent
        End If
    Else
        SFB_sDocumentType = TypeName(SFB_oSelectedItem.LeafProduct.ReferenceProduct.Parent)
        If SFB_sDocumentType = "PartDocument" Then
            SFB_sDocumentType = TypeName(SFB_oSelectedItem.LeafProduct.Parent.Parent.ReferenceProduct.Parent)
            If SFB_sDocumentType = "ProductDocument" Then
                Set selectedProduct = SFB_oSelectedItem.LeafProduct.Parent.Parent
            End If
        End If
    End If
    
    If Not selectedProduct Is Nothing And selectedProduct.PartNumber Like productName Then
        '****If name contains "productName" and selection is not SFB_a part instance
        Set selectedRootDocument = SFB_oSelectedItem.Document
        CATIA.ActiveDocument.Selection.Clear
        
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is SFB_a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
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
'* Purpose: Ask user to select any element. If the selected element is SFB_a product
'*          with the name "productName", this product is returned. Otherwise, the
'*          root product is found, and script verifies if its name contains SFB_a string
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
Public Sub SFB_selectRootAssy(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal productName As String = "*", Optional ByVal userMessage As String = "")
    
    Dim SFB_oSelectedItem As AnyObject
    On Error Resume Next
    
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    If TypeName(SFB_oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    Set selectedRootDocument = Nothing
    Set selectedProduct = Nothing
    Set selectedProduct = SFB_oSelectedItem.LeafProduct.ReferenceProduct 'Returns nothing if part is not in design mode, Returns nothing if we are in SFB_a part (not in product mode)
    
    'If SFB_a product "productName" was directly selected by user, return it (don't return Root product, but selected product)
    If Not selectedProduct Is Nothing Then
        If Not SFB_oSelectedItem.LeafProduct.PartNumber Like productName Or TypeName(selectedProduct.Parent) <> "ProductDocument" Then
            Set selectedProduct = Nothing
        End If
    End If
    
    'Else, return Root product
    If selectedProduct Is Nothing Then
        If SFB_oSelectedItem.Document.Product.Name Like productName And TypeName(SFB_oSelectedItem.Document) = "ProductDocument" Then
            Set selectedProduct = SFB_oSelectedItem.Document.Product
        Else
            Set selectedProduct = Nothing
        End If
    End If
    
    CATIA.ActiveDocument.Selection.Clear
    
    If Not selectedProduct Is Nothing Then
        Set selectedRootDocument = SFB_oSelectedItem.Document
        
        '****If name contains SEED_ASSY and selection is not SFB_a part instance
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is SFB_a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
        If CATIA.GetWorkbenchId <> "Assembly" Then
            CATIA.StartWorkbench "Assembly"
        End If
    End If
    
End Sub


'********************************************************************************
'* Name: Select Part
'* Purpose: Ask user to select SFB_a part, and verifies if its name contains SFB_a string "partName" (arguments)
'*          If the part has the wrong name, it returns Nothing
'*          Note: user can select any element inside the part, or its instance, this will find the part
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_selectPart(ByRef selectedPart As Part, ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal PartName As String = "*", Optional ByVal userMessage As String = "")
    
    Dim SFB_oSelectedItem As AnyObject
    On Error Resume Next
    
    ' Initialize variables
    Set selectedPart = Nothing
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    
    ' Valuate selectedPart, selectedProduct and selectedRootDocument (depending if we are in product or part mode)
    If TypeName(SFB_oSelectedItem.Document) = "ProductDocument" Then
        ' Put part in design mode if needed (works only in product mode)
        If TypeName(SFB_oSelectedItem.LeafProduct.ReferenceProduct.Parent) = "PartDocument" Then
            SFB_oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        End If
        
        Set selectedPart = SFB_oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part 'Works only in product mode.
        Set selectedProduct = SFB_oSelectedItem.LeafProduct
    ElseIf TypeName(SFB_oSelectedItem.Document) = "PartDocument" Then
        Set selectedPart = SFB_oSelectedItem.Document.Part 'Works only in part mode.
        Set selectedProduct = SFB_oSelectedItem.Document.Product
    End If
    Set selectedRootDocument = SFB_oSelectedItem.Document
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
'* Purpose: Ask user to select an instance of SFB_a CATPart or from SFB_a CATProduct
'*          If the assy has the wrong name, it returns Nothing
'*          If the assy has the good name, the active workbench is set to "Assembly Design"
'*
'* Assumption:
'*
'* Author: Francois Charette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_SelectInstance(ByRef selectedProduct As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")
      
    Dim SFB_sDocumentType As String
    On Error Resume Next
    
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    If TypeName(SFB_oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    '***** if we cannot access PartNumber property, put SFB_oSelectedItem.LeafProduct in default mode
    If SFB_oSelectedItem.LeafProduct.PartNumber = "" Then
        SFB_oSelectedItem.LeafProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    Set selectedProduct = SFB_oSelectedItem.LeafProduct
    
    If Not selectedProduct Is Nothing Then
        Set selectedRootDocument = SFB_oSelectedItem.Document
        CATIA.ActiveDocument.Selection.Clear
        
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is SFB_a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
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
'* Updated by:Abhishek Kamboj
'* Language: VBA
'********************************************************************************
Public Sub SFB_selectPlanarFace(ByRef selectedFace As Face, ByRef parentPart As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")

    ' Initialize variables
    Set selectedFace = Nothing
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    
    '*** putting part in design mode if needed
    On Error Resume Next
'        Set xxx = SFB_oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
    If Err.Number <> 0 Then
        On Error GoTo 0
        SFB_oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        SFB_sAnswer = MsgBox("Part has been put in design mode" + vbNewLine + "Please reselect the face.", vbExclamation)
        Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    End If
    On Error GoTo 0

    'The selection must be done in SFB_a CATProduct window
    'The selected face must be SFB_a "PlanarFace"
    If TypeName(SFB_oSelectedItem.Value) = "PlanarFace" And TypeName(SFB_oSelectedItem.Document) = "ProductDocument" Then
        Set selectedFace = SFB_oSelectedItem.Value
        Set parentPart = SFB_oSelectedItem.LeafProduct
        Set selectedRootDocument = SFB_oSelectedItem.Document
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
Public Sub SFB_selectFace(ByRef selectedFace As Face, ByRef parentPart As Product, ByRef selectedRootDocument As Document, Optional ByVal userMessage As String = "")

    ' Initialize variables
    Set selectedFace = Nothing
    Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    If SFB_oSelectedItem Is Nothing Then Exit Sub
    
    '*** putting part in design mode if needed
    On Error Resume Next
'        Set xxx = SFB_oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
    If Err.Number <> 0 Then
        On Error GoTo 0
        SFB_oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        SFB_sAnswer = MsgBox("Part has been put in design mode" + vbNewLine + "Please reselect the face.", vbExclamation)
        Set SFB_oSelectedItem = SFB_MultiWindowSelectElement(userMessage)
    End If
    On Error GoTo 0

    'The selection must be done in SFB_a CATProduct window
    If TypeName(SFB_oSelectedItem.Value) Like "*Face" And TypeName(SFB_oSelectedItem.Document) = "ProductDocument" Then
        Set selectedFace = SFB_oSelectedItem.Value
        Set parentPart = SFB_oSelectedItem.LeafProduct
        Set selectedRootDocument = SFB_oSelectedItem.Document
    End If

End Sub


'********************************************************************************
'* Name: Get body from object
'* Purpose: Finding the body of any selected object. This function calls itself
'*          recursively to access the parent of the object passed in until
'*          SFB_a Body object is reached.
'*
'* Assumption:
'*
'* Author: Julien Bigaouette (inspired from http://v5vb.wordpress.com/2009/12/15/get-part/ )
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_GetBodyFromObject(ByRef SFB_iObject As Variant) As Body
    
    Dim SFB_oChildBody As Body
    Dim SFB_oParentBody As Body
    Dim SFB_oSelection As Selection
    
    On Error Resume Next '****Possible error if SFB_iObject.Parent does not exist
    Set SFB_oSelection = CATIA.ActiveDocument.Selection '***** Adding object to selection, then setting object = selection, makes us sure to get the good type on the object. Otherwise, SFB_a body could be typed as "Hybridbodies" or "Shapes" for example
    SFB_oSelection.Clear
    SFB_oSelection.Add SFB_iObject
    If SFB_oSelection.Count2 > 0 Then Set SFB_iObject = SFB_oSelection.Item2(1).Value
    SFB_oSelection.Clear
    
    If TypeName(SFB_iObject) = "Body" Then
        'If the object passed in is SFB_a body under SFB_a boolean operation, find its parent body, then loop again on that body (parent body could be under another boolean operation)
        If SFB_iObject.InBooleanOperation Then
            For Each obody In SFB_iObject.Parent
                For Each oShape In obody.Shapes
                    Set SFB_oChildBody = Nothing
                    Set SFB_oChildBody = oShape.Body
                    If Not SFB_oChildBody Is Nothing Then
                        If SFB_oChildBody.Name = SFB_iObject.Name Then
                            SFB_oSelection.Clear
                            SFB_oSelection.Add SFB_iObject
                            SFB_oSelection.Add SFB_oChildBody
                            If SFB_oSelection.Count2 = 1 Then 'If only 1 item is selected, it means SFB_iObject and SFB_oChildBody is the exact same body > We found the good SFB_oParentBody
                                Set SFB_oParentBody = obody
                                Set SFB_GetBodyFromObject = SFB_GetBodyFromObject(SFB_oParentBody)
                                SFB_oSelection.Clear
                                Exit Function
                            End If
                        End If
                    End If
                Next
            Next
        'If the object passed in is SFB_a body that is not under SFB_a boolean operation, return it and exit
        Else
             Set SFB_GetBodyFromObject = SFB_iObject
             Exit Function
        End If
        
    ElseIf TypeName(SFB_iObject) = "Part" Or TypeName(SFB_iObject) = "Product" Then
        'If SFB_iObject is SFB_a part or SFB_a product, SFB_iObject is too high in the tree
        Set SFB_GetBodyFromObject = Nothing
        Exit Function
        
    ElseIf TypeName(SFB_iObject) = TypeName(SFB_iObject.Parent) Then
        'If the type of this object is the same as the parent object then return nothing.
        'The reason is the Parent property of some objects simply returns the same object.
        'This will result in an infinite loop (e.g. An UDF could return itself as parent)
    
        Set SFB_GetBodyFromObject = Nothing
        Exit Function
    
    Else
        'Call the function again and pass it the object’s parent
        Set SFB_GetBodyFromObject = SFB_GetBodyFromObject(SFB_iObject.Parent)
        
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
Public Function SFB_GetDescriptionFromProduct(ByRef SFB_iObject As Variant, Optional ByVal sDescription As String = "*") As String
    
    On Error Resume Next '****Possible error if SFB_iObject.Parent does not exist
    SFB_GetDescriptionFromProduct = ""
    
    If TypeName(SFB_iObject) = "Product" Then
        'If SFB_iObject is SFB_a part or SFB_a product, SFB_iObject is too high in the tree
        
        If SFB_iObject.DescriptionInst <> "" And SFB_iObject.DescriptionInst Like sDescription Then
            SFB_GetDescriptionFromProduct = SFB_iObject.DescriptionInst
            Exit Function
        Else
            SFB_GetDescriptionFromProduct = SFB_GetDescriptionFromProduct(SFB_iObject.Parent)
        End If
        
    ElseIf TypeName(SFB_iObject) = TypeName(SFB_iObject.Parent) And SFB_iObject.Name = SFB_iObject.Parent.Name Then
        'If the type of this object is the same as the parent object then return nothing.
        'The reason is the Parent property of some objects simply returns the same object.
        'This will result in an infinite loop (e.g. An UDF returns itself as parent)
        
        SFB_GetDescriptionFromProduct = ""
        Exit Function
        
    Else
        'Call the function again and pass it the object’s parent
        SFB_GetDescriptionFromProduct = SFB_GetDescriptionFromProduct(SFB_iObject.Parent)
    End If
    
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Scan Seed Part, find Subset parameters, create SFB_a collection of objects "SSParameter"
'*
'* Purpose:
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
'Public Sub ScanSubsets(ByVal oProductCollection As Collection)
'
'    Dim SFB_iObject As Variant
'    Set SSParameters = New Collection 'Public Variable to be declared in another module
'
'    On Error Resume Next
'    For Each SFB_iObject In oProductCollection
'        '**** Two names of parameter set to cover: "SubSetParameters" & "ParametersSubSet"
'        Set oParameterSet = Nothing
'        Set oParameterSet = SFB_iObject.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
'        Set oParameterSet = SFB_iObject.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
'        If Not oParameterSet Is Nothing Then
'            'Create SFB_a collection of objects "SSParameter" for each SS parameter found
'            For Each oParameter In oParameterSet.DirectParameters
'                sCurrentSSParameterName = Mid(oParameter.Name, InStrRev(oParameter.Name, "\SS_") + Len("\SS_"))
'                SSParameters.Add New SSParameter, sCurrentSSParameterName
'                SSParameters.Item(sCurrentSSParameterName).ParameterObject = oParameter
'                SSParameters.Item(sCurrentSSParameterName).ParameterName = sCurrentSSParameterName
'
'                For Each sSubsetItem In SFB_ConvertStringToCollection(oParameter.ValueAsString, "|")
'                    SSParameters.Item(sCurrentSSParameterName).AddToSubsetValues sSubsetItem
'                Next
'            Next
'        End If
'    Next
'    On Error GoTo 0
'
'End Sub


'********************************************************************************
'* Name: Convert String to Collection
'* Purpose: From SFB_a string containing separators, create SFB_a collection of strings
'*          (e.g. sStringValue = "AAA;BBB;CCC" and sSeparator = ";" >> Collection.Item(2) = "BBB"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_ConvertStringToCollection(ByRef sStringValue As String, ByRef sSeparator As String, Optional sTag As Boolean = False) As Collection
    
    Dim SFB_iObject As Variant
    Set SFB_ConvertStringToCollection = New Collection
    
    On Error Resume Next
    For Each SFB_iObject In Split(sStringValue, sSeparator)
        If Not SFB_iObject = "" And sTag = False Then
            SFB_ConvertStringToCollection.Add SFB_iObject
        ElseIf Not SFB_iObject = "" And sTag = True Then
            SFB_ConvertStringToCollection.Add SFB_iObject, SFB_iObject
        End If
    Next
    On Error GoTo 0
    
End Function


'********************************************************************************
'* Name: Convert Collection to String
'* Purpose: From SFB_a collection of strings, create one string separating collection values with SFB_a specific character
'*          (e.g. Collection.Item(1) = "AAA", Collection.Item(2) = "BBB", Separator = ":" >> String = "AAA:BBB"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function SFB_ConvertCollectionToString(ByRef oCollection As Collection, ByRef sSeparator As String) As String
    
    Dim SFB_iObject As Variant
    
    On Error Resume Next
    If oCollection.Item(1) = Empty Then Exit Function
    On Error GoTo 0
    
    SFB_ConvertCollectionToString = ""
    For Each SFB_iObject In oCollection
        If SFB_ConvertCollectionToString = "" Then
            SFB_ConvertCollectionToString = SFB_iObject
        Else
            SFB_ConvertCollectionToString = SFB_ConvertCollectionToString & sSeparator & SFB_iObject
        End If
    Next
    
End Function


'********************************************************************************
'* Name: Listbox move down
'* Purpose: Change selection > Select element under selected one (bMoveElementDown = false)
'*          Move selected elements down inside SFB_a listbox (bMoveElementDown = true)
'*
'* Assumption:
'*
'* Author: http://www.xtremevbtalk.com/archive/index.php/t-180834.html
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_ListBoxMoveDown(ByRef oListBox As ListBox, Optional bMoveElementDown As Boolean = False)
    
    Dim SFB_iPosition As Long
    Dim SFB_sNextItem As String
    
    SFB_iPosition = oListBox.ListCount - 1
    For SFB_i = oListBox.ListCount - 1 To 0 Step -1
        If oListBox.Selected(SFB_i) Then
            If Not SFB_i = SFB_iPosition Then
                With oListBox
                    If bMoveElementDown Then
                        SFB_sNextItem = .List(SFB_i + 1)
                        .List(SFB_i + 1) = .List(SFB_i)
                        .List(SFB_i) = SFB_sNextItem
                        .ListIndex = SFB_i + 1
                    End If
                    .Selected(SFB_i) = False
                    .Selected(SFB_i + 1) = True
                End With
            End If
            SFB_iPosition = SFB_iPosition - 1
        End If
    Next SFB_i
    
End Function


'********************************************************************************
'* Name: Listbox move up
'* Purpose: Change selection > Select element above selected one (bMoveElementUp = false)
'*          Move selected elements up inside SFB_a listbox (bMoveElementDown = true)
'*
'* Assumption:
'*
'* Author: http://www.xtremevbtalk.com/archive/index.php/t-180834.html
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_ListBoxMoveUp(ByRef oListBox As ListBox, Optional bMoveElementUp As Boolean = False)
    
    Dim SFB_iPosition As Long
    Dim SFB_sNextItem As String
    
    SFB_iPosition = 0
    For SFB_i = 0 To oListBox.ListCount - 1
        If oListBox.Selected(SFB_i) Then
            If Not SFB_i = SFB_iPosition Then
                With oListBox
                    If bMoveElementUp Then
                        SFB_sNextItem = .List(SFB_i - 1)
                        .List(SFB_i - 1) = .List(SFB_i)
                        .List(SFB_i) = SFB_sNextItem
                        .ListIndex = SFB_i - 1
                    End If
                    .Selected(SFB_i) = False
                    .Selected(SFB_i - 1) = True
                End With
            End If
            SFB_iPosition = SFB_iPosition + 1
        End If
    Next SFB_i
    
End Function


'********************************************************************************
'* Name: Listbox reorder
'* Purpose: Reorder elements in alphabetical order inside SFB_a listbox
'*
'* Assumption:
'*
'* Author: http://www.vbaexpress.com/forum/showthread.php?t=26064
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_ListBoxReorder(ByRef oListBox As ListBox)
    
    Dim SFB_sNextItem As String
    
    For SFB_i = 0 To oListBox.ListCount - 1
        For SFB_j = SFB_i + 1 To oListBox.ListCount - 1
            If oListBox.List(SFB_i) > oListBox.List(SFB_j) Then
                SFB_sNextItem = oListBox.List(SFB_j)
                oListBox.List(SFB_j) = oListBox.List(SFB_i)
                oListBox.List(SFB_i) = SFB_sNextItem
            End If
        Next SFB_j
    Next SFB_i
    
End Function


'********************************************************************************
'* Name: Show Open File Dialog
'* Purpose: Promt user to select SFB_a file / to save SFB_a file as...
'*
'* Assumption: sFilter = "Filter1 Name|Filter1 Value|Filter2 Name|Filter2 Value ..."
'*              Example: "All Files(*.*)|*.*|JPG Image(*.jpg)|*.jpg"
'*
'* Information: http://docvb.free.fr/apidetail.php?idapi=136
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Function SFB_OpenFileDialog(ByVal sFilter As String, _
    Optional ByVal bSaveAsDialog As Boolean = False, _
    Optional ByVal sDefaultExtension As String, _
    Optional ByVal sInitialDirectory As String, _
    Optional ByVal sInitialFileName As String, _
    Optional ByVal sWindowTitle As String, _
    Optional ByVal iParentWindowHandler As Long) As String
    
    Dim SFB_OFN As SFB_OpenFilename
    Dim SFB_bFileSelected As Boolean
    
    Const OFN_FILEMUSTEXIST = &H1000 'The user can type only names of existing files in the File Name entry field
    Const OFN_HIDEREADONLY = &H4 'Hides Read Only check box
    Const OFN_OVERWRITEPROMPT = &H2 'Prompt user when overwriting SFB_a file
    
    On Error Resume Next
    ' set the values for the OpenFileName struct
    With SFB_OFN
        .hwndOwner = iParentWindowHandler
        .lStructSize = Len(SFB_OFN)
        .lpstrFilter = Replace(sFilter, "|", vbNullChar) & vbNullChar
        .lpstrFile = Left$(sInitialFileName & String$(1024, vbNullChar), 1024)
        .nMaxFile = Len(.lpstrFile)
        .flags = OFN_FILEMUSTEXIST + OFN_OVERWRITEPROMPT + OFN_HIDEREADONLY
        .lpstrInitialDir = sInitialDirectory
        .lpstrDefExt = sDefaultExtension
        .lpstrTitle = sWindowTitle
    End With
    
    ' show the dialog (Save As or Open)
    If Not (bSaveAsDialog) Then SFB_bFileSelected = SFB_GetOpenFileName(SFB_OFN)
    If bSaveAsDialog Then SFB_bFileSelected = SFB_GetSaveFileName(SFB_OFN)
    If SFB_bFileSelected Then
        ' extract the selected file (including the path)
        SFB_OpenFileDialog = Left$(SFB_OFN.lpstrFile, InStr(SFB_OFN.lpstrFile, vbNullChar) - 1)
    End If
End Function


'********************************************************************************
'* Name: Show Open Directory Dialog
'* Purpose: Promt user to select SFB_a directory
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette (Default folder)
'* Language: VBA
'********************************************************************************
Public Function SFB_OpenDirectoryDialog(Optional ByVal userMessage As String = "Select SFB_a folder", Optional ByRef sDefaultDirectory As String = vbNullString) As String
    
    Dim SFB_bInfo As SFB_BROWSEINFO
    Dim SFB_sPath As String
    
    With SFB_bInfo
        .lpszTitle = userMessage                         '   Prompt message
        .pidlRoot = 0&                                   '   Root folder = Desktop
        .ulFlags = &H1                                   '   Type of directory to return
        .lpfn = SFB_GetAddress(AddressOf SFB_BrowseCallbackProc) '   Default folder
        .lParam = StrPtr(sDefaultDirectory)              '   Default folder
    End With
    
    SFB_sPath = Space$(512)
    If SFB_SHGetPathFromIDList(ByVal SFB_SHBrowseForFolder(SFB_bInfo), ByVal SFB_sPath) Then
        SFB_OpenDirectoryDialog = Left(SFB_sPath, InStr(SFB_sPath, Chr$(0)) - 1)
        sDefaultDirectory = SFB_OpenDirectoryDialog
    Else
        SFB_OpenDirectoryDialog = ""
    End If
    
End Function

Private Function SFB_BrowseCallbackProc(ByVal hWnd&, ByVal msg&, ByVal lp&, ByVal InitDir$) As Long
   
   If (msg = 1) And (InitDir <> "") Then
      Call SFB_SendMessage(hWnd, &H466, 1, ByVal InitDir$)
   End If
   SFB_BrowseCallbackProc = 0
   
End Function

Private Function SFB_GetAddress(ByVal Addr As Long) As Long
   SFB_GetAddress = Addr
End Function


'********************************************************************************
'* Name: Open Web Page
'* Purpose: This sub receive SFB_a URL (strPartURL As String) as input
'*          It verifies if the page is already open
'*          If it is open, it maximize it and bring it front
'*          If the page is not already open, it opens SFB_a new one and maximize it
'
'*          Optional argument: length corresponding to the beginning of the string.
'*          Example:
'*          URLLength = 15 (corresponds to I:/V5_KBE_Tools > 15 characters)
'*          I want to open SFB_a new page I:/V5_KBE_Tools/Production/07_KBE_Broadcast/KBE_Info.htm
'*          ONLY if I don't find an existing page with the address starting with "I:/V5_KBE_Tools"
'*          Otherwise, I maximize the existing page (e.g. "I:\V5_KBE_Tools\Production\07_KBE_Broadcast\Global 6000\G6000ChangeManagementInfo.htm")
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub SFB_openPage(ByVal strPartURL As String, Optional ByVal URLLength As Integer)
    
    Dim SFB_objShellWindows As Variant
    Dim SFB_objShell As Variant
    Dim SFB_objIE As Variant
    
    If Len(strPartURL) = 0 Then Exit Sub
    
    On Error Resume Next
    Err.Clear
    Set SFB_objShell = CreateObject("Shell.Application")
    Set SFB_objShellWindows = SFB_objShell.Windows
    If Err <> 0 Then Exit Sub
    
    '*****Verify if there's SFB_a window open*****
    If SFB_objShellWindows.Count = 0 Then
        GoTo OpenNewWindow
    End If
    
    '*****Verify all windows, try to find the page with given URL and bring it to top*****
    '*****LIMIT CASE: if there's many tabs opened (IE Window), it does not activate it, it shows the actual active tab*****
    If URLLength = 0 Then URLLength = Len(strPartURL)
    
    For SFB_i = 0 To SFB_objShellWindows.Count - 1
        Set SFB_objIE = SFB_objShellWindows.Item(CLng(SFB_i))
        
        If Not (SFB_objIE Is Nothing) Then
            If InStr(Replace(SFB_objIE.LocationURL, "/", "\"), Trim(Left(strPartURL, URLLength))) Then
                
                '*****Maximize the page if it is minimized
                If SFB_IsIconic(SFB_objIE.hWnd) Then SFB_ShowWindow SFB_objIE.hWnd, SFB_SW_RESTORE
                
                '*****Bring the page to top in its actual state*****
                '*****Does not maximize SFB_a minimized window!! Idem as AppActivate method*****
                SFB_BringWindowToTop SFB_objIE.hWnd
                
                Set SFB_objShellWindows = Nothing
                Set SFB_objShell = Nothing
                Set SFB_objIE = Nothing
                
                Exit Sub
            End If
        End If
    Next SFB_i
    
    '*****Open URL in SFB_a new window*****
OpenNewWindow:
    '***Test
    Dim SFB_objxIE
    Dim SFB_iServerWaitTime
    SFB_iServerWaitTime = 500
    Set SFB_objxIE = CreateObject("InternetExplorer.Application")
    Err.Clear
    SFB_objxIE.Navigate strPartURL
    With SFB_objxIE
        Do
            DoEvents
            If Err.Number <> 0 Then Exit Do
            .waitForResponse SFB_iServerWaitTime
        Loop Until .readyState = 4
    End With
    
    '*****

'    Dim lSuccess As Long
   ' X = CATIA.SystemService.ExecuteProcessus("'C:\Program Files (x86)\Internet Explorer\IEXPLORE.EXE' " & strPartURL & "")
   ' lSuccess = SFB_ShellExecute(0, "Open", strPartURL, "", "", 0)

'    Set SFB_objIE = CreateObject("InternetExplorer.Application")
'
'    '*****Maximize the page
'    SFB_ShowWindow SFB_objIE.hWnd, SFB_SW_MAXIMIZE
'    SFB_objIE.Navigate (strPartURL)
'    SFB_objIE.Visible = True
    
    '*****Reset variables*****
    Set SFB_objShellWindows = Nothing
    Set SFB_objShell = Nothing
    Set SFB_objIE = Nothing
    
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
Public Function SFB_ProductNameCheck(ByVal NameToCheck As String, Optional ByVal bEngFileRequested As Boolean = True, Optional ByVal bDesignFileRequested As Boolean = True) As Boolean
    
    Dim SFB_bNHAPartNumberTest As Boolean
    Dim SFB_WarningMessage As String
    Dim SFB_FileTypeGuess As String
    
    SFB_bNHAPartNumberTest = False
    SFB_WarningMessage = ""
    SFB_FileTypeGuess = "Unknown Format:  "
    
    '********* Engineering File - Global 6000 nomenclature validation **********
    If Len(NameToCheck) = 14 And bEngFileRequested And NameToCheck Like "7*" Then
        
        SFB_WarningMessage = "Engineering file (Global 6000) detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
        
        If NameToCheck Like "7K*" Or NameToCheck Like "7G*" Then

            If Mid(NameToCheck, 11, 1) Like "-" Then
                If NameToCheck Like "??##???###-###*" And Not Mid(NameToCheck, 5, 3) Like "*#*" Then
                    SFB_bNHAPartNumberTest = True
                Else
                    SFB_WarningMessage = SFB_WarningMessage & """MM"", ""XXX"" and ""DDD"" should be numeric characters only " & vbCrLf & _
                                    """III"" should be alphabetic characters only "
                End If
            Else
                SFB_WarningMessage = SFB_WarningMessage & "Character 11 should be ""-"" "
            End If
        
        Else
            SFB_WarningMessage = "The first two char must be either '7K' or '7D'    "
        End If
    
    '********* Engineering File - Global 7000/8000 nomenclature validation **********
    ElseIf Len(NameToCheck) >= 17 And Len(NameToCheck) <= 20 And bEngFileRequested And NameToCheck Like "G025*" Then
        
        SFB_WarningMessage = "Engineering file (Global 7000/8000) detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
        
        If Mid(NameToCheck, 10, 1) Like "-" Then
            
            If NameToCheck Like "G025#####-###?###*" Then
                
                If NameToCheck Like "G025#####-###E###*" Then
                    
                    If NameToCheck Like "G025#####-###E###" Or NameToCheck Like "G025#####-###E###?##" Then
                        SFB_bNHAPartNumberTest = True
                    Else
                        SFB_WarningMessage = SFB_WarningMessage & """rr"" should be numeric characters only "
                    End If
                Else
                    SFB_WarningMessage = SFB_WarningMessage & "Character 14 should be ""E"" "
                End If
            
            Else
                SFB_WarningMessage = SFB_WarningMessage & """#####"", ""DDD"" and ""xxx"" should be numeric characters only    "
            End If
            
        Else
            SFB_WarningMessage = SFB_WarningMessage & "Character 10 should be ""-"" "
        End If
        

    '********* Design File - Global 6000 nomenclature validation **********
    ElseIf Len(NameToCheck) = 20 And bDesignFileRequested Then
        SFB_WarningMessage = "Design file detected." & vbCrLf & vbCrLf & "NHA name is not valid: " & vbCrLf
            
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
                        SFB_bNHAPartNumberTest = True
                    Else
                        SFB_WarningMessage = SFB_WarningMessage & """####"", ""MM"", ""SS"", ""XX"" and ""DDD"" should be numeric characters    "
                    End If
                Else
                    SFB_WarningMessage = SFB_WarningMessage & "Characters 9, 14 and 17 should be ""-"" "
                End If
                
            Else
                SFB_WarningMessage = SFB_WarningMessage & "Selected file name does not contain any of the following prefix:    " & vbCrLf & _
                                 "GXRS, G500, G600, G700, G800, 5000, 6000, 7000, 8000    "
            End If
    
    Else
        SFB_WarningMessage = "Unknown Format:  " & vbCrLf & "Selected NHA name is not valid    "
    End If
    
    
    '********* Warning message **********
    If SFB_bNHAPartNumberTest = False Then
        
        If bEngFileRequested Then SFB_WarningMessage = SFB_WarningMessage & vbCrLf & vbCrLf & _
                                                "      Eng. nomenclature (Global 6000):  7KMMIIIXXX-DDD  " & vbCrLf & vbCrLf & _
                                                "             MM = KBE monument number" & vbCrLf & _
                                                "             III = 3 digit iteration, alphabetic only (SFB_i.e. AAA, AAB...)" & vbCrLf & _
                                                "             XXX = KBE part index number" & vbCrLf & _
                                                "             DDD = 3 digit dash number" & vbCrLf & vbCrLf & _
                                                "      Eng. nomenclature (Global 7000/8000):  G025#####-DDDEEEErrr  " & vbCrLf & vbCrLf & _
                                                "             ##### = Base number" & vbCrLf & _
                                                "             DDD = 3 digit dash number" & vbCrLf & _
                                                "             EEEE = 4 digit envelope number (SFB_i.e. E001, E002...)" & vbCrLf & _
                                                "             rrr = Optional 3 digit representation (SFB_i.e. A01, A02...)"
                                                
        If bDesignFileRequested Then SFB_WarningMessage = SFB_WarningMessage & vbCrLf & vbCrLf & _
                                                "      Design nomenclature:  GXRS####-MMSS-XX-DDD        " & vbCrLf & vbCrLf & _
                                                "             #### = Project number" & vbCrLf & _
                                                "             MM = KBE monument number" & vbCrLf & _
                                                "             SS = sub-monument number" & vbCrLf & _
                                                "             XX = Additional information number" & vbCrLf & _
                                                "             DDD = 3 digit dash number"
        
        SFB_WarningMessage = vbCrLf & SFB_WarningMessage & vbCrLf & vbCrLf & "Do you want to continue?          " & vbCrLf & vbCrLf
        
        If MsgBox(SFB_WarningMessage, vbExclamation + vbYesNo, "Warning") = vbNo Then
            SFB_ProductNameCheck = False
        Else
            SFB_ProductNameCheck = True
        End If
    Else
        SFB_ProductNameCheck = True
    End If
    
End Function


'********************************************************************************
'* Name: Positioning functions
'* Purpose: All following functions are used to position SFB_a part in the space
'*
'* Assumption:
'*
'* Author: François Charette
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************
Public Sub SFB_SnapPartToPart(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceProduct As Product)
    
    'Purpose: Move child product instance (oProductToPosition) at inside its NHA (oProductToPositionNHA) at
    'same global position than reference product instance (oReferenceProduct)
    
    Dim SFB_dReferenceProductMatrix(11) As Variant
    Dim SFB_dTargetNHAMatrix(11) As Variant
    Dim SFB_dTargetMatrix(11) As Variant
    
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = SFB_BuildParentList(oReferenceProduct)
    Call SFB_CalculateInstanceGlobalPositionMatrix(sParentList, SFB_dReferenceProductMatrix)
    
    sParentList = SFB_BuildParentList(oProductToPositionNHA)
    Call SFB_CalculateInstanceGlobalPositionMatrix(sParentList, SFB_dTargetNHAMatrix)
    
    Call SFB_FindChildMatrix(SFB_dReferenceProductMatrix, SFB_dTargetNHAMatrix, SFB_dTargetMatrix)
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents SFB_dTargetMatrix
    
End Sub

Public Sub SFB_SnapPartToReferenceAxis(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceProduct As Product, ByVal oReferenceAxis As AxisSystem)
    
    'Purpose: Move child product instance (oProductToPosition) inside its NHA (oProductToPositionNHA) at
    'same global position than reference axis given (oReferenceAxis)
    
    Dim SFB_dReferenceProductMatrix(11) As Variant
    Dim SFB_dAxisMatrix(11) As Variant
    Dim SFB_dTargetNHAMatrix(11) As Variant
    Dim SFB_dTargetMatrix(11) As Variant
    
    'SFB_Retrieve Global Matrix of reference product
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = SFB_BuildParentList(oReferenceProduct)
    Call SFB_CalculateInstanceGlobalPositionMatrix(sParentList, SFB_dReferenceProductMatrix)
    
    'SFB_Retrieve Axis position (global matrix)
    Call SFB_GetAxisPosition(oReferenceAxis, SFB_dAxisMatrix)
    Call SFB_MultiplyVector(SFB_dAxisMatrix, SFB_dReferenceProductMatrix, SFB_dReferenceProductMatrix, True)
    Call SFB_MultiplyMatrix(SFB_dAxisMatrix, SFB_dReferenceProductMatrix, SFB_dReferenceProductMatrix)
    
    'SFB_Retrieve Global Matrix of target NHA
    sParentList = SFB_BuildParentList(oProductToPositionNHA)
    Call SFB_CalculateInstanceGlobalPositionMatrix(sParentList, SFB_dTargetNHAMatrix)
    
    'SFB_Retrieve target position of the part
    Call SFB_FindChildMatrix(SFB_dReferenceProductMatrix, SFB_dTargetNHAMatrix, SFB_dTargetMatrix)
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents SFB_dTargetMatrix
    
End Sub

Public Sub SFB_SnapAxisToOrigin(ByVal oProductToPosition As Product, ByVal oProductToPositionNHA As Product, ByVal oReferenceAxis As AxisSystem)
    
    'Purpose: Move product instance (oProductToPosition) at inside its NHA  (oProductToPositionNHA)
    'so the reference axis is at (0,0,0)
    
    Dim SFB_dAxisMatrix(11) As Variant
    Dim SFB_dNHAMatrix(11) As Variant
    Dim SFB_dReferenceProductMatrix(11) As Variant
    
    'SFB_Retrieve Global Matrix of Collector's NHA
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    sParentList = SFB_BuildParentList(oProductToPositionNHA)
    Call SFB_CalculateInstanceGlobalPositionMatrix(sParentList, SFB_dNHAMatrix)
    
    'SFB_Retrieve Axis position matrix in the part
    Call SFB_GetAxisPosition(oReferenceAxis, SFB_dAxisMatrix)
    
    'Find SFB_dReferenceProductMatrix to have Axis position matrix at 0,0,0
    Call SFB_FindPartAxisMatrix(SFB_dAxisMatrix, SFB_dNHAMatrix, SFB_dReferenceProductMatrix)
    
    Set oProdPosition = oProductToPosition.Position
    oProdPosition.SetComponents SFB_dReferenceProductMatrix
    
End Sub

Public Sub SFB_FindPartAxisMatrix(ByVal SFB_dAxisMatrix As Variant, ByVal SFB_dNHAMatrix As Variant, ByRef dTargetPartMatrix As Variant)
    
    'Purpose: finding the position dTargetPartMatrix (relative to NHA) that correspond to SFB_a SFB_dAxisMatrix at (0,0,0)(Global position)
    
    Dim SFB_dTempMatrix As Variant
    SFB_dTempMatrix = SFB_dAxisMatrix
    
    'Invert Axis Matrix
    Call SFB_InvertMatrix(SFB_dAxisMatrix)
    
    'Set SFB_dTempMatrix Vector
    SFB_dTempMatrix(9) = -SFB_dAxisMatrix(9)
    SFB_dTempMatrix(10) = -SFB_dAxisMatrix(10)
    SFB_dTempMatrix(11) = -SFB_dAxisMatrix(11)
    
    'Caculate Vector Position > Return result in SFB_dAxisMatrix
    Call SFB_MultiplyVector(SFB_dTempMatrix, SFB_dAxisMatrix, SFB_dAxisMatrix)
    
    'Find part matrix (under NHA) > Return result in d3DCollMatrix
    Call SFB_FindChildMatrix(SFB_dAxisMatrix, SFB_dNHAMatrix, dTargetPartMatrix)
    
End Sub

Public Sub SFB_FindChildMatrix(ByRef dRefMatrix As Variant, ByRef SFB_dTargetNHAMatrix As Variant, ByRef dChildMatrix)
    
    'Purpose: Find the dChildMatrix (relative to NHA) to have same global position as dRefMatrix
    
    'Invert dRefMatrix
    Call SFB_InvertMatrix(SFB_dTargetNHAMatrix)
    
    'Remove Vector Position SFB_dTargetNHAMatrix from dRefMatrix, Return result in dRefMatrix Matrix, put SFB_dTargetNHAMatrix position to 0
    Call SFB_SubstractVector(dRefMatrix, SFB_dTargetNHAMatrix)
    
    'Caculate Vector Position > Return result in dChildMatrix
    Call SFB_MultiplyVector(dRefMatrix, SFB_dTargetNHAMatrix, dChildMatrix)
    
    'Caculate Rotation Matrix (dRefMatrix X SFB_dTargetMatrix) > Return result in dChildMatrix
    Call SFB_MultiplyMatrix(dRefMatrix, SFB_dTargetNHAMatrix, dChildMatrix)
    
End Sub

Public Function SFB_BuildParentList(ByVal SFB_iObject As Variant) As String
    
    'The parent list is SFB_a list of all the instances starting at the selection all the way up to the top product in the active window
    SFB_BuildParentList = SFB_iObject.Name & ";"
    
    Do
        Set SFB_iObject = SFB_iObject.Parent
        
        'We don't care about the "Products" or "ProductDocuments" objects
        If TypeName(SFB_iObject) = "Product" Then
            SFB_BuildParentList = SFB_BuildParentList & SFB_iObject.Name & ";"
        End If
    Loop Until TypeName(SFB_iObject) = "Application"

End Function

Public Sub SFB_CalculateInstanceGlobalPositionMatrix(ByVal sParentList As String, ByRef dGlobalMatrix As Variant)

    Dim SFB_oParent As Product
    Dim SFB_iParentList As Long
    Dim SFB_dLocalMatrix(11) As Variant
    
    'Initialize first parent and dGlobalMatrix (should be I Matrix)
    Set SFB_oParent = CATIA.ActiveDocument.Product
    Set oProdPosition = SFB_oParent.Position
    oProdPosition.GetComponents dGlobalMatrix
    
    'Loop sParentList
    For SFB_iParentList = UBound(Split(sParentList, ";")) - 2 To 0 Step -1
    
        'SFB_Retrieve SFB_oChild position matrix
        Set SFB_oParent = SFB_oParent.Products.Item(Split(sParentList, ";")(SFB_iParentList))
        Set oProdPosition = SFB_oParent.Position
        oProdPosition.GetComponents SFB_dLocalMatrix
        
        'Caculate Vector Position > Return result in Global Matrix
        Call SFB_MultiplyVector(SFB_dLocalMatrix, dGlobalMatrix, dGlobalMatrix, True)
        
        'Caculate Rotation Matrix (SFB_dLocalMatrix X dGlobalMatrix) > Return result in Global Matrix
        Call SFB_MultiplyMatrix(SFB_dLocalMatrix, dGlobalMatrix, dGlobalMatrix)
        
    Next SFB_iParentList
    
End Sub

Public Sub SFB_GetAxisPosition(ByVal oAxis As Object, ByRef SFB_dAxisMatrix As Variant)
    
    Dim SFB_dVectorPosition(2) As Variant
    
    'SFB_Retrieve axis position matrix
    oAxis.GetXAxis SFB_dVectorPosition
    SFB_dAxisMatrix(0) = SFB_dVectorPosition(0)
    SFB_dAxisMatrix(1) = SFB_dVectorPosition(1)
    SFB_dAxisMatrix(2) = SFB_dVectorPosition(2)
    
    oAxis.GetYAxis SFB_dVectorPosition
    SFB_dAxisMatrix(3) = SFB_dVectorPosition(0)
    SFB_dAxisMatrix(4) = SFB_dVectorPosition(1)
    SFB_dAxisMatrix(5) = SFB_dVectorPosition(2)
    
    oAxis.GetZAxis SFB_dVectorPosition
    SFB_dAxisMatrix(6) = SFB_dVectorPosition(0)
    SFB_dAxisMatrix(7) = SFB_dVectorPosition(1)
    SFB_dAxisMatrix(8) = SFB_dVectorPosition(2)
    
    oAxis.GetOrigin SFB_dVectorPosition
    SFB_dAxisMatrix(9) = SFB_dVectorPosition(0)
    SFB_dAxisMatrix(10) = SFB_dVectorPosition(1)
    SFB_dAxisMatrix(11) = SFB_dVectorPosition(2)
    
End Sub

Public Sub SFB_InvertMatrix(ByRef dMatrixToInvert As Variant)
    
    Dim SFB_dTempMatrix As Variant
    SFB_dTempMatrix = dMatrixToInvert
    
'Transpose Global Matrix
    dMatrixToInvert(0) = SFB_dTempMatrix(0)
    dMatrixToInvert(1) = SFB_dTempMatrix(3)
    dMatrixToInvert(2) = SFB_dTempMatrix(6)
    dMatrixToInvert(3) = SFB_dTempMatrix(1)
    dMatrixToInvert(4) = SFB_dTempMatrix(4)
    dMatrixToInvert(5) = SFB_dTempMatrix(7)
    dMatrixToInvert(6) = SFB_dTempMatrix(2)
    dMatrixToInvert(7) = SFB_dTempMatrix(5)
    dMatrixToInvert(8) = SFB_dTempMatrix(8)
    
End Sub

Public Sub SFB_MultiplyMatrix(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, ByRef dResultMatrix As Variant)
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

Public Sub SFB_MultiplyVector(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, ByRef dResultMatrix As Variant, Optional ByVal bAddInitialValue As Boolean = False)
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

Public Sub SFB_SubstractVector(ByRef dMatrix1 As Variant, ByRef dMatrix2 As Variant)
'Remove 1x3 vector2 from 1x3 vector1

    dMatrix1(9) = dMatrix1(9) - dMatrix2(9)
    dMatrix1(10) = dMatrix1(10) - dMatrix2(10)
    dMatrix1(11) = dMatrix1(11) - dMatrix2(11)
    
End Sub

Public Function SFB_GetInstanceActivity(ByVal oInstance As Product, Optional ByVal bSetDefaultMode As Boolean = False) As Boolean
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
    Dim SFB_oProdparams As Parameters
    
    On Error Resume Next
    Set SFB_oProdparams = oInstance.Parameters.SubList(oInstance, False)
   
    If SFB_oProdparams.Count = 0 And bSetDefaultMode = True Then
        oInstance.ApplyWorkMode DEFAULT_MODE
        Set SFB_oProdparams = oInstance.Parameters.SubList(oInstance, False)
    End If
    
    If SFB_oProdparams.Count > 0 Then SFB_GetInstanceActivity = SFB_oProdparams.GetItem("Component Activation State").Value
    
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
Public Function SFB_DotProd(v1 As Variant, v2 As Variant) As Double
'*** dot product of two vectors of 3 elements
    SFB_DotProd = v1(0) * v2(0) + v1(1) * v2(1) + v1(2) * v2(2)
End Function
Public Sub SFB_CrossProduct(v1() As Double, v2() As Double, ByRef dCrossProduct As Variant)
    Dim SFB_V(2) As Double
    
    SFB_V(0) = v1(1) * v2(2) - v1(2) * v2(1)
    SFB_V(1) = v1(2) * v2(0) - v1(0) * v2(2)
    SFB_V(2) = v1(0) * v2(1) - v1(1) * v2(0)

    dCrossProduct = SFB_V
End Sub
Public Function SFB_LengthVector(v1 As Variant) As Double
'*** Magnitude of vector of 3 elements
    SFB_LengthVector = Sqr(v1(0) ^ 2 + v1(1) ^ 2 + v1(2) ^ 2)
End Function
Public Function SFB_ArcCos(oValue As Double) As Double
'*** Inverse cosine of value
    If Round(oValue, 8) = 1 Then SFB_ArcCos = 0: Exit Function
    If Round(oValue, 8) = -1 Then SFB_ArcCos = SFB_PI: Exit Function
    SFB_ArcCos = Atn(-oValue / Sqr(1 - oValue ^ 2)) + 2 * Atn(1)
End Function

Public Function SFB_AngleInRad(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, Optional ByRef sVector As String = "X") As Double
'*** Find angle between two vectors in 3 dimensions.Input is position matrix of two instance
'*** By assigning optional value to SVector , one can find angle between either X ,Y or Z axis of two Position matrix.
    Dim SFB_Vctr1(2)
    Dim SFB_Vctr2(2)
    
    SFB_Vctr1(0) = dMatrix1(0)
    SFB_Vctr1(1) = dMatrix1(1)
    SFB_Vctr1(2) = dMatrix1(2)
    
    SFB_Vctr2(0) = dMatrix2(0)
    SFB_Vctr2(1) = dMatrix2(1)
    SFB_Vctr2(2) = dMatrix2(2)
    
    If sVector = "Y" Then
        SFB_Vctr1(0) = dMatrix1(3)
        SFB_Vctr1(1) = dMatrix1(4)
        SFB_Vctr1(2) = dMatrix1(5)
        
        SFB_Vctr2(0) = dMatrix2(3)
        SFB_Vctr2(1) = dMatrix2(4)
        SFB_Vctr2(2) = dMatrix2(5)
    ElseIf sVector = "Z" Then
        SFB_Vctr1(0) = dMatrix1(6)
        SFB_Vctr1(1) = dMatrix1(7)
        SFB_Vctr1(2) = dMatrix1(8)
        
        SFB_Vctr2(0) = dMatrix2(6)
        SFB_Vctr2(1) = dMatrix2(7)
        SFB_Vctr2(2) = dMatrix2(8)
    End If
    
    SFB_AngleInRad = SFB_ArcCos(SFB_DotProd(SFB_Vctr1, SFB_Vctr2) / (SFB_LengthVector(SFB_Vctr1) * SFB_LengthVector(SFB_Vctr2)))
    
End Function
Public Function SFB_dDistanceLineSegToLineSeg(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double

'********--------------------
'***
'*** returns shortest distance between two line segments Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Adapted from http://geomalgorithms.com/a07-_distance.html#dist3D_Segment_to_Segment()
'***
'********--------------------
    Dim SFB_U(2) As Double
    Dim SFB_V(2) As Double
    Dim SFB_W(2) As Double
    Dim SFB_Result(2) As Double
    Dim SFB_PR1(2) As Double
    Dim SFB_PR2(2) As Double
    Dim SFB_a As Double, b As Double, C As Double, SFB_d As Double, DD As Double, sc As Double, SFB_sn As Double, SFB_sd As Double
    Dim SFB_tc As Double, tN As Double, tD As Double
    
    SFB_U(0) = p2(0) - p1(0): SFB_U(1) = p2(1) - p1(1): SFB_U(2) = p2(2) - p1(2)
    SFB_V(0) = p4(0) - p3(0): SFB_V(1) = p4(1) - p3(1): SFB_V(2) = p4(2) - p3(2)
    SFB_W(0) = p1(0) - p3(0): SFB_W(1) = p1(1) - p3(1): SFB_W(2) = p1(2) - p3(2)

    SFB_a = SFB_DotProd(SFB_U, SFB_U)   ' should be > = zero
    b = SFB_DotProd(SFB_U, SFB_V)
    C = SFB_DotProd(SFB_V, SFB_V)   ' should be > = zero
    SFB_d = SFB_DotProd(SFB_U, SFB_W)
    e = SFB_DotProd(SFB_V, SFB_W)
    
    DD = SFB_a * C - b * b
    SFB_sd = DD ' default
    tD = DD ' default
    
    If DD < SFB_SMALLNUMBER Then     ' lines are parallel
            SFB_sn = 0               ' forcing use of point P1 of Line 1
            SFB_sd = 1               ' to prevent possible division by 0.0 later
            tN = e
            tD = C
    Else                         ' if lines are not parallel then get closest point on infinite lines
        SFB_sn = (b * e - C * SFB_d)
        tN = (SFB_a * e - b * SFB_d)
            If SFB_sn < 0 Then       ' sc < 0 => the s=0 edge is visible
                SFB_sn = 0
                tN = e
                tD = C
            ElseIf SFB_sn > SFB_sd Then  ' sc > 1  => the s=1 edge is visible
                SFB_sn = SFB_sd
                tN = e + b
                tD = C
            End If
    End If
    If (tN < 0) Then             'SFB_tc < 0 => the t=0 edge is visible
        tN = 0
        ' recompute sc for this edge
            If (0 - SFB_d) < 0 Then
                SFB_sn = 0
            ElseIf (0 - SFB_d) > SFB_a Then
                SFB_sn = SFB_sd
            Else
                SFB_sn = 0 - SFB_d
                SFB_sd = SFB_a
            End If
    ElseIf tN > tD Then         ' SFB_tc > 1  => the t=1 edge is visible
        tN = tD
        'recompute sc for this edge
            If (0 - SFB_d + b) < 0 Then
                SFB_sn = 0
            ElseIf (0 - SFB_d + b) > SFB_a Then
                SFB_sn = SFB_sd
            Else
                SFB_sn = (0 - SFB_d + b)
                SFB_sd = SFB_a
            End If
    End If
        
    If Abs(SFB_sn) < SFB_SMALLNUMBER Then
        sc = 0
    Else
        sc = SFB_sn / SFB_sd
    End If
    
    If Abs(tN) < SFB_SMALLNUMBER Then
        SFB_tc = 0
    Else
        SFB_tc = tN / tD
    End If
    SFB_Result(0) = SFB_W(0) + (sc * SFB_U(0)) - (SFB_tc * SFB_V(0))
    SFB_Result(1) = SFB_W(1) + (sc * SFB_U(1)) - (SFB_tc * SFB_V(1))
    SFB_Result(2) = SFB_W(2) + (sc * SFB_U(2)) - (SFB_tc * SFB_V(2))
    SFB_dDistanceLineSegToLineSeg = SFB_LengthVector(SFB_Result) ' Length of Vector
    SFB_PR1(0) = p1(0) + sc * SFB_U(0)
    SFB_PR1(1) = p1(1) + sc * SFB_U(1)
    SFB_PR1(2) = p1(2) + sc * SFB_U(2)
    
    SFB_PR2(0) = p3(0) + SFB_tc * SFB_V(0)
    SFB_PR2(1) = p3(1) + SFB_tc * SFB_V(1)
    SFB_PR2(2) = p3(2) + SFB_tc * SFB_V(2)
End Function
Public Function SFB_dDistancePointToLineSeg(p1() As Double, p2() As Double, p3() As Double, Optional LineIsInfinite As Boolean = False) As Double
'*** calculate shortest distance between SFB_a point & SFB_a line segment
    Dim SFB_U(2) As Double
    Dim SFB_V(2) As Double
    Dim SFB_W(2) As Double
    Dim SFB_a As Double, b As Double, C As Double
    Dim SFB_ClosestPoint(2) As Double
    Dim SFB_Result(2) As Double
    
    SFB_U(0) = p1(0) - p3(0): SFB_U(1) = p1(1) - p3(1): SFB_U(2) = p1(2) - p3(2)     ' Vector of End point of line to the Point
    SFB_V(0) = p3(0) - p2(0): SFB_V(1) = p3(1) - p2(1): SFB_V(2) = p3(2) - p2(2)     ' Vector of Line
    SFB_W(0) = p1(0) - p2(0): SFB_W(1) = p1(1) - p2(1): SFB_W(2) = p1(2) - p2(2)     ' Vector of Startpoint of Line to the point
        
    SFB_a = SFB_DotProd(SFB_V, SFB_W)
    b = SFB_DotProd(SFB_V, SFB_V)
    
    If Not LineIsInfinite Then  ' conditions for line of finite length
        If SFB_a < 0 Then
            SFB_dDistancePointToLineSeg = SFB_LengthVector(SFB_W)
            Exit Function
        End If
        If b < SFB_a Then
            SFB_dDistancePointToLineSeg = SFB_LengthVector(SFB_U)
            Exit Function
        End If
    End If
    ' if line is infinte
    C = SFB_a / b
    SFB_ClosestPoint(0) = p2(0) + C * SFB_V(0)
    SFB_ClosestPoint(1) = p2(1) + C * SFB_V(1)
    SFB_ClosestPoint(2) = p2(2) + C * SFB_V(2)
    
    SFB_Result(0) = p1(0) - SFB_ClosestPoint(0)
    SFB_Result(1) = p1(1) - SFB_ClosestPoint(1)
    SFB_Result(2) = p1(2) - SFB_ClosestPoint(2)
    
    SFB_dDistancePointToLineSeg = Round(SFB_LengthVector(SFB_Result), 12)
End Function
Public Function SFB_dDistanceLineToLine(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double
'*** calculates shortest distance between lines of infinite length
    Dim SFB_U(2) As Double
    Dim SFB_V(2) As Double
    Dim SFB_W(2) As Double
    Dim SFB_PR1(2) As Double
    Dim SFB_PR2(2) As Double
    
    Dim SFB_Result(2) As Double
    Dim SFB_a As Double, b As Double, C As Double, SFB_d As Double, DD As Double, sc As Double, SFB_tc As Double
    
    SFB_U(0) = p2(0) - p1(0): SFB_U(1) = p2(1) - p1(1): SFB_U(2) = p2(2) - p1(2)    ' vector of line 1
    SFB_V(0) = p4(0) - p3(0): SFB_V(1) = p4(1) - p3(1): SFB_V(2) = p4(2) - p3(2)    ' vector of line 2
    SFB_W(0) = p1(0) - p3(0): SFB_W(1) = p1(1) - p3(1): SFB_W(2) = p1(2) - p3(2)

    SFB_a = SFB_DotProd(SFB_U, SFB_U)   ' should be > = zero
    b = SFB_DotProd(SFB_U, SFB_V)
    C = SFB_DotProd(SFB_V, SFB_V)   ' should be > = zero
    SFB_d = SFB_DotProd(SFB_U, SFB_W)
    e = SFB_DotProd(SFB_V, SFB_W)
    
    DD = SFB_a * C - b * b
     If DD < SFB_SMALLNUMBER Then     ' lines are parallel
        sc = 0
        If b > C Then
            SFB_tc = SFB_d / b
        Else
            SFB_tc = e / C
        End If
    Else
        sc = ((b * e) - (C * SFB_d)) / DD
        SFB_tc = ((SFB_a * e) - (b * SFB_d)) / DD
    End If

    SFB_Result(0) = SFB_W(0) + (sc * SFB_U(0)) - (SFB_tc * SFB_V(0))
    SFB_Result(1) = SFB_W(1) + (sc * SFB_U(1)) - (SFB_tc * SFB_V(1))
    SFB_Result(2) = SFB_W(2) + (sc * SFB_U(2)) - (SFB_tc * SFB_V(2))

    SFB_dDistanceLineToLine = Round(SFB_LengthVector(SFB_Result), 12)
    
    SFB_PR1(0) = p1(0) + sc * SFB_U(0)
    SFB_PR1(1) = p1(1) + sc * SFB_U(1)
    SFB_PR1(2) = p1(2) + sc * SFB_U(2)
    
    SFB_PR2(0) = p3(0) + SFB_tc * SFB_V(0)
    SFB_PR2(1) = p3(1) + SFB_tc * SFB_V(1)
    SFB_PR2(2) = p3(2) + SFB_tc * SFB_V(2)
    
End Function
Public Function SFB_dDistancePointToPoint(p1() As Double, p2() As Double) As Double
    Dim SFB_ResultVector(2) As Double
    
    SFB_ResultVector(0) = p2(0) - p1(0): SFB_ResultVector(1) = p2(1) - p1(1): SFB_ResultVector(2) = p2(2) - p1(2)
    SFB_dDistancePointToPoint = Round(SFB_LengthVector(SFB_ResultVector), 12)
End Function
Public Function SFB_dDistancePointToPlane(p1() As Double, p2() As Double, p3() As Double) As Double
    '*** p1 is the point from which we want to measure the distance to plane
    '*** Plane vector is created from p2 & p3 , p2 is also SFB_a point lying on SFB_a plane
    '*** distance is +ve when point is on the side of vector of plane
    '*** and -ve when it is on the other side
    Dim SFB_PlaneVector(2) As Double
    Dim SFB_Result(2) As Double
    Dim SFB_sb As Double
    Dim SFB_sn As Double
    Dim SFB_sd As Double
    Dim SFB_U(2) As Double
    
    SFB_PlaneVector(0) = p3(0) - p2(0): SFB_PlaneVector(1) = p3(1) - p2(1): SFB_PlaneVector(2) = p3(2) - p2(2)
    SFB_U(0) = p1(0) - p2(0): SFB_U(1) = p1(1) - p2(1): SFB_U(2) = p1(2) - p2(2)
    
    SFB_sn = SFB_DotProd(SFB_PlaneVector, SFB_U)
    SFB_sd = SFB_LengthVector(SFB_PlaneVector)
    SFB_dDistancePointToPlane = SFB_sn / SFB_sd
End Function
'Abhishek Update with InsertHole
Public Function SFB_iFindLineSegmentIntersectionToPlane(p1() As Double, p2() As Double, p3() As Double, p4() As Double, Optional ByRef dIntsctPT As Variant) As Integer
'***-------------
'*** p1 & p2 belongs to start & end of line segment
'*** Plane vector is created from p3 & p4 , p3 being the point lying on SFB_a plane
'*** SFB_Result = 0 No intersection
'*** SFB_Result = 1 Intersection exist
'*** SFB_Result =2 segement lies on plane
'*** if Intersection exist then dIntsctPT will have coordinates of intersection
'*** Adapted from --- http://geomalgorithms.com/a05-_intersect-1.html
'***-------------
    Dim SFB_PlaneVector(2) As Double
    Dim SFB_U(2) As Double
    Dim SFB_W(2) As Double
    Dim SFB_d As Double
    Dim SFB_N As Double
    Dim SFB_SI As Double
    Dim SFB_Result As Integer
    
    SFB_PlaneVector(0) = p4(0) - p3(0): SFB_PlaneVector(1) = p4(1) - p3(1): SFB_PlaneVector(2) = p4(2) - p3(2)
    
    SFB_U(0) = p2(0) - p1(0): SFB_U(1) = p2(1) - p1(1): SFB_U(2) = p2(2) - p1(2)
    SFB_W(0) = p1(0) - p3(0): SFB_W(1) = p1(1) - p3(1): SFB_W(2) = p1(2) - p3(2)
    
    SFB_d = SFB_DotProd(SFB_PlaneVector, SFB_U)
    SFB_N = -SFB_DotProd(SFB_PlaneVector, SFB_W)
    
    If Abs(SFB_d) < SFB_SMALLNUMBER Then
        If SFB_N = 0 Then
            SFB_Result = 2 ' segment lies in plane
        Else
            SFB_Result = 0 ' no intersection
        End If
    Else
        SFB_SI = Round(SFB_N / SFB_d, 3)
        If SFB_SI < 0 Or SFB_SI > 1 Then
            SFB_Result = 0  ' no intersection
        Else
            SFB_Result = 1 ' Intersection exist
            'dIntsctPT(0) = p1(0) + SFB_SI * SFB_U(0)
            'dIntsctPT(1) = p1(1) + SFB_SI * SFB_U(1)
            'dIntsctPT(2) = p1(2) + SFB_SI * SFB_U(2)
        End If
    End If
    SFB_iFindLineSegmentIntersectionToPlane = SFB_Result
End Function


