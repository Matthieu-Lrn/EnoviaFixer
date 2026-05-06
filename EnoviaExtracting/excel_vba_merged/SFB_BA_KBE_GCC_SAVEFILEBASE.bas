Attribute VB_Name = "SFB_BA_KBE_GCC_SAVEFILEBASE"
Public SFB_sCATProductTemplateFile As String
Public SFB_sMetadataTemplateFile As String
Public SFB_sDTExportTemplateFile As String
Public SFB_bCancelAction As Boolean
Public SFB_oPasswordMgr As SFB_clsPasswordMgr
Public SFB_oPathDict
Public SFB_WebServiceAccessTool As Object
Public SFB_sActiveToolbarPath As String
Public SFB_sWebServiceAccessToolName As String

Public Sub SFB_CATMain(ByVal sKBEFilePathInput As String, ByVal sToolbarPathInput As String)
    
'Initialize
SFB_sKBEPathFile = sKBEFilePathInput
SFB_sActiveToolbarPath = sToolbarPathInput

    
Dim SFB_sAnswer As String
Dim SFB_sMessage As String
Dim SFB_dTimer As Double

SFB_bCancelAction = False

'Initialize
Set SFB_oPathDict = CreateObject("Scripting.Dictionary")
Call SFB_GetTagValueFromFile("", "", True)
    
'Transfer the content of SFB_sKBEPathFile in SFB_oPathDict
Call SFB_ReadSettingFile(SFB_oPathDict, SFB_sKBEPathFile)
    
'Get info from path file
SFB_sCATProductTemplateFile = SFB_oPathDict("Template CATProduct")
SFB_sMetadataTemplateFile = SFB_oPathDict("Metadata Template")
SFB_sDTExportTemplateFile = SFB_oPathDict("DTExport Template")
SFB_sWebServiceAccessToolName = SFB_oPathDict("WebServiceAccessToolName")

'Common variables
Call SFB_setCommonVariables

'Kill existing "SFB_WebServiceAccessTool.exe" process
Call SFB_KillProcess("SFB_WebServiceAccessTool.exe")

'Start "SFB_WebServiceAccessTool.exe" process
Call SFB_WebServiceExecute("Start")

'Make sure SFB_WebServiceAccessTool is not nothing before trying to use it
SFB_dTimer = Timer
Do
    If Not SFB_WebServiceAccessTool Is Nothing Then Exit Do
    DoEvents
    SFB_Sleep 250
    Debug.Print "Waiting"
    If Abs(Timer - SFB_dTimer) > 5 Then
        Exit Do
    End If
Loop

If SFB_WebServiceAccessTool Is Nothing Then
    Err.Raise 9999, "", "Cannot instantiate new WebService Object"
End If
SFB_WebServiceAccessTool.UsedbyApplication = "DDPToolBar"

'Start password manager
Set SFB_oPasswordMgr = New SFB_clsPasswordMgr
SFB_oPasswordMgr.ResetPassword

'Ask the user to select the action
AppActivate CATIA.Caption
DoEvents
SFB_sMessage = "Select the action to be performed:"
SFB_sMessage = SFB_sMessage & vbCrLf & " 1) Generate the file base 3D structure"
SFB_sMessage = SFB_sMessage & vbCrLf & " 2) SFB_Retrieve and save the CATDrawings"
SFB_sAnswer = InputBox(SFB_sMessage)

Select Case SFB_sAnswer

    Case "1"
        Call SFB_mdlSave3DStructure.SFB_StartProcess
    
    Case "2"
        Call SFB_mdlLoadSaveDrawings.SFB_StartProcess

    Case Else
        'Do nothing
        
End Select

'End
Call SFB_WebServiceExecute("Stop")

End Sub



Private Sub SFB_ReadSettingFile(ByRef SFB_oDict, ByVal SFB_sFilePath As String)

Dim SFB_i As Integer
Dim SFB_sEntireFile As String, SFB_aEntireFile() As String, sOneLine As String, sTag As String, sValue As String, SFB_sString As String

'Read Path file
SFB_sEntireFile = SFB_ReadTextFile(SFB_sFilePath, 0, True)

'Split with vbNewLine
SFB_aEntireFile() = Split(SFB_sEntireFile, vbNewLine)

'Extract each tag
For SFB_i = 0 To UBound(SFB_aEntireFile)
    
    'Get the line
    sOneLine = Trim(SFB_aEntireFile(SFB_i))
    
    If sOneLine Like "<*>*" Then
    
        'Get the tag
        sTag = Trim(Split(sOneLine, ">")(0))
        sTag = Trim(Mid(sTag, 2))
    
        'Get the value
        sValue = Trim(Split(sOneLine, ">")(1))
        sValue = Replace(sValue, Chr(9), "")
        
        'Add to SFB_oDict
        If Not SFB_oDict.Exists(sTag) Then
            Call SFB_oDict.Add(sTag, sValue)
        End If
        
    End If
Next

End Sub

Private Sub SFB_KillProcess(ByVal SFB_sName As String)

Dim SFB_objWMIService As Object
Dim SFB_oProcessColl As Object
Dim SFB_oProcess As Object

'Get processess
Set SFB_objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

'Get SFB_sName processess
Set SFB_oProcessColl = SFB_objWMIService.ExecQuery("SELECT * FROM Win32_Process where Name ='" & SFB_sName & "'", , 48)

'Kill
For Each SFB_oProcess In SFB_oProcessColl
  SFB_oProcess.Terminate
Next

End Sub



