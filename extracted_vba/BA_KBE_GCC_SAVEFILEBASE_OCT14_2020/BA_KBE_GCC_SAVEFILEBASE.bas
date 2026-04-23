Attribute VB_Name = "BA_KBE_GCC_SAVEFILEBASE"
Public sCATProductTemplateFile As String
Public sMetadataTemplateFile As String
Public sDTExportTemplateFile As String
Public bCancelAction As Boolean
Public oPasswordMgr As clsPasswordMgr
Public oPathDict
Public WebServiceAccessTool As Object
Public sActiveToolbarPath As String
Public sWebServiceAccessToolName As String

Public Sub CATMain(ByVal sKBEFilePathInput As String, ByVal sToolbarPathInput As String)
    
'Initialize
sKBEPathFile = sKBEFilePathInput
sActiveToolbarPath = sToolbarPathInput

    
Dim sAnswer As String
Dim sMessage As String
Dim dTimer As Double

bCancelAction = False

'Initialize
Set oPathDict = CreateObject("Scripting.Dictionary")
Call GetTagValueFromFile("", "", True)
    
'Transfer the content of sKBEPathFile in oPathDict
Call ReadSettingFile(oPathDict, sKBEPathFile)
    
'Get info from path file
sCATProductTemplateFile = oPathDict("Template CATProduct")
sMetadataTemplateFile = oPathDict("Metadata Template")
sDTExportTemplateFile = oPathDict("DTExport Template")
sWebServiceAccessToolName = oPathDict("WebServiceAccessToolName")

'Common variables
Call setCommonVariables

'Kill existing "WebServiceAccessTool.exe" process
Call KillProcess("WebServiceAccessTool.exe")

'Start "WebServiceAccessTool.exe" process
Call WebServiceExecute("Start")

'Make sure WebServiceAccessTool is not nothing before trying to use it
dTimer = Timer
Do
    If Not WebServiceAccessTool Is Nothing Then Exit Do
    DoEvents
    Sleep 250
    Debug.Print "Waiting"
    If Abs(Timer - dTimer) > 5 Then
        Exit Do
    End If
Loop

If WebServiceAccessTool Is Nothing Then
    Err.Raise 9999, "", "Cannot instantiate new WebService Object"
End If
WebServiceAccessTool.UsedbyApplication = "DDPToolBar"

'Start password manager
Set oPasswordMgr = New clsPasswordMgr
oPasswordMgr.ResetPassword

'Ask the user to select the action
AppActivate CATIA.Caption
DoEvents
sMessage = "Select the action to be performed:"
sMessage = sMessage & vbCrLf & " 1) Generate the file base 3D structure"
sMessage = sMessage & vbCrLf & " 2) Retrieve and save the CATDrawings"
sAnswer = InputBox(sMessage)

Select Case sAnswer

    Case "1"
        Call mdlSave3DStructure.StartProcess
    
    Case "2"
        Call mdlLoadSaveDrawings.StartProcess

    Case Else
        'Do nothing
        
End Select

'End
Call WebServiceExecute("Stop")

End Sub



Private Sub ReadSettingFile(ByRef oDict, ByVal sFilePath As String)

Dim i As Integer
Dim sEntireFile As String, aEntireFile() As String, sOneLine As String, sTag As String, sValue As String, sString As String

'Read Path file
sEntireFile = ReadTextFile(sFilePath, 0, True)

'Split with vbNewLine
aEntireFile() = Split(sEntireFile, vbNewLine)

'Extract each tag
For i = 0 To UBound(aEntireFile)
    
    'Get the line
    sOneLine = Trim(aEntireFile(i))
    
    If sOneLine Like "<*>*" Then
    
        'Get the tag
        sTag = Trim(Split(sOneLine, ">")(0))
        sTag = Trim(Mid(sTag, 2))
    
        'Get the value
        sValue = Trim(Split(sOneLine, ">")(1))
        sValue = Replace(sValue, Chr(9), "")
        
        'Add to oDict
        If Not oDict.Exists(sTag) Then
            Call oDict.Add(sTag, sValue)
        End If
        
    End If
Next

End Sub

Private Sub KillProcess(ByVal sName As String)

Dim objWMIService As Object
Dim oProcessColl As Object
Dim oProcess As Object

'Get processess
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

'Get sName processess
Set oProcessColl = objWMIService.ExecQuery("SELECT * FROM Win32_Process where Name ='" & sName & "'", , 48)

'Kill
For Each oProcess In oProcessColl
  oProcess.Terminate
Next

End Sub



