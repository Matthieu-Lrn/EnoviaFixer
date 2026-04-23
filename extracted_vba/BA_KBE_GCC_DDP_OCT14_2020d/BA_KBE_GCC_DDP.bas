Attribute VB_Name = "BA_KBE_GCC_DDP"
'********************************************************************************
'               SCRIPT VERSION
'********************************************************************************
Public Const sScriptVersion As String = "BA_KBE_GCC_DDP_OCT14_2020"
'********************************************************************************

Public sLogFileName As String
Public sLog As String
Public sParentList As String
Public bCancelAction As Boolean '***** Used to determine if user pressed Cancel in a progress bar
Public oSelection As Selection
Public oSelectedItem As AnyObject
Public oPart As Part
Public oProduct As Product
Public oDocument As Document
Public sSourceWindow As Window
Public sTargetWindow As Window
Public oParameter As Parameter
Public iObject As Variant
Public timeStart As Double
Public timeEnd As Double
Public iProgressCount1 As Double
Public iProgressCount1Max As Double
Public iProgressCount2 As Double
Public iProgressCount2Max As Double
Public SSParameters As Collection
Public oProdPosition As Object
Public oExcelGlobal As clsExcel
Public oTemplateData As New BATemplate
Public oCreateCapture As New CreateCapture
Public oRenameInstance As New RenameInstance
Public oInsertDrawingFormat As New InsertDrawingFormat
Public oSearchFindNo As New SearchFindNo
Public oSearchFlagNote As New SearchFlagNote
Public oFillFindNumber As New FillFindNb
Public oSwapActivityInsertHole As New SwapActivityInsertHole
Public oPlExtract As New PLExtract
Public oPlImport As New PLImport
Public oPlTransferToExcel As New PLTransferToExcel
Public oSDDImport As New SDDImport
Public oSDDTransferToExcel As New SDDTransfertToExcel
Public oEvolveParts As New EvolveParts
Public oCheckPanelAssy As New CheckPanelAssy
Public oImportTablefromExcel As New ImportTableFromExcel
Public oIDLExtract As IDLExtract
Public oResetDraftingColors As New ResetDraftingColors
Public oPasswordMgr As clsPasswordMgr
Public oTracking As clsTracking

'*******************************************************************************
'               NEEDED URLS
'********************************************************************************
Public oPathDict
Public sActiveToolbarPath As String         '***** Active Toolbar Path
Public sActiveToolbarModule As String       '***** Active Toolbar Module
Public sKbe_Hardwaretable As String         '***** Table having definition of all hardarwares to be checked
Public sKbe_ReportTemplate As String        '***** Template of report
Public sPanelAssyParametersFile As String   '***** parameters file for panel assy check
Public sCATVBAusedLogPath As String         '***** Path of CATVBA file used log file

Public sCATPartTemplateFile As String
Public sCATProductTemplateFile As String
Public sAttributesTemplateFile As String
Public sInsertHoleDefinition As String
Public sMandatoryCaptures As String
Public sFinishType As String
Public sPLTemplateFile As String
Public sSDDTemplateFile As String
Public sIDLCompareFile As String

Public sSoftMaterialList As String
Public sColorCodeDashNumber As String
Public sColorCodePartExclusionList As String
Public sColorCodeList As String
Public sMacroMode As String
'*******************************************
'           BDI webservices parameters
'*******************************************
Public Const bUseBDIDevtDatabase As Boolean = False
Public sWebServiceAccessToolName As String
Public WebServiceAccessTool As Object
'********************************************************************************
'               CONSTANT VARIABLES
'********************************************************************************
Public Const sNoSelection As String = "No Selection"
Public Const dDeltaPosition As Double = 0.00001

'*******************************************
'           For extracting user information
'*******************************************
Public Enum EXTENDED_NAME_FORMAT
    fNameUnknown = 0
    fNameFullyQualifiedDN = 1
    fNameSamCompatible = 2
    fNameDisplay = 3
    fNameUniqueId = 6
    fNameCanonical = 7
    fNameUserPrincipal = 8
    fNameCanonicalEx = 9
    fNameServicePrincipal = 10
    fNameDnsDomain = 12
End Enum


Public Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" ( _
                                       ByVal nFormat As EXTENDED_NAME_FORMAT, ByVal lpBuffer As String, ByRef nSize As Long) As Long

Public Const c_BNumber As Long = 2
Public Const c_FullName As Long = 3
Public Const c_Canonical As Long = 7
Public Const c_UserEmail As Long = 8
Public Const c_NameDnsDomain = 12

Public Sub CATMain(ByVal sKBEFilePathInput As String, ByVal sMacroModeInput As String, ByVal sToolbarPathInput As String)
    
    'Initialize
    sKBEPathFile = sKBEFilePathInput
    sMacroMode = sMacroModeInput
    sActiveToolbarPath = sToolbarPathInput
    sActiveToolbarModule = "BA_KBE_GCC_DDP"
    
    Set oPathDict = CreateObject("Scripting.Dictionary")
    Call GetTagValueFromFile("", "", True)
    
    'Transfer the content of sKBEPathFile in oPathDict
    Call ReadSettingFile(oPathDict, sKBEPathFile)
    
    'Transfer the content of oPathDict to the public variables. This is to speed up the process.
    'Note that all theses public variable could be replaced in the code by using the value from oPathDict
    sCATPartTemplateFile = oPathDict("Template CATPart")
    sCATProductTemplateFile = oPathDict("Template CATProduct")
    sAttributesTemplateFile = oPathDict("Part Attributes")
    sInsertHoleDefinition = oPathDict("Insert Definition")
    sMandatoryCaptures = oPathDict("MandatoryCaptures")
    sFinishType = oPathDict("Finish Type")
    sPLTemplateFile = oPathDict("PL Template")
    sSDDTemplateFile = oPathDict("WP4 SDD Template")
    sKbe_Hardwaretable = oPathDict("KBE_HardwareTable")
    sKbe_ReportTemplate = oPathDict("KBE_ReportTemplate")
    sPanelAssyParametersFile = oPathDict("CheckPanelAssy_Parameters")
    sSoftMaterialList = oPathDict("Soft Material List")
    sColorCodeDashNumber = oPathDict("Color Code Dash Number")
    sColorCodePartExclusionList = oPathDict("Color Code Part Exclusion List")
    sColorCodeList = oPathDict("Color Code List")
    sIDLCompareFile = oPathDict("IDL_Compare")
    sWebServiceAccessToolName = oPathDict("WebServiceAccessToolName")
    
    'Common variables
    Call setCommonVariables

    'Get soft material list
    sSoftMaterialList = GetTagValueFromFile(sSoftMaterialList, "Soft Material List")

    'Start Webservice access tool
    Call StartWebServiceTool
    
    'Start password manager
    Set oPasswordMgr = New clsPasswordMgr
    oPasswordMgr.ResetPassword

    '***** Update CATSettings - Make sure relations update and are synchronous
    Call enableRelationUpdate
    
    '***** Load Toolbar
    frmKBEMain.Show vbModeless
    
End Sub

Public Function GetNTDomainUser(Optional ByVal NameFormat As Long = c_BNumber, _
                                Optional ByVal Normalize As Variant) As String

' Retrieve the Domain name of the current user.
'
' The format varies according to the value of NameFormat:
'
'     0   NameUnknown            DON'T USE
'     1   NameFullyQualifiedDN   CN=LASTNAME Firstname,OU=users,OU=organisationalunit,DC=domainmajor,DC=domainminor
'     2   NameSamCompatible      DOMAINMAJOR-DOMAINMINOR\firstname.lastname
'     3   NameDisplay            LASTNAME firstname
'     6   NameUniqueId           {04d79494-5826-4443-9e3f-0f7087c85ab3}
'     7   NameCanonical          domainmajor.domainminor/organisationalunit/users/LASTNAME firstname
'     8   NameUserPrincipal      Firstname.LASNAME@domainmajor.domainminor
'     9   NameCanonicalEx        DON'T USE
'    10   NameServicePrincipal   domainmajor.domainminor/organisationalunit/users/LASTNAME Firstname
'    12   NameDnsDomain          DOMAINMAJOR-DOMAINMINOR\firstname.lastname
'
    Dim strUserName As String
    Dim lngUserNameSize As Long

    Select Case NameFormat
        Case 0, 9
            strUserName = "<DON'T USE>"
        Case 1, 2, 3, 6, 7, 8, 10, 12
            strUserName = String$(255, 0)
            lngUserNameSize = Len(strUserName)
            If GetUserNameEx(NameFormat, strUserName, lngUserNameSize) <> 0 Then
                strUserName = Left$(strUserName, lngUserNameSize)
            Else
                strUserName = ""
            End If
        Case Else
            Err.Raise 452
    End Select
    GetNTDomainUser = strUserName
    
End Function

Public Sub AddToLogFile(ByVal sFunctionOrSubName As String, _
                        Optional ByVal sDocNameWithExtn As String = "", _
                        Optional ByVal sDocNumber As String = "", _
                        Optional ByVal sDocRev As String = "", _
                        Optional ByVal iDocIteration As Integer = "0", _
                        Optional ByVal sStatus As String = "", _
                        Optional ByVal sLogType As String = "LOG", _
                        Optional ByVal sDumpString As String = "", _
                        Optional ByVal sOpArg2 As String = "", Optional ByVal sOpArg3 As String = "", Optional ByVal sOpArg4 As String = "", _
                        Optional ByVal sOpArg5 As String = "", Optional ByVal sOpArg6 As String = "", Optional ByVal sOpArg7 As String = "", _
                        Optional ByVal sOpArg8 As String = "", Optional ByVal sOpArg9 As String = "", Optional ByVal sOpArg10 As String = "", _
                        Optional ByVal sLogFilePath = "", _
                        Optional ByVal bGetIteration As Boolean = False)

    Dim sLogUsage As String
    Dim sMacroFileName As String
    
    '****get DocNumber and Rev from sDocNameWithExtn
    '*****
    If sDocNumber = "" And sDocRev = "" And Not (sDocNameWithExtn = "") Then
        If Split(sDocNameWithExtn, ".")(UBound(Split(sDocNameWithExtn, "."))) Like "CATDrawing" Or _
           Split(sDocNameWithExtn, ".")(UBound(Split(sDocNameWithExtn, "."))) Like "CATPart" Or _
           Split(sDocNameWithExtn, ".")(UBound(Split(sDocNameWithExtn, "."))) Like "CATProduct" Then
            sDocNumber = Left(sDocNameWithExtn, InStrRev(sDocNameWithExtn, ".") - 3)
            sDocRev = Mid(Left(sDocNameWithExtn, InStrRev(sDocNameWithExtn, ".") - 1), Len(Left(sDocNameWithExtn, InStrRev(sDocNameWithExtn, ".") - 1)) - 1, 2)
        End If
    End If
    '*****
    If bGetIteration Then
        If iDocIteration = "0" Then
            If sDocNumber <> "" And sDocRev <> "" Then
            'Get attributes of drawing document
            Dim sAttributes
            WebServiceAccessTool.ClearCache
            Set sAttributes = WebServiceAccessTool.GetENOVIADocumentAttributs(sDocNumber, sDocRev, bUseBDIDevtDatabase)
            iDocIteration = sAttributes.GetItem("DOCUMENT_ITERATION")
            End If
        End If
    End If

    On Error Resume Next
    Call WebServiceAccessTool.LogToolUsage(sLogType, _
                                    sStatus, _
                                    sScriptVersion, _
                                    sMacroMode, _
                                    sFunctionOrSubName, _
                                    sDocNumber, _
                                    sDocRev, _
                                    iDocIteration, _
                                    sDumpString, _
                                    sDocNameWithExtn, sOpArg2, sOpArg3, sOpArg4, sOpArg5, sOpArg6, sOpArg7, sOpArg8, sOpArg9, sOpArg10)

    On Error GoTo 0
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


Public Sub StartWebServiceTool()

Dim sString As String
Dim dTimer As Double

On Error GoTo eh
sString = WebServiceAccessTool.webServiceErrorMessage

Exit Sub

eh:
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

End Sub


