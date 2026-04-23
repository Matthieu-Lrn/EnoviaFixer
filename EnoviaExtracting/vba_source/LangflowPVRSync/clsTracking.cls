Attribute VB_Name = "clsTracking"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mNbLine As Integer
Private mText As String
Private Const MaxBuffer = 5
Private sPath As String

Public Sub AddHeader(ByVal sHeader As String)

mText = mText & vbCrLf & "######################################################################################################"
mText = mText & vbCrLf & "                                         " & sHeader
mText = mText & vbCrLf & "######################################################################################################"
mNbLine = mNbLine + 3
Call Save(False)

End Sub
Public Sub AddErrorText(ByVal sText As String)

mText = mText & vbCrLf & sText
mNbLine = mNbLine + 1

Call Save(False)

End Sub
Public Sub AddInformationLine(ByVal sFctName As String, ByVal sType As String, ByVal sMsg1 As String, Optional sMsg2 As String = "")

Dim sLine As String

'Date
sLine = Format(Now(), "yyyy/MM/dd hh:mm:ss")
sLine = AppendEmptySpace(sLine, 27)

'Type
sLine = sLine & sType
sLine = AppendEmptySpace(sLine, 41)

'Function Name
sLine = sLine & sFctName
sLine = AppendEmptySpace(sLine, 81)

'Msg1
sLine = sLine & sMsg1
sLine = AppendEmptySpace(sLine, 121)

'Msg2
sLine = sLine & sMsg2
sLine = AppendEmptySpace(sLine, 151)

'Append to whole text
mText = mText & vbCrLf & Trim(sLine)
mNbLine = mNbLine + 1

Call Save(False)


'CodeError
If sType = "CODEERROR" Then
    Call Save(True)
    Call MsgBox("Error was found. Process will be terminated. For more information refer to tracking file " & Environ("temp") & "\PVRSyncTracking.txt", vbCritical)
    End
End If

End Sub

Public Sub Save(ByVal bSaveNow As Boolean)

Dim iFreeFile
    If mNbLine > MaxBuffer Or bSaveNow = True Then

        iFreeFile = FreeFile
        Open sPath For Binary As #iFreeFile
        Put #iFreeFile, LOF(iFreeFile) + 1, mText
        Close iFreeFile
        mText = ""
        mNbLine = 0
    End If

End Sub


Public Sub CreateEmptyTrackFile()

    sPath = Environ("temp") & "\PVRSyncTracking.txt"

    Dim iFreeFile
    
    iFreeFile = FreeFile
    Open sPath For Output As #iFreeFile
    Print #iFreeFile, ""
    Close iFreeFile
End Sub

Private Function AppendEmptySpace(ByVal sText As String, ByVal iLength As Integer) As String

Dim i As Integer
Dim iOriginalLength As Integer

iOriginalLength = Len(sText)

If iOriginalLength < iLength Then
    For i = iOriginalLength To iLength
        sText = sText & " "
    Next
End If

'Return
AppendEmptySpace = sText
End Function


Public Sub WebServiceResultAnalysis(ByVal sWebServiceName As String, Optional ByVal oData, Optional ByVal Arg1 As String = "", Optional ByVal Arg2 As String = "", Optional ByVal bStopOnError = True)
Dim sErrorType As String
Dim sErrorMsg As String

Select Case sWebServiceName

    Case "GetAssyPositionMatricies"
        If UCase(WebServiceAccessTool.WebResponseExceptionStatus) Like "BAD GATEWAY*" Then
            sErrorType = "DataQuality"
        ElseIf WebServiceAccessTool.GeneralException <> "" Or _
            WebServiceAccessTool.WebResponseExceptionStatus <> "" Or _
            WebServiceAccessTool.WebClientReturnStatus <> "" Then
            sErrorType = "Network"
        Else
            sErrorType = "NoError"
        End If

    Case "GetPLMActionDataFromDocument", "GetENOVIADocumentAttributs"
        If UCase(WebServiceAccessTool.WebResponseExceptionStatus) Like "NOT FOUND*" Then
            sErrorType = "DataQuality"
        ElseIf WebServiceAccessTool.GeneralException <> "" Or _
            WebServiceAccessTool.WebResponseExceptionStatus <> "" Or _
            WebServiceAccessTool.WebClientReturnStatus <> "" Then
            sErrorType = "Network"
        Else
            sErrorType = "NoError"
        End If

    Case "GetPLMActionDocuments"
        If WebServiceAccessTool.GeneralException <> "" Or _
            WebServiceAccessTool.WebResponseExceptionStatus <> "" Or _
            WebServiceAccessTool.WebClientReturnStatus <> "" Then
            sErrorType = "Network"
        Else
            sErrorType = "NoError"
        End If


    Case "GetDocumentByBaseNumber"
        If WebServiceAccessTool.GeneralException <> "" Or _
            WebServiceAccessTool.WebResponseExceptionStatus <> "" Or _
            WebServiceAccessTool.WebClientReturnStatus <> "" Then
            sErrorType = "Network"
        ElseIf oData.Count = 0 Then
            sErrorType = "DataQuality"
        Else
            sErrorType = "NoError"
        End If
        
    Case "JSON_GETBDIID", "JSON_BDIPartInfo", "JSON_BDIDocumentInfo", "JSON_BDIPartInfoFromPnRev"
        If WebServiceAccessTool.GeneralException <> "" Or _
            WebServiceAccessTool.WebResponseExceptionStatus <> "" Or _
            WebServiceAccessTool.WebClientReturnStatus <> "" Then
            sErrorType = "Network"
        Else
            sErrorType = "NoError"
        End If
                
        
End Select

'Errors with data
If sErrorType = "DataQuality" Then
    sErrorMsg = "WebService: " & sWebServiceName
    sErrorMsg = IIf(Arg1 <> "", sErrorMsg & ", Arg1: " & Arg1, sErrorMsg)
    sErrorMsg = IIf(Arg2 <> "", sErrorMsg & ", Arg2: " & Arg2, sErrorMsg)
    sErrorMsg = sErrorMsg & " - Data could not be found."
    
    If bStopOnError = True Then
        Call oTracking.AddInformationLine("ScanConfProductStructure", "CODEERROR", sErrorMsg)
    Else
        Call oTracking.AddInformationLine("ScanConfProductStructure", "WARNING", sErrorMsg)
    End If

'Errors with data
ElseIf sErrorType = "Network" Then
    sErrorMsg = "WebService: " & sWebServiceName
    sErrorMsg = IIf(Arg1 <> "", sErrorMsg & ", Arg1: " & Arg1, sErrorMsg)
    sErrorMsg = IIf(Arg2 <> "", sErrorMsg & ", Arg2: " & Arg2, sErrorMsg)
    sErrorMsg = sErrorMsg & " - Request could not be performed. It looks like a network issue."
    If bStopOnError = True Then
        Call oTracking.AddInformationLine("ScanConfProductStructure", "CODEERROR", sErrorMsg)
    Else
        Call oTracking.AddInformationLine("ScanConfProductStructure", "WARNING", sErrorMsg)
    End If
End If

End Sub








