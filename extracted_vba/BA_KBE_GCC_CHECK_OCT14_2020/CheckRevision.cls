Attribute VB_Name = "CheckRevision"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long

'*************************************************************************
'* The goal of this class is to hide the check the revision of the
'* different KBE* documents
'*************************************************************************
Public Sub SelectSeedAssy()
    
    On Error Resume Next
    If oSeedAssy Is Nothing Then
        Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    Else
        If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    End If
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If (oSeedAssy Is Nothing Or oSeedAssy.Products.Count = 0) Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical, "No Seed Parts found..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0
    
    If Not oSeedAssy Is Nothing Then Call CheckRevision
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub CheckRevision()
   
Dim oProduct As Product
Dim sPartNumber As String
Dim sRevision() As String
Dim sMessage As String
Dim i As Integer

'Set Selection
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""
sMessage = " ***************** Revision of all loaded document is the same ***************"
i = 0

'We loop the Seed Assy looking for all "KBE*" parts
For Each oProduct In oSeedAssy.Products
    
    'Check instance name
    If oProduct.Name Like "KBE*" Then
    
        'Retrieve Part Number
        sPartNumber = ""
        On Error Resume Next
        sPartNumber = oProduct.PartNumber
        On Error GoTo 0

        'If we can Retrieve the Part Number the part should be in design mode
        If sPartNumber Like "KBE*" Then
        
            i = i + 1
            ReDim Preserve sRevision(i)
            sRevision(i) = Right(Split(oProduct.ReferenceProduct.Parent.Name, ".")(0), 2)
            sLog = sLog & vbCrLf & "Revision for " & oProduct.ReferenceProduct.Parent.Name & " is: " & sRevision(i)
        
        Else
            sLog = sLog & vbCrLf & "Following part may be in cgr mode: " & oProduct.Name
        End If
    End If
    
Next

'Check revision of all loaded document
If i <> 0 Then
    For i = 1 To UBound(sRevision) - 1
        If sRevision(i) <> sRevision(i + 1) Then
            sMessage = " ***************** Revision of all loaded document is the NOT the same ***************"
            Exit For
        End If
    Next
End If
sLog = sLog & vbCrLf & vbCrLf & sMessage

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "Here is the revision of the different documents:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckRevision.txt"
    Call WriteTextFile(sLogFileName, sLog)
    If OpenWithNotepad <> 0 Then
        MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
    End If
Else
    MsgBox ("No actions were performed")
End If

CATIA.DisplayFileAlerts = True
frmKBEMain.resetToolbar

End Sub






