Attribute VB_Name = "CheckRules"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'*************************************************************************
'* The goal of this class is to check that the rule of the Seed Assy do
'* not contain the '!' character
'*
'* This problem was found in R21. The Component Activation State can't be linked to the publication of a parameter:

'* OK - if `KBE_M55_SEED_CONFIG-V06\ParametersConfiguration\F01_SampleFloorplanParameter` == "C01.00-Schedule A"
'* KO - if `..!KBE_M55_SEED_CONFIG-V06!F01_SampleFloorplanParameter` == "C01.00-Schedule A"
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
    
    If Not oSeedAssy Is Nothing Then Call CheckRules
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub CheckRules()
        
    Dim oRule As Relation
    Dim sLog As String
    Dim sRuleText As String
    Dim iProgress As Integer
    Dim iNbRelations As Integer
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    iNbRelations = oSeedAssy.Relations.Count
    
    'Show progress bar
    iProgress = 0
    Call frmProgress.progressBarInitialize("Check Seed Assy Rules")
    
    'We loop all the relations of the Seed Assy
    For Each oRule In oSeedAssy.Relations
        
        'Progresse bar
        iProgress = iProgress + 1
        Call frmProgress.progressBarRepaint("Checking Seed Assy Relations (" & iProgress & " of " & iNbRelations & ")", iNbRelations, iProgress)
        
        'Check only the relations in the Seed Assy
        If oRule.Parent.Parent.Name = oSeedAssy.PartNumber Then
        
            'Only check the rules
            If TypeName(oRule) = "Rule" Then
        
                sRuleText = oRule.Value
        
                If InStr(sRuleText, "!") <> 0 Then
                    sLog = sLog & vbCrLf & " - Forbidden character '!' was found in " & oSeedAssy.Name & "/" & oRule.Name
                End If
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
           
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckRules.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    Else
        MsgBox ("No error were found")
    End If
    
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar
    
End Sub






