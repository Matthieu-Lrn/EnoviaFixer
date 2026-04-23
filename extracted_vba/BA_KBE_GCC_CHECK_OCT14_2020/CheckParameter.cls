Attribute VB_Name = "CheckParameter"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'*************************************************************************
'* The goal of this class is to check the value of the following parameters
'*  - ICD
'*  - CCD
'*  - Maintenance
'*  - AutomationMessagesDisabled
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
    
    If Not oSeedAssy Is Nothing Then Call CheckParameterValue
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub CheckParameterValue()
   
Dim oProduct As Product
Dim sPartNumber As String
Dim oParam As BoolParam

'Set Selection
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""


'We loop the Seed Assy looking for the Seed Option or Seed Config
For Each oProduct In oSeedAssy.Products
    
    'Check instance name
    If oProduct.Name Like "*SEED_OPTION*" Or oProduct.Name Like "*SEED_CONFIG*" Then
    
        'Apply Design Mode
        oProduct.ApplyWorkMode DESIGN_MODE
        
                   
        'ICD
        Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=ICD),sel", oProduct, oSelection)

        If oSelection.Count > 0 Then
            Set oParam = oSelection.Item(1).Value
            If oParam.Value Then
                oParam.Value = False
                sLog = sLog & vbCrLf & "ICD parameter value was change to False"
            End If
        Else
            sLog = sLog & vbCrLf & "ICD parameter can't be found in Seed Option"
        End If
    
        'CCD
        Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=CCD),sel", oProduct, oSelection)

        If oSelection.Count > 0 Then
            Set oParam = oSelection.Item(1).Value
            If oParam.Value Then
                oParam.Value = False
                sLog = sLog & vbCrLf & "CCD parameter value was change to False"
            End If
        Else
            sLog = sLog & vbCrLf & "CCD parameter can't be found in Seed Option"
        End If
            
            
         'Maintenance
        Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=Maintenance),sel", oProduct, oSelection)

        If oSelection.Count > 0 Then
            Set oParam = oSelection.Item(1).Value
            If oParam.Value Then
                oParam.Value = False
                sLog = sLog & vbCrLf & "Maintenance parameter value was change to False"
            End If
        Else
            sLog = sLog & vbCrLf & "Maintenance parameter can't be found in Seed Option"
        End If
           
            
         'AutomationMessagesDisabled
        Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=AutomationMessagesDisabled),sel", oProduct, oSelection)

        If oSelection.Count > 0 Then
            Set oParam = oSelection.Item(1).Value
            If oParam.Value Then
                oParam.Value = False
                sLog = sLog & vbCrLf & "AutomationMessagesDisabled parameter value was change to False"
            End If
        Else
            sLog = sLog & vbCrLf & "AutomationMessagesDisabled parameter can't be found in Seed Option"
        End If
            
   End If
Next

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "The following actions were performed:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckParameter.txt"
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


