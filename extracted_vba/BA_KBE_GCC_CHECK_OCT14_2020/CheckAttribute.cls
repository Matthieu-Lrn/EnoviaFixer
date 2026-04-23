Attribute VB_Name = "CheckAttribute"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'*************************************************************************
'* The goal of this class is to check the value of the following attributs
'*  - Dataset Type
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
    
    If Not oSeedAssy Is Nothing Then Call CheckAttributeValue
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub CheckAttributeValue()
   
Dim oProduct As Product
Dim sPartNumber As String
Dim oParameters As Parameters
Dim oStrParam As StrParam

'Set Selection
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""


'We loop the selected Seed Part check the value of the different attributs
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
            
            On Error Resume Next
            Set oParameters = oProduct.ReferenceProduct.UserRefProperties
        
            'String Parameter
            Set oStrParam = oParameters.GetItem("Dataset Type")
        
            If Err.Number <> 0 Then
                sLog = sLog & vbCrLf & "Dataset Type attribut could not be found for " & oProduct.PartNumber
            ElseIf oStrParam.Value <> "KBE" Then
                oStrParam.Value = "KBE"
                sLog = sLog & vbCrLf & "Dataset Type attribut was modified to KBE for " & oProduct.PartNumber
            End If
            
            On Error GoTo 0
        Else
            sLog = sLog & vbCrLf & "Dataset Type attribut could not be found for " & oProduct.Name
        End If
    End If
    
Next

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "The following actions were performed:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckAttribute.txt"
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

