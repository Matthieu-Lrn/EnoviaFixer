Attribute VB_Name = "CheckParamPath"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to check that parameters in the Root Parameter Set
'* have a local name
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
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy, , , True)
    
    If sSeedPartsList <> "" Then
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else
noSelection:
        Set oSeedAssy = Nothing
        Set sSourceWindow = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set oSeedPartCollection = New Collection
        Exit Sub
    End If
    
    Call ScanParameters
    
End Sub

Private Sub ScanParameters()
   
Dim oProduct As Product
Dim sPartNumber As String
Dim oRootParamSet As ParameterSet
Dim oPart As Part
Dim oParam As Parameter
Dim sOriginalPath As String
Dim sLog As String


'Set Selection
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""


'We loop all the instances in the collection
For Each oProduct In oSeedPartCollection
    
    'Retrieve Part Number
    sPartNumber = ""
    On Error Resume Next
    sPartNumber = oProduct.PartNumber
    On Error GoTo 0

    'If we can Retrieve the Part Number the part should be in design mode
    If Not sPartNumber Like "" Then
        
        Set oPart = oProduct.ReferenceProduct.Parent.Part
        Set oRootParamSet = oPart.Parameters.RootParameterSet
        
        sOriginalPath = sPartNumber & "\"
        Call CheckParameterPath(oRootParamSet, sOriginalPath, sLog)
    Else
        sLog = sLog & vbCrLf & " " & oProduct.Name & " may be in cgr mode"
    End If
    
    
Next

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "The following problems were found:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckParameterPath.txt"
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

Private Sub CheckParameterPath(oParamSet As ParameterSet, ByVal sPath As String, ByRef sLog As String)

Dim oParam As Parameter
Dim oParameterSet As ParameterSet
Dim sNewPath As String

'Direct parameters
For Each oParam In oParamSet.DirectParameters
    If Not oParam.Name Like sPath & "*" Then
        sLog = sLog & vbCrLf & " In " & Split(sPath, "\")(0) & ", " & oParam.Name & " should be modified"
    End If
Next

'Child Parameter Set
For Each oParameterSet In oParamSet.ParameterSets
    sNewPath = sPath & oParameterSet.Name & "\"
    Call CheckParameterPath(oParameterSet, sNewPath, sLog)
Next

End Sub

