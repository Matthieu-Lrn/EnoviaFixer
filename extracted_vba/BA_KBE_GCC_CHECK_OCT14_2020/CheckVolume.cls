Attribute VB_Name = "CheckVolume"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oKBEBodyCollection As New Collection
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to measure the volume of each KBE body
'* and report it when volume = 0
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
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy)
    
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
    
    Call MeasureVolume
    
End Sub

Private Sub PopulateKBEBodyCollection(oActiveSeedPart As Product)
'This Sub will populate the oKBEBodyCollection with an array or the KBE Body and the parent Part
    
    Dim i As Integer
    Dim oBody As Body
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("CATPrtSearch.BodyFeature.Name_CAP=KBE_*,sel", oActiveSeedPart, oSelection)
    
    'Populate the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
        
            Set oBody = oSelection.Item(i).Value
            
            'In the collection we store an array of the body and the parent part
            oKBEBodyCollection.Add KBEBody_Part_Array(oBody, oActiveSeedPart.ReferenceProduct.Parent.Part)
            
        Next
    End If
    
End Sub

Private Sub MeasureVolume()
    
    Dim oSeedPartInstance As Product
    Dim oObject As Variant
    Dim oBody As Body
    Dim oPart As Part
    Dim oSPAWorkbench As Workbench
    Dim oMeasurable As Measurable
    Dim oReference As Reference
    Dim dVolume As Double
    Dim iProgressCount As Integer
    Dim sKBEBodySeedPartName() As String
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Checking KBE Body Volume")
        
    Set oKBEBodyCollection = Nothing
    
    'We loop the selected Seed Part to Retrieve the KBE Bodies and populate the collection
    For Each oSeedPartInstance In oSeedPartCollection
                   
        Call frmProgress.progressBarRepaint("Retrieving KBE Bodies", 2, 1)
                    
        'Populate the oKBEBodyCollection
        Call PopulateKBEBodyCollection(oSeedPartInstance)
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'List all KBE Bodies along with their respective parameters (Boolean and ATB)
    iProgressCount = 0
    For Each oObject In oKBEBodyCollection
    
        iProgressCount = iProgressCount + 1
        Call frmProgress.progressBarRepaint("Checking volume", 2, 2, iProgressCount & " of " & oKBEBodyCollection.Count, oKBEBodyCollection.Count, iProgressCount)
    
        'Retrieve the part and the body from the array
        Set oBody = oObject(0)
        Set oPart = oObject(1)
            
        'Retrieve the SPA workbench
        Set oSPAWorkbench = oPart.Parent.GetWorkbench("SPAWorkbench")
        
        'Define a reference with the body
        Set oReference = oPart.CreateReferenceFromObject(oBody)
    
        'Retrieve measurable object of the body
        Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
    
        On Error Resume Next
        dVolume = 1
        dVolume = oMeasurable.Volume
        If Err.Number <> 0 Or dVolume = 0 Then
            sLog = sLog & vbCrLf & "- Volume of " & oBody.Name & " in " & oPart.Name & " is null"
        End If
        On Error GoTo 0
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
            
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckVolume.txt"
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

Private Function KBEBody_Part_Array(oBody As Body, oPart As Part) As Variant
'We create and array to keep both the body and the parent part in the same object
    
    Dim oArray(1) As Variant
    Set oArray(0) = oBody
    Set oArray(1) = oPart
        
    KBEBody_Part_Array = oArray
    
End Function




