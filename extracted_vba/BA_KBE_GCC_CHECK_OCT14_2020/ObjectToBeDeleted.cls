Attribute VB_Name = "ObjectToBeDeleted"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oSeedPartCollection As New Collection
Private oDeleteManagementCollection As New Collection
Private oBodyCollection As New Collection

'*************************************************************************
'* The goal of this class is to check if objected that should be deleted
'* still exist:
'*  - "DeleteManagement" parameter set    (except the one in a body named BodyShifter*)
'*  - "VBScript_DeleteManagement" reaction  (except the one in a body named BodyShifter*)
'*  - "DeleteManagement" rule  (except the one in a body named BodyShifter*)
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
    
    Call CheckObjectToBeDeleted
    
End Sub

Private Sub CheckObjectToBeDeleted()
        
    Dim oSeedPartInstance As Product
    Dim iProgress1 As Integer
    Dim sLog As String
    Dim oObject As AnyObject
    Dim oPartDocument As PartDocument
    Dim oPart As Part
    Dim oHybridShapeFactory As HybridShapeFactory
    Dim oBody As Body
    Dim i As Integer
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    
    sLog = ""
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Check for object to be deleted")
    
    'We loop the selected Seed Part
    iProgress1 = 0
    For Each oSeedPartInstance In oSeedPartCollection
        
        'Progress Bar
        iProgress1 = iProgress1 + 1
        Call frmProgress.progressBarRepaint("Scanning " & oSeedPartInstance.Name, oSeedPartCollection.Count, iProgress1)
        
        '*******Delete: "DeleteManagement" parameter set and rule, "VBScript_DeleteManagement" reaction
        
        'Initial set
        Set oPartDocument = oSeedPartInstance.GetMasterShapeRepresentation(False)
        Set oPart = oPartDocument.Part
        Set oHybridShapeFactory = oPart.HybridShapeFactory
    
        'Loop all bodies
        For Each oBody In oPart.Bodies
        
            If Not oBody.Name Like "BodyShifter*" Then
                
                'Search for "*DeleteManagement" objects under the body
                oSelection.Clear
                oSelection.Add oBody
                oSelection.Search "Name_CAP=*DeleteManagement,sel"
                
                'Report and delete
                If oSelection.Count > 0 Then
                    For i = 1 To oSelection.Count
                        sLog = sLog & vbCrLf & " - " & TypeName(oSelection.Item(i).Value) & " named " & oSelection.Item(i).Value.Name & " was deleted from " & oBody.Name
                    Next
                    oSelection.Delete
                End If
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following actions were performed:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_ObjectToBeDeleted.txt"
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

Private Sub PopulateDeleteManagementCollection(oActiveSeedPart As Product)
'SUB NOT USED ANYMORE
'This Sub will populate the oDeleteManagementCollection with all the "DeleteManagement"
    
    Dim i As Integer
    Dim oObject As AnyObject
    
    Set oDeleteManagementCollection = Nothing
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("Name_CAP=*DeleteManagement,sel", oActiveSeedPart, oSelection)
    
    'Transfer elements in the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
        
            Set oObject = oSelection.Item(i).Value
            oDeleteManagementCollection.Add oObject
        Next
    End If
    
End Sub
