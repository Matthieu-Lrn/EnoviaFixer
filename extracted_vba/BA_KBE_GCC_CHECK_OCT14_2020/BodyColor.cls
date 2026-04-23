Attribute VB_Name = "BodyColor"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oSeedPartCollection As New Collection
Private oKBEBodyCollection As New Collection

'*************************************************************************
'* The goal of this class is to check the color of the KBE bodies
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
    
    Call CheckBodyColor
    
End Sub

Private Sub PopulateKBEBodyCollection(oActiveSeedPart As Product)
'This Sub will populate the oKBEBodyCollection with all the KBE Bodies
    
    Dim i As Integer
    Dim oBody As Body
    
    Set oKBEBodyCollection = Nothing
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("CATPrtSearch.BodyFeature.Name_CAP=KBE_*,sel", oActiveSeedPart, oSelection)
    
    'Transfer Body in the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
        
            Set oBody = oSelection.Item(i).Value
            oKBEBodyCollection.Add oBody
        Next
    End If
    
End Sub

Private Sub CheckBodyColor()
    
    Dim oSeedPartInstance As Product
    Dim iProgress1 As Integer
    Dim iProgress2 As Integer
    Dim oBody As Body
    Dim oShape As Shape
    Dim sLog As String
    Dim sWarning As String
    Dim sBodyColorName As String
    Dim sShapeColorName As String
    Dim oVisProperties As VisPropertySet
    Dim oPart As Part
    Dim oPartDocument As PartDocument
    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    
    sLog = ""
    
    'Time set for Progress Bar
    dStartTime = DateTime.Time
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Check the color of KBE Bodies")
    Call frmProgress.progressBarRepaint("Processing...")
        
    'We loop the selected Seed Part to Retrieve the KBE Bodies
    iProgress1 = 0
    For Each oSeedPartInstance In oSeedPartCollection
        
        'Progress Bar
        iProgress1 = iProgress1 + 1
        Call frmProgress.progressBarRepaint("Scanning " & oSeedPartInstance.Name, oSeedPartCollection.Count, iProgress1)
        
        'Populate the oKBEBodyCollection
        Call PopulateKBEBodyCollection(oSeedPartInstance)
    
        'Loop all KBE Bodies
        iProgress2 = 0
        For Each oBody In oKBEBodyCollection
            
            'Progress Bar
            iProgress2 = iProgress2 + 1
            Call frmProgress.progressBarRepaint("Scanning " & oSeedPartInstance.Name, oSeedPartCollection.Count, iProgress1, oBody.Name & " (" & iProgress2 & " of " & oKBEBodyCollection.Count & ")", oKBEBodyCollection.Count, iProgress2)
        
            'Retrieve color
            oSelection.Clear
            oSelection.Add oBody
            Set oVisProperties = oSelection.VisProperties
            Call oVisProperties.GetRealColor(iR, iG, iB)
            
            sBodyColorName = FindColorName(iR, iG, iB)
            
            'KBE_xxxx_F should be "Sandy Brown"
            If Right(oBody.Name, 2) = "_F" And sBodyColorName = "Sandy Brown" Then
                sWarning = "OK"
                
            '"*Molding*" should be "Burgundy"
            ElseIf oBody.Name Like "*Molding*" And sBodyColorName = "Burgundy" Then
                sWarning = "OK"
               
            '"*Molding*" should be "Burgundy"
            ElseIf oBody.Name Like "*Grill*" And sBodyColorName = "Grill Grey" Then
                sWarning = "OK"
                
            ElseIf sBodyColorName = "Default Grey" Then
                sWarning = "OK"
            Else
                sWarning = "WARNING"
            End If
            
            'Populate log
            sLog = sLog & vbCrLf & "- " & oBody.Name & " found in " & oSeedPartInstance.PartNumber & " : " & sBodyColorName & " : " & sWarning
    
    
            'Loop all shape under the KBE Body. It seems that sometime a Body can contain empty shapes (not visible in the CATIA tree).
            On Error Resume Next
            Err.Clear
            For Each oShape In oBody.Shapes
                
                If Err.Number <> 0 Then GoTo Nextshape
                oSelection.Clear
                oSelection.Add oShape
                Set oVisProperties = oSelection.VisProperties
                Call oVisProperties.GetRealColor(iR, iG, iB)
                
                sShapeColorName = FindColorName(iR, iG, iB)
                
                'Shape under a body should be the same color as the body
                If sShapeColorName <> sBodyColorName Then
                    sLog = sLog & vbCrLf & "  - " & oShape.Name & " : Color is not the same as the body - Warning"
                End If
                
Nextshape:
                Err.Clear
            Next
            On Error GoTo 0
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
           
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "KBE Body Color Report:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_BodyColor.txt"
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

Private Function FindColorName(ByVal iR As Long, ByVal iG As Long, ByVal iB As Long) As String
    
    If iR = 210 And iG = 210 And iB = 255 Then
        FindColorName = "Default Grey"
    ElseIf iR = 211 And iG = 178 And iB = 125 Then
        FindColorName = "Sandy Brown"
    ElseIf iR = 255 And iG = 255 And iB = 255 Then
        FindColorName = "White"
    ElseIf iR = 128 And iG = 0 And iB = 0 Then
        FindColorName = "Burgundy"
    ElseIf iR = 193 And iG = 196 And iB = 192 Then
        FindColorName = "Grill Grey"
    Else
        FindColorName = "No Name"
    End If
    
End Function

