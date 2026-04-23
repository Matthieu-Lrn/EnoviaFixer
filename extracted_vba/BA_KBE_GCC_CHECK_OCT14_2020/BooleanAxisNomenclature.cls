Attribute VB_Name = "BooleanAxisNomenclature"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private oBoolParamCollection As New Collection
Private oAxisCollection As New Collection
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to check the nomenclature of the Boolean Parameters and
'* axis systems of Std Parts.
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
    
    Call SearchCompareBooleanVsAxis
    
End Sub

Private Sub PopulateAxisCollection(oActiveSeedPart As Product)
'This Sub will populate the oAxisCollection
'A Std Part axis is any element that:
'Start with "SP_*" but doesn't end with "_Boolean"
'or
'Start with "SA_*" but doesn't end with "_Boolean"
    
    Dim i As Integer
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("((Name_CAP=SP_* - Name_CAP=*_Boolean) + (Name_CAP=SA_* - Name_CAP=*_Boolean)),sel", oActiveSeedPart, oSelection)
    
    'Populate the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
            oAxisCollection.Add AxisArray(oSelection.Item(i).Value, oActiveSeedPart.PartNumber)
        Next
    End If
    
End Sub

Private Sub PopulateBoolParamCollection(oActiveSeedPart As Product)
'This Sub will populate the oBoolParamCollection with an array of the Boolean Parameter, it's local name and the parent Part's name
    
    Dim i As Integer
    Dim oParam As Parameter
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean),sel", oActiveSeedPart, oSelection)
    
    'Populate the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
            Set oParam = oSelection.Item(i).Value
            oBoolParamCollection.Add BoolParamArray(oParam, oActiveSeedPart.PartNumber)
        Next
    End If
    
End Sub

Private Function RetrieveAxis(ByVal sParamName As String) As Variant
    
    Dim oObject As Variant
    RetrieveAxis = oObject
    
    For Each oObject In oAxisCollection
    
        If sParamName Like oObject(0).Name & "_*" Then
            RetrieveAxis = oObject
            Exit For
        End If
    Next
    
End Function

Private Sub SearchCompareBooleanVsAxis()
    
    Dim oSeedPartInstance As Product
    Dim sLog As String
    Dim iProgressCount1 As Integer
    Dim oAxis As Variant
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    
    'Time set for Progress Bar
    dStartTime = DateTime.Time
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Comparing Boolean Parameters and SP/SA Axis")
    Call frmProgress.progressBarRepaint("Initializing...")
        
    iProgressCount1 = 0
    Set oBoolParamCollection = Nothing
    Set oAxisCollection = Nothing
    
    'We loop the selected Seed Part to Retrieve the axis and boolean parameters and populate the collections
    For Each oSeedPartInstance In oSeedPartCollection
        
        iProgressCount1 = iProgressCount1 + 1
        Call frmProgress.progressBarRepaint("Retrieving Boolean Parameters and Axis", 4, 1, oSeedPartInstance.PartNumber, oSeedPartCollection.Count, iProgressCount1)
        
        'Populate the oBoolParamCollection
        Call PopulateBoolParamCollection(oSeedPartInstance)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        
        'Populate the oAxisCollection
        Call PopulateAxisCollection(oSeedPartInstance)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Check for duplicate Bool Param
    For i = 1 To oBoolParamCollection.Count - 1
        
        Call frmProgress.progressBarRepaint("Check for duplicate Boolean Parameter", 4, 2, "Boolean Parameter (" & i & "/" & oBoolParamCollection.Count & ")", oBoolParamCollection.Count, i)
    
        For j = i + 1 To oBoolParamCollection.Count
                
            If oBoolParamCollection.Item(i)(1) = oBoolParamCollection.Item(j)(1) Then
                sLog = sLog & vbCrLf & "- Duplicate Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2) & " and " & oBoolParamCollection.Item(j)(2)
                Exit For
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    
    'Check for duplicate Axis
    For i = 1 To oAxisCollection.Count - 1
    
        Call frmProgress.progressBarRepaint("Check for duplicate Axis", 4, 3, "Axis (" & i & "/" & oAxisCollection.Count & ")", oAxisCollection.Count, i)
    
        For j = i + 1 To oAxisCollection.Count
            If oAxisCollection.Item(i)(0).Name = oAxisCollection.Item(j)(0).Name Then
                sLog = sLog & vbCrLf & "- Duplicate Axis: " & oAxisCollection.Item(i)(0).Name & " : Axis found in " & oAxisCollection.Item(i)(1) & " and " & oAxisCollection.Item(j)(1)
                Exit For
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Scan the list of all Boolean parameters
    For i = 1 To oBoolParamCollection.Count
    
        Call frmProgress.progressBarRepaint("Find matching Boolean Parameter / Axis", 4, 4, "Boolean Parameter (" & i & "/" & oBoolParamCollection.Count & ")", oBoolParamCollection.Count, i)
    
        'Check the Monument number in the Boolean Parameter local name. Where the Monument number is "M55" in "KBE_M55_SEED_PART_AA-V06"
        If Not Split(oBoolParamCollection.Item(i)(2), "_")(1) = Split(oBoolParamCollection.Item(i)(1), "_")(1) Then
            sLog = sLog & vbCrLf & "- Monument Number is wrong for Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
        End If
    
        'Check the Seed Part ID in the Boolean Parameter local name. Where the Seed Part ID is "_AA" in "KBE_M55_SEED_PART_AA-V06"
        If Not oBoolParamCollection.Item(i)(1) Like "*" & Right(Split(oBoolParamCollection.Item(i)(2), "-")(0), 3) & "*" Then
            sLog = sLog & vbCrLf & "- Seed Part ID is wrong for Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
        End If
    
        'Retrieve the matching axis system
        oAxis = RetrieveAxis(oBoolParamCollection.Item(i)(1))
    
        'Matching axis was not found
        If IsEmpty(oAxis) Then
            sLog = sLog & vbCrLf & "- The axis matching parameter " & oBoolParamCollection.Item(i)(1) & " could not be found" & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
    
        'Matching axis was found
        Else
    
            'Check that the Bool Param and the matching axis are in the same Seed Part
            If Not oBoolParamCollection.Item(i)(2) = oAxis(1) Then
                sLog = sLog & vbCrLf & "- Matching Boolean Parameter and Axis are not in the same Seed Part: " & oAxis(0).Name & " in " & oAxis(1) & " while " & oBoolParamCollection.Item(i)(1) & " in " & oBoolParamCollection.Item(i)(2)
            End If
            'At this point there is no need to check the nomenclature of the axis since the nomenclature of it's matching Boolean Param is ok
        End If
        On Error GoTo 0
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLog = sLog & vbCrLf & vbCrLf & "*************************************************"
        sLog = sLog & vbCrLf & " After correcting the errors run the check again"
        sLog = sLog & vbCrLf & " Correcting one error may trigger a new one"
        sLog = sLog & vbCrLf & "*************************************************"
        
        sLogFileName = sTempDirectory & "LogFile_KBECheck_BooleanVsAxis.txt"
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

Private Function BoolParamArray(oParam As Parameter, sPartName As String) As Variant
'We create and array to keep the boolean parameter, the local name and the the parent part name in the same object
    
    Dim oArray(2) As Variant
    
    Set oArray(0) = oParam
    oArray(1) = Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\")))
    oArray(2) = sPartName
    
    BoolParamArray = oArray
    
End Function

Private Function AxisArray(oAxis As Variant, sPartName As String) As Variant
'We create and array to keep the axis and the the parent part name in the same object
    
    Dim oArray(1) As Variant
    
    Set oArray(0) = oAxis
    oArray(1) = sPartName
    
    AxisArray = oArray
    
End Function

