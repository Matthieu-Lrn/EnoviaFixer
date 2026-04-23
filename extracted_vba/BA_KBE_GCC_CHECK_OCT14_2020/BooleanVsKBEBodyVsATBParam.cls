Attribute VB_Name = "BooleanVsKBEBodyVsATBParam"
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
Private oATBParamCollection As New Collection
Private oKBEBodyCollection As New Collection
Private oKBEBodyCollection2 As New Collection
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to check the nomenclature of the Boolean Parameters,
'* ATB Parameters and KBE bodies
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
    
    Call SearchCompareBooleanVsKBEBody
    
End Sub

Private Sub SearchCompareBooleanVsKBEBody()
    
    Dim oSeedPartInstance As Product
    Dim sLog As String
    Dim sList As String
    Dim iProgressCount1 As Integer
    Dim oBody As Variant
    Dim bOK As Boolean
    Dim iNbOfParam As Integer
    Dim iTypeFinishNb As Integer
    Dim bParamFound As Boolean
    Dim sOtherParamName As String
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    
    
    'Time set for Progress Bar
    dStartTime = DateTime.Time
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Comparing Boolean Parameters, KBE Bodies and ATB Parameters")
    Call frmProgress.progressBarRepaint("Initializing...")
        
    Set oBoolParamCollection = Nothing
    Set oKBEBodyCollection = Nothing
    Set oKBEBodyCollection2 = Nothing
    Set oATBParamCollection = Nothing
    
    'We loop the selected Seed Part to Retrieve the KBE Bodies, Boolean Parameters and ATB Parameters and populate the collections
    iProgressCount1 = 0
    For Each oSeedPartInstance In oSeedPartCollection
                   
        iProgressCount1 = iProgressCount1 + 1
        Call frmProgress.progressBarRepaint("Retrieving ATB and Boolean Parameters and KBE Bodies", 7, 1, oSeedPartInstance.PartNumber, oSeedPartCollection.Count, iProgressCount1)
        
        'Populate the oPopulateBoolParamCollection
        Call PopulateBoolParamCollection(oSeedPartInstance)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        
        'Populate the oKBEBodyCollection
        Call PopulateKBEBodyCollection(oSeedPartInstance)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        
        'Populate the oATBParamCollection
        Call PopulateATBParamCollection(oSeedPartInstance)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    
    'List all KBE Bodies along with their respective parameters (Boolean and ATB)
    sList = "This is the list of all KBE Bodies along with their respectives ATB and Boolean parameters" & vbCrLf & vbCrLf
    
    For i = 1 To oKBEBodyCollection.Count
        
        Call frmProgress.progressBarRepaint("Generate parameters list of KBE Bodies", 7, 2, "KBE Body (" & i & "/" & oKBEBodyCollection.Count & ")", oKBEBodyCollection.Count, i)
    
        sList = sList & oKBEBodyCollection(i)(0).Name & " : Body found in " & oKBEBodyCollection(i)(1) & vbCrLf
        
        'Scan oBoolParamCollection to retrieve correponding parameters
        For j = 1 To oBoolParamCollection.Count
            If oBoolParamCollection(j)(1) = oKBEBodyCollection(i)(0).Name & "_Boolean" And oBoolParamCollection(j)(2) = oKBEBodyCollection(i)(1) Then
                sList = sList & " - " & oBoolParamCollection(j)(1) & vbCrLf
            End If
        Next
    
        'Scan ATB parameters list
        For j = 1 To oATBParamCollection.Count
            sParamNameModified = Replace(oATBParamCollection(j)(1), "ATB", "KBE")
            If sParamNameModified Like oKBEBodyCollection(i)(0).Name & "_M*" And oATBParamCollection(j)(2) = oKBEBodyCollection(i)(1) Then
                sList = sList & " - " & oATBParamCollection(j)(1) & vbCrLf
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Check for duplicate Bool Param
    For i = 1 To oBoolParamCollection.Count - 1
            
        Call frmProgress.progressBarRepaint("Check for duplicate Boolean Parameter", 7, 3, "Boolean Parameter (" & i & "/" & oBoolParamCollection.Count & ")", oBoolParamCollection.Count, i)
    
        For j = i + 1 To oBoolParamCollection.Count
            If oBoolParamCollection.Item(i)(1) = oBoolParamCollection.Item(j)(1) Then
                sLog = sLog & vbCrLf & "- Duplicate Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2) & " and " & oBoolParamCollection.Item(j)(2)
                Exit For
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    
    
    'Check for duplicate ATB Param
    For i = 1 To oATBParamCollection.Count
    
        Call frmProgress.progressBarRepaint("Check ATB Parameter", 7, 4, "ATB Parameter (" & i & "/" & oATBParamCollection.Count & ")", oATBParamCollection.Count, i)
    
        'Check for duplicate parameter
        For j = i + 1 To oATBParamCollection.Count
            If oATBParamCollection.Item(i)(1) = oATBParamCollection.Item(j)(1) Then
                sLog = sLog & vbCrLf & "- Duplicate ATB Parameter: " & oATBParamCollection.Item(i)(1) & " : Parameter found in " & oATBParamCollection.Item(i)(2) & " and " & oATBParamCollection.Item(j)(2)
                Exit For
            End If
        Next
    
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Check for duplicate KBE Bodies
    For i = 1 To oKBEBodyCollection.Count - 1
    
        Call frmProgress.progressBarRepaint("Check for duplicate KBE Bodies", 7, 5, "KBE Body (" & i & "/" & oKBEBodyCollection.Count & ")", oKBEBodyCollection.Count, i)
    
        For j = i + 1 To oKBEBodyCollection.Count
            If oKBEBodyCollection.Item(i)(0).Name = oKBEBodyCollection.Item(j)(0).Name Then
                sLog = sLog & vbCrLf & "- Duplicate KBE Body: " & oKBEBodyCollection.Item(i)(0).Name & " : Bodies found in " & oKBEBodyCollection.Item(i)(1) & " and " & oKBEBodyCollection.Item(j)(1)
                Exit For
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Find matching ATB Parameter and KBE Body
    For i = 1 To oATBParamCollection.Count
    
        Call frmProgress.progressBarRepaint("Find matching ATB Parameter / KBE Body", 7, 6, "ATB Parameter (" & i & "/" & oATBParamCollection.Count & ")", oATBParamCollection.Count, i)
    
        'Check the Monument number in the ATB Parameter local name. Where the Monument number is "M55" in "KBE_M55_SEED_PART_AA-V06"
        If Not Split(oATBParamCollection.Item(i)(2), "_")(1) = Split(oATBParamCollection.Item(i)(1), "_")(1) Then
            sLog = sLog & vbCrLf & "- Monument Number is wrong for ATB Parameter: " & oATBParamCollection.Item(i)(1) & " : Parameter found in " & oATBParamCollection.Item(i)(2)
        End If
    
        'Check the Seed Part ID in the ATB Parameter local name. Where the Seed Part ID is "_AA" in "KBE_M55_SEED_PART_AA-V06"
        If Not oATBParamCollection.Item(i)(1) Like "*" & Right(Split(oATBParamCollection.Item(i)(2), "-")(0), 3) & "*" Then
            sLog = sLog & vbCrLf & "- Seed Part ID is wrong for ATB Parameter: " & oATBParamCollection(i)(1) & " : Parameter found in " & oATBParamCollection(i)(2)
        End If
    
        'Retrieve the matching KBE Body
        oBody = RetrieveBodyFromATBParam(oATBParamCollection(i)(1))
    
        'Matching KBE Body was not found
        If IsEmpty(oBody) Then
            sLog = sLog & vbCrLf & "- The KBE Body matching parameter " & oATBParamCollection(i)(1) & " could not be found" & " : Parameter found in " & oATBParamCollection(i)(2)
    
        'Matching KBE Body was found
        Else
            'Check that the ATB Param and the matching KBE Body are in the same Seed Part
            If oATBParamCollection(i)(2) <> oBody(1) Then
                sLog = sLog & vbCrLf & "- Matching ATB Parameter and KBE Body are not in the same Seed Part: " & oATBParamCollection(i)(1) & " / " & oBody(0).Name & " : Parameter found in " & oATBParamCollection(i)(2) & " and body found in " & oBody(1)
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'For each ATB parameter remove the matching KBE body from the collection
    For i = 1 To oATBParamCollection.Count
        Call EmptyBodyCollection(oATBParamCollection(i)(1))
    Next
    
    'If the oKBEBodyCollection2 is not empty it is because some KBE Bodies don't have a matching ATB Parameter
    If oKBEBodyCollection2.Count <> 0 Then
        For Each oBody In oKBEBodyCollection2
            sLog = sLog & vbCrLf & "- The following KBE Body doesn't have a matching ATB parameter: " & oBody(0).Name & " : Body found in " & oBody(1)
        Next
    End If
    
    
    'Find matching Boolean Parameter and KBE Body
    For i = 1 To oBoolParamCollection.Count
    
        Call frmProgress.progressBarRepaint("Find matching Boolean Parameter / KBE Body", 7, 7, "Boolean Parameter (" & i & "/" & oBoolParamCollection.Count & ")", oBoolParamCollection.Count, i)
    
        
        'Check the Monument number in the Boolean Parameter name
        If Not Split(oBoolParamCollection.Item(i)(2), "_")(1) = Split(oBoolParamCollection.Item(i)(1), "_")(1) Then
            sLog = sLog & vbCrLf & "- Monument Number is wrong for Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
        End If
    
        'Check the Seed Part ID in the Boolean Parameter local name. Where the Seed Part ID is "_AA" in "KBE_M55_SEED_PART_AA-V06"
        If Not oBoolParamCollection.Item(i)(1) Like "*" & Right(Split(oBoolParamCollection.Item(i)(2), "-")(0), 3) & "*" Then
            sLog = sLog & vbCrLf & "- Seed Part ID is wrong for Boolean Parameter: " & oBoolParamCollection.Item(i)(1) & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
        End If
    
        'Retrieve the matching KBE Body
        oBody = RetrieveBodyFromBoolParam(oBoolParamCollection.Item(i)(1))
    
        'Matching KBE Body was not found
        If IsEmpty(oBody) Then
            sLog = sLog & vbCrLf & "- The KBE Body matching parameter " & oBoolParamCollection.Item(i)(1) & " could not be found" & " : Parameter found in " & oBoolParamCollection.Item(i)(2)
    
        'Matching KBE Body was found
        Else
    
            'Check that the Boolean Param and the matching KBE Body are in the same Seed Part
            If oBoolParamCollection.Item(i)(2) <> oBody(1) Then
                sLog = sLog & vbCrLf & "- Matching Boolean Parameter and KBE Body are not in the same Seed Part: " & oBoolParamCollection.Item(i)(1) & " / " & oBody(0).Name & " : Parameter found in " & oBoolParamCollection.Item(i)(2) & " and body found in " & oBody(1)
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    
    'If the oKBEBodyCollection is not empty it is because some KBE Bodies don't have a matching Boolean Parameter
    If oKBEBodyCollection.Count <> 0 Then
        For Each oBody In oKBEBodyCollection
            sLog = sLog & vbCrLf & "- The following KBE Body doesn't have a matching Boolean parameter: " & oBody(0).Name & " : Body found in " & oBody(1)
        Next
    End If
            
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_Compare_KBEBody_BoolParam_ATBParam.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If
    
    'Display List file in Notepad
    If sList <> "" Then
        sLogFileName = sTempDirectory & "LogFile_KBECheck_KBEBody_Param_List.txt"
        Call WriteTextFile(sLogFileName, sList)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sTempDirectory & sLogFileName, vbExclamation
        End If
    End If
    
    If sLog = "" And sList = "" Then MsgBox ("No error were found")
    
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar
    
End Sub

Private Sub PopulateBoolParamCollection(oActiveSeedPart As Product)
'This Sub will populate the oBoolParamCollection with an array of the Boolean Parameter, it's local name and the parent Part's name

Dim i As Integer
Dim oParam As Parameter

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
Set oSelection = KBESearch("CATKnowledgeSearch.InternalParameter.Name_CAP=KBE_*_Boolean,sel", oActiveSeedPart, oSelection)

'Transfer Boolean parameter in the collection
If oSelection.Count > 0 Then
    For i = 1 To oSelection.Count
    
        Set oParam = oSelection.Item(i).Value
        oBoolParamCollection.Add BoolParamArray(oParam, oActiveSeedPart.PartNumber)
        
    Next
End If

End Sub

Private Sub PopulateATBParamCollection(oActiveSeedPart As Product)
'This Sub will populate the oATBParamCollection with an array of the ATB Parameter, it's local name and the parent Part's name

Dim i As Integer
Dim oParam As Parameter

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
Set oSelection = KBESearch("CATKnowledgeSearch.InternalParameter.Name_CAP=ATB_*_MType,sel", oActiveSeedPart, oSelection)

'Transfer ATB parameter in the collection
If oSelection.Count > 0 Then
    For i = 1 To oSelection.Count
    
        Set oParam = oSelection.Item(i).Value
        oATBParamCollection.Add ATBParamArray(oParam, oActiveSeedPart.PartNumber)
        
    Next
End If

End Sub

Private Sub PopulateKBEBodyCollection(oActiveSeedPart As Product)
'This Sub will populate the oKBEBodyCollection with an array or the KBE Body and the parent Part name

Dim i As Integer
Dim oBody As Body

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
Set oSelection = KBESearch("CATPrtSearch.BodyFeature.Name_CAP=KBE_*,sel", oActiveSeedPart, oSelection)

'Transfer Body in the collection
If oSelection.Count > 0 Then
    For i = 1 To oSelection.Count
    
        'We populate 2 collections since one of them will be emptied
        Set oBody = oSelection.Item(i).Value
        oKBEBodyCollection.Add KBEBodyArray(oBody, oActiveSeedPart.PartNumber)
        oKBEBodyCollection2.Add KBEBodyArray(oBody, oActiveSeedPart.PartNumber)
        
    Next
End If

End Sub

Private Function RetrieveBodyFromBoolParam(ByVal sParamName As String) As Variant

Dim oObject As Variant
Dim i As Integer

i = 0
For Each oObject In oKBEBodyCollection
    
    i = i + 1
    If sParamName = oObject(0).Name & "_Boolean" Then
        RetrieveBodyFromBoolParam = oObject
        oKBEBodyCollection.Remove (i)
        Exit For
    End If
Next

End Function

Private Function RetrieveBodyFromATBParam(ByVal sParamName As String) As Variant

Dim oObject As Variant
Dim sParamNameModified As String

RetrieveBodyFromATBParam = oObject

'In the param name we replace the "ATB" string with the "KBE" string and search for matching body
sParamNameModified = Replace(sParamName, "ATB", "KBE")

For Each oObject In oKBEBodyCollection2

    If sParamNameModified Like oObject(0).Name & "*" Then
        RetrieveBodyFromATBParam = oObject
        Exit For
    End If
Next

End Function

Private Sub EmptyBodyCollection(ByVal sParamName As String)

Dim oObject As Variant
Dim sParamNameModified As String
Dim i As Integer

'In the param name we replace the "ATB" string with the "KBE" string and search for matching body
sParamNameModified = Replace(sParamName, "ATB", "KBE")

i = 0
For Each oObject In oKBEBodyCollection2

    i = i + 1
    If sParamNameModified Like oObject(0).Name & "*" Then
        oKBEBodyCollection2.Remove (i)
        Exit For
    End If
Next

End Sub

Function RetrieveATBParam(ByVal sATMParamName As String) As Boolean
'The goal of this macro is to scan the list sATBParamLocalName and look for a specific parameter

Dim i As Integer
Dim bParamFound As Boolean

bParamFound = False
For i = 1 To oATBParamCollection.Count
    If oATBParamCollection.Item(i)(1) = sATMParamName Then
        bParamFound = True
        Exit For
    End If
Next

RetrieveATBParam = bParamFound

End Function

Private Function BoolParamArray(oParam As Parameter, sPartName As String) As Variant
'We create and array to keep the boolean parameter, the local name and the the parent part name in the same object

Dim oArray(2) As Variant

Set oArray(0) = oParam
oArray(1) = Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\")))
oArray(2) = sPartName
   
BoolParamArray = oArray

End Function

Private Function ATBParamArray(oParam As Parameter, sPartName As String) As Variant
'We create and array to keep the ATB parameter, the local name and the the parent part name in the same object

Dim oArray(2) As Variant

Set oArray(0) = oParam
oArray(1) = Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\")))
oArray(2) = sPartName
   
ATBParamArray = oArray

End Function

Private Function KBEBodyArray(oBody As Body, sPartName As String) As Variant
'We create and array to keep both the body and the parent part name in the same object

Dim oArray(1) As Variant

Set oArray(0) = oBody
oArray(1) = sPartName
    
KBEBodyArray = oArray

End Function

