Attribute VB_Name = "frmSubsetsManagement"
Attribute VB_Base = "0{12C66DC8-5767-4AB0-8DD5-F80FC963A2D6}{DACEE913-112E-423D-9B81-9AD027F84D68}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private SSParameters As Collection
Private SSParametersToDelete As Collection
Private SSParametersToCreate As Collection
Private oAllSelectedItems As Collection
Private oAllAvailableItems As Collection
Private oSelectedBodiesCollection As Collection
Private oSelectedHybridBodiesCollection As Collection
Private sKBEBodiesCollection As Collection
Private sBooleanParamCollection As Collection
Private oCurrentSSParameter As StrParam
Private sCurrentSSParameterName As String
Private oSelectionBody As Body
Private sProductName As String
Private sInstanceDescription As String
Private iAvailableItemIndex As Integer
Private iSubsetIndex As Integer



Public Sub initializeSubSetsForm()
    
    '***** Window size according to number of shifters
    Me.StartUpPosition = Manual
    Me.Height = dTitleBarHeight + 300
    Me.Width = 760
    Me.fraSubsetList.Height = Me.Height - dTitleBarHeight - 82
    Me.lbSPSubsets.Height = Me.fraSubsetList.Height - 75
    Me.fraSubsetManagement.Height = Me.Height - dTitleBarHeight - 40
    Me.lbAvailableItems.Height = Me.fraSubsetManagement.Height - 38
    Me.lbSubsetValues.Height = Me.fraSubsetManagement.Height - 38
    Me.imgDSGlobe.Top = Me.Height - dTitleBarHeight - 28
    Me.imgDSGlobe.Left = -3
    Me.cmdCancel.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdCancel.Left = Me.Width - Me.cmdCancel.Width - 12
    Me.cmdApply.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdApply.Left = Me.Width - Me.cmdApply.Width - Me.cmdCancel.Width - 18
    Me.cmdOK.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdOK.Left = Me.Width - Me.cmdOK.Width - cmdApply.Width - Me.cmdCancel.Width - 24
    Me.lbAvailableItems.Height = Me.fraSubsetManagement.Height - 38
    Me.lbSubsetValues.Height = Me.fraSubsetManagement.Height - 38
    Me.Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - Me.Width - 85
    Me.Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - Me.Height - 85
    Me.Caption = "Subsets Management"
        
    '***** Default values
    Call resetUserform
    Call setDefaultFocus
    Me.Show vbModeless

End Sub

'********************************************************************************
'* Name: OK Button
'* Purpose: When user clicks OK, update SS Parameters in tree and close form
'********************************************************************************
Private Sub cmdOK_Click()
    Call createSSParametersInTree
    Call valuateSSParametersInTree
    Call deleteSSParametersInTree
    Set SSParameters = Nothing
    Call cmdCancel_Click
End Sub

'********************************************************************************
'* Name: Apply Button
'* Purpose: When user clicks Apply, update SS Parameters in tree
'********************************************************************************
Private Sub cmdApply_Click()
    Call createSSParametersInTree
    Call valuateSSParametersInTree
    Call deleteSSParametersInTree
    Call setDefaultFocus
End Sub

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

'********************************************************************************
'* Name: Select Seed Part
'* Purpose: When user press "..." then ask user to select a seed part.
'*          User can select the part, the instance, or anything inside part.
'*          The application will find the part. If SS Parameters exist in
'*          selected part, the form is filled according to what is found.
'********************************************************************************
Private Sub cmdSelectSeedPart_Click()
    
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSelectionSP.Text = sNoSelection
    Call resetUserform
    
    On Error Resume Next
    Call selectPart(oSeedPart, oSeedPartInstance, oSeedPartRootDocument, "*", "Please select a Seed Part!  ")
    If Not oSeedPart.Name Like "*_SEED_PART_*" And Not oSeedPart.Name Like "*-*COL??" And Not oSeedPart.Name Like "*-*COL???" Then Set oSeedPart = Nothing
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSeedPart Is Nothing Then
            MsgBox "Please select a valid Seed Part!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    
    If Not oSeedPart Is Nothing Then
        Me.txtSelectionSP.Text = oSeedPart.Name
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
    Else
noSelection:
        Me.txtSelectionSP.Text = sNoSelection
        Set oSeedPart = Nothing
        Set oSeedPartInstance = Nothing
        Set oSeedPartRootDocument = Nothing
    End If
    
    'Scan the reference Part for the SubSetparameters with the prefix of SS_
    '**** Two names of parameter set to cover: "SubSetParameters" & "ParametersSubSet"
    Set oParameterSet = Nothing
    Set oParameterSet = oSeedPart.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
    Set oParameterSet = oSeedPart.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
    If Not (oParameterSet Is Nothing) Then
        For i = 1 To oParameterSet.DirectParameters.Count
            Set oCurrentSSParameter = oParameterSet.DirectParameters.Item(i)
            sCurrentSSParameterName = Mid(oCurrentSSParameter.Name, InStrRev(oCurrentSSParameter.Name, "\SS_") + Len("\SS_"))
            
            SSParameters.Add New SSParameter, sCurrentSSParameterName
            SSParameters.Item(sCurrentSSParameterName).ParameterObject = oCurrentSSParameter
            SSParameters.Item(sCurrentSSParameterName).ParameterName = sCurrentSSParameterName
            
            For Each sSubSetItem In ConvertStringToCollection(oCurrentSSParameter.ValueAsString, "|")
                SSParameters.Item(sCurrentSSParameterName).AddToSubsetValues sSubSetItem
            Next
            
            Me.lbSPSubsets.AddItem sCurrentSSParameterName
        Next
    End If
    Me.lbSPSubsets.Selected(0) = True
    If Me.lbSPSubsets.ListCount = 0 Then Call frameButtonsActivity(fraSubsetManagement, False)
    
    On Error GoTo 0
        
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
    End If
    
    Call setDefaultFocus
    
End Sub

'********************************************************************************
'* Name: Selected Seed Part - Textbox change
'* Purpose: Detect if user select / unselect a seed part, and enables / disables
'*          buttons and textboxes according to the choice
'********************************************************************************
Private Sub txtSelectionSP_Change()
    
    Call frameButtonsActivity(fraSubsetList, True)
    
    If Me.txtSelectionSP.Text = sNoSelection Then
        Call frameButtonsActivity(fraSubsetList, False)
        Call frameButtonsActivity(fraSubsetManagement, False)
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdApply.ForeColor = -2147483631
        Me.cmdApply.Locked = True
        Me.txtSSName.Text = ""
        Me.txtSSName.Enabled = False
    Else
        Call frameButtonsActivity(fraSubsetList, True)
        '***** Manage fraSubsetManagement buttons activity according to Subsets list
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdApply.ForeColor = -2147483630
        Me.cmdApply.Locked = False
        Me.txtSSName.Enabled = True
    End If
    
End Sub

'********************************************************************************
'* Name: Create Button
'* Purpose: Create a new SS Parameter with the name found in txtSSName.
'*          Note: The SS Parameter is created in userform, but not in tree.
'*          It will be created in tree if user presses OK or Apply.
'*          If the given SS Parameter already exist, it will just highlight it.
'*          If the user gives a name starting with "SS_", the application will
'*          remove it in the userform.
'********************************************************************************
Private Sub cmdCreateSet_Click()
    
    'Remove prefix "SS_" if user puts it
    If Mid(Me.txtSSName.Text, 1, 3) = "SS_" Then Me.txtSSName.Text = Mid(Me.txtSSName.Text, InStrRev(Me.txtSSName.Text, "SS_") + Len("SS_"))
    If Me.txtSSName.Text = "" Then Exit Sub

    On Error Resume Next
    SSParameters.Add New SSParameter, Me.txtSSName.Text
    SSParametersToCreate.Add SSParameters.Item(Me.txtSSName.Text), Me.txtSSName.Text
    
    '***** Remove SSParameter and re-add it at the end of the collection SSParameters, so it appears at the end of the list
    '***** Remove it from the collection of Subsets to delete
    SSParameters.Remove (Me.txtSSName.Text)
    SSParameters.Add SSParametersToCreate.Item(Me.txtSSName.Text), Me.txtSSName.Text
    SSParameters.Item(Me.txtSSName.Text).ParameterName = Me.txtSSName.Text
    SSParametersToDelete.Remove (Me.txtSSName.Text)

    '***** Refresh Subsets listbox
    Me.lbSPSubsets.Clear
    For Each SSParameter In SSParameters
        If SSParametersToDelete.Item(SSParameter.ParameterName) Is Nothing Then Me.lbSPSubsets.AddItem SSParameter.ParameterName
    Next
    
    '***** Select the element created. If user tries to create a SS parameter with the same name, highlight it.
    For i = 0 To lbSPSubsets.ListCount - 1
        If Me.lbSPSubsets.List(i) = Me.txtSSName.Text Then lbSPSubsets.Selected(i) = True
    Next
    On Error GoTo 0
    
    '***** Empty New Set Name field
    Me.txtSSName.Text = ""
    
End Sub

'********************************************************************************
'* Name: Delete Button
'* Purpose: Delete a SS Parameter in the list lbSPSubsets.
'*          Note: The SS Parameter is deleted in userform, but not in tree.
'*          It will be deleted in tree if user presses OK or Apply.
'*          SS Parameter is still kept in memory. If user creates it back with
'*          the same name, it will retrieve its original values
'********************************************************************************
Private Sub cmdDeleteSet_Click()
    
    On Error Resume Next
    If Me.lbSPSubsets.ListIndex <> -1 Then
        SSParametersToDelete.Add SSParameters.Item(Me.lbSPSubsets.Text), Me.lbSPSubsets.Text
        SSParametersToCreate.Remove (Me.lbSPSubsets.Text)
        Me.lbSPSubsets.RemoveItem (Me.lbSPSubsets.ListIndex)
        If Me.lbSPSubsets.ListIndex = -1 Then
            Me.lbAvailableItems.Clear
            Me.lbSubsetValues.Clear
            Call frameButtonsActivity(fraSubsetManagement, False)
        End If
    End If
    On Error GoTo 0
    
    Call setDefaultFocus
    
    End Sub

'********************************************************************************
'* Name: Subsets - Listbox change
'* Purpose: Detect if user change his selection in list lbSPSubsets.
'*          It highlights the parameter in tree if it exist, and populate
'*          listboxes (Available items + Subset values)
'********************************************************************************
Private Sub lbSPSubsets_Change()
    
    'Populate available items list with previously selected objects
    If Me.lbSPSubsets.Text <> "" Then
        SSParameters.Item(Me.lbSPSubsets.Text).CurrentSelectedItems = New Collection
        For Each iObject In oAllAvailableItems
            SSParameters.Item(Me.lbSPSubsets.Text).AddToSelectedItems iObject, False
        Next
    End If
    
    Call PopulateSubsetValues
    Call PopulateAvailableItems
    Call frameButtonsActivity(fraSubsetManagement, True)
    Call setDefaultFocus
    
    'Resize ListBox to see last element
    Me.lbSPSubsets.IntegralHeight = False
    Me.lbSPSubsets.IntegralHeight = True
    
End Sub

'********************************************************************************
'* Name: New Subset Name - Textbox Keydown
'* Purpose: 1-If user presse "Enter", it will create a SS Parameter with the given name
'*          2-If user presses Down, it will change the selection in lbSPSubsets
'*          3-If user presses Up, it will change the selection in lbSPSubsets
'********************************************************************************
Private Sub txtSSName_KeyDown(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyAscii = 13 And txtSSName.Text <> "" Then
        KeyAscii = 0
        Call cmdCreateSet_Click
        Me.txtSSName.SetFocus
    ElseIf KeyAscii = vbKeyDown Then
        KeyAscii = 0
        Call ListBoxMoveDown(lbSPSubsets)
    ElseIf KeyAscii = vbKeyUp Then
        KeyAscii = 0
        Call ListBoxMoveUp(lbSPSubsets)
    End If
End Sub

'********************************************************************************
'* Name: Subsets - Listbox Keydown
'* Purpose: 1-If user presses Down, it will change the selection in lbSPSubsets
'*          2-If user presses Up, it will change the selection in lbSPSubsets
'********************************************************************************
Private Sub lbSPSubsets_KeyDown(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyAscii = vbKeyDown Then
        KeyAscii = 0
        Call ListBoxMoveDown(lbSPSubsets)
    ElseIf KeyAscii = vbKeyUp Then
        KeyAscii = 0
        Call ListBoxMoveUp(lbSPSubsets)
    End If
End Sub

'********************************************************************************
'* Name: Add Button - Add selected elements to "Available Items" list
'* Purpose: Adding the selected items to "Available Items" list. It finds all the
'*          KBE Bodies and Std Parts Boolean parameters selected
'*          Note 1: If a geometrical set is selected, the application can find
'*          boolean parameters under it.
'*          Note 2: User can select on the tree or interactive. In interactive,
'*          it is recommended to do a bounding box to select multiple elements.
'*          Selecting many "Single elements" may take more time (Performance problem
'*          when trying to retrive the parent of a face or an edge).
'*          Note 3: Possible selections in Seed parts, KBE Parts, Std parts (SP)
'*          and Std assy (SA)
'********************************************************************************
Private Sub cmdAddSelection_Click()
    
    On Error Resume Next
    
    Set sKBEBodiesCollection = New Collection
    Set oSelectedBodiesCollection = New Collection
    Set oSelectedHybridBodiesCollection = New Collection
    Set sBooleanParamCollection = New Collection
    Set oSelectedItems = New Collection
    
    CATIA.Interactive = False 'Do not allow user to click something
    frmKBEMain.Enabled = False 'Do not allow user to click something
    
    Call frmProgress.progressBarInitialize("Add Selected Items")
    Call frmProgress.progressBarRepaint("Scanning Selection...")
    
    Set oSelection = CATIA.ActiveDocument.Selection
    For i = 1 To oSelection.Count2
        oSelectedItems.Add oSelection.Item2(i)
    Next i
    oSelection.Clear
    
    iProgressCount1 = 1
    iProgressCount1Max = 3
    iProgressCount2 = 0
    iProgressCount2Max = oSelectedItems.Count
    For Each iObject In oSelectedItems
            
        iProgressCount2 = iProgressCount2 + 1
             
        If TypeName(iObject.Document) = "ProductDocument" Then sProductName = iObject.LeafProduct.Name: sInstanceDescription = GetDescriptionFromProduct(iObject.LeafProduct, "*_Boolean")
        If TypeName(iObject.Document) = "PartDocument" Then sProductName = iObject.Document.Product.Name: sInstanceDescription = ""
        
        'Adding KBE, DTL and ASY bodies
        If sProductName Like "KBE*" Or sProductName Like "*-*COL??" Or sProductName Like "*-*COL???" Or sInstanceDescription = "" Then
            Set oSelectionBody = GetBodyFromObject(iObject.Value)
            
            If oSelectionBody Is Nothing And TypeName(iObject.Value) = "HybridBody" Then
                oSelectedHybridBodiesCollection.Add iObject.Value
            End If
            oSelectedBodiesCollection.Add oSelectionBody, oSelectionBody.Name
        'Adding Selected Std Parts
        ElseIf TypeName(iObject.Document) = "ProductDocument" Then
            If sInstanceDescription Like "*_Boolean" And _
            (sInstanceDescription Like "SP_*" Or _
            sInstanceDescription Like "SA_*") Then
                sBooleanParamCollection.Add sInstanceDescription, sInstanceDescription
                oAllSelectedItems.Add iObject.LeafProduct, sInstanceDescription
            End If
        End If
                   
        Call frmProgress.progressBarRepaint("Scanning Selection...", iProgressCount1Max, iProgressCount1, "Finding bodies and instances...", iProgressCount2Max, iProgressCount2, "(" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")")
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
                   
    Next
    
    '***** Creating KBE Bodies name collection
    
    iProgressCount1 = 2
    iProgressCount2 = 0
    iProgressCount2Max = oSelectedBodiesCollection.Count
    For Each iObject In oSelectedBodiesCollection
        
        iProgressCount2 = iProgressCount2 + 1
        
        '***** Look if it is a KBE Body. If not (ASY or DTL), search it it for KBE Bodies
        If iObject.Name Like "KBE*" Then
            sKBEBodiesCollection.Add iObject.Name, iObject.Name
            oAllSelectedItems.Add iObject, iObject.Name
        Else
            Set oSelection = KBESearch("Name_CAP=*KBE_* & CATPrtSearch.BodyFeature,sel", iObject)
            For i = 1 To oSelection.Count2
                sKBEBodiesCollection.Add oSelection.Item2(i).Value.Name, oSelection.Item2(i).Value.Name
                oAllSelectedItems.Add iObject, oSelection.Item2(i).Value.Name
            Next i
        End If
        
        'Searching for Std Parts/Asy in KBE, DTL and ASY
        Set oSelection = KBESearch("CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean,sel", iObject)
        For i = 1 To oSelection.Count2
            If oSelection.Item2(i).Value.Name Like "*\SP_*" Then
                sInstanceDescription = Mid(oSelection.Item2(i).Value.Name, InStrRev(oSelection.Item2(i).Value.Name, "\SP_") + 1)
            ElseIf InStroSelection.Item2(i).Value.Name Like "*\SA_*" Then
                sInstanceDescription = Mid(oSelection.Item2(i).Value.Name, InStrRev(oSelection.Item2(i).Value.Name, "\SA_") + 1)
            End If
            sBooleanParamCollection.Add sInstanceDescription, sInstanceDescription
            oAllSelectedItems.Add oSelection.Item2(i).Value, sInstanceDescription
        Next i
                
        Call frmProgress.progressBarRepaint("Scanning Selection...", iProgressCount1Max, iProgressCount1, "Scanning KBE bodies for boolean parameters...", iProgressCount2Max, iProgressCount2, "(" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")")
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
    Next
    
    '***** Finding boolean parameters in selected geometrical sets
    iProgressCount1 = 3
    iProgressCount2 = 1
    iProgressCount2Max = 1
    Call frmProgress.progressBarRepaint("Scanning Selection...", iProgressCount1Max, iProgressCount1, "Scanning Geometrical Sets for boolean parameters...", iProgressCount2Max, iProgressCount2, "(" & CStr(oSelectedHybridBodiesCollection.Count) & "/" & CStr(oSelectedHybridBodiesCollection.Count) & ")")
    
    oSelection.Clear
    For Each iObject In oSelectedHybridBodiesCollection
        oSelection.Add iObject
    Next
    
    Set oSelection = KBESearch("CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean,sel")
    
    For i = 1 To oSelection.Count2
        If oSelection.Item2(i).Value.Name Like "*\SP_*" Then
            sInstanceDescription = Mid(oSelection.Item2(i).Value.Name, InStrRev(oSelection.Item2(i).Value.Name, "\SP_") + 1)
        ElseIf oSelection.Item2(i).Value.Name Like "*\SA_*" Then
            sInstanceDescription = Mid(oSelection.Item2(i).Value.Name, InStrRev(oSelection.Item2(i).Value.Name, "\SA_") + 1)
        End If
        sBooleanParamCollection.Add sInstanceDescription, sInstanceDescription
        oAllSelectedItems.Add oSelection.Item2(i).Value, sInstanceDescription
    Next i
        
endsub:
    '***** Adding KBE Bodies and Boolean Parameters to selected CurrentSelectedItems collection
    'KBE Bodies to collection
    For Each iObject In sKBEBodiesCollection
        SSParameters.Item(Me.lbSPSubsets.Text).AddToSelectedItems iObject, False
        oAllAvailableItems.Add iObject, iObject
    Next
    'Add Boolean parameters to collection
    For Each iObject In sBooleanParamCollection
        SSParameters.Item(Me.lbSPSubsets.Text).AddToSelectedItems iObject, False
        oAllAvailableItems.Add iObject, iObject
    Next
    
    oSelection.Clear
    frmProgress.Hide
    CATIA.Interactive = True
    frmKBEMain.Enabled = True
    CATIA.StatusBar = "Selection succesfully added! "
    Call PopulateAvailableItems
    On Error GoTo 0
    
End Sub

'********************************************************************************
'* Name: Clear Button
'* Purpose: Clear the list of "Available Items".
'********************************************************************************
Private Sub cmdClearSelection_Click()
    SSParameters.Item(Me.lbSPSubsets.Text).CurrentSelectedItems = New Collection
    Set oAllAvailableItems = New Collection
    Call PopulateAvailableItems
End Sub

'********************************************************************************
'* Purpose: Highlight elements that were added to "Available Items".
'********************************************************************************
Private Sub cmdHighlightItems_Click()
    
    Dim bNothingSelected As Boolean
    On Error Resume Next
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    bNothingSelected = True
    For i = 0 To Me.lbSubsetValues.ListCount - 1
        If Me.lbSubsetValues.Selected(i) = True Then
            oSelection.Add oAllSelectedItems.Item(Me.lbSubsetValues.List(i))
            bNothingSelected = False
        End If
    Next
    
    For i = 0 To Me.lbAvailableItems.ListCount - 1
        If Me.lbAvailableItems.Selected(i) = True Then
            oSelection.Add oAllSelectedItems.Item(Me.lbAvailableItems.List(i))
            bNothingSelected = False
        End If
    Next
    
    If bNothingSelected Then
        For Each iObject In oAllSelectedItems
            oSelection.Add iObject
        Next
    End If
    On Error GoTo 0
    
End Sub

'********************************************************************************
'* Name: All left Button
'* Purpose: Put all "Subset values" in "Available Items"
'********************************************************************************
Private Sub cmdAllLeft_Click()
    If Not Me.lbSPSubsets.Text = "" Then
        For Each iObject In SSParameters.Item(Me.lbSPSubsets.Text).SubsetValues
            SSParameters.Item(Me.lbSPSubsets.Text).AddToSelectedItems iObject, True
            On Error Resume Next
            oAllAvailableItems.Add iObject, iObject
            On Error GoTo 0
        Next
        
        Call PopulateSubsetValues
        Call PopulateAvailableItems
    End If
End Sub

'********************************************************************************
'* Name: Selection left Button
'* Purpose: Put selected "Subset values" in "Available Items"
'********************************************************************************
Private Sub cmdSelectionLeft_Click()
    If Not Me.lbSPSubsets.Text = "" Then
        For i = 0 To Me.lbSubsetValues.ListCount - 1
            If Me.lbSubsetValues.Selected(i) = True Then
                SSParameters.Item(Me.lbSPSubsets.Text).AddToSelectedItems Me.lbSubsetValues.List(i), True
                On Error Resume Next
                oAllAvailableItems.Add Me.lbSubsetValues.List(i), Me.lbSubsetValues.List(i)
                On Error GoTo 0
            End If
        Next
        
        Call PopulateSubsetValues
        Call PopulateAvailableItems
    End If
End Sub

'********************************************************************************
'* Name: Selection right Button
'* Purpose: Put selected "Available Items" in "Subset values"
'********************************************************************************
Private Sub cmdSelectionRight_Click()
    If Not Me.lbSPSubsets.Text = "" Then
        For i = 0 To Me.lbAvailableItems.ListCount - 1
            If Me.lbAvailableItems.Selected(i) = True Then
                SSParameters.Item(Me.lbSPSubsets.Text).AddToSubsetValues Me.lbAvailableItems.List(i)
                oAllAvailableItems.Remove Me.lbAvailableItems.List(i)
            End If
        Next
        
        Call PopulateSubsetValues
        Call PopulateAvailableItems
    End If
End Sub

'********************************************************************************
'* Name: All right Button
'* Purpose: Put all "Available Items" in "Subset values"
'********************************************************************************
Private Sub cmdAllRight_Click()
    If Not Me.lbSPSubsets.Text = "" Then
        For Each iObject In SSParameters.Item(Me.lbSPSubsets.Text).CurrentSelectedItems
            SSParameters.Item(Me.lbSPSubsets.Text).AddToSubsetValues iObject
            oAllAvailableItems.Remove iObject
        Next
        
        Call PopulateSubsetValues
        Call PopulateAvailableItems
    End If
End Sub

'********************************************************************************
'* Name: Move down Button
'* Purpose: Move down selected "Subset values" in the list
'********************************************************************************
Private Sub cmdMoveDown_Click()
    Call ListBoxMoveDown(lbSubsetValues, True)
    Call resetSSValuescollection
End Sub

'********************************************************************************
'* Name: Move up Button
'* Purpose: Move up selected "Subset values" in the list
'********************************************************************************
Private Sub cmdMoveUp_Click()
    Call ListBoxMoveUp(lbSubsetValues, True)
    Call resetSSValuescollection
End Sub

'********************************************************************************
'* Name: Reorder Button
'* Purpose: Alphabetically reorder "Subset values" in the list
'********************************************************************************
Private Sub cmdReorder_Click()
    Call ListBoxReorder(lbSubsetValues)
    Call resetSSValuescollection
End Sub

'********************************************************************************
'* Name: Populate "Subset values"
'* Purpose: Populate list "Subset values" according to selected SSParameter
'********************************************************************************
Private Sub PopulateSubsetValues()
    
    lbSubsetValues.Clear
    If Not Me.lbSPSubsets.Text = "" Then
        For Each iObject In SSParameters.Item(Me.lbSPSubsets.Text).SubsetValues
            lbSubsetValues.AddItem iObject
        Next
    End If
    Call setDefaultFocus
    
End Sub

'********************************************************************************
'* Name: Populate "Available Items"
'* Purpose: Populate list "Available Items" according to selected SSParameter
'********************************************************************************
Private Sub PopulateAvailableItems()
    
    Me.lbAvailableItems.Clear
    If Not Me.lbSPSubsets.Text = "" Then
        For Each iObject In SSParameters.Item(Me.lbSPSubsets.Text).CurrentSelectedItems
            lbAvailableItems.AddItem iObject
        Next
    End If
    Call setDefaultFocus
    
End Sub

'********************************************************************************
'* Name: Create parameters in tree
'* Purpose: If it does not exist, create a parameter set "SubSetParameters"
'*          Create in tree newly created parameters in userform (collection SSParametersToCreate)
'********************************************************************************
Private Sub createSSParametersInTree()
    '***** Create SubSetParameters Set if it does not exist
    If oParameterSet Is Nothing And SSParametersToCreate.Count > 0 Then
        oSeedPart.Parameters.RootParameterSet.ParameterSets.CreateSet ("ParametersSubSet")
        Set oParameterSet = oSeedPart.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
    End If

    '***** Create a String Parameter under the SubSetParameters Node
    For Each iObject In SSParametersToCreate
        oParameterSet.DirectParameters.CreateString "SS_" & iObject.ParameterName, ""
        SSParameters.Item(iObject.ParameterName).ParameterObject = oParameterSet.DirectParameters.Item(oParameterSet.DirectParameters.Count)
        SSParametersToCreate.Remove iObject.ParameterName
    Next
End Sub

'********************************************************************************
'* Name: Valuate parameters in tree
'* Purpose: For each SS Parameter, with its "Subset values", create a
'*          string with all values, separated with "|", and valuate all Parameters
'********************************************************************************
Private Sub valuateSSParametersInTree()
    If oParameterSet Is Nothing Then Exit Sub
    
    For Each oParameter In oParameterSet.DirectParameters
        sCurrentSSParameterName = Mid(oParameter.Name, InStrRev(oParameter.Name, "\SS_") + Len("\SS_"))
        oParameter.Value = ConvertCollectionToString(SSParameters.Item(sCurrentSSParameterName).SubsetValues, "|")
    Next
End Sub

'********************************************************************************
'* Name: Delete parameters in tree
'* Purpose: Delete in tree all parameters that user chose to delete
'*          and reset collection SSParametersToDelete. If all parameters are
'*          deleted, delete also the parameter set "SubSetParameters"
'********************************************************************************
Private Sub deleteSSParametersInTree()
    '***** Delete in tree all parameters in SSParametersToDelete, then reset collections
    If oParameterSet Is Nothing Then Exit Sub
    
    For i = oParameterSet.DirectParameters.Count To 1 Step -1
        For Each iObject In SSParametersToDelete
            sCurrentSSParameterName = ""
            On Error Resume Next '***** If iObject.ParameterObject does not exist (trying to delete a newly created parameter)
            sCurrentSSParameterName = iObject.ParameterObject.Name
            On Error GoTo 0
            
            If oParameterSet.DirectParameters.Item(i).Name = sCurrentSSParameterName Then
                oParameterSet.DirectParameters.Remove (i)
                SSParameters.Remove (iObject.ParameterName)
                SSParametersToDelete.Remove (iObject.ParameterName)
                Exit For
            End If
        Next
    Next
    
    If SSParameters.Count = 0 Then
    
        Me.lbAvailableItems.Clear '***** Clear Current Selection List
        Me.lbSubsetValues.Clear '***** Clear Subset Values List
        
        If Not oParameterSet Is Nothing Then '***** Delete oParameterSet in tree if it is empty
            
            '***** If we are in the context of a product, we need to activate the product to be able to delete
            '***** Switching to Assembly Design makes us sure that we are active on product
            If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
            sPreviousWorkbench = CATIA.GetWorkbenchId
            If CATIA.GetWorkbenchId <> "Assembly" And TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
                CATIA.StartWorkbench "Assembly"
            End If
            
            oSeedPart.HybridShapeFactory.DeleteObjectForDatum oSeedPart.CreateReferenceFromObject(oParameterSet)
            Set oParameterSet = Nothing
            
            '***** Going back to previous workbench
            '***** If we are in a product, we can select one (and only one) element inside a part and switch to
            '***** a "Part Workbench" to activate this part
            If sPreviousWorkbench <> "Assembly" And TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
                oSelection.Clear
                oSelection.Add oSeedPart
                CATIA.StartWorkbench sPreviousWorkbench
                oSelection.Clear
            End If
        End If
    End If
End Sub

'********************************************************************************
'* Name: Reset "Subset values" collection
'* Purpose: Reorder collection SubsetValues. After reordering the listbox
'*          "Subset Values", empty the collection SubsetValues, and
'*          populate it again with the same order shown in the listbox
'********************************************************************************
Private Sub resetSSValuescollection()
    SSParameters.Item(Me.lbSPSubsets.Text).SubsetValues = New Collection
    For i = 0 To lbSubsetValues.ListCount - 1
        SSParameters.Item(Me.lbSPSubsets.Text).AddToSubsetValues lbSubsetValues.List(i)
    Next
End Sub

'********************************************************************************
'* Name: Frame Buttons activity
'* Purpose: Activate or deactivate all cmd buttons in a frame passed in arguments
'*          (oFrame) according to the boolean passed in arguments (bActivity)
'********************************************************************************
Private Sub frameButtonsActivity(ByVal oFrame As Frame, ByVal bActivity As Boolean)
    For Each oControl In oFrame.Controls
        If TypeName(oControl) = "CommandButton" Then
            If bActivity Then
                oControl.ForeColor = -2147483630
                oControl.Locked = False
            Else
                oControl.ForeColor = -2147483631
                oControl.Locked = True
            End If
        End If
    Next
End Sub

'********************************************************************************
'* Name: Reset Userform
'* Purpose: Empty all listboxes, textboxes and collections
'*          Deactivate buttons in frames
'********************************************************************************
Private Sub resetUserform()
    Me.txtSelectionSP.Text = sNoSelection
    Me.txtSSName = ""
    Me.lbSPSubsets.Clear
    Me.lbAvailableItems.Clear
    Me.lbSubsetValues.Clear
    Set SSParameters = New Collection
    Set SSParametersToDelete = New Collection
    Set SSParametersToCreate = New Collection
    Set oAllSelectedItems = New Collection
    Set oAllAvailableItems = New Collection
    Call frameButtonsActivity(fraSubsetList, False)
    Call frameButtonsActivity(fraSubsetManagement, False)
End Sub

'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if there's no seed part selected
'*          Otherwise, set the focus on OK button
'********************************************************************************
Private Sub setDefaultFocus()
    
    'Set focus
    If Me.txtSelectionSP.Text = sNoSelection Then
        Me.cmdCancel.Default = True
        Me.cmdCancel.SetFocus
    Else
        Me.cmdOK.Default = True
        Me.cmdOK.SetFocus
        Me.txtSSName.SetFocus
    End If
    
End Sub

'********************************************************************************
'* Name: Reset Focus when clicking
'* Purpose: Reset buttons focus when clicking in userform
'********************************************************************************
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub txtSelectionSP_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub fraSeedPart_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub lbSPSubsets_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
    
    '***** Right click: Highlight selected object in tree if it exist, and center graph on it
    On Error Resume Next
    oSelectionFromSeedPart.Clear
    If Button = 2 Then
        oSelectionFromSeedPart.Add SSParameters.Item(lbSPSubsets.Text).ParameterObject
        CATIA.StartCommand "CATCafCenterGraphOnCSOHdr"
    End If
    On Error GoTo 0
    
End Sub
Private Sub fraSubsetManagement_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub lbSubsetValues_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub lbSubsetValues_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'Allow to unselect an element in list
    Dim iSelected As Integer
    iSelected = 0
    For i = 0 To Me.lbSubsetValues.ListCount - 1
        If Me.lbSubsetValues.Selected(i) Then iSelected = iSelected + 1
        If iSelected > 1 Then iSubsetIndex = -1: Exit For
    Next i
    If iSelected = 1 And Me.lbSubsetValues.ListIndex = iSubsetIndex Then
        Me.lbSubsetValues.Selected(Me.lbSubsetValues.ListIndex) = False
        iSubsetIndex = -1
    ElseIf iSelected = 1 Then
        iSubsetIndex = Me.lbSubsetValues.ListIndex
    End If
End Sub
Private Sub lbSubsetValues_Change()
    'Resize ListBox to see last element
    Me.lbSubsetValues.IntegralHeight = False
    Me.lbSubsetValues.IntegralHeight = True
End Sub
Private Sub lbAvailableItems_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call setDefaultFocus
End Sub
Private Sub lbAvailableItems_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'Allow to unselect an element in list
    Dim iSelected As Integer
    iSelected = 0
    For i = 0 To Me.lbAvailableItems.ListCount - 1
        If Me.lbAvailableItems.Selected(i) Then iSelected = iSelected + 1
        If iSelected > 1 Then iAvailableItemIndex = -1: Exit For
    Next i
    If iSelected = 1 And Me.lbAvailableItems.ListIndex = iAvailableItemIndex Then
        Me.lbAvailableItems.Selected(Me.lbAvailableItems.ListIndex) = False
        iAvailableItemIndex = -1
    ElseIf iSelected = 1 Then
        iAvailableItemIndex = Me.lbAvailableItems.ListIndex
    End If
End Sub
Private Sub lbAvailableItems_Change()
    'Resize ListBox to see last element
    Me.lbAvailableItems.IntegralHeight = False
    Me.lbAvailableItems.IntegralHeight = True
End Sub
