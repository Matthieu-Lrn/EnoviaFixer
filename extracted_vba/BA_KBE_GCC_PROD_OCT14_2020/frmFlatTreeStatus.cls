Attribute VB_Name = "frmFlatTreeStatus"
Attribute VB_Base = "0{FF970091-7572-4068-8B4A-005098E5B302}{A63C6242-0FD4-49E3-9C56-E415A734D95E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private iPartsToModifyIndex As Integer
Private iPartsNotFoundIndex As Integer

Public Sub initializeFlatTreeStatusForm(oInputCollection)
    
    '***** Window size according to number of shifters
    Me.StartUpPosition = Manual
    Me.Height = dTitleBarHeight + 205
    Me.Width = 375
    Me.lbPartsNotFound.Height = Me.Height - dTitleBarHeight - 50
    Me.lbPartsToModify.Height = Me.Height - dTitleBarHeight - 50
    Me.imgDSGlobe.Top = Me.Height - dTitleBarHeight - 28
    Me.imgDSGlobe.Left = -3
    Me.cmdCancel.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdCancel.Left = Me.Width - Me.cmdCancel.Width - 12
    Me.cmdOK.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdOK.Left = Me.Width - Me.cmdOK.Width - Me.cmdCancel.Width - 24
    Me.Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - Me.Width - 85
    Me.Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - Me.Height - 85
    Me.Caption = "Clean Flat Tree"
        
    '***** Default values
    Me.lbPartsNotFound.Clear
    Me.lbPartsToModify.Clear
    
    For Each iObject In oInputCollection
        Me.lbPartsNotFound.AddItem iObject.ValueAsString
    Next
    
    Me.cmdOK.SetFocus
    If Not Me.Visible Then Me.Show
    
End Sub

'********************************************************************************
'* Name: Cancel Button
'********************************************************************************
Private Sub cmdCancel_Click()
    Call cmdAllLeft_Click
    Me.Hide
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

'********************************************************************************
'* Name: OK Button
'********************************************************************************
Private Sub cmdOK_Click()
    Me.Hide
End Sub

'********************************************************************************
'* Name: All left Button
'********************************************************************************
Private Sub cmdAllLeft_Click()
    For i = Me.lbPartsToModify.ListCount - 1 To 0 Step -1
        Me.lbPartsNotFound.AddItem (Me.lbPartsToModify.List(0))
        Me.lbPartsToModify.RemoveItem (0)
    Next
End Sub

'********************************************************************************
'* Name: Selection left Button
'********************************************************************************
Private Sub cmdSelectionLeft_Click()
    For i = 0 To Me.lbPartsToModify.ListCount - 1
        If i > Me.lbPartsToModify.ListCount - 1 Then Exit For
        If Me.lbPartsToModify.Selected(i) = True Then
            Me.lbPartsNotFound.AddItem (Me.lbPartsToModify.List(i))
            Me.lbPartsToModify.RemoveItem (i)
            i = i - 1
        End If
    Next
End Sub

'********************************************************************************
'* Name: Selection right Button
'********************************************************************************
Private Sub cmdSelectionRight_Click()
    For i = 0 To Me.lbPartsNotFound.ListCount - 1
        If i > Me.lbPartsNotFound.ListCount - 1 Then Exit For
        If Me.lbPartsNotFound.Selected(i) = True Then
            Me.lbPartsToModify.AddItem (Me.lbPartsNotFound.List(i))
            Me.lbPartsNotFound.RemoveItem (i)
            i = i - 1
        End If
    Next
End Sub

'********************************************************************************
'* Name: All left Button
'********************************************************************************
Private Sub cmdAllRight_Click()
    For i = Me.lbPartsNotFound.ListCount - 1 To 0 Step -1
        Me.lbPartsToModify.AddItem (Me.lbPartsNotFound.List(0))
        Me.lbPartsNotFound.RemoveItem (0)
    Next
End Sub

Private Sub lbPartsNotFound_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'Allow to unselect an element in list
    Dim iSelected As Integer
    iSelected = 0
    For i = 0 To Me.lbPartsNotFound.ListCount - 1
        If Me.lbPartsNotFound.Selected(i) Then iSelected = iSelected + 1
        If iSelected > 1 Then iPartsNotFoundIndex = -1: Exit For
    Next i
    If iSelected = 1 And Me.lbPartsNotFound.ListIndex = iPartsNotFoundIndex Then
        Me.lbPartsNotFound.Selected(Me.lbPartsNotFound.ListIndex) = False
        iPartsNotFoundIndex = -1
    ElseIf iSelected = 1 Then
        iPartsNotFoundIndex = Me.lbPartsNotFound.ListIndex
    End If
    
End Sub
Private Sub lbPartsNotFound_Change()
    'Resize ListBox to see last element
    Me.lbPartsNotFound.IntegralHeight = False
    Me.lbPartsNotFound.IntegralHeight = True
End Sub

Private Sub lbPartsToModify_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'Allow to unselect an element in list
    Dim iSelected As Integer
    iSelected = 0
    For i = 0 To Me.lbPartsToModify.ListCount - 1
        If Me.lbPartsToModify.Selected(i) Then iSelected = iSelected + 1
        If iSelected > 1 Then iPartsToModifyIndex = -1: Exit For
    Next i
    If iSelected = 1 And Me.lbPartsToModify.ListIndex = iPartsToModifyIndex Then
        Me.lbPartsToModify.Selected(Me.lbPartsToModify.ListIndex) = False
        iPartsToModifyIndex = -1
    ElseIf iSelected = 1 Then
        iPartsToModifyIndex = Me.lbPartsToModify.ListIndex
    End If
    
End Sub
Private Sub lbPartsToModify_Change()
    'Resize ListBox to see last element
    Me.lbPartsToModify.IntegralHeight = False
    Me.lbPartsToModify.IntegralHeight = True
End Sub

