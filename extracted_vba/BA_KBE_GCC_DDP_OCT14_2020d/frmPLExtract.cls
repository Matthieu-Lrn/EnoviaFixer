Attribute VB_Name = "frmPLExtract"
Attribute VB_Base = "0{7042E28F-D735-4073-909B-37DE128D950E}{ECF5377D-0CCD-4B43-97FF-B8F8C44C0EF3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Dim oNHA_LHS As Product
Dim oNHA_RHS As Product
Dim oRootDocument As Document


Public Sub InitializeForm()

With Me
    .Caption = "Extract Part List"

    .FrameLHS.Left = 6
    .FrameLHS.Top = 6
    
    .FrameRHS.Left = .FrameLHS.Left
    .FrameRHS.Top = .FrameLHS.Top + .FrameLHS.Height + 6
    
    .Width = .FrameLHS.Left + .FrameLHS.Width + 8
    .Height = dTitleBarHeight + .FrameRHS.Top + .FrameRHS.Height + 30
        
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
    .Show vbModless
End With


Set oNHA_LHS = Nothing
Set oNHA_RHS = Nothing

End Sub


Private Sub cmdNHA_LHS_Click()


    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtNHA_LHS.Text = sNoSelection
    
    On Error Resume Next
    Call selectAssy(oNHA_LHS, oRootDocument, , "Please select the NHA !   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oNHA_LHS Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'Update form with selected object
    If Not oNHA_LHS Is Nothing Then
        Me.txtNHA_LHS.Text = oNHA_LHS.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtNHA_LHS.Text = sNoSelection
        Set oNHA_LHS = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdNHA_RHS_Click()


    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtNHA_RHS.Text = sNoSelection
    
    On Error Resume Next
    Call selectAssy(oNHA_RHS, oRootDocument, , "Please select the NHA !   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oNHA_RHS Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'Update form with selected object
    If Not oNHA_RHS Is Nothing Then
        Me.txtNHA_RHS.Text = oNHA_RHS.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtNHA_RHS.Text = sNoSelection
        Set oNHA_RHS = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdCancel_Click()
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
    Call oPlExtract.LaunchExtraction(oNHA_LHS, oNHA_RHS)
    Call resetUserform
    Call frmKBEMain.resetToolbar
    
End Sub

Private Sub resetUserform()
        
With Me
    .txtNHA_LHS.Text = sNoSelection
    .txtNHA_RHS.Text = sNoSelection
End With
     
End Sub

Private Sub cmdOKActivation()
    
On Error Resume Next
If Not (Me.txtNHA_LHS.Text = sNoSelection) Then
    Me.cmdOK.ForeColor = -2147483630
    Me.cmdOK.Locked = False
    Me.cmdOK.Default = True
    Me.cmdOK.SetFocus
Else
    Me.cmdOK.ForeColor = -2147483631
    Me.cmdOK.Locked = True
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
End If
On Error GoTo 0
    
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

