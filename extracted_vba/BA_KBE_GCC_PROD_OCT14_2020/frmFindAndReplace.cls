Attribute VB_Name = "frmFindAndReplace"
Attribute VB_Base = "0{273CD59A-DE6C-48D3-9929-DE731045CE4F}{28EC7052-0F7A-4150-A145-1E1EAABEF422}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private StrToFind As String
Private dClipboardContent As DataObject


Public Function FormInitialize()
    
    Dim StrToFind As String
    Set dClipboardContent = New DataObject
    
    With Me
        .Left = dPointsToPixelRatioH * CATIA.Left + 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - Me.Height - 105
        .Width = 275
        
        .txtStringToFind.Value = ""
        .lblStringToFind.Top = 6
        .lblStringToFind.Left = 8
        .txtStringToFind.Top = .lblStringToFind.Top + 12
        .txtStringToFind.Left = 8
        .txtStringToFind.Width = .Width - 20
        .txtStringToFind.Height = 18
        
        On Error Resume Next
        Err.Clear
        dClipboardContent.GetFromClipboard
        StrToFind = ""
        StrToFind = dClipboardContent.GetText
        If Len(StrToFind) = 0 Or Len(StrToFind) > 30 Or Err > 0 Then StrToFind = "RENAME"
        On Error GoTo 0
        .txtStringToFind.Text = StrToFind
        .txtStringToFind.SelStart = 0
        .txtStringToFind.SelLength = Len(.txtStringToFind.Text)
        .txtStringToFind.SetFocus
        
        .lblReplacingString.Top = .txtStringToFind.Top + .txtStringToFind.Height + 6
        .lblReplacingString.Left = 8
        .txtReplacingString.Top = .lblReplacingString.Top + 12
        .txtReplacingString.Left = 8
        .txtReplacingString.Width = .Width - 20
        .txtReplacingString.Height = 18
        .txtReplacingString.Text = ""
        
        .Height = .txtReplacingString.Top + .txtReplacingString.Height + dTitleBarHeight + 32
        
        .cbRenameInstances.Top = .Height - dTitleBarHeight - 24
        .cbRenameInstances.Left = 45
        .cbRenameInstances.Value = False
        
        '***** PUT CHECK BOX NOT VISIBLE IN PRODUCTION TOOLBAR - Not suitable for users, must be used only on Local files, not in Enovia
        .cbRenameInstances.Visible = False
        
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        .Show
    End With

End Function

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
    Me.txtStringToFind.Value = ""
    Me.txtReplacingString.Value = ""
    Call frmKBEMain.resetToolbar
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Private Sub txtStringToFind_Change()
    Call cmdOKActivation
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if there's no seed part selected
'*          Otherwise, set the focus on OK button
'********************************************************************************
Private Sub cmdOKActivation()
    
    On Error Resume Next
    If Me.txtStringToFind = "" Then
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdCancel.Default = True
    Else
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
    End If
    On Error GoTo 0
    
End Sub


