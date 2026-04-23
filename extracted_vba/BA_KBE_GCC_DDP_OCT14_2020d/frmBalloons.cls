Attribute VB_Name = "frmBalloons"
Attribute VB_Base = "0{55027723-45C8-4491-BCEE-A9937D635005}{488EFF45-C3B8-402E-8A2F-9C107993BE59}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public bfrmBalloonBusy As Boolean
'Option Explicit

Public Sub InitializeForm()
    With Me
        .lstProgram.AddItem "G7000"
        .lstProgram.AddItem "G8000"
        .lstProgram.AddItem "G6000"
        .lstProgram.AddItem "G5000"
        .lstProgram.AddItem "Others"
        .lstProgram.Value = "G7000"
        .lstProgram.Enabled = True
        .Cmb_Precision.AddItem "0.00"
        .Cmb_Precision.AddItem "0.000"
        .Cmb_Precision.Value = "0.00"
        
        .Height = dTitleBarHeight + 138
        .Width = 196
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
    bfrmBalloonBusy = False
    Call cmdOKActivation
End Sub

Private Sub cmdOK_Click()
    Dim sdummy As String
    Me.lstProgram.Enabled = False
    bfrmBalloonBusy = True
    
    Call AddToLogFile("Balloon Update")
    Call cmdOKActivation
    Do While bfrmBalloonBusy
        Call mdlBallonRef.CreateFSBallon(sdummy)
        If bStopMultiSelection Or bToolbarCommandSelected Then Call cmdCancel_Click
    Loop
    Call cmdOKActivation
End Sub

Private Sub cmdCancel_Click()
    'Call resetUserform
    bfrmBalloonBusy = False
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOKActivation()
    If Not bfrmBalloonBusy Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdCancel.Default = True
    End If
End Sub
'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub
