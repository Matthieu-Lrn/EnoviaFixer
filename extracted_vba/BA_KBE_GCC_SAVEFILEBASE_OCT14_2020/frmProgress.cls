Attribute VB_Name = "frmProgress"
Attribute VB_Base = "0{BDA63770-5D2C-4885-ADF6-7AE68F5A4BF9}{71BE1D3A-00CC-475A-A0FA-8432F9390829}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'********************************************************************************
'* Name: Progress bar
'*
'* When modifying this code please update in all toolbars
'*
'* Purpose: Show and Initialize progress bar
'*          Arguments are optional, and defines the shape of the toolbar
'*          e.g. If no iAdvancement is given, no progress bar is shown (only the window with an optional message)
'*          e.g. If no secondary iAdvancement is given, the secondary progress bar is not shown
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette, Abhishek Kamboj
'* Language: VBA
'********************************************************************************

Public Sub progressBarInitialize(Optional ByVal sCaption As String = "")
    
    bCancelAction = False
    
    With Me
        .Height = dTitleBarHeight + 60
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - Me.Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - Me.Height - 105
        .Caption = sCaption
        .pbProgressMain.Min = 0
        .pbProgressMain.Value = 0
        .pbProgressMain.Visible = False
        .pbProgress.Min = 0
        .pbProgress.Value = 0
        .pbProgress.Visible = False
        .MousePointer = fmMousePointerHourGlass
        .lblMessageMain.Caption = ""
        .lblMessage.Caption = ""
        .lblTimer.Caption = ""
    End With
    
End Sub

Public Sub progressBarRepaint(Optional ByVal sMsgMain As String = "", Optional ByVal dMaxValueMain As Double = 100, Optional ByVal iAvancementMain As Double = 0, Optional ByVal sMsg As String = "", Optional ByVal dMaxValue As Double = 100, Optional ByVal iAvancement As Double = 0, Optional ByVal sTimer As String = "")
    
    DoEvents    ' to allow user to exit procedure by pressing escape key
    If bCancelAction = False Then
        With Me
            
            'Main
            .pbProgressMain.Top = 18
            .lblMessageMain.Top = .pbProgressMain.Top - 12
            .pbProgressMain.Max = dMaxValueMain
            .lblMessageMain.Visible = True
            If iAvancementMain > dMaxValueMain Then
                .pbProgressMain.Value = dMaxValueMain
            Else
                .pbProgressMain.Value = iAvancementMain
            End If

            'Secondary
            .pbProgress.Top = .pbProgressMain.Top + 28
            .lblMessage.Top = pbProgress.Top - 12
            .pbProgress.Max = dMaxValue
            If iAvancement > dMaxValue Then
                .pbProgress.Value = dMaxValue
            Else
                .pbProgress.Value = iAvancement
            End If
            
            If iAvancement <> 0 Then
                .Height = .pbProgress.Top + dTitleBarHeight + 48
                .pbProgress.Visible = True
                .lblMessage.Visible = True
                .pbProgressMain.Visible = True
            ElseIf iAvancementMain <> 0 And iAvancement = 0 Then
                .Height = .pbProgressMain.Top + dTitleBarHeight + 48
                .pbProgress.Visible = False
                .lblMessage.Visible = False
                .pbProgressMain.Visible = True
            Else
                .Height = .pbProgressMain.Top + dTitleBarHeight + 48
                .pbProgress.Visible = False
                .lblMessage.Visible = False
                .pbProgressMain.Visible = False
            End If
            
            .imgDSGlobe.Top = Me.Height - dTitleBarHeight - 28
            .imgDSGlobe.Left = -3
            .lblTimer.Top = Me.Height - dTitleBarHeight - 35
            .cmdCancel.Top = Me.Height - dTitleBarHeight - 24
            
            .lblMessageMain.Caption = sMsgMain
            .lblMessage.Caption = sMsg
            .lblTimer.Caption = sTimer
            .Repaint
            
            If .Visible = False Then .Show vbModeless
            
        End With
    End If
    
End Sub

Private Sub cmdCancel_Click()
    bCancelAction = True
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

