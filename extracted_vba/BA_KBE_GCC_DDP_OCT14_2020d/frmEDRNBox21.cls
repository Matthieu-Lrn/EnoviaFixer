Attribute VB_Name = "frmEDRNBox21"
Attribute VB_Base = "0{6C333070-5479-4B73-A7CD-AABE199CE38C}{88FBD260-FCE1-4EAD-AC0D-5B501A7D894A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private bfrmEDRNBusy As Boolean
Private Const sInitialPNbrTxt As String = "Enter PartNumber"
Private Const sInitialRevTxt As String = "Enter Rev"

'Option Explicit

Public Sub InitializeForm()
    With Me
        .txtPartNumber.Text = sInitialPNbrTxt
        .txtRevision.Text = sInitialRevTxt
        
        .Height = dTitleBarHeight + 118
        .Width = 188
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - 12
        .cmdCancel.Left = .Width + 12
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .Show vbModless
    End With
    bfrmEDRNBusy = False
    Call cmdOKActivation
End Sub

Private Sub cmdOK_Click()
    If bfrmEDRNBusy = False Then
        bfrmEDRNBusy = True
        
        Call cmdOKActivation
        
        Dim sEDRNtxt As String

        Call StartWebServiceTool
        sEDRNtxt = WebServiceAccessTool.GetEDRNBox21(txtPartNumber.Text, txtRevision.Text, chkHistory.Value)
        '**** Log usage
        Call AddToLogFile("EDRNBOX 21", , txtPartNumber.Text, txtRevision.Text, , , , "History=" & CStr(chkHistory.Value) & vbCrLf & "Result=" & vbCrLf & sEDRNtxt)
        '****
        If Not (Trim(UCase(sEDRNtxt)) Like "*[0-9,A-Z]*") Then MsgBox "No result retrived for " & txtPartNumber.Text & txtRevision.Text & " from BDI.", vbExclamation
        '*** Transfer text to Clipboard
        On Error Resume Next
            Dim oData As DataObject
            Set oData = New DataObject
            If Err.Number = 0 And Not oData Is Nothing Then
            oData.Clear
            oData.SetText sEDRNtxt
            oData.PutInClipboard
            DoEvents
            End If
            Set oData = Nothing
        On Error GoTo 0
        
        bfrmEDRNBusy = False
        Call cmdCancel_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOKActivation()
    If Not bfrmEDRNBusy And _
            Not (txtPartNumber.Text = sInitialPNbrTxt Or txtPartNumber.Text = "") And _
            Not (txtRevision.Text = sInitialRevTxt Or txtRevision.Text = "") Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
    End If
End Sub
'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub
Private Sub txtPartNumber_Change()
    Dim sChr  As String
    Dim bCheck As Boolean
    
    If bfrmEDRNBusy Then Exit Sub
    
    If Not (txtPartNumber.Text = "" Or txtPartNumber.Text = sInitialPNbrTxt) Then
        On Error Resume Next
        txtPartNumber.Text = UCase(txtPartNumber.Text)
        On Error GoTo 0
        
        For i = 1 To Len(txtPartNumber.Text)
            sChr = Mid(txtPartNumber.Text, i, 1)
            If sChr Like "[-,A-Z,_]" Then
                bCheck = True
            ElseIf sChr Like "[0-9]" Then
                bCheck = True
            Else
                bCheck = False
            End If
            If Not bCheck Then
                txtPartNumber.Text = Left(txtPartNumber.Text, i - 1) & Replace(txtPartNumber.Text, sChr, "", i, 1)
            End If
        Next
    End If
    Call cmdOKActivation
End Sub
Private Sub txtRevision_Change()
    Dim sChr  As String
    Dim bCheck As Boolean
    
    If bfrmEDRNBusy Then Exit Sub
    
    If Not (txtRevision.Text = "" Or txtRevision.Text = sInitialRevTxt) Then
        On Error Resume Next
        txtRevision.Text = UCase(txtRevision.Text)
        On Error GoTo 0
        
        For i = 1 To Len(txtRevision.Text)
            sChr = Mid(txtRevision.Text, i, 1)
            If sChr Like "[-,A-Z]" Then
                bCheck = True
            Else
                bCheck = False
            End If
            If Not bCheck Then
                txtRevision.Text = Left(txtRevision.Text, i - 1) & Replace(txtRevision.Text, sChr, "", i, 1)
            End If
        Next
        If Len(txtRevision.Text) > 2 Then txtRevision.Text = Left(txtRevision.Text, 2)
    End If
    Call cmdOKActivation
End Sub
Private Sub txtRevision_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not bBusy Then
        If txtRevision.Text = sInitialRevTxt Then txtRevision.Text = ""
        txtRevision.ForeColor = &H80000008
    End If
End Sub
Private Sub txtPartNumber_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not bBusy Then
        If txtPartNumber.Text = sInitialPNbrTxt Then txtPartNumber.Text = ""
        txtPartNumber.ForeColor = &H80000008
    End If
End Sub
