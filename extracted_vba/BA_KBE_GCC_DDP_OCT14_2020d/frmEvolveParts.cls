Attribute VB_Name = "frmEvolveParts"
Attribute VB_Base = "0{256BCE96-F480-4B32-98C9-67D607D36803}{EDC1EB04-9C88-46A1-A8D2-E7A08AAD5454}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



'********************************************************
'* SCRIPT VERSION: July 19th 2016
'*
'* When modifying this code please update in all toolbars
'********************************************************
Private oSelectedProduct As Product
Private oDestination As Product
Private oRootDocument As Document

Public Sub InitializeForm()
   
'*** Position Controls
With Me

    .fraPart.Left = 6
    .fraPart.Top = 6
       
    .fraMethods.Left = .fraPart.Left
    .fraMethods.Top = .fraPart.Top + .fraPart.Height + 4
    
    .fraOptions.Left = .fraPart.Left
    .fraOptions.Top = .fraMethods.Top + .fraMethods.Height + 4
    
    .Height = .fraOptions.Top + .fraOptions.Height + dTitleBarHeight + 40
    .Width = fraPart.Left + .fraPart.Width + 10
    
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18

    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
           
    .chkFindAndReplace.Value = True
    Call chkDashNumber_Change
    .chkEvolveAllLevels.Value = True
    .chkDefiningPart.Value = True
    
    .txtBaseNumber.BackColor = &H80000016
    
    .optAskQuestion.Value = True
    
    .Show vbModeless
End With


End Sub



Private Sub cmdPartToBeEvolved_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtPartToBeEvolved.Text = sNoSelection
    
    On Error Resume Next
    Call SelectInstance(oSelectedProduct, oRootDocument, "Please select the part or NHA to be evolved !   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSelectedProduct Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    
    'Update form with selected object
    If Not oSelectedProduct Is Nothing Then
        Me.txtPartToBeEvolved.Text = oSelectedProduct.Name
        Me.txtBaseNumber.Text = GetBaseNumber(oSelectedProduct.Name)
        Set sTargetWindow = CATIA.ActiveWindow
        If Me.chkFindAndReplace.Value = True Then
            Me.txtSourceString = oSelectedProduct.PartNumber
            Me.txtTargetString = oSelectedProduct.PartNumber
        End If
    Else
noSelection:
        Me.txtPartToBeEvolved.Text = ""
        Me.txtBaseNumber.Text = ""
        Set oSelectedProduct = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdOKActivation()

If (Me.txtPartToBeEvolved.Text <> sNoSelection) Then
    Me.cmdOK.ForeColor = -2147483630
    Me.cmdOK.Locked = False
    Me.cmdOK.Default = True
Else
    Me.cmdOK.ForeColor = -2147483631
    Me.cmdOK.Locked = True
    Me.cmdCancel.Default = True
End If

End Sub

Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOK_Click()
    
Dim bAllLevels As Boolean
Dim iIncrement As Integer
Dim sBase As String
Dim sExistingPartOption As String
Dim sLogDump As String

Me.Hide

'Log File
If Me.chkEvolveAllLevels.Value = True Then
    sLogDump = "Evolve All Level = True"
Else
    sLogDump = "Evolve All Level = False"
End If
If Me.chkDefiningPart.Value = True Then
    sLogDump = sLogDump & " | Evolve Defining Part = True"
Else
    sLogDump = sLogDump & " | Evolve Defining Part = False"
End If

Call AddToLogFile("Evolve Part", , txtPartToBeEvolved.Value, , , , , sLogDump)

'** Convert values from Userform
bAllLevels = IIf(Me.chkEvolveAllLevels.Value = True, True, False)
If Me.txtDash.Text = "" Then
    iIncrement = 0
ElseIf Me.txtDash.Text = "-" Then
    iIncrement = 0
ElseIf Me.txtDash.Text = "-0" Then
    iIncrement = 0
ElseIf Left(Me.txtDash.Text, 1) = "-" Then
    iIncrement = -CInt(Mid(Me.txtDash.Text, 2))
Else
    iIncrement = CInt(Me.txtDash.Text)
End If
sBase = IIf(Me.chkDashWithBase.Value = True, txtBaseNumber.Text, "")

If Me.optAskQuestion.Value = True Then
    sExistingPartOption = "Ask Question"
ElseIf Me.optSkipEvolution.Value = True Then
    sExistingPartOption = "Skip Evolution"
Else
    sExistingPartOption = "Reuse"
End If


Call oEvolveParts.LaunchEvolveParts(oSelectedProduct, Me.txtSourceString.Text, Me.txtTargetString.Text, iIncrement, bAllLevels, oRootDocument, sBase, sExistingPartOption, Me.chkDefiningPart.Value)
Call frmKBEMain.resetToolbar
    
End Sub

Private Sub chkFindAndReplace_Change()

If Me.chkFindAndReplace.Value = True Then
    Me.txtSourceString.Locked = False
    Me.txtSourceString.BackColor = &H80000005
    Me.txtTargetString.Locked = False
    Me.txtTargetString.BackColor = &H80000005
Else
    Me.txtSourceString.Locked = True
    Me.txtSourceString.BackColor = &H80000016
    Me.txtTargetString.Locked = True
    Me.txtTargetString.BackColor = &H80000016
End If

With Me
    .txtSourceString.Text = ""
    .txtTargetString.Text = ""
End With

End Sub

Private Sub chkDashNumber_Change()

If Me.chkDashNumber.Value = True Then
    Me.txtDash.Locked = False
    Me.txtDash.BackColor = &H80000005
    Me.chkDashWithBase.Enabled = True
    
Else
    Me.txtDash.Locked = True
    Me.txtDash.BackColor = &H80000016
    Me.chkDashWithBase.Value = False
    Me.chkDashWithBase.Enabled = False
End If

Me.txtDash.Text = ""
End Sub

Private Sub txtDash_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
If Me.txtDash.Text = "" Then

    'Only acceptable character are: 0-9 and -
    KeyAscii = KeyValidNumericAndNegative(KeyAscii)
Else
    'Only acceptable character are: 0-9 and
    KeyAscii = KeyValidNumeric(KeyAscii)
End If

End Sub

Private Function KeyValidNumeric(ByVal iKey As Integer) As String

'0 to 9
If iKey >= 48 And iKey <= 57 Then
    KeyValidNumeric = iKey
Else
    KeyValidNumeric = 0
End If

End Function
Private Function KeyValidNumericAndNegative(ByVal iKey As Integer) As String

'0 to 9
If iKey >= 48 And iKey <= 57 Then
    KeyValidNumericAndNegative = iKey
ElseIf iKey = 45 Then
    KeyValidNumericAndNegative = iKey
Else
    KeyValidNumericAndNegative = 0
End If

End Function



'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub


Private Function GetBaseNumber(ByVal sName As String) As String

Dim sBasenumber As String

'Keep left of the "."
sBasenumber = Split(sName, ".")(0)

'Check if we have a "-"
If InStrRev(sName, "-") = 0 Then
    GetBaseNumber = sBasenumber
    Exit Function
End If

'Get Base Number
GetBaseNumber = Left(sBasenumber, InStrRev(sBasenumber, "-") - 1)

End Function
