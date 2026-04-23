Attribute VB_Name = "frmFinish"
Attribute VB_Base = "0{BCD5D103-253D-468A-A8DC-F82C69F3D4A2}{1CEF37B9-D37E-4B7E-9C15-9EDC08CF2198}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oFinishProd As Product
Private oFinishPart As Part
Private oRootDocument As Document
Private oFirstRootDocument As Document
Private oFinishTargetPart As New clsFinishTargetPart
Private oCreateFinish As New CreateFinish
Private oInsertHole As Variant
Private oExcel As clsExcel

Public Sub InitializeForm()
'********************************************************************************
'* Name: Finish
'* Purpose: Create a finish pad
'*
'* Assumption:
'*
'* Author: Francois Charette
'* Language: VBA
'********************************************************************************

With Me

    .fraFinishPart.Left = 6
    .fraFinishPart.Top = 9
    
    .fraType.Left = .fraFinishPart.Left
    .fraType.Top = .fraFinishPart.Top + .fraFinishPart.Height + 4.5
    
    .fraOption.Left = .fraFinishPart.Left
    .fraOption.Top = .fraType.Top + .fraType.Height + 4.5
    
    .Height = dTitleBarHeight + .fraOption.Top + .fraOption.Height + 30
    .Width = .fraFinishPart.Left + .fraFinishPart.Width + 10


    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
    .chkDisableInsertHole.Value = True
    .chkDisableSpecificDiatHole.Value = True
    
    .txtLowerLimit.Text = 0
    .txtUpperLimit.Text = 0.1
End With

'Open Excel file and populate the combo box
Call PopulateComboFinishType


'** Retrieve the Inset Hole definition
'Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetNewExcel
Call oExcel.OpenExcelFile(sInsertHoleDefinition)

'Transfer data to oInsertHole
iLastRow = 2
Do While Trim(oExcel.App.Cells(iLastRow, 1).Value) <> ""
    iLastRow = iLastRow + 1
Loop
iLastRow = iLastRow - 1
oInsertHole = oExcel.App.Range(oExcel.App.Cells(2, 1), oExcel.App.Cells(iLastRow, 2)).Value

'Close Excel workbook and kill process
Call oExcel.CloseAlreadyOpenWorkbook(sInsertHoleDefinition)
Set oExcel = Nothing

Set oFirstRootDocument = Nothing

Me.Show vbModless
End Sub

Private Sub chkDisableSpecificDiatHole_Change()

If Me.chkDisableSpecificDiatHole.Value = True Then
    Me.txtLowerLimit.Locked = False
    Me.txtLowerLimit.BackColor = &H80000005
    Me.txtUpperLimit.Locked = False
    Me.txtUpperLimit.BackColor = &H80000005
    
Else
    Me.txtLowerLimit.Locked = True
    Me.txtLowerLimit.BackColor = &H80000016
    Me.txtUpperLimit.Locked = True
    Me.txtUpperLimit.BackColor = &H80000016

End If

Call cmdOKActivation
End Sub

Private Sub cmdFinishPart_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtFinishPart.Text = sNoSelection
    
    On Error Resume Next
    Call selectPart(oFinishPart, oFinishProd, oRootDocument, , "Please select the finish part!   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oFinishPart Is Nothing Then
            MsgBox "Please select the finish part!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    
    'Check that all selection are made from the same window
    If CheckRootDocument = "KO" Then
        MsgBox "All selection must be done from the same window!  ", vbCritical, "Wrong Selection..."
        GoTo noSelection
    End If
    
    'Update form with selected object
    If Not oFinishPart Is Nothing Then
        Me.txtFinishPart.Text = oFinishPart.Name
        Set sTargetWindow = CATIA.ActiveWindow
        
        'Define oFinishTargetPart object
        Set oFinishTargetPart.Product = oFinishProd
        Call oFinishTargetPart.Initialize
        oFinishTargetPart.ParentList = BuildParentList(oFinishProd)

    Else
noSelection:
        Me.txtFinishPart.Text = sNoSelection
        Set oFinishPart = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdOK_Click()

Me.Hide

'Log File
Call AddToLogFile("Create Finish", , txtFinishPart.Value)

Call oCreateFinish.Launch(oFinishTargetPart, oRootDocument, CDbl(Me.txtFinishThk.Text), Me.chkDisableInsertHole.Value, _
                          oInsertHole, Me.chkDisableSpecificDiatHole.Value, CDbl(Me.txtLowerLimit.Text) * 25.4, CDbl(Me.txtUpperLimit.Text) * 25.4)

Me.Show vbModless

End Sub


Private Sub PopulateComboFinishType()

'In the Combo Box we have three columns
' - Finish Type: Laminate, Veneer, ......
' - Thickness value
' - A concatenation of all values
'Only the first one is visible to the user

Dim sFirstItem As String
Dim iRow As Integer


'Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetExcel
Call oExcel.OpenExcelFile(sFinishType)

'Clear All
sFirstItem = ""
comboFinishType.Clear

'Transfer all values to combo
iRow = 2

'Transfer value from Excel file to ComboBox
Do While oExcel.App.Cells(iRow, 1).Value <> ""

    With comboFinishType
        .AddItem oExcel.App.Cells(iRow, 1).Value
        .List(.ListCount - 1, 1) = oExcel.App.Cells(iRow, 2).Value

        .List(.ListCount - 1, 2) = .List(.ListCount - 1, 0) & "|" & _
                                   .List(.ListCount - 1, 1) & "|"
        If sFirstItem = "" Then
            sFirstItem = .List(.ListCount - 1, 2)
        End If
    End With
    
    iRow = iRow + 1
Loop

'Add value "Other" to ComboBox
With comboFinishType
    .AddItem "Other"
    .List(.ListCount - 1, 1) = "1.0"

    .List(.ListCount - 1, 2) = .List(.ListCount - 1, 0) & "|" & _
                               .List(.ListCount - 1, 1) & "|"
    If sFirstItem = "" Then
        sFirstItem = .List(.ListCount - 1, 2)
    End If
End With

comboFinishType.Value = sFirstItem

'Close Excel file
If oExcel.WasOpen = False Then
    oExcel.CloseExcelFile
End If

End Sub

Private Function CheckRootDocument() As String
'Check that all selection are made from the same window

Dim sStatus As String

If oFirstRootDocument Is Nothing Then
    Set oFirstRootDocument = oRootDocument
    sStatus = "OK"
ElseIf oFirstRootDocument.Name <> oRootDocument.Name Then
    sStatus = "KO"
Else
    sStatus = "OK"
End If

CheckRootDocument = sStatus

End Function

Private Sub cmdOKActivation()
    
On Error Resume Next
If Not (Me.txtFinishPart.Text = sNoSelection) Then
    
    If chkDisableSpecificDiatHole.Value = True Then
        If CDbl(Me.txtLowerLimit.Text) <= CDbl(Me.txtUpperLimit.Text) Then
            Me.cmdOK.ForeColor = -2147483630
            Me.cmdOK.Locked = False
            Me.cmdOK.Default = True
'            Me.cmdOK.SetFocus
        Else
            Me.cmdOK.ForeColor = -2147483631
            Me.cmdOK.Locked = True
            Me.cmdCancel.Default = True
'            Me.cmdCancel.SetFocus
        End If
    Else
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
'        Me.cmdOK.SetFocus
    End If
Else
    Me.cmdOK.ForeColor = -2147483631
    Me.cmdOK.Locked = True
    Me.cmdCancel.Default = True
'    Me.cmdCancel.SetFocus
End If
On Error GoTo 0
    

End Sub

Private Sub cmdCancel_Click()
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub resetUserform()
        
With Me
    .txtFinishPart.Text = sNoSelection
End With
     
End Sub

Private Sub comboFinishType_Change()

If Me.comboFinishType.Value <> "" Then

    Me.txtFinishThk.Text = Split(Me.comboFinishType.Value, "|")(1)
    
    If Split(Me.comboFinishType.Value, "|")(0) = "Other" Then
            Me.txtFinishThk.Locked = False
            Me.txtFinishThk.BackColor = &H80000005
    Else
            Me.txtFinishThk.Locked = True
            Me.txtFinishThk.BackColor = &H80000016
    End If
    
End If
End Sub

Private Sub txtLowerLimit_Change()

Call cmdOKActivation

End Sub

Private Sub txtUpperLimit_Change()

Call cmdOKActivation

End Sub

Private Sub txtLowerLimit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'0 to 9
If KeyAscii >= 48 And KeyAscii <= 57 Then
    KeyAscii = KeyAscii

'Only one "."
ElseIf KeyAscii = 46 And Not Me.txtLowerLimit.Text Like "*.*" Then
    KeyAscii = 46
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtUpperLimit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'0 to 9
If KeyAscii >= 48 And KeyAscii <= 57 Then
    KeyAscii = KeyAscii

'Only one "."
ElseIf KeyAscii = 46 And Not Me.txtUpperLimit.Text Like "*.*" Then
    KeyAscii = 46
Else
    KeyAscii = 0
End If
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

