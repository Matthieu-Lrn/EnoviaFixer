Attribute VB_Name = "frmCreateFlatPanel"
Attribute VB_Base = "0{3567FFF9-5CCE-41B3-A7BE-5171B641AA64}{669FCDE9-B8D5-496E-9F53-2B9CC50684A2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oPanelProd As Product
Private oPanelPart As Part
Private oSketchFace As PlanarFace
Private oSketchProd As Product
Private oThkFace As PlanarFace
Private oThkProd As Product
Private oFirstRootDocument As Document
Private oRootDocument As Document
Private oCreateFlatPanel As New CreateFlatPanel

Public Sub InitializeForm()
'********************************************************************************
'* Name: Create Flat Panel
'* Purpose: Creates a pad based on a face selected on another body (in a different part).
'*
'* Assumption: The Z axis of the part where the pad will be created must be normal
'*             to the selected face
'*
'* Author: Francois Charette
'* Language: VBA
'********************************************************************************

With Me
    .Height = dTitleBarHeight + 230
    .Width = 190
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
End With

Set oFirstRootDocument = Nothing

Me.Show vbModless

End Sub

Private Sub cmdCancel_Click()
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
    'Log File
    Call AddToLogFile("Create Flat Panel", , txtFlatPanelPart.Value)
   
    'Progress bar
    Call frmProgress.progressBarInitialize("Creating flat panel")
    Call frmProgress.progressBarRepaint("Creating flat panel")

    CATIA.RefreshDisplay = False
    Call oCreateFlatPanel.LaunchCreationFlatPanel(oRootDocument, oPanelProd, oSketchFace, oSketchProd, oThkFace, oThkProd, chkCreateConstraints.Value)
    CATIA.RefreshDisplay = True

    frmProgress.Hide
    
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub resetUserform()
        
With Me
    .txtFlatPanelPart.Text = sNoSelection
    .txtSketchSurface.Text = sNoSelection
    .txtThkSurface.Text = sNoSelection
End With
     
End Sub

Private Sub cmdOKActivation()
    
On Error Resume Next
If Not (Me.txtFlatPanelPart.Text = sNoSelection Or Me.txtSketchSurface.Text = sNoSelection Or Me.txtThkSurface.Text = sNoSelection) Then
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

Private Sub cmdFlatPanelPart_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtFlatPanelPart.Text = sNoSelection
    
    On Error Resume Next
    Call selectPart(oPanelPart, oPanelProd, oRootDocument, , "Please select the panel part!   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oPanelPart Is Nothing Then
            MsgBox "Please select a panel part!  ", vbCritical, "Wrong Selection..."
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
    
    'Check that PartBody is empty
    If CheckPartBody(oPanelPart) = "KO" Then
        MsgBox "Partbody must be empty!  ", vbCritical, "Wrong Selection..."
        GoTo noSelection
    End If
    
    'Update form with selected object
    If Not oPanelPart Is Nothing Then
        Me.txtFlatPanelPart.Text = oPanelPart.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtFlatPanelPart.Text = sNoSelection
        Set oPanelPart = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdSketchSurface_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSketchSurface.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oSketchFace, oSketchProd, oRootDocument, "Please select the sketch planar face of the new flat panel!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSketchFace Is Nothing Then
            MsgBox "Please select a planar face!  ", vbCritical, "Wrong Selection..."
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
    If Not oSketchFace Is Nothing Then
        Me.txtSketchSurface.Text = oSketchFace.DisplayName
    Else
noSelection:
        Me.txtSketchSurface.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If
  

End Sub

Private Sub cmdThkSurface_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtThkSurface.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oThkFace, oThkProd, oRootDocument, "Please select the thickness planar face of the new flat panel!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oThkFace Is Nothing Then
            MsgBox "Please select a planar face!  ", vbCritical, "Wrong Selection..."
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
    If Not txtThkSurface Is Nothing Then
        Me.txtThkSurface.Text = oThkFace.DisplayName
    Else
noSelection:
        Me.txtThkSurface.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Function CheckRootDocument() As String

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

Private Function CheckPartBody(ByVal oPart As Part) As String

Dim sStatus As String
Dim oPartBody As Body

sStatus = "OK"
Set oPartBody = oPart.Bodies.Item("PartBody")

If oPartBody.Shapes.Count <> 0 Then sStatus = "KO"

CheckPartBody = sStatus

End Function


'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

