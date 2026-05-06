Attribute VB_Name = "frmPVRSync3"
Attribute VB_Base = "0{1AEBC485-957F-4F8D-9815-95472E2E0B18}{D4D170B9-8927-468A-8ED9-76597BBEBA24}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Option Explicit

Public Answer As String
Private Const Manual As Integer = 0
Private sReportPath As String

Public Sub InitializeForm(ByVal sProposedDwgAs As String, ByVal sProposedPVR As String, ByVal sLoadedPVR As String, ByVal sPath As String, ByVal sTail As String)

Dim sText As String

sText = "For "
sText = sText & sTail
sText = sText & " the effective document is "
sText = sText & sProposedDwgAs
sText = sText & " and the related PVR is "
sText = sText & sProposedPVR
sText = sText & vbCrLf
sText = sText & "Do you want to keep using "
sText = sText & sLoadedPVR
sText = sText & " ?"


'Initial set
sReportPath = sPath

With Me
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
    .lblText.Caption = sText
    .cmdYes.SetFocus
    .Show
End With

End Sub



Private Sub cmdHelp_Click()

Dim lSuccess

lSuccess = ShellExecute(0, "Open", sReportPath, , , 1)

End Sub

Private Sub cmdNo_Click()

Answer = "No"
Me.Hide
End Sub

Private Sub cmdYes_Click()

Answer = "Yes"
Me.Hide
End Sub
