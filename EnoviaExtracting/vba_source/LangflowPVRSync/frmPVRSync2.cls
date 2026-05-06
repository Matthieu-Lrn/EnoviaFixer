Attribute VB_Name = "frmPVRSync2"
Attribute VB_Base = "0{A498B1B3-384D-4C6E-A44F-BA9823CB355E}{850DA623-A6DB-469B-93A8-F9CAC12C613D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Option Explicit
Public Project As String
Public SyncWithBSF As Boolean
Public Cancel As Boolean
Public CIOption As String
Public NonCIOption As String

Public Sub InitializeForm(ByVal sPartNumber As String)

Me.Cancel = False
Me.optBSF.Value = True
Me.lblTitle.Caption = Replace(Me.lblTitle.Caption, "XXX", sPartNumber)
Me.Show

End Sub

Private Sub cmdOK_Click()

Me.Project = txtProject.Text
Me.Hide

'CI option
If Me.optCIAlways.Value = True Then
    CIOption = "DisplayAlways"
ElseIf Me.optCINonRel.Value = True Then
    CIOption = "DisplayNonRel"
ElseIf Me.optCINever.Value = True Then
    CIOption = "DisplayNever"
Else
    CIOption = "N/A"
End If

'Non CI option
If Me.optNonCILatestRel.Value = True Then
    NonCIOption = "LatestRel"
ElseIf Me.optNonCIBSF.Value = True Then
    NonCIOption = "BSF"
Else
    NonCIOption = "N/A"
End If

End Sub

Private Sub optBSF_Click()

Me.SyncWithBSF = True

Me.txtProject.Text = "SXXXX"
Me.txtProject.Enabled = False

Me.optCIAlways.Value = False
Me.optCINever.Value = False
Me.optCINonRel.Value = False

Me.optCIAlways.Enabled = False
Me.optCINever.Enabled = False
Me.optCINonRel.Enabled = False

Me.optNonCIBSF.Value = False
Me.optNonCILatestRel.Value = False

Me.optNonCIBSF.Enabled = False
Me.optNonCILatestRel.Enabled = False

End Sub

Private Sub optTail_Click()

Me.txtProject.Enabled = True
Me.SyncWithBSF = False

Me.optCIAlways.Enabled = True
Me.optCINever.Enabled = True
Me.optCINonRel.Enabled = True

Me.optCIAlways.Value = True

Me.optNonCIBSF.Enabled = True
Me.optNonCILatestRel.Enabled = True

Me.optNonCILatestRel.Value = True
End Sub

Private Sub cmdCancel_Click()

Me.Cancel = True
Me.Hide

End Sub
Private Sub txtProject_Enter()

If txtProject.Text = "SXXXX" Then
    txtProject.Text = ""
End If

End Sub


