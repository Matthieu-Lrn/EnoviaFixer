Attribute VB_Name = "frmSelectAttributeFamily"
Attribute VB_Base = "0{E322BB93-DCC0-49D5-BCC0-A6239D4E9B22}{81BFE5FD-4858-4FB5-8832-007D6EC36D47}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public DocName As String
Public DocTitle As String
Public DocumentAttributFamilyType As String

Private Sub ComboBoxSelect_Change()
   DocumentAttributFamilyType = Me.ComboBoxSelect.Text
   'Unload
   Me.Hide
End Sub

Private Sub UserForm_Activate()
    With Me
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width / 2) - .Width / 2
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height / 2) - .Height / 2
    End With
    lblPartName.Caption = DocName & vbCrLf & DocTitle
    ComboBoxSelect.DropDown
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If Me.ComboBoxSelect.Text = "" Then Cancel = 1
   DocumentAttributFamilyType = Me.ComboBoxSelect.Text
End Sub

