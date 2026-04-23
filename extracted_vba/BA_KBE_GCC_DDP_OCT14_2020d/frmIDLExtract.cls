Attribute VB_Name = "frmIDLExtract"
Attribute VB_Base = "0{F1AE1C22-846E-4AF4-8EC4-DC8B71B777C1}{F9D13810-FF25-4F03-8561-1EC5858CC82F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub cmdSelect_Click()
Call oIDLExtract.HightlightInCATIA
End Sub

Private Sub cmdSelectExcel_Click()
Call oIDLExtract.HightlightInExcel
End Sub


'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set oIDLExtract = Nothing
    If CloseMode = 0 Then Call frmKBEMain.resetToolbar
End Sub

