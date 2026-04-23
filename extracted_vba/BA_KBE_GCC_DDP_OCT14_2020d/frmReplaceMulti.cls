Attribute VB_Name = "frmReplaceMulti"
Attribute VB_Base = "0{B95D0937-075F-4043-B33A-CCBA48EAE00E}{D68FA6F5-0B24-4625-8914-C498CC63BBC7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSelCollection As New Collection

Public Sub InitializeForm()

With Me
    .Height = dTitleBarHeight + 318
    .Width = 208
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
    Call RetrieveSelection
    Call UpdateListBox
    
    .Show vbModless
End With


End Sub

Private Sub RetrieveSelection()

'Transfer the selected instances in the collection
Set oSelCollection = Nothing
For i = 1 To oSelection.Count
    oSelCollection.Add oSelection.Item(i).Value
Next

End Sub

Private Sub UpdateListBox()

Dim i As Integer
Dim oDocument As Document

ListBox_Documents.Clear

For i = 1 To CATIA.Documents.Count

    Set oDocument = CATIA.Documents.Item(i)
    
    If oDocument.Name Like "*.CATProduct" Or oDocument.Name Like "*.CATPart" Then
        
        If txtFilter.Text = "" Then
            ListBox_Documents.AddItem oDocument.Name
        Else
            
            If oDocument.Name Like txtFilter.Text & "*" Then
                ListBox_Documents.AddItem oDocument.Name
            End If
        End If
    End If
Next

End Sub

Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdApply_Click()
    Call UpdateListBox
End Sub


Private Sub cmdOKActivation()
    
On Error Resume Next
If Not (Me.ListBox_Documents.Value = "") Then
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

Private Sub cmdOK_Click()

Dim i As Integer
Dim oParent As Product
Dim sResult As String

Me.Hide

'Log File
Call AddToLogFile("Replace Multiple Components")


'Progress bar
Call frmProgress.progressBarInitialize("Replacing Part")
  
For i = 1 To oSelCollection.Count

    Call frmProgress.progressBarRepaint("Replacing Part", oSelCollection.Count, i)

    'Find parent instance
    Set oParent = oSelCollection.Item(i).Parent

    'Replace
    Set Nothing1 = oParent.Products.ReplaceProduct(oSelCollection.Item(i), CATIA.Documents.Item(ListBox_Documents.Value).Product, False)
    
    Call oRenameInstance.RenameSelectedInstance(oSelCollection.Item(i))
    
Next

frmProgress.Hide

sResult = MsgBox("If the replaced instances are located under a component" & vbCrLf & "you may need to run the BAReorderStdPart macro.", vbExclamation)

Call frmKBEMain.resetToolbar

End Sub

Private Sub ListBox_Documents_Change()
    Call cmdOKActivation
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub



