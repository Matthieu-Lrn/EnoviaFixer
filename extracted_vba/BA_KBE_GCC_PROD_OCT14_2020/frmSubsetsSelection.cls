Attribute VB_Name = "frmSubsetsSelection"
Attribute VB_Base = "0{829D78A9-84DA-48AA-981C-D5B9DD6D1FB5}{EBABF660-3F55-4532-8302-D52CC5B6B3B0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private i As Long

Public Sub FormInitialize(ByVal positionLeft As Double, ByVal positionTop As Double)
    
    If SSParameters.Count > 0 Then
        
        Me.Left = positionLeft
        Me.Top = positionTop
        Me.SSListBox.Clear
        
        '**** Fill listbox
        For Each iObject In SSParameters
            Me.SSListBox.AddItem iObject.ParameterName
        Next
        
        '**** Highlight first SS
        Me.SSListBox.Selected(0) = True
        
        '***** Window size according to number of seed parts
        Me.SSListBox.Height = Me.SSListBox.ListCount * 11 + 5
        Me.cbInvertSelection.Top = Me.SSListBox.Top + Me.SSListBox.Height + 5
        Me.cbInvertSelection.Left = Me.SSListBox.Left
        Me.Height = Me.SSListBox.Height + dTitleBarHeight + 62
        Me.imgDSGlobe.Top = Me.Height - dTitleBarHeight - 28
        Me.imgDSGlobe.Left = -3
        Me.cmdCancel.Top = Me.Height - dTitleBarHeight - 24
        Me.cmdCancel.Left = Me.Width - Me.cmdCancel.Width - 12
        Me.cmdOK.Top = Me.Height - dTitleBarHeight - 24
        Me.cmdOK.Left = Me.Width - Me.cmdOK.Width - Me.cmdCancel.Width - 18
        Me.cmdCancel.SetFocus
        Me.Show
    Else '***** No Subset found
        MsgBox "No Subset found!" & vbCrLf & "Please make sure you have a least one SSParameter, or    " & vbCrLf & "uncheck ""Subsets"" before launching application!     ", vbExclamation, "Subset not found..."
        sCurrentSSParameterName = ""
        bCancelAction = True
        Exit Sub
    End If

End Sub

Private Sub cmdCancel_Click()
    sCurrentSSParameterName = ""
    bCancelAction = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
'********************************************************************************
'* Name: cmdOK_Click
'*
'*Original: User was able to select only one selection set from the list box
'*     New: User can select multiple subsets from the list box & selected subsets
'*          will be synchronised / created in 3D / SP collector.
'* Author:
'* Updated by: Abhishek Kamboj
'* Language: VBA
'******************************************************************************
    
    Me.Hide
    sCurrentSSParameterName = "Temporary"
    bCancelAction = True
    Set SelectedSSParameters = New Collection
    Dim oTempSubSetParameter As New SSParameter
    oTempSubSetParameter.ParameterName = sCurrentSSParameterName
    For i = 0 To Me.SSListBox.ListCount - 1
        If Me.SSListBox.Selected(i) Then
            SelectedSSParameters.Add Me.SSListBox.List(i), Me.SSListBox.List(i)
            If SSParameters.Item(Me.SSListBox.List(i)).SubsetValues.Count > 0 Then
                For Each sSubSetItem In SSParameters.Item(Me.SSListBox.List(i)).SubsetValues
                    oTempSubSetParameter.AddToSubsetValues sSubSetItem
                Next
                bCancelAction = False
            End If
        End If
    Next i

    SSParameters.Add oTempSubSetParameter, sCurrentSSParameterName
    If Not (oTempSubSetParameter.SubsetValues.Count > 0) Then
        MsgBox "Selected Subset is empty!" & vbCrLf & "Please make sure to valuate SSParameter, or uncheck    " & vbCrLf & """Subsets"" before launching 3D Collector!     ", vbExclamation, "Subset empty..."
    End If
End Sub

Private Sub SSListBox_Change()
    
    '***** To avoid unselect all elements
    For i = 0 To Me.SSListBox.ListCount - 1
        If Me.SSListBox.Selected(i) Then Exit Sub
    Next i

    Me.SSListBox.Selected(Me.SSListBox.ListIndex) = True
    
End Sub

