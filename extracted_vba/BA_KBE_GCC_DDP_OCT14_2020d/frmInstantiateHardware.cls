Attribute VB_Name = "frmInstantiateHardware"
Attribute VB_Base = "0{6619C3C7-55D0-4427-97C2-EABCB2DBBDDD}{A7F78FA6-71EF-4400-AD34-775BDC18B8D2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSPInsert As Product
Private oSPOther As Product
Private oSPRootDocument As Document
Private oHardware1 As Product
Private oHardware2 As Product
Private oHardware3 As Product
Private oInstantiateHardware As New InstantiateHardware
Private sInsertsList As String
Private oExcel As clsExcel

Public Sub InitializeForm()
'********************************************************************************
'********************************************************************************

Dim i As Integer
Dim sFirstItem As String

With Me
    .Height = dTitleBarHeight + 430
    .Width = 309
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdExit.Top = .Height - dTitleBarHeight - 24
    .cmdExit.Left = .Width - .cmdExit.Width - 12
    .cmdInstantiate.Top = .Height - dTitleBarHeight - 24
    .cmdInstantiate.Left = .Width - .cmdInstantiate.Width - .cmdExit.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
End With

Set oHardware1 = Nothing
Set oHardware2 = Nothing
Set oHardware3 = Nothing
Set oSPRootDocument = Nothing

Call cmdInstantiateActivation
Call PopulateInsertsList

Me.Show vbModless

End Sub

Private Sub cmdExit_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdInstantiate_Click()

Me.Hide

'Log File
Call AddToLogFile("Instantiate Hardware")


Call oInstantiateHardware.Launch(oSPOther, oSPInsert, oHardware1, oHardware2, oHardware3, oSPRootDocument, sInsertsList)

Me.Show vbModless

End Sub

Private Sub cmdSPInsertSelect_Click()

Dim oRootDocument As Document

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.Hide
    Me.txtSPInsert.Text = sNoSelection
    
    On Error Resume Next
    Call selectAssy(oSPInsert, oRootDocument, , "Please select the SP Collector")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSPInsert Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'The two SP must be in the same window
    If oSPRootDocument Is Nothing Then
        Set oSPRootDocument = oRootDocument
    Else
        If oSPRootDocument.Name <> oRootDocument.Name Then
            MsgBox "All SPs must in the same window!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    End If
    
            
    'Update form with selected object
    If Not oSPInsert Is Nothing Then
        Me.txtSPInsert.Text = oSPInsert.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtSPInsert.Text = sNoSelection
        Set oSPInsert = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdInstantiateActivation
    End If
    
End Sub

Private Sub cmdInsertClear_Click()

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.txtSPInsert.Text = sNoSelection
    Set oSPInsert = Nothing
    
End Sub

Private Sub cmdSPOtherSelect_Click()

Dim oRootDocument As Document

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.Hide
    Me.txtSPOther.Text = sNoSelection
    
    On Error Resume Next
    Call selectAssy(oSPOther, oRootDocument, , "Please select the SP Collector")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSPOther Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'The two SP must be in the same window
    If oSPRootDocument Is Nothing Then
        Set oSPRootDocument = oRootDocument
    Else
        If oSPRootDocument.Name <> oRootDocument.Name Then
            MsgBox "All SPs must in the same window!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    End If
    
            
    'Update form with selected object
    If Not oSPOther Is Nothing Then
        Me.txtSPOther.Text = oSPOther.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtSPOther.Text = sNoSelection
        Set oSPOther = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdInstantiateActivation
    End If
    
    Stop
End Sub

Private Sub cmdHardware1Select_Click()

Dim oRootDocument As Document

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.Hide
    Me.txtHardware1.Text = sNoSelection
    
    On Error Resume Next
    Call SelectInstance(oHardware1, oRootDocument, "Please select an instance of Hardware 1   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oHardware1 Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'Update form with selected object
    If Not oHardware1 Is Nothing Then
        Me.txtHardware1.Text = oHardware1.PartNumber
    Else
noSelection:
        Me.txtHardware1.Text = sNoSelection
        Set oHardware1 = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdInstantiateActivation
    End If
        
End Sub

Private Sub cmdHardware1Clear_Click()

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.txtHardware1.Text = sNoSelection
    Set oHardware1 = Nothing
    
    Call cmdInstantiateActivation
      
End Sub

Private Sub cmdHardware2Select_Click()

Dim oRootDocument As Document

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.Hide
    Me.txtHardware2.Text = sNoSelection
    
    On Error Resume Next
    Call SelectInstance(oHardware2, oRootDocument, "Please select an instance of Hardware 2  ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oHardware2 Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'Update form with selected object
    If Not oHardware2 Is Nothing Then
        Me.txtHardware2.Text = oHardware2.PartNumber
    Else
noSelection:
        Me.txtHardware2.Text = sNoSelection
        Set oHardware2 = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdInstantiateActivation
    End If
        
End Sub

Private Sub cmdHardware2Clear_Click()

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.txtHardware2.Text = sNoSelection
    Set oHardware2 = Nothing
    
    Call cmdInstantiateActivation
      
End Sub

Private Sub cmdHardware3Select_Click()

Dim oRootDocument As Document

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.Hide
    Me.txtHardware3.Text = sNoSelection
    
    On Error Resume Next
    Call SelectInstance(oHardware3, oRootDocument, "Please select an instance of Hardware 3     ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oHardware3 Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
            
    'Update form with selected object
    If Not oHardware3 Is Nothing Then
        Me.txtHardware3.Text = oHardware3.PartNumber
    Else
noSelection:
        Me.txtHardware3.Text = sNoSelection
        Set oHardware3 = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdInstantiateActivation
    End If
        
End Sub

Private Sub cmdHardware3Clear_Click()

    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
    Me.txtHardware3.Text = sNoSelection
    Set oHardware3 = Nothing
    
   Call cmdInstantiateActivation
   
End Sub

Private Sub cmdInstantiateActivation()
    
On Error Resume Next
If Not (Me.txtSPOther.Text = sNoSelection) And Not (Me.txtHardware1 = sNoSelection And Me.txtHardware2 = sNoSelection And Me.txtHardware3 = sNoSelection) Then
    Me.cmdInstantiate.ForeColor = -2147483630
    Me.cmdInstantiate.Locked = False
    Me.cmdInstantiate.Default = True
    Me.cmdInstantiate.SetFocus
Else
    Me.cmdInstantiate.ForeColor = -2147483631
    Me.cmdInstantiate.Locked = True
    Me.cmdExit.Default = True
    Me.cmdExit.SetFocus
End If
On Error GoTo 0
    
End Sub


Private Sub PopulateInsertsList()

Dim iColumn As Integer
Dim iRow As Integer

'Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetNewExcel
Call oExcel.OpenExcelFile(sInsertHoleDefinition)

'Initial set
sInsertsList = ""

'Transfer data from Excel file to sFinishTypeList()
iColumn = 1
iRow = 2
Do While oExcel.App.Cells(iRow, iColumn).Value <> ""
    
    If iRow = 2 Then
        sInsertsList = oExcel.App.Cells(iRow, iColumn).Value
    Else
        sInsertsList = sInsertsList & "|" & oExcel.App.Cells(iRow, iColumn).Value
    End If
    
    iRow = iRow + 1
Loop

'Close Excel workbook and kill process
Call oExcel.CloseAlreadyOpenWorkbook(sInsertHoleDefinition)
Set oExcel = Nothing

End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdExit_Click
End Sub

