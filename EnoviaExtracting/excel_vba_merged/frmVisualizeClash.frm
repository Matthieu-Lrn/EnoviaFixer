Attribute VB_Name = "frmVisualizeClash"
Attribute VB_Base = "0{688D213C-9EFB-4F78-BC48-E7D4E77888FD}{1366E790-308E-4093-9C8F-B2AB4521C754}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long
Private Declare Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) _
  As Long
               
Private Declare Function SetWindowLong _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) _
  As Long

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

Private oClashAnalysis As New ClashAnalysis
Public aTableList



Private Sub cmdCancel_Click()

Call oClashAnalysis.CleanAssy

Call frmKBEMain.resetToolbar

End Sub

Private Sub cmdOK_Click()

    
    Dim firstPart As String
    Dim secondPart As String
    Dim clashValue As Double
    Dim firstPath As String
    Dim secondPath As String
    Dim firstClashPoint As String
    Dim secondClashPoint As String
    
    

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then

            firstPart = ListBox1.Column(1, i)
            secondPart = ListBox1.Column(2, i)
            clashValue = ListBox1.Column(3, i)
            firstPath = ListBox1.Column(5, i)
            secondPath = ListBox1.Column(6, i)
            firstClashPoint = ListBox1.Column(7, i)
            secondClashPoint = ListBox1.Column(8, i)
        End If
    Next i
    
    partList = Split(firstPath, ";")
    

    Set oParent = CATIA.ActiveDocument.Product
    
    On Error Resume Next
    
    For i = UBound(partList) - 1 To 1 Step -1
        If oParent.Name <> partList(i) Then
            MsgBox ("Cant find Part, Please review your inputs")
            Exit Sub
        End If
        
        Set oParent = oParent.Products.Item(partList(i - 1))
        
        
    Next
    
    On Error GoTo 0
    
    
    Call oClashAnalysis.DefineClash(firstPart, secondPart, clashValue, firstPath, secondPath, firstClashPoint, secondClashPoint)


End Sub

Private Sub ExpExcel_Click()
    
    Call oClashAnalysis.ExportToExcel(frmClashAnalysis.cTableList)

End Sub


Private Sub UserForm_Initialize()

Dim aTableList
Dim a As Integer

a = frmClashAnalysis.cTableList.Count
   Dim columnNoWidth, column1stPartWidth, column2ndPartWidth, columnClashValueWidth, columnNotesWidth As Integer
   
Dim columnHeaders(0, 4) As Variant
columnHeaders(0, 0) = "#"
columnHeaders(0, 1) = "1st Part"
columnHeaders(0, 2) = "2nd Part"
columnHeaders(0, 3) = "Clash Value"
columnHeaders(0, 4) = "Notes"


With Me
    .Width = 408
    .Height = 318
    .imgDSGlobe.Top = .Height - dTitleBarHeight - 37
    .imgDSGlobe.Left = -3
    .cmdCancel.Left = Me.Width - 66
    .cmdCancel.Top = Me.Height - 54
    .ExpExcel.Top = Me.Height - 54
    .ExpExcel.Left = 30
    .cmdOK.Top = Me.Height - 54
    .cmdOK.Left = Me.Width / 2 - cmdOK.Width / 2
    
    
    With .ListBox2
        .Height = 20
        columnNoWidth = 20
        column1stPartWidth = (.Width - columnNoWidth) / 4
        column2ndPartWidth = (.Width - columnNoWidth) / 4
        columnClashValueWidth = (.Width - columnNoWidth) / 4
        columnNotesWidth = .Width - columnClashValueWidth - column2ndPartWidth - column1stPartWidth - columnNoWidth - 5
        .ColumnCount = 9
        .List = columnHeaders
        .ColumnWidths = columnNoWidth & ";" & column1stPartWidth & ";" & column2ndPartWidth & ";" & columnClashValueWidth & ";" & columnNotesWidth & ";0;0;0;0"
    End With

End With


If a > 0 Then
    With ListBox1
    
        .ColumnCount = 5
        columnNoWidth = 20
        column1stPartWidth = (.Width - columnNoWidth) / 4
        column2ndPartWidth = (.Width - columnNoWidth) / 4
        columnClashValueWidth = (.Width - columnNoWidth) / 4
        columnNotesWidth = .Width - columnClashValueWidth - column2ndPartWidth - column1stPartWidth - columnNoWidth - 5
        .ColumnWidths = columnNoWidth & ";" & column1stPartWidth & ";" & column2ndPartWidth & ";" & columnClashValueWidth & ";" & columnNotesWidth & ";0;0;0;0"
        
        For j = 1 To frmClashAnalysis.cTableList.Count
            .AddItem j
            .List(ListBox1.ListCount - 1, 1) = frmClashAnalysis.cTableList.Item(j)(0)
            .List(ListBox1.ListCount - 1, 2) = frmClashAnalysis.cTableList.Item(j)(1)
            .List(ListBox1.ListCount - 1, 3) = Format(frmClashAnalysis.cTableList.Item(j)(2), "###0.0000")
            .List(ListBox1.ListCount - 1, 4) = frmClashAnalysis.cTableList.Item(j)(7)
            .List(ListBox1.ListCount - 1, 5) = frmClashAnalysis.cTableList.Item(j)(3)
            .List(ListBox1.ListCount - 1, 6) = frmClashAnalysis.cTableList.Item(j)(4)
            .List(ListBox1.ListCount - 1, 7) = frmClashAnalysis.cTableList.Item(j)(5)
            .List(ListBox1.ListCount - 1, 8) = frmClashAnalysis.cTableList.Item(j)(6)
                    
        Next j
    End With
        

End If



End Sub
Public Sub MakeFormResizable()

  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal
  
    hWnd = GetForegroundWindow
  
    'Get the basic window style
    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME

    'Set the basic window styles
    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub
Private Sub UserForm_Resize()

    With Me
    
        If .Width < 204 Then .Width = 204
        If .Height < 100 Then .Height = 100
        
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 37
        .imgDSGlobe.Left = -3
        .ListBox1.Left = 6
        .ListBox2.Left = 6
        .ListBox1.Width = .Width - .ListBox1.Left - 18
        .ListBox2.Width = .Width - .ListBox2.Left - 18
        .ListBox1.Height = Me.Height - 120
        .ListBox2.Height = 20
        
        columnNoWidth = 20
        column1stPartWidth = (.Width - columnNoWidth) / 4
        column2ndPartWidth = (.Width - columnNoWidth) / 4
        columnClashValueWidth = (.Width - columnNoWidth) / 4
        columnNotesWidth = .Width - columnClashValueWidth - column2ndPartWidth - column1stPartWidth - columnNoWidth - 30
        .ListBox1.ColumnWidths = columnNoWidth & ";" & column1stPartWidth & ";" & column2ndPartWidth & ";" & columnClashValueWidth & ";" & columnNotesWidth & ";0;0;0;0"
        .ListBox2.ColumnWidths = columnNoWidth & ";" & column1stPartWidth & ";" & column2ndPartWidth & ";" & columnClashValueWidth & ";" & columnNotesWidth & ";0;0;0;0"


        .cmdCancel.Left = Me.Width - 66
        .cmdCancel.Top = Me.Height - 54
        .ExpExcel.Top = Me.Height - 54
        .ExpExcel.Left = 30
        .cmdOK.Top = Me.Height - 54
        .cmdOK.Left = Me.Width / 2 - cmdOK.Width / 2
    End With
End Sub
