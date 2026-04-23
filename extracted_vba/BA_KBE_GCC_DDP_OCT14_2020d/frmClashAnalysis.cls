Attribute VB_Name = "frmClashAnalysis"
Attribute VB_Base = "0{971CE3BA-689D-4A97-A34A-74CBD032F40D}{814481C8-530A-4BC1-B37B-F4367158C1BE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oClashAnalysis As New ClashAnalysis
Private iMinClearance As Double
Public cTableList As New Collection
Public cTargetAssys As New Collection
Private oClash As ClashAnalysis
Private oExcel As clsExcel

Public Sub initializeClashAnalysisForm()

    Call enableRelationUpdate

    With Me
        '***** Window size according to number of shifters
        .Height = 280
        .Width = 250
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdCancel.SetFocus
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        iMinClearance = Val(.TextBox1.Value)

        .Show vbModeless
    End With

End Sub

Private Sub cmdSelectTargetSA_Click()

    Me.Hide
    Me.txtTargetSA.Text = sNoSelection


    Dim Document, HybridShapeFactory, HybridBodies, HybridBody, OriginElements, Plane, PlaneReference, Status
    Dim Selection 'As Selection
    Dim InputObjectType(0), PointIndex, PointReference, HybridShapeSymmetry
    Dim MyDoc As ProductDocument
    Set MyDoc = CATIA.ActiveDocument
    Set Sel = MyDoc.Selection
    
    InputObjectType(0) = "Product"
    Status = Sel.SelectElement3(InputObjectType, "Select Products", True, CATMultiSelTriggWhenUserValidatesSelection, False)
    If (Status = "Cancel") Then Exit Sub
    
    Me.txtTargetSA.Text = Str(Sel.Count) + " Product(s) Selected"
    
    
    For i = 1 To Sel.Count
        cTargetAssys.Add Sel.Item(i).Value
    Next
    Sel.Clear
    Me.Show vbModeless

End Sub

Private Sub cmdOK_Click()

Dim testType As String
Dim Item As Product

Me.Hide

'Log File
Call AddToLogFile("Clash Analysis")

   
CATIA.DisplayFileAlerts = False
CATIA.DisplayFileAlerts = True
CATIA.Interactive = True
CATIA.RefreshDisplay = True


'Clear cache
Call StartWebServiceTool
WebServiceAccessTool.ClearCache

If OptionButtonSelectTargetSA.Value = True Then


    If cTargetAssys.Count = 0 Or iMinClearance < 0 Then
        GoTo noSelection
    End If

    Call frmProgress.progressBarInitialize("Performing clash analysis")
        
    If ClashNoHarware = True Then
        
        Call oClashAnalysis.ClashAnalysisWithoutHarware(cTargetAssys, HardwareClash)
        
    End If
    
    If HardwareClash = True Then
        
        Call oClashAnalysis.Pins_InsertsClashAnalysis(cTargetAssys, ClashNoHarware, iMinClearance)
        
    End If
    
    
    
    Call oClashAnalysis.CompileClash
    
    Dim sStringKeep As String
      
        
    For Each iObject In UserForms
        If Not iObject.Name Like "*KBEMain*" And Not iObject.Name Like "*frmClashAnalysis*" Then Unload iObject
    Next
    
    frmVisualizeClash.Show vbModeless
    Call frmVisualizeClash.MakeFormResizable
End If


If OptionButtonSelectTargetExcel.Value = True Then
    Call oClashAnalysis.ImportFromToExcel(oExcel, cTableList)
    frmVisualizeClash.Show vbModeless
End If
    
    'Call frmKBEMain.resetToolbar
    
noSelection:


End Sub

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdSelectTargetExcel_Click()

    Dim strFileToOpen As String
    Set oExcel = New clsExcel
    
    Me.txtTargetExcel.Text = sNoSelection
    
    Call oExcel.GetExcel
    Call oExcel.OpenExcelFile(, "Please choose a file to open")
    
    Me.txtTargetExcel.Text = oExcel.App.Workbooks.Item(1).Name
    
    
End Sub

Private Sub OptionButtonSelectTargetSA_Click()

    cmdSelectTargetSA.Enabled = True
    cmdSelectTargetExcel.Enabled = False
    
    txtTargetSA.BackColor = &H80000005
    txtTargetExcel.BackColor = &H80000003
    
End Sub
Private Sub OptionButtonSelectTargetExcel_Click()

    cmdSelectTargetSA.Enabled = False
    cmdSelectTargetExcel.Enabled = True
    
    txtTargetSA.BackColor = &H80000003
    txtTargetExcel.BackColor = &H80000005

End Sub

Private Sub PannelClashCheck_Click()
Dim panelClashCheck As Boolean
    panelClashCheck = OptionButton1

End Sub

Private Sub TextBox1_Change()

Dim sMinClearance As String


    sMinClearance = TextBox1.Value
    If sMinClearance <> "" Then
        iMinClearance = Val(sMinClearance)
    Else
        iMinClearance = -1
    End If

End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Public Function BuildParentListWithRefProducts(ByVal iObject As Variant) As String

    'This is a function similar to Build Parent List but we build Parent list with Reference product name, instead of instance name
    BuildParentListWithRefProducts = iObject.Name & ";"

    Do
        Set iObject = iObject.Parent

        'We don't care about the "Products" or "ProductDocuments" objects
        If TypeName(iObject) = "Product" Then
            BuildParentListWithRefProducts = BuildParentListWithRefProducts & iObject.ReferenceProduct.Name & ";"
        End If
    Loop Until TypeName(iObject) = "Application"

End Function




