Attribute VB_Name = "frmNewPartTemplate"
Attribute VB_Base = "0{DFB5DBF6-4A61-40C3-8774-1E4FDC96A1B1}{179B02C0-BE1E-40C9-8FC2-C279E39890A5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private oNHAProd As Product
Private oPositionSourceProd As Product
Private oZ0Face As PlanarFace
Private oZ0Prod As Product
Private oZPositiveFace As PlanarFace
Private oZPositiveProd As Product
Private oYAxisFace As Face
Private oYAxisProd As Product
Private oOriginFace As Face
Private oOriginProd As Product
Private oFirstRootDocument As Document
Private oRootDocument As Document
Private oNewPartTemplate As New NewPartTemplate
Private sTemplateList() As String
Private oExcel As clsExcel

Public Sub InitializeForm()
'********************************************************************************
'* Name: Insert New Part
'* Purpose: Add a new part under a selected NHA and positions it based on faces
'*          selected on other bodies.
'*
'* Assumption: The macro will use the BA "TEMPLATE_CATPART.CATPart" as template.
'*
'*
'* Author: Francois Charette
'* Language: VBA
'********************************************************************************

Dim i As Integer
Dim sFirstItem As String

With Me

    .fraNHA.Top = 9
    .fraNHA.Left = 6
    
    .fraPartInfo.Top = .fraNHA.Top + .fraNHA.Height + 4.5
    .fraPartInfo.Left = 6

    .fraPosition.Top = .fraPartInfo.Top + .fraPartInfo.Height + 4.5
    .fraPosition.Left = 6
    
    .Height = dTitleBarHeight + .fraPosition.Top + .fraPosition.Height + 30
    .Width = .fraPartInfo.Left + .fraPartInfo.Width + 10
    
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

'Retrieve data from Excel file to populate sTemplateList
Call PopulateTemplateList

'Initial value
optOnePart.Value = True
optNHAPosition.Value = True

Me.Show vbModless

Set oFirstRootDocument = Nothing

End Sub

Private Sub cmdOK_Click()

Dim sDocType As String 'The associated document type of the selection made in comboTemplate. Can be found in row 2
                       'of the Excel file. Values are "CATPart" or "CATProduct"
Dim iColumn As Integer 'The column number, in the Excel file, representing the document type selected in comboTemplate_X
Dim oNewProduct As Product
Dim oDetailProduct As Product
Dim oCollectorProduct As Product
Dim oDtlAssyProduct As Product
Dim oARMProduct As Product

Dim oAssyProduct As Product
Dim oPart1 As Product
Dim oPart2 As Product
Dim sLogDump As String

Me.Hide

'Log File
If optOnePart.Value = True Then
    sLogDump = "One Part"
ElseIf optFabAssy.Value = True Then
    sLogDump = "Fab Part"
ElseIf optAssy.Value = True Then
    sLogDump = "Assy with Collector"
End If
Call AddToLogFile("Insert New Part", , txtBaseNumber_1 & "-" & txtDashNumber_1, , , , , sLogDump)

'Progress bar
Call frmProgress.progressBarInitialize("Inserting new part/assembly")
Call frmProgress.progressBarRepaint("Inserting new part/assembly")

'**********Insert one part**********
If optOnePart.Value = True Then

    sNewDocType = Split(Split(comboTemplate_1.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_1.Value, "||")(2)

    'Add new part and edit attributes as per the Excel file
    Set oNewProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                             txtBaseNumber_1.Text & "-" & txtDashNumber_1.Text, _
                                                             txtTitle_1.Text, _
                                                             sNewDocType, _
                                                             iColumn, _
                                                             "Open/Close")

    'Position the new part at origin of NHA
    If optNHAPosition.Value = True Then
        Call oNewPartTemplate.PositionPartAtOriginNHA(oNewProduct)

    'Position the new part using 3D geometry
    ElseIf optDefinePosition.Value = True Then
        Call oNewPartTemplate.PositionPartWithGeometry(oNewProduct, oZ0Face, oZ0Prod, oZPositiveFace, oZPositiveProd, _
                                                       oYAxisFace, oYAxisProd, oOriginFace, oOriginProd)

    'Position the new part using position of another part
    Else
        Call oNewPartTemplate.PositionPartToPart(oNewProduct, oNHAProd, oPositionSourceProd)
    End If

'**********Insert a Fab Assy**********
ElseIf optFabAssy = True Then

    '***Insert Detail
    sNewDocType = Split(Split(comboTemplate_3.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_3.Value, "||")(2)

    Set oNewProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                             txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text, _
                                                             txtTitle_3.Text, _
                                                             sNewDocType, _
                                                             iColumn, _
                                                             "Open")

    'The position of the detail will be used as reference to position the Sub Assy and the Collector
    Set oDetailProduct = oNewProduct
    
    'Position Detail at origin of NHA
    If optNHAPosition.Value = True Then
        Call oNewPartTemplate.PositionPartAtOriginNHA(oNewProduct)

    'Position Detail using 3D geometry
    ElseIf optDefinePosition.Value = True Then
        Call oNewPartTemplate.PositionPartWithGeometry(oNewProduct, oZ0Face, oZ0Prod, oZPositiveFace, oZPositiveProd, _
                                                       oYAxisFace, oYAxisProd, oOriginFace, oOriginProd)

    'Position Detail using position of another part
    Else
        Call oNewPartTemplate.PositionPartToPart(oNewProduct, oNHAProd, oPositionSourceProd)
    End If


    '***Insert Sub Assy
    sNewDocType = Split(Split(comboTemplate_1.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_1.Value, "||")(2)

    Set oDtlAssyProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                                 txtBaseNumber_1.Text & "-" & txtDashNumber_1.Text, _
                                                                 txtTitle_1.Text, _
                                                                 sNewDocType, _
                                                                 iColumn, _
                                                                 "NoAction")


    'Position Sub Assy at same position as Detail
    Call oNewPartTemplate.PositionPartToPart(oDtlAssyProduct, oNHAProd, oDetailProduct)

    '***Insert ARM
    sNewDocType = Split(Split(comboTemplate_4.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_4.Value, "||")(2)

    Set oARMProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                                 txtBaseNumber_4.Text & "-" & txtDashNumber_4.Text, _
                                                                 txtTitle_4.Text, _
                                                                 sNewDocType, _
                                                                 iColumn, _
                                                                 "NoAction")


    'Position Sub Assy at same position as Detail
    Call oNewPartTemplate.PositionPartToPart(oARMProduct, oNHAProd, oDetailProduct)


    '***Insert Collector
    sNewDocType = Split(Split(comboTemplate_2.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_2.Value, "||")(2)

    Set oCollectorProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                                   txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text, _
                                                                   txtTitle_2.Text, _
                                                                   sNewDocType, _
                                                                   iColumn, _
                                                                   "Close")
   
    'Position Collector at same position as Detail
    Call oNewPartTemplate.PositionPartToPart(oCollectorProduct, oNHAProd, oDetailProduct)


    'Cut/Paste Detail and Collector in Detail Assy
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    oSelection.Add oCollectorProduct
    oSelection.Add oARMProduct
    oSelection.Add oDetailProduct
    oSelection.Cut
    
    oSelection.Clear
    oSelection.Add oDtlAssyProduct
    oSelection.Paste
    
'**********Insert a Assy**********
ElseIf Me.optAssy = True Then

    '***Insert Assy
    sNewDocType = Split(Split(comboTemplate_1.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_1.Value, "||")(2)

    'Add new part and edit attributes as per the Excel file
    Set oNewProduct = oNewPartTemplate.InsertNewPartTemplate(oNHAProd, _
                                                             txtBaseNumber_1.Text & "-" & txtDashNumber_1.Text, _
                                                             txtTitle_1.Text, _
                                                             sNewDocType, _
                                                             iColumn, _
                                                             "Open")


    'The position of the detail will be used as reference to position the Sub Assy and the Collector
    Set oAssyProduct = oNewProduct
    
    'Position Assy at origin of NHA
    If optNHAPosition.Value = True Then
        Call oNewPartTemplate.PositionPartAtOriginNHA(oNewProduct)

    'Position Assy using position of another part
    Else
        Call oNewPartTemplate.PositionPartToPart(oNewProduct, oNHAProd, oPositionSourceProd)
    End If



    '***Insert Collector
    sNewDocType = Split(Split(comboTemplate_2.Value, "||")(1), "|")(0)
    iColumn = Split(comboTemplate_2.Value, "||")(2)

    Set oNewProduct = oNewPartTemplate.InsertNewPartTemplate(oAssyProduct, _
                                                                   txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text, _
                                                                   txtTitle_2.Text, _
                                                                   sNewDocType, _
                                                                   iColumn, _
                                                                   "Close")
   
    'Position Collector at origin of Assy
    Call oNewPartTemplate.PositionPartAtOriginNHA(oNewProduct)



    
End If
    
frmProgress.Hide
Call ResetUserformReload
Me.Show vbModless

End Sub

Private Sub cmdCancel_Click()
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub



Private Sub optCopyPosition_Change()
    Call ManageButtonsPosition
    Call cmdOKActivation
End Sub

Private Sub optDefinePosition_Change()
    Call ManageButtonsPosition
    Call cmdOKActivation
End Sub

Private Sub optNHAPosition_Change()
    Call ManageButtonsPosition
    Call cmdOKActivation
End Sub

Private Sub ManageButtonsPosition()

If optNHAPosition.Value = True Then
    cmdZ0Surface.Enabled = False
    cmdZPositiveSurface.Enabled = False
    cmdSurfaceYAxis.Enabled = False
    cmdSurfaceOrigin.Enabled = False
    Me.txtZ0Surface.Text = sNoSelection
    Me.txtZPositiveSurface.Text = sNoSelection
    Me.txtSurfaceYAxis.Text = sNoSelection
    Me.txtSurfaceOrigin.Text = sNoSelection
    cmdSourcePart.Enabled = False
    Me.txtSourcePart.Text = sNoSelection
ElseIf optCopyPosition.Value = True Then
    cmdZ0Surface.Enabled = False
    cmdZPositiveSurface.Enabled = False
    cmdSurfaceYAxis.Enabled = False
    cmdSurfaceOrigin.Enabled = False
    Me.txtZ0Surface.Text = sNoSelection
    Me.txtZPositiveSurface.Text = sNoSelection
    Me.txtSurfaceYAxis.Text = sNoSelection
    Me.txtSurfaceOrigin.Text = sNoSelection
    cmdSourcePart.Enabled = True
Else
    cmdZ0Surface.Enabled = True
    cmdZPositiveSurface.Enabled = True
    cmdSurfaceYAxis.Enabled = True
    cmdSurfaceOrigin.Enabled = True
    cmdSourcePart.Enabled = False
    Me.txtSourcePart.Text = sNoSelection
End If

End Sub

Private Sub optOnePart_Click()

Dim oControl As Control

For Each oControl In fraPartInfo.Controls
    
    'Combo Box
    If Left(oControl.Name, 5) = "combo" Then
        If Right(oControl.Name, 2) = "_1" Then
            Call PopulateComboTemplate(oControl, "All", "Populate") 'Populate comboTemplate_1
            oControl.Locked = False
            oControl.BackColor = &H80000005
        Else
            Call PopulateComboTemplate(oControl, "All", "Clear") 'Clear comboTemplate_2, 3 and 4
            oControl.Locked = True
            oControl.BackColor = &H80000016
        End If
    
    'Text Box: set background color and lock
    ElseIf Left(oControl.Name, 3) = "txt" Then
        If Right(oControl.Name, 2) = "_1" Then
            oControl.Locked = False
            oControl.BackColor = &H80000005
        Else
            oControl.Locked = True
            oControl.BackColor = &H80000016
            oControl.Text = ""
        End If
    End If
Next
    
Call cmdOKActivation

End Sub

Private Sub optFabAssy_Click()

Dim oControl As Control

For Each oControl In fraPartInfo.Controls
    
    'Lock and BackColor
    If Left(oControl.Name, 5) = "combo" Or Left(oControl.Name, 3) = "txt" Then
        oControl.Locked = False
        oControl.BackColor = &H80000005
    End If
    
    'Combo Box
    If Left(oControl.Name, 5) = "combo" Then
        If Right(oControl.Name, 2) = "_1" Then
            Call PopulateComboTemplate(oControl, "DETAIl/Fab Assy", "Populate") 'Populate comboTemplate_1
        ElseIf Right(oControl.Name, 2) = "_2" Then
            Call PopulateComboTemplate(oControl, "Collector", "Populate") 'Populate comboTemplate_2
        ElseIf Right(oControl.Name, 2) = "_3" Then
            Call PopulateComboTemplate(oControl, "CATPart|DETAIL", "Populate") 'Populate comboTemplate_3
        Else
            Call PopulateComboTemplate(oControl, "Skeleton - Finish, ARM", "Populate") 'Populate comboTemplate_4
        End If
    End If
Next

Call cmdOKActivation

End Sub

Private Sub optAssy_Click()

Dim oControl As Control

For Each oControl In fraPartInfo.Controls
    
    'Combo Box
    If Left(oControl.Name, 5) = "combo" Then
        If Right(oControl.Name, 2) = "_1" Then
            Call PopulateComboTemplate(oControl, "CATProduct|NONE", "Populate") 'Populate comboTemplate_1
            oControl.Locked = False
            oControl.BackColor = &H80000005
        ElseIf Right(oControl.Name, 2) = "_2" Then
            Call PopulateComboTemplate(oControl, "Collector", "Populate") 'Populate comboTemplate_2
            oControl.Locked = False
            oControl.BackColor = &H80000005
        Else
            Call PopulateComboTemplate(oControl, "All", "Clear") 'Clear comboTemplate_3 and 4
            oControl.Locked = True
            oControl.BackColor = &H80000016
        End If
    
    'Text Box: set background color and lock
    ElseIf Left(oControl.Name, 3) = "txt" Then
        If Right(oControl.Name, 2) = "_1" Or Right(oControl.Name, 2) = "_2" Then
            oControl.Locked = False
            oControl.BackColor = &H80000005
        Else
            oControl.Locked = True
            oControl.BackColor = &H80000016
            oControl.Text = ""
        End If
    End If
Next
    
Call cmdOKActivation
End Sub
Private Sub PopulateComboTemplate(ByRef oCombo As Control, ByVal sFilter As String, ByVal sAction As String)

'In the Combo Box we have four columns
' - Template Type: Panel, Assy, Collector.....
' - Document Type: CATPart or CATProduct
' - Excel column: The Excel column where the data is found
' - A concatenation of all values
'Only the first one is visible to the user

Dim sFirstItem As String


sFirstItem = ""
oCombo.Clear

If sAction = "Clear" Then Exit Sub

For i = 0 To UBound(sTemplateList, 2)
    
    If UCase(sTemplateList(0, i)) = UCase(sFilter) Or UCase(sTemplateList(1, i)) = UCase(sFilter) Or sFilter = "All" Then
        With oCombo
            .AddItem sTemplateList(0, i)
            .List(.ListCount - 1, 1) = sTemplateList(1, i)
            .List(.ListCount - 1, 2) = sTemplateList(2, i)
            .List(.ListCount - 1, 3) = .List(.ListCount - 1, 0) & "||" & .List(.ListCount - 1, 1) & "||" & .List(.ListCount - 1, 2)
            If sFirstItem = "" Then
                sFirstItem = .List(.ListCount - 1, 3)
            End If
        End With
    End If
Next

oCombo.Value = sFirstItem

End Sub

Private Sub txtBaseNumber_1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtDashNumber_1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtBaseNumber_2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtDashNumber_2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtBaseNumber_3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtDashNumber_3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtBaseNumber_4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Sub txtDashNumber_4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Only acceptable character are: a-z, A-Z and 0-9
    KeyAscii = KeyValid(KeyAscii)
End Sub

Private Function KeyValid(ByVal iKey As Integer) As String

'0 to 9
If iKey >= 48 And iKey <= 57 Then
    KeyValid = iKey
    
'A to Z
ElseIf iKey >= 65 And iKey <= 90 Then
    KeyValid = iKey
    
'a to z
ElseIf iKey >= 97 And iKey <= 122 Then
    KeyValid = iKey - 32
    
'TAB
ElseIf iKey = 11 Then
    KeyValid = iKey

Else
    KeyValid = 0
End If

End Function

Private Sub txtTitle_1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Upper case Title
If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
End If

End Sub

Private Sub txtTitle_2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Upper case Title
If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
End If

End Sub

Private Sub txtTitle_3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Upper case Title
If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
End If

End Sub

Private Sub txtTitle_4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Upper case Title
If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
End If

End Sub


Private Sub txtBaseNumber_1_AfterUpdate()

'Check that the Part Number is not already loaded in session
If txtBaseNumber_1.Text <> "" And txtDashNumber_1.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_1.Text & "-" & txtDashNumber_1.Text)
End If

'Transfer value to txtBaseNumber_2 and txtBaseNumber_3
If optFabAssy.Value = True Then
    txtBaseNumber_2.Text = txtBaseNumber_1.Text
    txtBaseNumber_3.Text = txtBaseNumber_1.Text
    txtBaseNumber_4.Text = txtBaseNumber_1.Text
End If

'Transfer value to txtBaseNumber_2
If Me.optAssy.Value = True Then
    txtBaseNumber_2.Text = txtBaseNumber_1.Text
End If

If txtBaseNumber_2.Text <> "" And txtDashNumber_2.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text)
End If

If txtBaseNumber_3.Text <> "" And txtDashNumber_3.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text)
End If

If txtBaseNumber_4.Text <> "" And txtDashNumber_4.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_4.Text & "-" & txtDashNumber_4.Text)
End If

Call cmdOKActivation

End Sub

Private Sub txtDashNumber_1_AfterUpdate()

Dim sDocType As String

'Check that the Part Number is not already loaded in session
If txtBaseNumber_1.Text <> "" And txtDashNumber_1.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_1.Text & "-" & txtDashNumber_1.Text)
End If

'Check the dash number only when value is numeric
sDocType = Split(Split(comboTemplate_1.Value, "||")(1), "|")(0)
On Error Resume Next
If IsNumeric(CInt(txtDashNumber_1.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_1.Text))
End If
On Error GoTo 0

'Transfer value to txtDashNumber_2 and txtDashNumber_3
If optFabAssy.Value = True Then
    txtDashNumber_2.Text = txtDashNumber_1.Text & "SP01"
    txtDashNumber_4.Text = txtDashNumber_1.Text & "ARM01"

    On Error Resume Next
    If IsNumeric(CInt(txtDashNumber_1.Text)) Then
        txtDashNumber_3.Text = CInt(txtDashNumber_1.Text) + 100
    End If
    On Error GoTo 0
End If

'Transfer value to txtDashNumber_2
If Me.optAssy.Value = True Then
    txtDashNumber_2.Text = txtDashNumber_1.Text & "SP01"
End If

If txtBaseNumber_2.Text <> "" And txtDashNumber_2.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text)
End If

If txtBaseNumber_3.Text <> "" And txtDashNumber_3.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text)
End If

If txtBaseNumber_4.Text <> "" And txtDashNumber_4.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_4.Text & "-" & txtDashNumber_4.Text)
End If

Call cmdOKActivation

End Sub

Private Sub txtBaseNumber_2_AfterUpdate()

'Check that the Part Number is not already loaded in session
If txtBaseNumber_2.Text <> "" And txtDashNumber_2.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text)
End If

Call cmdOKActivation

End Sub

Private Sub txtDashNumber_2_AfterUpdate()

Dim sDocType As String

'Check that the Part Number is not already loaded in session
If txtBaseNumber_2.Text <> "" And txtDashNumber_2.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_2.Text & "-" & txtDashNumber_2.Text)
End If

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_2.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_2.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_2.Text))
End If
On Error GoTo 0

Call cmdOKActivation

End Sub

Private Sub txtBaseNumber_3_AfterUpdate()

'Check that the Part Number is not already loaded in session
If txtBaseNumber_3.Text <> "" And txtDashNumber_3.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text)
End If

Call cmdOKActivation

End Sub

Private Sub txtDashNumber_3_AfterUpdate()

Dim sDocType As String

'Check that the Part Number is not already loaded in session
If txtBaseNumber_3.Text <> "" And txtDashNumber_3.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text)
End If

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_3.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_3.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_3.Text))
End If
On Error GoTo 0

Call cmdOKActivation

End Sub

Private Sub txtBaseNumber_4_AfterUpdate()

'Check that the Part Number is not already loaded in session
If txtBaseNumber_3.Text <> "" And txtDashNumber_3.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_3.Text & "-" & txtDashNumber_3.Text)
End If

Call cmdOKActivation

End Sub

Private Sub txtDashNumber_4_AfterUpdate()

Dim sDocType As String

'Check that the Part Number is not already loaded in session
If txtBaseNumber_4.Text <> "" And txtDashNumber_4.Text <> "" Then
    Call CheckPartNumber(txtBaseNumber_4.Text & "-" & txtDashNumber_4.Text)
End If

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_4.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_4.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_4.Text))
End If
On Error GoTo 0

Call cmdOKActivation

End Sub
Private Sub comboTemplate_1_Change()

Dim sDocType As String

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_1.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_1.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_1.Text))
End If
On Error GoTo 0

'Refer to Class NewPartTemplate.PositionPartWithGeometry
'When selecting the option "Define Position" the macro will first create a new part and position it at the origin.
'It will then extract the selected surfaces and calculate the final origin of this new part. The extracted surfaces must be
'stored in a geo set of a CATPart. For this reason I can't use the "Define Position" option to create a CATProduct. I would
'need a dummy CATPart but i'm too lazy to code it !!!!!! So I will just block it.

'I can use the "Define Position" option to create a Fab Assy because I'm first creating the CATPart and using it to
'store the extracted surfaces
If comboTemplate_1.Value <> "Null" Then
    If sDocType = "CATPart" Or Me.optFabAssy.Value = True Then
        optDefinePosition.Enabled = True
    Else
        optDefinePosition.Enabled = False
        If optDefinePosition.Value = True Then
            Me.optNHAPosition.Value = True
        End If
    End If
End If

End Sub

Private Sub comboTemplate_2_Change()

Dim sDocType As String

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_2.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_2.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_2.Text))
End If
On Error GoTo 0

End Sub

Private Sub comboTemplate_3_Change()


Dim sDocType As String

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_3.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_3.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_3.Text))
End If
On Error GoTo 0

End Sub

Private Sub comboTemplate_4_Change()

Dim sDocType As String

'Check the dash number only when value is numeric
On Error Resume Next
sDocType = Split(Split(comboTemplate_4.Value, "||")(1), "|")(0)
If IsNumeric(CInt(txtDashNumber_4.Text)) Then
    Call CheckDashNumber(sDocType, CInt(txtDashNumber_4.Text))
End If
On Error GoTo 0

End Sub

Private Sub CheckPartNumber(ByVal sPartNumber As String)
'Checks that no other loaded document has the same part number

Dim i As Integer

For i = 1 To CATIA.Documents.Count

    If CATIA.Documents.Item(i).Name Like sPartNumber & "*" Then
        sResponse = MsgBox("The following document is already loaded in CATIA:" & vbCrLf & " - " & CATIA.Documents.Item(i).Name & vbCrLf & vbCrLf & "You may want to change the Part Number", vbExclamation)
        Exit For
    End If
Next

End Sub

Private Sub cmdOKActivation()
    
Dim bActivateOK As Boolean

bActivateOK = False

On Error Resume Next

If txtPartNHA.Text = sNoSelection Then GoTo SkipTo
If optOnePart.Value = True And (txtBaseNumber_1.Text = "" Or txtDashNumber_1 = "") Then GoTo SkipTo
If optFabAssy.Value = True And (txtBaseNumber_1.Text = "" Or txtDashNumber_1 = "" Or _
                                txtBaseNumber_2.Text = "" Or txtDashNumber_2 = "" Or _
                                txtBaseNumber_3.Text = "" Or txtDashNumber_3 = "" Or _
                                txtBaseNumber_4.Text = "" Or txtDashNumber_4 = "") Then GoTo SkipTo
If Me.optAssy.Value = True And (txtBaseNumber_1.Text = "" Or txtDashNumber_1 = "" Or _
                                txtBaseNumber_2.Text = "" Or txtDashNumber_2 = "") Then GoTo SkipTo
                                
If optCopyPosition.Value = True And txtSourcePart.Text = sNoSelection Then GoTo SkipTo
If optDefinePosition.Value = True And (txtZ0Surface.Text = sNoSelection Or _
                                       txtZPositiveSurface.Text = sNoSelection Or _
                                       txtSurfaceYAxis.Text = sNoSelection Or _
                                       txtSurfaceOrigin.Text = sNoSelection) Then GoTo SkipTo

bActivateOK = True

SkipTo:

If bActivateOK = True Then
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

Private Sub cmdPartNHA_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtPartNHA.Text = sNoSelection
    
    On Error Resume Next
    Call selectAssy(oNHAProd, oRootDocument, , "Please select the NHA of the new part!   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oNHAProd Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
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
    If Not oNHAProd Is Nothing Then
        Me.txtPartNHA.Text = oNHAProd.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtPartNHA.Text = sNoSelection
        Set oNHAProd = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdSourcePart_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSourcePart.Text = sNoSelection
    
    On Error Resume Next
    Call SelectInstance(oPositionSourceProd, oRootDocument, "Please select an instance to copy the position from !   ")
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oPositionSourceProd Is Nothing Then
            MsgBox "Please select a product!  ", vbCritical, "Wrong Selection..."
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
    If Not oPositionSourceProd Is Nothing Then
        Me.txtSourcePart.Text = oPositionSourceProd.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtSourcePart.Text = sNoSelection
        Set oPositionSourceProd = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdZ0Surface_Click()


    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtZ0Surface.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oZ0Face, oZ0Prod, oRootDocument, "Please select a planar face representing the Z=0 of the new part!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oZ0Face Is Nothing Then
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
    If Not oZ0Face Is Nothing Then
        Me.txtZ0Surface.Text = oZ0Face.DisplayName
    Else
noSelection:
        Me.txtZ0Surface.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub
Private Sub cmdZPositiveSurface_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtZPositiveSurface.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oZPositiveFace, oZPositiveProd, oRootDocument, "Please select a planar face representing the Z(+) direction of the new part!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oZPositiveFace Is Nothing Then
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
    If Not oZPositiveFace Is Nothing Then
        Me.txtZPositiveSurface.Text = oZPositiveFace.DisplayName
    Else
noSelection:
        Me.txtZPositiveSurface.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdSurfaceYAxis_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSurfaceYAxis.Text = sNoSelection
    
    On Error Resume Next
    Call selectFace(oYAxisFace, oYAxisProd, oRootDocument, "Please select an intersect surface to define the Y axis of the new part!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oYAxisFace Is Nothing Then
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
    If Not oYAxisFace Is Nothing Then
        Me.txtSurfaceYAxis.Text = oYAxisFace.DisplayName
    Else
noSelection:
        Me.txtSurfaceYAxis.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdSurfaceOrigin_Click()

    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSurfaceOrigin.Text = sNoSelection
    
    On Error Resume Next
    Call selectFace(oOriginFace, oOriginProd, oRootDocument, "Please select an intersect surface to define the origin of the new part!   ")

    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oOriginFace Is Nothing Then
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
    If Not oOriginFace Is Nothing Then
        Me.txtSurfaceOrigin.Text = oOriginFace.DisplayName
    Else
noSelection:
        Me.txtSurfaceOrigin.Text = sNoSelection
    End If
    On Error GoTo 0
    
      ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If
End Sub

Private Function CheckRootDocument() As String
'Check that all selection are made from the same window

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

Private Sub PopulateTemplateList()

Dim iColumn As Integer
Dim i As Integer

ReDim sTemplateList(2, 0)

'Open Excel file
Set oExcel = New clsExcel
Call oExcel.GetNewExcel
oExcel.App.ScreenUpdating = False
oExcel.App.DisplayAlerts = False
Call oExcel.OpenExcelFile(sAttributesTemplateFile)

'Transfer data from Excel file to sTemplateList()
iColumn = 3
i = 0
Do While oExcel.App.Cells(1, iColumn).Value <> ""
    If oExcel.App.Cells(2, iColumn).Value Like "CATP*" Then
        If sTemplateList(0, i) = "" Then
            sTemplateList(0, i) = oExcel.App.Cells(1, iColumn).Value
            sTemplateList(1, i) = oExcel.App.Cells(2, iColumn).Value
            sTemplateList(2, i) = iColumn

        Else
            i = i + 1
            ReDim Preserve sTemplateList(2, i)
            sTemplateList(0, i) = oExcel.App.Cells(1, iColumn).Value
            sTemplateList(1, i) = oExcel.App.Cells(2, iColumn).Value
            sTemplateList(2, i) = iColumn
        End If
        
    End If
    iColumn = iColumn + 1
Loop


'Close Excel workbook and kill process
Call oExcel.CloseAlreadyOpenWorkbook(sAttributesTemplateFile)
Set oExcel = Nothing

End Sub
Private Sub ResetUserformReload()

With Me
    .txtBaseNumber_1.Text = ""
    .txtDashNumber_1.Text = ""
    .txtBaseNumber_2.Text = ""
    .txtDashNumber_2.Text = ""
    .txtBaseNumber_3.Text = ""
    .txtDashNumber_3.Text = ""
    .txtBaseNumber_4.Text = ""
    .txtDashNumber_4.Text = ""
    .txtTitle_1.Text = ""
    .txtTitle_2.Text = ""
    .txtTitle_3.Text = ""
    .txtTitle_4.Text = ""
    .optNHAPosition.Value = True
End With

Call cmdOKActivation

End Sub
Private Sub resetUserform()
        
With Me
    .txtPartNHA.Text = sNoSelection
    .txtZ0Surface.Text = sNoSelection
    .txtZPositiveSurface.Text = sNoSelection
    .txtSurfaceYAxis.Text = sNoSelection
    .txtSurfaceOrigin.Text = sNoSelection
End With
     
End Sub


Private Sub CheckDashNumber(ByVal sDocType As String, ByVal iDash As Integer)
'Check validity of dash number vs the selected type: detail or assy

If iDash > 99 And sDocType = "CATProduct" Then
    sResponse = MsgBox("Dash number of an assembly should be smaller than 100", vbExclamation)
End If

If iDash < 100 And sDocType = "CATPart" Then
    sResponse = MsgBox("Dash number of a detail should be greater than 100", vbExclamation)
End If


End Sub



'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub



