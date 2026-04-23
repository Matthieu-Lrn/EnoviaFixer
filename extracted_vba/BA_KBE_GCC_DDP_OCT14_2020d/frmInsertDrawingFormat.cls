Attribute VB_Name = "frmInsertDrawingFormat"
Attribute VB_Base = "0{3A256ACB-953B-4FCC-9B5E-4E72C5E73DA9}{FEC25575-25C0-4D9E-9ED5-693CE049857E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Sub InitializeForm()
'********************************************************************************
'********************************************************************************

With Me
    .Height = dTitleBarHeight + 396
    .Width = 438
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

Call PopulateComboBox

Me.Show vbModless

End Sub

Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub PopulateComboBox()

Dim oComboBox As ComboBox
Dim i As Integer
Dim sNumber As String

For i = 1 To 20
    If i < 10 Then
        sNumber = "0" & i
    Else
        sNumber = i
    End If
    
    Set oComboBox = Me.Controls.Item("SH" & sNumber & "_Combo")

    oComboBox.AddItem " "
    oComboBox.AddItem "EO/NIEO BT0204"
    oComboBox.AddItem "NIEO BT0209"
    oComboBox.AddItem "A Portrait"
    oComboBox.AddItem "A Landscape"
    oComboBox.AddItem "B"
    oComboBox.AddItem "C"
    oComboBox.AddItem "D"
    oComboBox.AddItem "E"
    oComboBox.AddItem "J66"
    oComboBox.AddItem "J88"
    oComboBox.AddItem "J110"
    oComboBox.AddItem "J132"
    oComboBox.AddItem "J154"
    oComboBox.AddItem "J176"
    oComboBox.AddItem "Clear BKGD View"
    oComboBox.ListRows = 16
Next

For i = 1 To 10
    
    If i < 10 Then
        sNumber = "0" & i
    Else
        sNumber = i
    End If
    
    Set oComboBox = Controls.Item("PL" & sNumber & "_Combo")

    oComboBox.AddItem " "
    oComboBox.AddItem "A Landscape"
    oComboBox.AddItem "Clear BKGD View"
    oComboBox.ListRows = 3
Next


'Populate comboCATDrawing
Template_ComboBox.AddItem " "
For i = 1 To CATIA.Documents.Count
    If CATIA.Documents.Item(i).Name Like "*CATDrawing" And CATIA.Documents.Item(i).Name <> CATIA.ActiveDocument.Name Then
        Template_ComboBox.AddItem CATIA.Documents.Item(i).Name
    End If
Next



End Sub

Private Sub cmdOK_Click()


Dim i As Integer
Dim oComboBox As ComboBox
Dim sTemplateSheetName(30) As String
Dim dPageWidth(30) As Double
Dim dPageHeight(30) As Double
Dim oTemplateDoc As DrawingDocument

Me.Hide

'Log File
Call AddToLogFile("Insert Drawing Format", CATIA.ActiveDocument.Name)


For i = 1 To 20
    If i < 10 Then
        sNumber = "0" & i
    Else
        sNumber = i
    End If
    
    Set oComboBox = Controls.Item("SH" & sNumber & "_Combo")

    Select Case oComboBox.Value
    
    Case ""
        dPageWidth(i) = 0
        dPageHeight(i) = 0

    Case "EO/NIEO BT0204"
        dPageWidth(i) = (11 * 25.4)
        dPageHeight(i) = (8.5 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-EO/NIEO.BT0204-01"
        Else
            sTemplateSheetName(i) = "SH02-EO/NIEO.BT0204-02"
        End If


    Case "NIEO BT0209"
        dPageWidth(i) = (11 * 25.4)
        dPageHeight(i) = (8.5 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-NIEO.BT0209-01"
        Else
            sTemplateSheetName(i) = "SH02-NIEO.BT0209-02"
        End If

    Case "A Portrait"
        dPageWidth(i) = (8.5 * 25.4)
        dPageHeight(i) = (11 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-A-PORTRAIT.BT0500-1"
        Else
            sTemplateSheetName(i) = "SH02-A-PORTRAIT.BT0500-21"
        End If
        
    Case "A Landscape"
        dPageWidth(i) = (11 * 25.4)
        dPageHeight(i) = (8.5 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-A-LANDSCAPE.BT0500-12"
        Else
            sTemplateSheetName(i) = "SH02-A-LANDSCAPE.BT0500-32"
        End If
  
    Case "B"
        dPageWidth(i) = (17 * 25.4)
        dPageHeight(i) = (11 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-B.BT0500-2"
        Else
            sTemplateSheetName(i) = "SH02-B.BT0500-22"
        End If
        
    Case "C"
        dPageWidth(i) = (22 * 25.4)
        dPageHeight(i) = (17 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-C.BT0500-3"
        Else
            sTemplateSheetName(i) = "SH02-C.BT0500-23"
        End If

    Case "D"
        dPageWidth(i) = (34 * 25.4)
        dPageHeight(i) = (22 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-D.BT0500-4"
        Else
            sTemplateSheetName(i) = "SH02-D.BT0500-24"
        End If
    
    Case "E"
        dPageWidth(i) = (44 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-E.BT0500-5"
        Else
            sTemplateSheetName(i) = "SH02-E.BT0500-25"
        End If
    
    Case "J66"
        dPageWidth(i) = (66 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J66.BT0500-6"
        Else
            sTemplateSheetName(i) = "SH02-J66.BT0500-26"
        End If
        
    Case "J88"
        dPageWidth(i) = (88 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J88.BT0500-7"
        Else
            sTemplateSheetName(i) = "SH02-J88.BT0500-27"
        End If
    
  Case "J110"
        dPageWidth(i) = (110 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J110.BT0500-8"
        Else
            sTemplateSheetName(i) = "SH02-J110.BT0500-28"
        End If

  Case "J132"
        dPageWidth(i) = (132 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J132.BT0500-9"
        Else
            sTemplateSheetName(i) = "SH02-J132.BT0500-29"
        End If

  Case "J154"
        dPageWidth(i) = (154 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J154.BT0500-10"
        Else
            sTemplateSheetName(i) = "SH02-J154.BT0500-30"
        End If

  Case "J176"
        dPageWidth(i) = (176 * 25.4)
        dPageHeight(i) = (34 * 25.4)
        If i = 1 Then
            sTemplateSheetName(i) = "SH01-J176.BT0500-11"
        Else
            sTemplateSheetName(i) = "SH02-J176.BT0500-31"
        End If

  Case "Clear BKGD View"
        dPageWidth(i) = 1
        dPageHeight(i) = 1

  End Select

Next

For i = 21 To 30
    

    If i < 30 Then
        sNumber = "0" & i - 20
    Else
        sNumber = i - 20
    End If
    
    Set oComboBox = Controls.Item("PL" & sNumber & "_Combo")

    Select Case oComboBox.Value
    
    Case ""
        dPageWidth(i) = 0
        dPageHeight(i) = 0

    Case "A Landscape"
        dPageWidth(i) = (11 * 25.4)
        dPageHeight(i) = (8.5 * 25.4)
        If i = 21 Then
            sTemplateSheetName(i) = "PL01.BT0501-1"
        Else
            sTemplateSheetName(i) = "PL02.BT0501-2"
        End If
    
    Case "Clear BKGD View"
        dPageWidth(i) = 1
        dPageHeight(i) = 1

    End Select
Next

Set oTemplateDoc = CATIA.Documents.Item(Me.Template_ComboBox.Value)

Call oInsertDrawingFormat.Launch(dPageWidth, dPageHeight, sTemplateSheetName, oTemplateDoc)

Call frmKBEMain.resetToolbar

End Sub

Private Sub Template_ComboBox_Change()

Call cmdOKActivation

End Sub

Private Sub cmdOKActivation()
    
On Error Resume Next
If Not Me.Template_ComboBox.Value = " " Then
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

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

