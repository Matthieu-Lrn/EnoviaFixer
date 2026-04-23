Attribute VB_Name = "frmSeedPartsList"
Attribute VB_Base = "0{10BE6A96-1050-485E-B759-AEAB074FCBE6}{B82B9C15-7244-48FF-8DFD-B33FBF82F662}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private sPreviousSeedPartsList As String

Public Sub ListBoxInitialize(ByRef oSeedAssy As Product, Optional ByVal positionLeft As Double = -1, Optional ByVal positionTop As Double = -1, Optional ByVal bConfigPart As Boolean = False)
    
    Me.seedPartsListBox.Clear
    sPreviousSeedPartsList = sSeedPartsList
    Call ExtractSeedParts(bConfigPart)
    
    If sSeedPartsList = "" Then '***** No Seed Part found under Seed Assy
        MsgBox "At least one (1) Seed Part must be in Design mode!      ", vbExclamation, "Seed Parts not found..."
        Exit Sub
    End If
    
    '***** Window size according to number of seed parts
    Me.seedPartsListBox.Height = Me.seedPartsListBox.ListCount * 11 + 5
    If Me.seedPartsListBox.Height > 550 Then Me.seedPartsListBox.Height = 550
    
    'Resize ListBox to see last element
    Me.seedPartsListBox.IntegralHeight = False
    Me.seedPartsListBox.IntegralHeight = True
    
    Me.Height = Me.seedPartsListBox.Height + dTitleBarHeight + 42
    Me.imgDSGlobe.Top = Me.Height - dTitleBarHeight - 28
    Me.imgDSGlobe.Left = -3
    Me.cmdOK.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdCancel.Top = Me.Height - dTitleBarHeight - 24
    Me.cmdCancel.SetFocus
    
    Me.Left = positionLeft
    If Me.Left < 0 Or Me.Left + Me.Width > dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) Then Me.Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - Me.Width - 85
    
    Me.Top = positionTop
    If Me.Top < 0 Or Me.Top + Me.Height > dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) Then Me.Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - Me.Height - 85
    
    '**** Highlight previously selected seed parts
    For i = 0 To Me.seedPartsListBox.ListCount - 1
        If InStr(1, sPreviousSeedPartsList, Me.seedPartsListBox.List(i), vbTextCompare) > 0 Then
            Me.seedPartsListBox.Selected(i) = True
        End If
    Next i
    Me.Show
    
    sSeedPartsList = ""
    '***** sSeedPartsList = selected parts in List Box
    For i = 0 To Me.seedPartsListBox.ListCount - 1
        If Me.seedPartsListBox.Selected(i) Then sSeedPartsList = sSeedPartsList & Me.seedPartsListBox.List(i) & ";"
    Next i
    
    '***** if nothing selected > sSeedPartsList = all parts in List Box
    If sSeedPartsList = "" Then
        For i = 0 To Me.seedPartsListBox.ListCount - 1
            sSeedPartsList = sSeedPartsList & Me.seedPartsListBox.List(i) & ";"
        Next i
    End If
    
End Sub

Private Sub ExtractSeedParts(Optional ByVal bConfigPart As Boolean = False)
'********************************************************************************
'* Name: ExtractSeedParts
'* Purpose:
'*
'* Assumption:
'*
'* Author:
'* Updated by:Abhishek Kamboj
'* Language: VBA
'********************************************************************************
    Dim bInstActivity As Boolean
    sSeedPartsList = ""
    Set oSeedAssyToolVersion = Nothing
    Set oSeedAssyConfigName = Nothing
    For Each oProduct In oSeedAssy.Products
        
        On Error Resume Next
        sSeedPartName = ""
        sSeedPartName = oProduct.PartNumber
        
        If sSeedPartName Like "KBE*_SEED_PART_*" Or ((sSeedPartName Like "*SEED_OPTION*" Or sSeedPartName Like "*SEED_CONFIG*") And bConfigPart) Then
            
            bInstActivity = False
            If Not TypeName(oProduct.GetMasterShapeRepresentation(False)) = "Document" Then '**** used to check if instance is in visualization mode
                    bInstActivity = GetInstanceActivity(oProduct)
            End If
            
            If bInstActivity Then               '***** Consider only parts that are in "DESIGN" mode and are active
               If oSeedAssyToolVersion Is Nothing Then
                    '**** Two names of parameter set to cover: "TemplateParameters" & "ParametersTemplate"
                    Set oSeedAssyToolVersion = oProduct.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("TemplateParameters").DirectParameters.Item("ToolVersion_DT")
                    Set oSeedAssyToolVersion = oProduct.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("ParametersTemplate").DirectParameters.Item("ToolVersion_DT")
                    sSeedAssyToolVersion = ""
                    sSeedAssyToolVersion = oSeedAssyToolVersion.ValueAsString
                End If
                
                If oSeedAssyConfigName Is Nothing Then
                    '**** Two names of parameter set to cover: "TemplateParameters" & "ParametersTemplate"
                    Set oSeedAssyConfigName = oProduct.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("TemplateParameters").DirectParameters.Item("ConfigurationName_DT")
                    Set oSeedAssyConfigName = oProduct.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("ParametersTemplate").DirectParameters.Item("ConfigurationName_DT")
                    sSeedAssyConfigName = ""
                    sSeedAssyConfigName = oSeedAssyConfigName.ValueAsString
                End If
                
                If sSeedPartName Like "KBE*_SEED_PART_*" Then
                    sSeedPartsList = sSeedPartsList & "SEED_PART_" & Split(Split(oProduct.Name, "SEED_PART_")(1), "-")(0) & ";"
                    Me.seedPartsListBox.AddItem "SEED_PART_" & Split(Split(oProduct.Name, "SEED_PART_")(1), "-")(0)
                ElseIf sSeedPartName Like "KBE*_SEED_OPTION*" Then
                    sSeedPartsList = sSeedPartsList & "SEED_OPTION" & ";"
                    Me.seedPartsListBox.AddItem "SEED_OPTION"
                ElseIf sSeedPartName Like "KBE*_SEED_CONFIG*" Then
                    sSeedPartsList = sSeedPartsList & "SEED_CONFIG" & ";"
                    Me.seedPartsListBox.AddItem "SEED_CONFIG"
                End If
            End If
        End If
        
        On Error GoTo 0
    Next
    
End Sub

Private Sub cmdCancel_Click()
    Me.seedPartsListBox.Clear
    bCancelAction = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub


'********************************************************************************
'* Name: Populate Seed Part collection
'* Purpose:
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************

Public Function PopulateSeedPartCollection() As Collection
'This Sub will populate the oSeedPartCollection with the seleted Seed Part

    Dim sString As String
    Set PopulateSeedPartCollection = New Collection
    
    'Loop all text from the sSeedPartsList
    For i = 0 To UBound(Split(sSeedPartsList, ";")) - 1
    
        'Retrieve text(i)
        sString = Split(sSeedPartsList, ";")(i)
    
        'Loop instances under the Seed Assy
        For Each oProduct In oSeedAssy.Products
            If oProduct.Name Like "*_" & sString & "-*" Then
                PopulateSeedPartCollection.Add oProduct, oProduct.PartNumber
                Exit For
            End If
        Next
    Next
    
End Function


