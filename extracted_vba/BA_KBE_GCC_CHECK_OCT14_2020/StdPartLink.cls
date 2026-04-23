Attribute VB_Name = "StdPartLink"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private oSeedPartCollection As New Collection
Private oCoincidenceConstraintCollection As New Collection
Private oBoolParamCollection As New Collection
Private oAxisCollection As New Collection
Private oActivityRuleCollection As New Collection
Private oExcel As clsExcel

'*************************************************************************
'* The goal of this class is to check the link between different objects:
'*  - Coincidence Constraints
'*  - Std Parts
'*  - SP/SA Boolean parameters
'*  - Axis
'*************************************************************************
Public Sub SelectSeedAssy()
    
    On Error Resume Next
    If oSeedAssy Is Nothing Then
        Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    Else
        If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    End If
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If (oSeedAssy Is Nothing Or oSeedAssy.Products.Count = 0) Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical, "No Seed Parts found..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy)
    
    If sSeedPartsList <> "" Then
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else
noSelection:
        Set oSeedAssy = Nothing
        Set sSourceWindow = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set oSeedPartCollection = New Collection
        Exit Sub
    End If
    
    Call RetrieveLinkedObjects
    
End Sub

Private Sub RetrieveLinkedObjects()
    
    Dim sLog As String
    Dim oSeedPartInstance As Product
    Dim iProgressCount As Integer
    Dim oConstraint As Constraint
    Dim iExcelRow As Integer
    Dim i As Integer
    Dim oReference1 As Reference
    Dim oReference2 As Reference
    Dim sCstElemSeedPart As String
    Dim sCstElemStdPart As String
    Dim oStdPart As Product
    Dim sCstNewName As String
    Dim oObject As AnyObject
    Dim sString1 As String
    Dim sString2 As String
    Dim bError As Boolean
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    Set oBoolParamCollection = Nothing
    Set oAxisCollection = Nothing
    Set oActivityRuleCollection = Nothing
    Set oCoincidenceConstraintCollection = Nothing
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Retrieve Std Part linked Objects")
    
    'Put Std Part in Design Mode
    Call frmProgress.progressBarRepaint("Put Std Part in Design Mode", 4, 1)
    Call PutStdPartInDesignMode(oSeedAssy)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
    'Retrieve Coincidence Constraints
    Call frmProgress.progressBarRepaint("Retrieving Coincidence Constraints", 4, 2)
    Call PopulateConstraintCollection
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
    'We loop the selected Seed Part to Retrieve the axis, boolean parameters and Activity rules and populate the collections
    For Each oSeedPartInstance In oSeedPartCollection
                   
        iProgressCount = iProgressCount + 1
        Call frmProgress.progressBarRepaint("Retrieving Boolean Parameters and Axis", 4, 3, oSeedPartInstance.PartNumber, oSeedPartCollection.Count, iProgressCount)
             
        'Populate the oBoolParamCollection
        bError = False
        Call PopulateBoolParamCollection(bError, oSeedPartInstance)
        If bError = True Then GoTo endSub
           
        'Populate the oAxisCollection
        bError = False
        Call PopulateAxisCollection(bError, oSeedPartInstance)
        If bError = True Then GoTo endSub
        
        'Populate the oActivityRuleCollection
        Call PopulateActivityRuleCollection(oSeedPartInstance)
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Open Excel
    Set oExcel = New clsExcel
    Call oExcel.GetExcel
    Call oExcel.NewExcelFile
    
    'Excel File Header
    oExcel.App.Cells(1, 1).Value = "Coincidence Constraint(before running the macro)"
    oExcel.App.Cells(1, 2).Value = "Coincidence Constraint(after running the macro)"
    oExcel.App.Cells(1, 3).Value = "Std Part Instance Name"
    oExcel.App.Cells(1, 4).Value = "Std Part Part Number"
    oExcel.App.Cells(1, 5).Value = "Std Part Instance Description(before running the macro)"
    oExcel.App.Cells(1, 6).Value = "Std Part Instance Description(after running the macro)"
    oExcel.App.Cells(1, 7).Value = "Axis Name"
    oExcel.App.Cells(1, 8).Value = "Seed Part"
    oExcel.App.Cells(1, 9).Value = "Boolean Parameter(as found in Activity Rule)"
    oExcel.App.Cells(1, 10).Value = "Check Duplicated Std Part-Axis"
    
    'Loop Coincidence Constraint Collection
    iExcelRow = 1
    Call frmProgress.progressBarRepaint("Retrieving linked elements", 4, 4)
    
    For Each oConstraint In oCoincidenceConstraintCollection
    
        iExcelRow = iExcelRow + 1
    
        'Initial set
        sString1 = ""
        sString2 = ""
        Set oReference1 = Nothing
        Set oReference2 = Nothing
        
        'Retrieve constraint elements
        On Error Resume Next
        Set oReference1 = oConstraint.GetConstraintElement(1)
        Set oReference2 = oConstraint.GetConstraintElement(2)
        sString1 = oReference1.DisplayName
        sString2 = oReference2.DisplayName
        
        If sString1 Like "*SEED_PART*" Then
            sCstElemSeedPart = sString1
            sCstElemStdPart = sString2
        Else
            sCstElemSeedPart = sString2
            sCstElemStdPart = sString1
        End If
    
    
        'Retrieve Std Part
        If sCstElemStdPart <> "" Then
            Set oStdPart = oSeedAssy.Products.Item(Split(sCstElemStdPart, "/")(1))
        Else
            Set oStdPart = Nothing
        End If
        
        'Coincidence Constraint(before running the macro)
        oExcel.App.Cells(iExcelRow, 1).Value = oConstraint.Name
        
        'Std Part Instance Name
        If Not oStdPart Is Nothing Then
            oExcel.App.Cells(iExcelRow, 3).Value = oStdPart.Name
        Else
            oExcel.App.Cells(iExcelRow, 3).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 3).Interior.ColorIndex = 3
        End If
    
        'Std Part Part Number
        If Not oStdPart Is Nothing Then
            oExcel.App.Cells(iExcelRow, 4).Value = oStdPart.PartNumber
        Else
            oExcel.App.Cells(iExcelRow, 4).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 4).Interior.ColorIndex = 3
        End If
    
        'Std Part Instance Description(before running the macro)
        If Not oStdPart Is Nothing Then
            oExcel.App.Cells(iExcelRow, 5).Value = oStdPart.DescriptionInst
        Else
            oExcel.App.Cells(iExcelRow, 5).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 5).Interior.ColorIndex = 3
        End If
    
        'Axis Name
        oExcel.App.Cells(iExcelRow, 7).Value = Split(sCstElemSeedPart, "/!")(1)
        On Error Resume Next
        oAxisCollection.Remove (oExcel.App.Cells(iExcelRow, 7).Value)
        On Error GoTo 0
        
        'Seed Part
        oExcel.App.Cells(iExcelRow, 8).Value = Split(Split(sCstElemSeedPart, "/")(1), ".")(0)
    
        'Boolean Parameter(as found in Activity Rule)
        If Not oStdPart Is Nothing Then
            oExcel.App.Cells(iExcelRow, 9).Value = RetrieveBooleanParameterName(oStdPart.Name, Split(Split(sCstElemSeedPart, "/")(1), ".")(0))
            On Error Resume Next
            oBoolParamCollection.Remove (oExcel.App.Cells(iExcelRow, 9).Value)
            On Error GoTo 0
    
            If oExcel.App.Cells(iExcelRow, 9).Value = "Parameter Not Found" Then
                oExcel.App.Cells(iExcelRow, 9).Interior.ColorIndex = 45
            End If
        Else
            oExcel.App.Cells(iExcelRow, 9).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 9).Interior.ColorIndex = 3
        End If
        
        'Check Duplicated Axis-Std Part
        'All Seed Part Axis / Std Part Part Number must be unique
        If Not oStdPart Is Nothing Then
            oExcel.App.Cells(iExcelRow, 10).Value = oStdPart.PartNumber & "|" & oExcel.App.Cells(iExcelRow, 7).Value
            
            If iExcelRow > 2 Then
                For i = 2 To iExcelRow - 1
                    If oExcel.App.Cells(iExcelRow, 10).Value = oExcel.App.Cells(i, 10).Value Then
                        oExcel.App.Cells(iExcelRow, 10).Interior.ColorIndex = 7
                        Exit For
                    End If
                Next
            End If
        Else
            oExcel.App.Cells(iExcelRow, 10).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 10).Interior.ColorIndex = 3
        End If
        
        'Coincidence Constraint(after running the macro)
        If Not oStdPart Is Nothing Then
            If oConstraint.Name <> oStdPart.Name & " | " & oExcel.App.Cells(iExcelRow, 7).Value & " | " & oExcel.App.Cells(iExcelRow, 8).Value Then
                oConstraint.Name = oStdPart.Name & " | " & oExcel.App.Cells(iExcelRow, 7).Value & " | " & oExcel.App.Cells(iExcelRow, 8).Value
                oExcel.App.Cells(iExcelRow, 2).Value = oStdPart.Name & " | " & oExcel.App.Cells(iExcelRow, 7).Value & " | " & oExcel.App.Cells(iExcelRow, 8).Value
                oExcel.App.Cells(iExcelRow, 2).Interior.ColorIndex = 6
            Else
                oExcel.App.Cells(iExcelRow, 2).Value = "No change"
                oExcel.App.Cells(iExcelRow, 2).Interior.ColorIndex = 50
            End If
        Else
            oExcel.App.Cells(iExcelRow, 2).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 2).Interior.ColorIndex = 3
        End If
        
        'Std Part Instance Description(after running the macro)
        If Not oStdPart Is Nothing Then
        
            If oExcel.App.Cells(iExcelRow, 9).Value = "Parameter Not Found" Then
                oExcel.App.Cells(iExcelRow, 6).Value = "Parameter Not Found"
                oExcel.App.Cells(iExcelRow, 6).Interior.ColorIndex = 45
            ElseIf oStdPart.DescriptionInst <> oExcel.App.Cells(iExcelRow, 9).Value Then
                oStdPart.DescriptionInst = oExcel.App.Cells(iExcelRow, 9).Value
                oExcel.App.Cells(iExcelRow, 6).Value = oExcel.App.Cells(iExcelRow, 9).Value
                oExcel.App.Cells(iExcelRow, 6).Interior.ColorIndex = 6
            Else
                oExcel.App.Cells(iExcelRow, 6).Value = "No change"
                oExcel.App.Cells(iExcelRow, 6).Interior.ColorIndex = 50
            End If
        Else
            oExcel.App.Cells(iExcelRow, 6).Value = "Std Part Not Found"
            oExcel.App.Cells(iExcelRow, 6).Interior.ColorIndex = 3
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
    'Axis not being used by a Std Part
    For Each oObject In oAxisCollection
        iExcelRow = iExcelRow + 1
        oExcel.App.Cells(iExcelRow, 1).Value = "Axis not being used by a constraint: " & oObject.Name
    Next
    
    'Boolean not being used by a Std Part
    For Each oObject In oBoolParamCollection
        iExcelRow = iExcelRow + 1
        oExcel.App.Cells(iExcelRow, 1).Value = "Boolean parameter not being used by an Activity Rule: " & oObject.Name
    Next
    
    Call oExcel.ShowExcelWindow
    
endSub:
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar

End Sub

Private Sub PopulateAxisCollection(ByRef bError As Boolean, oActiveSeedPart As Product)
'This Sub will populate the oAxisCollection
'A Std Part axis is any element that:
'Start with "SP_*" but doesn't end with "_Boolean"
'or
'Start with "SA_*" but doesn't end with "_Boolean"
    
    Dim i As Integer
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("((Name_CAP=SP_* - Name_CAP=*_Boolean) + (Name_CAP=SA_* - Name_CAP=*_Boolean)),sel", oActiveSeedPart, oSelection)
    
    'Transfer all the selected elements in oAxisCollection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
        
            On Error Resume Next
            oAxisCollection.Add oSelection.Item(i).Value, oSelection.Item(i).Value.Name
            
            If Err.Number <> 0 Then
                MsgBox "Two axis are named: " & oSelection.Item(i).Value.Name & _
                       vbCrLf & "Correct error before running the macro"
                bError = True
                Exit Sub
            End If

        Next
    End If
    
End Sub

Private Sub PopulateBoolParamCollection(ByRef bError As Boolean, oActiveSeedPart As Product)
'This Sub will populate the oBoolParamCollection with all the Boolean parameter
    
    Dim i As Integer
    Dim oParam As Parameter
    
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    Set oSelection = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean),sel", oActiveSeedPart, oSelection)
    
    'Transfer Boolean parameter in the collection
    If oSelection.Count > 0 Then
        For i = 1 To oSelection.Count
        
            Set oParam = oSelection.Item(i).Value
            
            On Error Resume Next
            oBoolParamCollection.Add oParam, Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\")))
            
            If Err.Number <> 0 Then
                MsgBox "Two boolean parameters are named: " & Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\"))) & _
                       vbCrLf & "Correct error before running the macro"
                bError = True
                Exit Sub
            End If
        Next
    End If
    
End Sub

Private Sub PopulateActivityRuleCollection(oActiveSeedPart As Product)
'This Sub will populate the oActivityRuleCollection with the corresponding Activity Rule
    
    Dim oRelation As Relation
    
    bRuleFound = False
    
    For Each oRelation In oSeedAssy.Relations
        If TypeName(oRelation) = "Rule" And oRelation.Name = oActiveSeedPart.PartNumber & "_StdPartsActivityRule" Then
            oActivityRuleCollection.Add oRelation, oActiveSeedPart.PartNumber & "_StdPartsActivityRule"
            bRuleFound = True
            Exit For
        End If
    Next
    
End Sub

Private Sub PopulateConstraintCollection()
    
    Dim oConstraints As Constraints
    Dim oConstraint As Constraint
    Dim oSeedPart As Product
    Dim sString1 As String
    Dim sString2 As String
    Dim oReference1 As Reference
    Dim oReference2 As Reference
    Dim sCstElemSeedPart As String
    
    'Initial set
    Set oConstraints = oSeedAssy.Connections("CATIAConstraints")
    
    'Loop all the constraints
    If Not oConstraints.Count = 0 Then
        For Each oConstraint In oConstraints
        
            'We only check the Coincidence constraints
            If oConstraint.Type = 2 Then
                            
                'Initial set
                sString1 = ""
                sString2 = ""
                Set oReference1 = Nothing
                Set oReference2 = Nothing
                
                'Retrieve constraint elements
                On Error Resume Next
                Set oReference1 = oConstraint.GetConstraintElement(1)
                Set oReference2 = oConstraint.GetConstraintElement(2)
                sString1 = oReference1.DisplayName
                sString2 = oReference2.DisplayName
                
                If sString1 Like "*SEED_PART*" Then
                    sCstElemSeedPart = sString1
                Else
                    sCstElemSeedPart = sString2
                End If
    
                'Loop the Seed Part collection
                For Each oSeedPart In oSeedPartCollection
        
                    'We only keep constraints related to selected Seed Part
                    If sCstElemSeedPart Like "*" & oSeedPart.PartNumber & "*" Then
                        oCoincidenceConstraintCollection.Add oConstraint
                        Exit For
                    End If
                Next
                On Error GoTo 0
            End If
        Next
    End If
    
End Sub

Private Sub PutStdPartInDesignMode(ByVal oParentProd As Product)
'This Sub will put all Std Parts in Design Mode
    
    Dim oProduct As Product
    Dim sPartNumber As String
    
    'Loop instances under the Seed Assy
    For Each oProduct In oParentProd.Products
        If Not (oProduct.Name Like "KBE_*" Or oProduct.Name Like "MONUMENTORIGINPART*") Then
                
            sPartNumber = ""
            On Error Resume Next
            sPartNumber = oProduct.PartNumber
            If sPartNumber = "" Then oProduct.ApplyWorkMode DESIGN_MODE
            
            'Recursive Sub for Standard Assemblies
            If TypeName(oProduct.ReferenceProduct.Parent) = "ProductDocument" Then Call PutStdPartInDesignMode(oProduct)
        End If
    Next
    
End Sub

Private Function RetrieveBooleanParameterName(ByVal sStdPartID As String, ByVal sSeedPart As String) As String
    
    Dim oRule As Relation
    Dim sRuleText As String
    Dim sReturnValue As String
    
    'Initial set
    Set oRule = Nothing
    sRuleText = ""
            
    'Retrieve rule from collection
    On Error Resume Next
    Set oRule = oActivityRuleCollection(sSeedPart & "_StdPartsActivityRule")
    sRuleText = oRule.Value
    On Error GoTo 0
            
    'If the rule was found we scan it line by line
    sReturnValue = "Parameter Not Found"
    If Not oRule Is Nothing Then
        For j = 0 To UBound(Split(sRuleText, vbCrLf))
            sOneLine = Split(sRuleText, vbCrLf)(j)
            
            'We look for the Std Part instance name and the Seed Part
            If sOneLine Like "*" & sStdPartID & "*" And sOneLine Like "*" & sSeedPart & "*" Then
                
                sReturnValue = Split(sOneLine, "_Boolean")(0)
                sReturnValue = Mid(sReturnValue, InStrRev(sReturnValue, "\") + 1)
                sReturnValue = sReturnValue & "_Boolean"
        
                Exit For
            End If
        Next
    
    End If
        
    RetrieveBooleanParameterName = sReturnValue
    
End Function
