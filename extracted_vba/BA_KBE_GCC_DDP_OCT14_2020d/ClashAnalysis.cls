Attribute VB_Name = "ClashAnalysis"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public cPinsSel As New Collection
Public cPartsSel As New Collection
Public cPartsSelNoHW As New Collection
Public GrpPanelsClashWithPins As New Collection
Public cClashesDetected As New Collection
Public ColInserts As New Collection
Dim cClashes As Clashes
Public objSel As Selection
Public visPropertySet1 As VisPropertySet

Private oExcelApp As Object

Private bUseCRTDatabse As Boolean

Public oDict As Variant
Private oAttList As New clsAttributesList

Public Sub ClashAnalysisWithoutHarware(ByVal cTargetAssys As Collection, HardwareClash As Boolean)

    ' Create Enovia Connection Object:
    Dim SysService
    Set SysService = CATIA.SystemService
    
    'Set oDict = CreateObject("Scripting.Dictionary")

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    bUseCRTDatabse = False

    Dim sAttributs, oAttributeDictionary
    
    Dim sPNTitle As String
    Dim sPNRevisionOrg As String
    
    
    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")

    Dim oClash As Clash

    
'----------------------------------------------------------------------------------------------
    
    Dim productDocument1 As ProductDocument
    Set productDocument1 = CATIA.ActiveDocument
    
    Dim cPartSel As New Collection
    
    
    Dim selection1 As Selection
    Set selection1 = productDocument1.Selection
    
    For i = 1 To cTargetAssys.Count
        '################ LOOP HERE FOR ALL oSeedAssy ########################
        selection1.Add cTargetAssys.Item(i) '.Value
        selection1.Search "(CATAsmSearch.Product.Name=*),sel"
        
        For j = 1 To selection1.Count
            cPartSel.Add selection1.Item(j)
        Next
        selection1.Clear
        '######################################################################
    Next
    'Call frmProgress.progressBarInitialize("Performing clash analysis")
    

    
    Dim d As Integer
    d = 1
    If HardwareClash = True Then
        d = 2
    End If
        
    For i = 1 To cPartSel.Count
        
        
        If cPartSel.Item(i).Type = "Product" Then
            If cPartSel.Item(i).Value.HasAMasterShapeRepresentation Then
                On Error Resume Next
                testPartNumber = cPartSel.Item(i).Value.PartNumber
                'testPartNumberRevision = cPartSel.Item(i).Value.Revision
                


                If (Err.Number <> 0 And testPartNumber = "") Then
                    cPartSel.Item(i).Value.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
                End If
                Err.Clear
                
                testPartNumber = cPartSel.Item(i).Value.PartNumber
                If Err.Number <> 0 Then
                    GoTo NextItem   'TO CHANGE********************
                End If
                              
                
                If oAttList.GetEnoviaAttributes(testPartNumber, "--", False, "Revision Organization") <> "Part or Attribute doesn't exist" Then
                    
                    sPNRevisionOrg = oAttList.GetEnoviaAttributes(testPartNumber, "--", False, "Revision Organization")
                    sPNTitle = oAttList.GetEnoviaAttributes(testPartNumber, "--", False, "Title")
                    
                    On Error GoTo 0
        
                    
                    If Not ((sPNTitle Like "*SCREW*" Or sPNTitle Like "*WASHER*" Or sPNTitle Like "*PIN*" Or sPNTitle Like "*BOLT*" Or sPNTitle Like "*INSERT*" Or sPNTitle Like "*RIVET*") And sPNRevisionOrg Like "*STANDARD*") Then
        '                If testPartNumber Like "SC0300364*" Then
        '                    a = oAttributeDictionary("Title")
        '                End If
                        'If Not (sPNRevisionOrg Like "*STANDARDS") Then
                            cPartsSelNoHW.Add cPartSel.Item(i)
                        'End If
                    End If
               End If
            End If
        Call frmProgress.progressBarRepaint("Performing Panel Clash...", cPartSel.Count * d, i, cPartSel.Item(i).Value.PartNumber, cPartSel.Count, i, "")
        End If

NextItem:
        testPartNumber = ""
        
        
        
        
       

        '***if user presses cancel on progress toolbar
        If bCancelAction = True Then
            GoTo cleanExit:
        End If
    
    Next
    
    
    Dim objGroups As Groups
    Set objGroups = MyProduct.GetTechnologicalObject("Groups")
    
    Dim oClasheResult As ClashResults
    Set oClasheResult = MyProduct.GetTechnologicalObject("ClashResults")
    
    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")
    Set oClash = cClashes.Add()
    oClash.Name = "Macro-Clash Temp"
    
    Dim grpFirst1 As Group
    Set grpFirst1 = objGroups.Add()
    grpFirst1.Name = "Macro-Clash Temp"

           
    For i = 1 To cPartsSelNoHW.Count
        grpFirst1.AddExplicit cPartsSelNoHW.Item(i).Value
    Next i
        
    oClash.FirstGroup = grpFirst1
    oClash.ComputationType = catClashComputationTypeInsideOne

    oClash.Compute
    Set xx = oClash.AnnotatedViews

    Dim vFirstClashPoint(2)
    Dim vSecondClashPoint(2)
    Dim oConflict 'As Conflict
    
    For i = 1 To oClash.Conflicts.Count
        If oClash.Conflicts.Item(i).Type = catConflictTypeClash Then
            cClashesDetected.Add oClash.Conflicts.Item(i)
            
            Set oConflict = oClash.Conflicts.Item(i)
            oConflict.GetFirstPointCoordinates vFirstClashPoint
            oConflict.GetSecondPointCoordinates vSecondClashPoint
            
        End If
    Next i
    

cleanExit:
a = 1
'Call CleanAssy

End Sub

Public Sub ClashAnalysisBetweenTwo(Item As Product, cPartsSel As Collection, GrpPanelsClashWithPins As Collection)

    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Dim cClashes As Clashes
    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")
    Dim oClash As Clash
    Set oClash = cClashes.Add()
    Dim objGroups As Groups
    Set objGroups = MyProduct.GetTechnologicalObject("Groups")

    
    
    Dim grpFirst As Group
    Dim grpSecond As Group
    
    
    Set grpFirst = objGroups.Add()
    Set grpSecond = objGroups.Add()

    
    grpFirst.AddExplicit Item
    
    For i = 1 To cPartsSel.Count
        grpSecond.AddExplicit cPartsSel.Item(i)
    Next i
        
    
    oClash.FirstGroup = grpFirst
    oClash.SecondGroup = grpSecond

    oClash.ComputationType = catClashComputationTypeBetweenTwo
   

    oClash.Compute


    For i = 1 To oClash.Conflicts.Count
        GrpPanelsClashWithPins.Add oClash.Conflicts.Item(i).secondProduct
    Next i
   
    

    For i = 1 To objGroups.Count
        objGroups.Remove (1)
    Next
   
    
End Sub
Public Sub ClashAnalysisBetweenPinAndInserts(Item As Product, ColInserts As Collection, GrpPanelsClashWithPins As Collection, iMinClearance As Double)

    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Dim cClashes As Clashes
    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")
    Dim oClash As Clash
    Set oClash = cClashes.Add()
    Dim objGroups As Groups
    Set objGroups = MyProduct.GetTechnologicalObject("Groups")

    
    
    Dim grpFirst As Group
    Dim grpSecond As Group
    
    
    Set grpFirst = objGroups.Add()
    Set grpSecond = objGroups.Add()

    
    grpFirst.AddExplicit Item
    
    For i = 1 To ColInserts.Count
        grpSecond.AddExplicit ColInserts.Item(i)
    Next i
        
    
    oClash.FirstGroup = grpFirst
    oClash.SecondGroup = grpSecond

    oClash.ComputationType = catClashComputationTypeBetweenTwo
    oClash.InterferenceType = catClashInterferenceTypeClearance
    'oClash.InterferenceType = catClashInterferenceTypeContact
    oClash.Clearance = iMinClearance * 25.4
    oClash.Compute
    
    Dim Count As Integer
    Count = cClashesDetected.Count + 1
    
    Dim vFirstClashPoint(2)
    Dim vSecondClashPoint(2)
    Dim oConflict 'As Conflict
    
    
    For i = 1 To oClash.Conflicts.Count
        cClashesDetected.Add oClash.Conflicts.Item(i)
           
        Set oConflict = oClash.Conflicts.Item(i)
        oConflict.GetFirstPointCoordinates vFirstClashPoint
        oConflict.GetSecondPointCoordinates vSecondClashPoint
        
    Next i
   
End Sub
Public Sub Pins_InsertsClashAnalysis(ByVal cTargetAssys As Collection, ClashNoHarware As Boolean, iMinClearance As Double)

Dim testType As String
Dim Item As Product

'Set oDict = CreateObject("Scripting.Dictionary")

Set objFSO = CreateObject("Scripting.FileSystemObject")

bUseCRTDatabse = False


    
    testType = "HW"
    Call FindAllPins(cTargetAssys)
    Dim i As Integer
    i = 1
    
    Dim total As Integer
    total = 0
    If ClashNoHarware = True Then
        total = cPinsSel.Count
    End If
    
    For i = 1 To cPinsSel.Count
        Set GrpPanelsClashWithPins = New Collection
        
        Call ClashAnalysisBetweenTwo(cPinsSel.Item(i).Value, cPartsSel, GrpPanelsClashWithPins)
        
        If GrpPanelsClashWithPins.Count > 0 Then
            Call FindInsertsForPanel(GrpPanelsClashWithPins)
            Call ClashAnalysisBetweenPinAndInserts(cPinsSel.Item(i).Value, ColInserts, GrpPanelsClashWithPins, iMinClearance)
            
            
            For j = 1 To ColInserts.Count
                ColInserts.Remove (1)
            Next
            
            
        End If
        
        Call frmProgress.progressBarRepaint("Performing Pin/Insert Clash...", cPinsSel.Count + total, total + i, "", cPinsSel.Count, i, "")
'        For i = 1 To GrpPanelsClashWithPins.Count
'          GrpPanelsClashWithPins.Remove 1 ' Remove first item
'        Next i
        'i = i + 1
    Next
    
End Sub
Public Sub FindAllPins(ByVal cTargetAssys As Collection)
    
'    Dim cPinsSel As New Collection
    Dim testPartNumber As String
    
    CATIA.ActiveDocument.Selection.Clear

    Dim productDocument1 As ProductDocument
    Set productDocument1 = CATIA.ActiveDocument
    
    'Set pinsSel = CATIA.ActiveDocument.Selection
    

    Dim selection1 As Selection
    Set selection1 = productDocument1.Selection
    
    For i = 1 To cTargetAssys.Count
        selection1.Add cTargetAssys.Item(i)
    Next
    
    selection1.Search "(CATAsmSearch.Product.Name=*SC0300042*),sel"


    For i = 1 To selection1.Count
        cPinsSel.Add selection1.Item(i)
    Next
    
    selection1.Clear
    
    For i = 1 To cPinsSel.Count
        On Error Resume Next
        testPartNumber = cPinsSel.Item(i).Value.PartNumber
        
        If Err.Number <> 0 Then
            cPinsSel.Item(i).Value.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
        End If
        Err.Clear
        testPartNumber = cPinsSel.Item(i).Value.PartNumber
        If Err.Number <> 0 Then
            'Get rid of object
        End If
        On Error GoTo 0
        
        'cPinsSel.Add pinsSel.Item(i).Value
        
    Next i


    testType = "noHW" '****************to be changed!!!!!!!!*************
    Call GenerateItemlist(cTargetAssys, "noHW")



End Sub
Private Sub FindInsertsForPanel(GrpPanelsClashWithPins)

Dim parentProduct As Product
Dim strItem As String
Dim strPartent As String
Dim testPartNumber As String

For Each Item In GrpPanelsClashWithPins

    Set parentProduct = Item.Parent.Parent
   
    On Error Resume Next
    testPartNumber = Item.PartNumber
    
    If Err.Number <> 0 Then
        Item.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
    End If
    Err.Clear
    testPartNumber = parentProduct.PartNumber
    
    If Err.Number <> 0 Then
        parentProduct.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
    End If
    Err.Clear
    testPartNumber = Item.PartNumber
    If Err.Number <> 0 Then
        'Get rid of object
    End If
    testPartNumber = parentProduct.PartNumber
    If Err.Number <> 0 Then
        'Get rid of object
    End If
    On Error GoTo 0
            
            
            
    strItem = Left(Item.PartNumber, InStr(Item.Name, "-") - 1)
    strPartent = Left(parentProduct.PartNumber, InStr(parentProduct.Name, "-") - 1)
   
    If (strItem Like strPartent) Then
        Call findPanelInserts(parentProduct, ColInserts)
    End If
Next


End Sub
Public Sub GenerateItemlist(ByVal cTargetAssy As Collection, testType As String)

    Dim oInstance As Product
    Dim sParentList As String
    Dim TestArray() As String
    Dim sPartType As String


    Dim oShapeRepresentation As CATBaseDispatch

    Dim oselectionList As Selection
    Set oselectionList = CATIA.ActiveDocument.Selection
    oselectionList.Clear

If testType = "noHW" Then
    For i = 1 To cTargetAssy.Count
        For Each oInstance In cTargetAssy.Item(i).Products
    
            If oInstance.Name Like "*SP0*" And Not oInstance.HasAMasterShapeRepresentation Then
                GoTo NextIteration
            End If
    
            If oInstance.Products.Count > 0 Then
                Dim cRecursiveColection As New Collection
                cRecursiveColection.Add oInstance
                Call GenerateItemlist(cRecursiveColection, testType) ' recursion to scan all sub branches
            Else
    
                If oInstance.HasAMasterShapeRepresentation Then
                
                    'oselectionList.Add oInstance
    
                    ''Call pinsSel.Search(searchPartName)
                    'cPartsSel.Add oselectionList.Item(oselectionList.Count)
                    cPartsSel.Add oInstance
    
    '                If (OptionButton1 = True) Then
    '                    oInstance.ApplyWorkMode DESIGN_MODE
    '                End If
                End If
    
            End If
NextIteration:
        Next oInstance

    Next i

End If

If testType = "HW" Then


    For Each oInstance In oAssembly.Products

        If oInstance.Name Like "*SP0*" And Not oInstance.HasAMasterShapeRepresentation Then
            GoTo NextIterationHW
        End If

        If oInstance.Products.Count > 0 Then
            Call GenerateItemlist(oInstance, testType) ' recursion to scan all sub branches
        Else

            If oInstance.HasAMasterShapeRepresentation Then
                oselectionList.Add oInstance
'                If (OptionButton1 = True) Then
'                    oInstance.ApplyWorkMode DESIGN_MODE
'                End If
            End If

        End If


NextIterationHW:
    Next

End If

End Sub
Public Sub findPanelInserts(oAssembly As Product, ColInserts As Collection)
Dim oInstance As Product
Dim testPartNumber As String


For Each oInstance In oAssembly.Products
    
    If oInstance.Products.Count > 0 Then
        Call findPanelInserts(oInstance, ColInserts)
    Else
        If oInstance.Name Like "TYE2*" Then
        
        
            On Error Resume Next
            testPartNumber = pinsSel.Item(i).Value.PartNumber

            If (Err.Number <> 0 And testPartNumber = "") Then
                'pinsSel.Item(i).Value.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
                oInstance.ApplyWorkMode (CatWorkModeType.DEFAULT_MODE)
            End If
            Err.Clear
            testPartNumber = oInstance.PartNumber
            If Err.Number <> 0 Then
                'Get rid of object
            End If
            On Error GoTo 0
        
        
        
        
        
            ColInserts.Add oInstance
   
        End If
    End If


Next

End Sub
Public Sub ExportToExcel(cTableList)
    
    Dim oExcel As clsExcel
    
    Set oExcel = New clsExcel
    Call oExcel.GetExcel
    oExcel.NewExcelFile
    
    iRow = 1
    iColumn = 1
    
    oExcel.App.Cells(iRow, iColumn).Value = "Clash Number"
    oExcel.App.Cells(iRow, iColumn + 1).Value = "First Product"
    oExcel.App.Cells(iRow, iColumn + 2).Value = "Second Product"
    oExcel.App.Cells(iRow, iColumn + 3).Value = "Clash Value (in)"
    oExcel.App.Cells(iRow, iColumn + 4).Value = "First Product Path"
    oExcel.App.Cells(iRow, iColumn + 5).Value = "Second Product Path"
    oExcel.App.Cells(iRow, iColumn + 6).Value = "First Clash Coordinates"
    oExcel.App.Cells(iRow, iColumn + 7).Value = "Second Clash Coordinates"
    oExcel.App.Cells(iRow, iColumn + 8).Value = "Notes"
  
  On Error Resume Next

  
    If IsNull(cTableList.Item(1)) = True Then
        On Error GoTo 0
        Exit Sub
    End If
      
    On Error GoTo 0

    a = 1
    For Each Element In cTableList
    
        iRow = iRow + 1
        
        oExcel.App.Cells(iRow, iColumn).Value = a
        oExcel.App.Cells(iRow, iColumn + 1).Value = cTableList.Item(a)(0)
        oExcel.App.Cells(iRow, iColumn + 2).Value = cTableList.Item(a)(1)
        oExcel.App.Cells(iRow, iColumn + 3).Value = cTableList.Item(a)(2)
        oExcel.App.Cells(iRow, iColumn + 4).Value = cTableList.Item(a)(3)
        oExcel.App.Cells(iRow, iColumn + 5).Value = cTableList.Item(a)(4)
        oExcel.App.Cells(iRow, iColumn + 6).Value = cTableList.Item(a)(5)
        oExcel.App.Cells(iRow, iColumn + 7).Value = cTableList.Item(a)(6)
        oExcel.App.Cells(iRow, iColumn + 8).Value = cTableList.Item(a)(7)

        a = a + 1
    Next Element

    oExcel.App.Range("A1:I1").Select
    oExcel.App.Selection.EntireColumn.AutoFit
    oExcel.App.Range("D1:H1").Select
    oExcel.App.Selection.EntireColumn.Hidden = True

    
    oExcel.App.Range("A1:I1").Select

    With oExcel.App.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatics
        .ThemeColor = 5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    oExcel.ShowExcelWindow
    

End Sub
Public Sub ImportFromToExcel(oExcel, cTableList)



    Dim RowClash(7) As Variant
    
    i = 2
    While oExcel.App.Cells(i, 1).Value <> ""
    
        RowClash(0) = oExcel.App.Cells(i, 2).Value
        RowClash(1) = oExcel.App.Cells(i, 3).Value
        RowClash(2) = oExcel.App.Cells(i, 4).Value
        RowClash(3) = oExcel.App.Cells(i, 5).Value
        RowClash(4) = oExcel.App.Cells(i, 6).Value
        RowClash(5) = oExcel.App.Cells(i, 7).Value
        RowClash(6) = oExcel.App.Cells(i, 8).Value
        RowClash(7) = oExcel.App.Cells(i, 9).Value
        
        frmClashAnalysis.cTableList.Add (RowClash)
        i = i + 1
    Wend
    


End Sub
Public Sub CompileClash()

    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")

    Dim pFirstProduct As Product

    Dim firstCoordinates(2)
    Dim secondCoordinates(2)
    Dim oConflict 'As Conflict
    
    iRow = iRow + 1
    
    Dim vFirstClashPoint(2)
    Dim vSecondClashPoint(2)
    
    Dim partNameSplit 'As String
    
    'objTextFile = GetTagValueFromFile("C:\Users\b0474786\Desktop\ClashMacro\ClashMacro_Soft Material List.txt", "Soft Material List")
    
    objTextFile = Replace(sSoftMaterialList, "\", "/")
    
    Dim SoftMaterialList 'As Variant
    SoftMaterialList = Split(objTextFile, "|")
    
    Dim sPNMaterialSpecification(1) As String

    
    For i = 1 To cClashesDetected.Count

        Set oConflict = cClashesDetected.Item(i)


        oConflict.GetFirstPointCoordinates vFirstClashPoint
        oConflict.GetSecondPointCoordinates vSecondClashPoint
        
        If (oConflict.Type = catConflictTypeClash Or catConflictTypeClearance) Then
            If (oConflict.Value <> 0) Then
                 
 
                partNameSplit = Split(oConflict.firstProduct.ReferenceProduct.Parent.Name, ".")
                Dim rev As String
                rev = Right(partNameSplit(0), 2)

                sPNMaterialSpecification(0) = oAttList.GetEnoviaAttributes(oConflict.firstProduct.PartNumber, rev, False, "Material Specification Production")
                If sPNMaterialSpecification(0) = "Part or Attribute doesn't exist" Then
                    sPNMaterialSpecification(0) = ""
                End If


                partNameSplit = Split(oConflict.secondProduct.ReferenceProduct.Parent.Name, ".")
                rev = Right(partNameSplit(0), 2)

                sPNMaterialSpecification(1) = oAttList.GetEnoviaAttributes(oConflict.secondProduct.PartNumber, rev, False, "Material Specification Production")
                If sPNMaterialSpecification(1) = "Part or Attribute doesn't exist" Then
                    sPNMaterialSpecification(1) = ""
                End If
                
                Dim RowClash(7) As Variant
                RowClash(0) = oConflict.firstProduct.PartNumber
                RowClash(1) = oConflict.secondProduct.PartNumber
                RowClash(2) = oConflict.Value / 25.4
                RowClash(3) = BuildParentList(oConflict.firstProduct)
                RowClash(4) = BuildParentList(oConflict.secondProduct)
                RowClash(5) = CStr(vFirstClashPoint(0)) + "," + CStr(vFirstClashPoint(1)) + "," + CStr(vFirstClashPoint(2))
                RowClash(6) = CStr(vSecondClashPoint(0)) + "," + CStr(vSecondClashPoint(1)) + "," + CStr(vSecondClashPoint(2))
                RowClash(7) = "Clash"
                
                For k = 0 To 1
                    For Each Material In SoftMaterialList
                        If Material = sPNMaterialSpecification(k) Then
                            RowClash(7) = "Soft Material Clash"
                        End If
                    Next Material
                Next k
                
                frmClashAnalysis.cTableList.Add (RowClash)
    
                iRow = iRow + 1
                              
            End If

            ElseIf (oConflict.Type = catConflictTypeContact) Then
                oConflict.Status = catConflictStatusIrrelevant
        End If
    Next
    

    

    For i = 1 To cClashes.Count
        cClashes.Remove (1)
    Next
End Sub

Public Sub DefineClash(firstPart As String, secondPart As String, clashValue As Double, firstPath As String, secondPath As String, firstClashPoint As String, secondClashPoint As String)

    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Dim cClashes As Clashes
    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")

    
    Dim objGroups As Groups
    Set objGroups = MyProduct.GetTechnologicalObject("Groups")
    
    
    Dim productDocument1 As ProductDocument
    Set productDocument1 = CATIA.ActiveDocument

    Dim navigatorWorkbench1 As NavigatorWorkbench
    Set navigatorWorkbench1 = productDocument1.GetWorkbench("NavigatorWorkbench")
    
    CATIA.ActiveDocument.Selection.Clear
    
    Set objSel = CATIA.ActiveDocument.Selection
    

    
    Call CleanAssy

    'Dim oClash As Clash
    Set oClash = cClashes.Add()
    oClash.Name = "Macro-Clash Interference"

    Dim oRow As Integer
    Dim firstProduct As String
    Dim secondProduct As String
    Dim firstProductPath As String
    Dim secondProductPath As String
    Dim searchPartName As String
    Dim FirstPoint()  As String
    Dim SecondPoint()  As String
    
    Dim PinPath() As String
    Dim sPath(1) As String
    
    Dim oAxisPinArray(11) As Variant
    Dim vClashVector(2)  As Double
    Dim vZVector As Variant
    Dim vYVector As Variant
    
    Dim U(2) As Double
    Dim V(2) As Double
    
    Dim oSectionComponents(11) As Variant
    
    sAssyName = ""

    On Error Resume Next
        firstProduct = firstPart
        secondProduct = secondPart
        firstProductPath = firstPath
        secondProductPath = secondPath
        FirstPoint = Split(firstClashPoint, ",")
        SecondPoint = Split(secondClashPoint, ",")
        
        vClashVector(0) = Val(FirstPoint(0)) - Val(SecondPoint(0))
        vClashVector(1) = Val(FirstPoint(1)) - Val(SecondPoint(1))
        vClashVector(2) = Val(FirstPoint(2)) - Val(SecondPoint(2))
        
        
        
        sPath(0) = firstProductPath
        sPath(1) = secondProductPath
        
        
    On Error GoTo 0
    

Call CalculateInstanceGlobalPositionMatrix(firstProductPath, oAxisPinArray)


Dim Vnorm As Double
Vnorm = Sqr(oAxisPinArray(6) ^ 2 + oAxisPinArray(7) ^ 2 + oAxisPinArray(8) ^ 2)

V(0) = oAxisPinArray(6)
V(1) = oAxisPinArray(7)
V(2) = oAxisPinArray(8)
Call CrossProduct(vClashVector, V, vZVector)
V(0) = vZVector(0)
V(1) = vZVector(1)
V(2) = vZVector(2)
Call CrossProduct(vClashVector, V, vYVector)


oSectionComponents(0) = vClashVector(0)
oSectionComponents(1) = vClashVector(1)
oSectionComponents(2) = vClashVector(2)
oSectionComponents(3) = vYVector(0)
oSectionComponents(4) = vYVector(1)
oSectionComponents(5) = vYVector(2)
oSectionComponents(6) = vZVector(0)
oSectionComponents(7) = vZVector(1)
oSectionComponents(8) = vZVector(2)
oSectionComponents(9) = oAxisPinArray(9)
oSectionComponents(10) = oAxisPinArray(10)
oSectionComponents(11) = oAxisPinArray(11)

Check = 0
For i = 0 To 11
    Check = Check + oSectionComponents(i)
Next
If Check = 0 Then
    MsgBox "No suitable part found. Review your selection"
    Exit Sub
End If








    If sAssyName = "" Then
        'MsgBox ("Please select the row with assembly name in Excel")
    Else
  '      Call SelectAssemblyInCatia(sAssyName)
    End If



    Dim oTargetProduct As Product

    Dim grpFirst As Group
    Dim grpSecond As Group

    
    
    Set grpFirst = objGroups.Add()
    Set grpSecond = objGroups.Add()
    
For k = 0 To 1
    Set oTargetProduct = productDocument1.Product
    PinPath = Split(sPath(k), ";")
    For i = 1 To UBound(PinPath)

        For j = 1 To oTargetProduct.Products.Count
            If oTargetProduct.Products.Item(j).Name = PinPath(UBound(PinPath) - i) Then

               Set oTargetProduct = oTargetProduct.Products.Item(j)
               Exit For
            End If
        Next
    Next
    objSel.Add oTargetProduct
    If k = 0 Then
        grpFirst.AddExplicit oTargetProduct
    End If

    If k = 1 Then
        grpSecond.AddExplicit oTargetProduct
    End If


Next

    Dim oGroup As Group
    Set oGroup = objGroups.AddFromSel
    oGroup.Name = "Macro-Clash Components"
    
    oGroup.FillSelWithInvert
    Call KBESearch("CATAsmSearch.Product.Visibility=Shown,sel")
    Set oGroup = objGroups.AddFromSel
    oGroup.Name = "Macro-Clash Hiden Components"
    
    objSel.Clear
    objSel.Add oGroup
    
    

    Set visPropertySet1 = objSel.VisProperties
    visPropertySet1.SetShow 1
    
    objSel.Clear


    oClash.FirstGroup = grpFirst
    oClash.SecondGroup = grpSecond

    oClash.ComputationType = catClashComputationTypeBetweenTwo

    oClash.Compute
    
    objSel.Add grpFirst
    objSel.Add grpSecond
    objSel.Delete
    
    
    Dim marker3Ds1 As Marker3Ds
    Set marker3Ds1 = navigatorWorkbench1.Marker3Ds
    
    Dim arrayOfVariantOfDouble1(2)
    arrayOfVariantOfDouble1(0) = Val(FirstPoint(0))
    arrayOfVariantOfDouble1(1) = Val(FirstPoint(1))
    arrayOfVariantOfDouble1(2) = Val(FirstPoint(2))
    Dim arrayOfVariantOfDouble2(2)
    arrayOfVariantOfDouble2(0) = Val(FirstPoint(0))
    arrayOfVariantOfDouble2(1) = Val(FirstPoint(1))
    arrayOfVariantOfDouble2(2) = Val(FirstPoint(2))

    Dim reference1 As Reference
    Set reference1 = productDocument1.Product

    Dim marker3D1 As Marker3D
    Set marker3Ds1Variant = marker3Ds1
    Set marker3D1 = marker3Ds1Variant.Add3DText(arrayOfVariantOfDouble1, "<---", arrayOfVariantOfDouble2, reference1)

    marker3D1.TextFont = "Swiss.pfb"
    marker3D1.TextSize = 4#
    marker3D1.Name = "Macro-Clash Pointer"
    marker3D1.Update

    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutOrigin Array(arrayOfVariantOfDouble1(0), arrayOfVariantOfDouble1(1), arrayOfVariantOfDouble1(2))
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutUpDirection Array(vYVector(0), vYVector(1), vYVector(2))
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutSightDirection Array(vZVector(0), vZVector(1), vZVector(2))
    

    CATIA.StartCommand ("Reframe On")
    CATIA.ActiveWindow.ActiveViewer.Update



End Sub

Public Sub CleanAssy()

    Dim MyDoc As Document
    Set MyDoc = CATIA.ActiveDocument

    Dim MyProduct As Product
    Set MyProduct = MyDoc.Product

    Dim cClashes As Clashes
    Set cClashes = MyProduct.GetTechnologicalObject("Clashes")

    Dim objGroups As Groups
    Set objGroups = MyProduct.GetTechnologicalObject("Groups")
    
    Dim productDocument1 As ProductDocument
    Set productDocument1 = CATIA.ActiveDocument

    Dim navigatorWorkbench1 As NavigatorWorkbench
    Set navigatorWorkbench1 = productDocument1.GetWorkbench("NavigatorWorkbench")
    
    Dim showstate As CatVisPropertyShow
    
    Set objSel = CATIA.ActiveDocument.Selection
    
    
    For Each Item In objGroups
        If Item.Name Like "Macro-Clash*" Then
            objSel.Add Item
            Set visPropertySet1 = objSel.VisProperties
           
            visPropertySet1.GetShow showstate

            If showstate = 1 Then
                visPropertySet1.SetShow catVisPropertyShowAttr
            End If
                
            objSel.Delete
        End If
    Next
    
    For Each Item In cClashes
        If Item.Name Like "Macro-Clash*" Then
            objSel.Add Item
            objSel.Delete
        End If
    Next
    
    For Each Item In navigatorWorkbench1.Marker3Ds
        If Item.Name Like "Macro-Clash*" Then
            objSel.Add Item
            objSel.Delete
        End If
    Next

End Sub

Public Sub ImportSelection()

    Dim sDocumentType As String
    Dim oSelectedItem As AnyObject
    On Error Resume Next
    
    Set selectedProduct = Nothing
    Set selectedRootDocument = Nothing
    Set oSelectedItem = MultiWindowSelectElement(userMessage)
    If oSelectedItem Is Nothing Then Exit Sub
    If TypeName(oSelectedItem.Document) <> "ProductDocument" Then Exit Sub
    
    '***** if we cannot access PartNumber property, put oSelectedItem.LeafProduct in default mode
    If oSelectedItem.LeafProduct.PartNumber = "" Then
        oSelectedItem.LeafProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    sDocumentType = ""
    sDocumentType = TypeName(oSelectedItem.Value.ReferenceProduct.Parent)
    If sDocumentType = "ProductDocument" Then
        Set selectedProduct = oSelectedItem.Value
    ElseIf sDocumentType = "PartDocument" Then
        If TypeName(oSelectedItem.Value.Parent.Parent.ReferenceProduct.Parent) = "ProductDocument" Then
            Set selectedProduct = oSelectedItem.Value.Parent.Parent
        End If
    Else
        sDocumentType = TypeName(oSelectedItem.LeafProduct.ReferenceProduct.Parent)
        If sDocumentType = "PartDocument" Then
            sDocumentType = TypeName(oSelectedItem.LeafProduct.Parent.Parent.ReferenceProduct.Parent)
            If sDocumentType = "ProductDocument" Then
                Set selectedProduct = oSelectedItem.LeafProduct.Parent.Parent
            End If
        End If
    End If
    
    If Not selectedProduct Is Nothing And selectedProduct.PartNumber Like productName Then
        '****If name contains "productName" and selection is not a part instance
        Set selectedRootDocument = oSelectedItem.Document
        CATIA.ActiveDocument.Selection.Clear
        
        If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
        If CATIA.GetWorkbenchId <> "Assembly" Then
            CATIA.StartWorkbench "Assembly"
        End If
    Else
        Set selectedProduct = Nothing
        Set selectedRootDocument = Nothing
        CATIA.ActiveDocument.Selection.Clear
    End If



End Sub
