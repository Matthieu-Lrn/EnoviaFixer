Attribute VB_Name = "CreateFinish"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oRefSurface As New clsFinishRefSurfaces
Private oRootDocument As Document

Public Sub Launch(ByVal oFinishTargetPart As clsFinishTargetPart, ByVal oFinishPartRootDoc As Document, _
                  ByVal dFinishThk As Double, ByVal bDisableInsertHoles As Boolean, ByVal oInsertHole As Variant, _
                  ByVal bDisableSpecificHoles As Boolean, ByVal dLowerLimit As Double, ByVal dUpperLimit As Double)

Dim obody As Body
Dim oSelectedFace As AnyObject
Dim iBodyID As Integer
Dim iSurfaceID As Integer
Dim sSelectionStatus As String
Dim iLastRow As Integer
Dim oSPAWorkbench As Workbench

'Swap to Assembly wkb
Call SwapToAssemblyWorkbench

Set oSPAWorkbench = oFinishPartRootDoc.GetWorkbench("SPAWorkbench")

'Retreive next available ID
iBodyID = 0
For Each obody In oFinishTargetPart.Part.Bodies
    If obody.Name Like "FinishBody_*" Then
        If IsNumeric(Split(obody.Name, "_")(1)) Then
            If CInt(Split(obody.Name, "_")(1)) >= iBodyID Then
                iBodyID = CInt(Split(obody.Name, "_")(1)) + 1
            End If
        End If
    End If
Next

'Add Body and Geoset
oFinishTargetPart.AddNewFinishBody (iBodyID)

'Select face
sSelectionStatus = "NoSelection"

iSurfaceID = 0
Do

    iSurfaceID = iSurfaceID + 1
    
    'Select face
    Set oSelectedItem = MultiWindowSelectElement("Please select a face for the finish or press ESC !")
    
    'Check conditions
    If oSelectedItem Is Nothing Then
        
        If sSelectionStatus = "NoSelection" Then
            sSelectionStatus = "KO"
        Else
            sSelectionStatus = "OK"
        End If
        Exit Do
    End If
    
    If Not (TypeName(oSelectedItem.Value) Like "*Face" And oSelectedItem.Document.Name = oFinishPartRootDoc.Name) Then
        sResult = MsgBox("Selected element is not a face." & vbCrLf & "Make sure part is in Design Mode.", vbExclamation)
        sSelectionStatus = "KO"
        Exit Do
    Else
        sSelectionStatus = "OK"
    End If
    
    
    'Initialize
    Set oRefSurface.Product = oSelectedItem.LeafProduct
    Call oRefSurface.Initialize
    oRefSurface.ParentList = BuildParentList(oRefSurface.Product)

    'Create extract surface on selected face
    Call oRefSurface.CreateExtractSurface(oSelectedItem.Value, "_FaceExtract")
    
    'Create Construction Geoset
    Call oFinishTargetPart.AddContructionGeoset
    
    'Copy/Paste the extract surface in Finish Part
    Call CopyPasteElements(oRefSurface.ParentList, oFinishTargetPart.ParentList, "FinishBody_" & iBodyID, "ExtractedFace_" & iSurfaceID, "_FaceExtract")
    
    'Rename pasted surface
    oFinishTargetPart.RefGeoset.HybridBodies.Item("Construction_" & iSurfaceID).HybridBodies.Item("ExtractedFace_" & iSurfaceID).HybridShapes.Item(1).Name = "Face" & iSurfaceID
    
    'Delete Geoset of extracted surface
    oRefSurface.DeleteGeoset
    
Loop

If sSelectionStatus = "OK" Then


    'Analyze the extracted face and removes:
    ' - Inserts Holes
    ' - Holes with a specific diameter
    Call oFinishTargetPart.ManageHoles(oSPAWorkbench, bDisableInsertHoles, oInsertHole, bDisableSpecificHoles, _
                                            dLowerLimit, dUpperLimit)
  
    'Create a join between all the "cleaned" selected surfaces
    If oFinishTargetPart.RefGeoset.HybridShapes.Count > 1 Then
        oFinishTargetPart.CreateJoin
    End If

    'Create Thick Surface
    oFinishTargetPart.CreateThickSurface dFinishThk * 25.4

    'Hide geoset
    Call HideObject(oFinishTargetPart.RefGeoset)

Else

    'Delete Body
    Call oFinishTargetPart.DeleteBody(oFinishPartRootDoc)

End If

End Sub

Private Sub CopyPasteElements(sSourceParentList As String, sTargetParentList As String, sTargetBodyName As String, sTargetGeosetName As String, ByVal sElementName As String)

'When there is more that one instance of the source or the target part it is required to active the right part
'for the copy and paste operation


Dim sWkb As String
Dim i As Integer
Dim oProduct As Product

'Initial set
Set oSelection = CATIA.ActiveDocument.Selection

'Retrieve parent instance of element to copy
Set oProduct = RetrieveInstanceFromList(sSourceParentList)

'Copy
oSelection.Clear
oSelection.Add oProduct
oSelection.Search "Name=" & sElementName & ",sel"
CATIA.StartWorkbench "PrtCfg"
oSelection.Copy
CATIA.StartWorkbench "Assembly"

'Retrieve parent instance of target element
Set oProduct = RetrieveInstanceFromList(sTargetParentList)

'Paste
oSelection.Clear
oSelection.Add oProduct
oSelection.Search "Name=" & sTargetBodyName & ",sel"
oSelection.Search "Name=" & sTargetGeosetName & ",sel"
CATIA.StartWorkbench "PrtCfg"
oSelection.PasteSpecial "CATPrtResultWithOutLink"
CATIA.StartWorkbench "Assembly"
oSelection.Clear



End Sub

Private Function RetrieveInstanceFromList(sList As String) As Product

'From a parent list we can retrieve a specific instance in the tree

Dim oSelection
Dim i As Integer
Dim oProduct As Product
Dim oProdDocument As ProductDocument

'Initial set
Set oProdDocument = CATIA.ActiveDocument
Set oProduct = oProdDocument.Product
Set oSelection = oProdDocument.Selection

'Remove the last caracter which is a ";"
If Right(sList, 1) = ";" Then
    sList = Left(sList, Len(sList) - 1)
End If

'We select the instance in which the element to copy is found
For i = UBound(Split(sList, ";")) - 1 To 0 Step -1                  'We skip the first element in the list since it is the top product
    Set oProduct = oProduct.Products.Item(Split(sList, ";")(i))
Next

Set RetrieveInstanceFromList = oProduct

End Function

Private Sub HideObject(oObject As Variant)

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
oSelection.Add oObject

Dim oVisu As VisPropertySet

Set oVisu = oSelection.VisProperties
oVisu.SetShow 1

End Sub

Private Sub SwapToAssemblyWorkbench()

Dim sWkb As String

sWkb = CATIA.GetWorkbenchId
If sWkb <> "Assembly" Then
    CATIA.StartWorkbench "Assembly"
End If

End Sub

