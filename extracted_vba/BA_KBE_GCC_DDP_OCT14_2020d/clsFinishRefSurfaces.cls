Attribute VB_Name = "clsFinishRefSurfaces"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oProduct As Product
Private oPart As Part
Private sParentList As String
Private oHybridShapeFactory As HybridShapeFactory
Private oGeoset As HybridBody

'Product
Public Property Set Product(ByVal oValue As Product)
Set oProduct = oValue
End Property

Public Property Get Product() As Product
Set Product = oProduct
End Property

'Part
Public Property Get Part() As Part
Set Part = oPart
End Property

'Parent list
Public Property Get ParentList() As String
     ParentList = sParentList
End Property

Public Property Let ParentList(ByVal oValue As String)
    sParentList = oValue
End Property

Public Sub Initialize()

'Retrieve Part
Set oPart = oProduct.ReferenceProduct.Parent.Part

'Retrieve HybridShapeFactory
Set oHybridShapeFactory = oPart.HybridShapeFactory

'Add dummy geoset
Set oGeoset = oPart.HybridBodies.Add
oGeoset.Name = "_DummyGeoset"

End Sub

Public Sub CreateExtractSurface(ByVal oSelectedFace As AnyObject, ByVal sExtractName As String)

Dim oExtractFace As HybridShapeExtract

Set oExtractFace = oHybridShapeFactory.AddNewExtract(oSelectedFace)
oExtractFace.PropagationType = 3
oExtractFace.ComplementaryExtract = False
oExtractFace.IsFederated = False
oGeoset.AppendHybridShape oExtractFace
oExtractFace.Name = sExtractName

oPart.UpdateObject oGeoset
End Sub

Public Sub DeleteGeoset()

Dim oReference As Reference

Set oReference = oPart.CreateReferenceFromObject(oGeoset)
oHybridShapeFactory.DeleteObjectForDatum oReference
oPart.Update

End Sub
