Attribute VB_Name = "clsPanelCheck"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public PanelInstance As Product
Public PanelPart As Part
Public PanelInstanceStatus As String    '
Public PanelAssyInstance As Product
Public PanelAssyStatus As EnumPanelAssyStatus
Public PanelThicknessStatus As EnumPanelThicknessStatus
Public PanelThickness As Double
Public PanelMaterial As String
Public FlipPlane1Dir As Boolean
Public FlipPlane2Dir As Boolean
Public PanelPlane1 As Plane
Public PanelPlane2 As Plane
Private dArPlane1(8) 'As Double
Private dArPlane2(8) 'As Double
Public oPanelThicknessGeometery As HybridBody
Public InsertkeyCollection As Collection
Public oListOfSPCollectors As Collection    ': update 28Jan2018
Public InsertReportCollection As Collection

Private Sub Class_Initialize()
    Set InsertkeyCollection = New Collection
    Set InsertReportCollection = New Collection
    Set oListOfSPCollectors = New Collection
End Sub
Public Sub DeleteThicknessGeo()
    Dim oHSF As HybridShapeFactory
    Set oHSF = PanelPart.HybridShapeFactory
    
    If Not oPanelThicknessGeometery Is Nothing Then oHSF.DeleteObjectForDatum PanelPart.CreateReferenceFromObject(oPanelThicknessGeometery)
    
End Sub
Public Sub SetArPlane1(ByVal oValue As Variant)
    dArPlane1(0) = oValue(0)
    dArPlane1(1) = oValue(1)
    dArPlane1(2) = oValue(2)
    dArPlane1(3) = oValue(3)
    dArPlane1(4) = oValue(4)
    dArPlane1(5) = oValue(5)
    dArPlane1(6) = oValue(6)
    dArPlane1(7) = oValue(7)
    dArPlane1(8) = oValue(8)
End Sub
Public Property Get ArPlane1()
    ArPlane1 = dArPlane1
End Property
Public Sub SetArPlane2(ByVal oValue As Variant)
    dArPlane2(0) = oValue(0)
    dArPlane2(1) = oValue(1)
    dArPlane2(2) = oValue(2)
    dArPlane2(3) = oValue(3)
    dArPlane2(4) = oValue(4)
    dArPlane2(5) = oValue(5)
    dArPlane2(6) = oValue(6)
    dArPlane2(7) = oValue(7)
    dArPlane2(8) = oValue(8)
End Sub
Public Property Get ArPlane2()
    ArPlane2 = dArPlane2
End Property
Public Sub GetMaterial()
    Dim oMatman 'As MaterialManager
    Dim oAppliedMaterial 'As Material
    Dim oDoc As Document
    Dim sAppliedMaterial As String
    Dim oPartBody As Body
    
    Set oDoc = PanelInstance.ReferenceProduct.Parent
    Set oMatman = oDoc.Part.GetItem("CATMatManagerVBExt")
    Set oPartBody = oDoc.Part.MainBody
    oMatman.GetMaterialOnBody oPartBody, oAppliedMaterial
    '*** Verify that Material is applied
    If Not (oAppliedMaterial Is Nothing) Then PanelMaterial = oAppliedMaterial.Name Else PanelMaterial = ""
End Sub
Public Sub GetSPs() ': update 28Jan2018
    Dim oInstance As Product
    Dim oTest As AnyObject
    
    For Each oInstance In PanelAssyInstance.Products
        On Error Resume Next
        Err.Clear
        Set oTest = oInstance.ReferenceProduct
        If Err.Number <> 0 Then oInstance.ApplyWorkMode DEFAULT_MODE
        Err.Clear
        Set oTest = oInstance.ReferenceProduct  '*** if instance cannot be loaded
        If Err.Number = 0 Then
            On Error GoTo 0
            If oInstance.ReferenceProduct.Name Like "*SP##" Or oInstance.ReferenceProduct.Name Like "*SP###" Then
                oListOfSPCollectors.Add oInstance
            End If
        End If
        On Error GoTo 0
    Next
End Sub
