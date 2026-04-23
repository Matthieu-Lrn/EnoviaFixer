Attribute VB_Name = "clsFastener"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private sUniqueNameInAssy As String
Private oInstance As Product
Private sReferenceProductName As String
Private dInstanceGlobalMatrix As Variant
Private dStartPoint(2) As Double
Private dEndPoint(2) As Double
Private dThreadPoint(2) As Double
Private dFastenerLength As Double
Private dStackupOffset As Double
Private dStackupOffsetMin As Double
Private sFastenerType As String
Private bIsMale As Boolean
Private oDummy As Variant
Private sDia As String
Private bBlind As Boolean
Private bThreaded As Boolean
Private bCSK As Boolean
Private bHasGripLength As Boolean
Private sTag As String

Public Property Let UniqueNameInAssembly(ByVal oValue As String)
    sUniqueNameInAssy = oValue
End Property
Public Property Get UniqueNameInAssembly() As String
    UniqueNameInAssembly = sUniqueNameInAssy
End Property
Public Property Let Tag(ByVal oValue As String)
    sTag = oValue
End Property
Public Property Get Tag() As String
    Tag = sTag
End Property
Public Property Set InstanceObject(ByVal oValue As Product)
    Set oInstance = oValue
End Property
Public Property Get InstanceObject() As Product
    Set InstanceObject = oInstance
End Property
Public Property Let GlobalPositionMatrix(ByRef oValue As Variant)
    dInstanceGlobalMatrix = oValue
End Property
Public Property Get GlobalPositionMatrixIndex(oValue As Integer) As Double
    GlobalPositionMatrixIndex = dInstanceGlobalMatrix(oValue)
End Property
Public Property Get GlobalPositionMatrix() As Variant
    GlobalPositionMatrix = dInstanceGlobalMatrix
End Property
Public Property Get ThreadPoint() As Double()
    ThreadPoint = dThreadPoint
End Property
Public Property Get StartPoint() As Double()
    StartPoint = dStartPoint
End Property
Public Property Get EndPoint() As Double()
    EndPoint = dEndPoint
End Property
Public Property Let FastenerType(ByVal oValue As String)
    sFastenerType = oValue
End Property
Public Property Get FastenerType() As String
    FastenerType = sFastenerType
End Property

Public Property Let IsMale(ByVal oValue As Boolean)
    bIsMale = oValue
End Property
Public Property Get IsMale() As Boolean
    IsMale = bIsMale
End Property

Public Property Let ReferenceProductName(ByVal oValue As String)
    sReferenceProductName = oValue
End Property
Public Property Get ReferenceProductName() As String
    ReferenceProductName = sReferenceProductName
End Property
Public Property Let Dia(ByVal oValue As String)
    sDia = UCase(oValue)
End Property
Public Property Get Dia() As String
    Dia = sDia
End Property
Public Property Let IsBlind(ByVal oValue As Boolean)
     bBlind = oValue
End Property
Public Property Get IsBlind() As Boolean
     IsBlind = bBlind
End Property
Public Property Let IsThreaded(ByVal oValue As Boolean)
    bThreaded = oValue
End Property
Public Property Get IsThreaded() As Boolean
    IsThreaded = bThreaded
End Property
Public Property Let IsCounterSunk(ByVal oValue As Boolean)
     bCSK = oValue
End Property
Public Property Get IsCounterSunk() As Boolean
     IsCounterSunk = bCSK
End Property
Public Property Let Length(ByVal oValue As Double)
    dFastenerLength = oValue
End Property
Public Property Get Length() As Double
    Length = dFastenerLength
End Property

Public Property Let StackupOffset(ByVal oValue As Double)
    dStackupOffset = oValue
End Property
Public Property Get StackupOffset() As Double
    StackupOffset = dStackupOffset
End Property

Public Property Let StackupOffsetMin(ByVal oValue As Double)
    dStackupOffsetMin = oValue
End Property
Public Property Get StackupOffsetMin() As Double
    StackupOffsetMin = dStackupOffsetMin
End Property
Public Property Let HasGripLength(ByVal oValue As Boolean)
     bHasGripLength = oValue
End Property
Public Property Get HasGripLength() As Boolean
     HasGripLength = bHasGripLength
End Property
Public Sub SetEndPoint(ByVal oValue As Variant)
    dEndPoint(0) = oValue(9)
    dEndPoint(1) = oValue(10)
    dEndPoint(2) = oValue(11)
End Sub
Public Sub SetStartPoint(ByVal oValue As Variant)
    dStartPoint(0) = oValue(9)
    dStartPoint(1) = oValue(10)
    dStartPoint(2) = oValue(11)
End Sub
Public Sub SetThreadPoint(ByVal oValue As Variant)
    dThreadPoint(0) = oValue(9)
    dThreadPoint(1) = oValue(10)
    dThreadPoint(2) = oValue(11)
End Sub
