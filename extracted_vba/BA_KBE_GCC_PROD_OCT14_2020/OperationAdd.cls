Attribute VB_Name = "OperationAdd"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSelectedBody As Body
Private oParameterOAValue As Body
Private sSeedPartLetter As String
Private oParameterOA As StrParam

'*************************************************************
'************* OA properties ***********************
'*************************************************************
Public Property Get BodyOA() As Body
    Set BodyOA = oSelectedBody
End Property
Public Property Let BodyOA(oBodyOA As Body)
    Set oSelectedBody = oBodyOA
End Property


Public Property Get ImpactedBody() As Body
    Set ImpactedBody = oParameterOAValue
End Property
Public Property Let ImpactedBody(oKBEBody As Body)
    Set oParameterOAValue = oKBEBody
End Property


Public Property Get SeedPartLetter() As String
    SeedPartLetter = sSeedPartLetter
End Property
Public Property Let SeedPartLetter(sSPLetter As String)
    sSeedPartLetter = sSPLetter
End Property


Public Property Get ParameterOA() As StrParam
    Set ParameterOA = oParameterOA
End Property
Public Property Let ParameterOA(oStrParam As StrParam)
    Set oParameterOA = oStrParam
End Property
