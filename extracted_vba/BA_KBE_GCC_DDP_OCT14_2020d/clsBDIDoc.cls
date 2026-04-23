Attribute VB_Name = "clsBDIDoc"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public DocNumber As String
Public DocRev As String
Public DocID As String
Public Status As String
Public Effectivity As clsCollection
Private Sub Class_Initialize()
    Set Effectivity = New clsCollection
End Sub
