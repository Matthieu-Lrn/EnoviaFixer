Attribute VB_Name = "SFB_clsBDIPart"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public PartId As String
Public PartCompleteName As String
Public PartNumber As String
Public PartRev As String
Public PartTitle As String
Public IsCI As Boolean
Public Status As String
Public PrimaryDoc As SFB_clsBDIDoc
Public NextDoc As SFB_clsBDIDoc

Private Sub Class_Initialize()
    Set PrimaryDoc = New SFB_clsBDIDoc
    Set NextDoc = New SFB_clsBDIDoc
End Sub
Public Property Get Effectivity() As SFB_clsCollection
    Dim oCL As New SFB_clsCollection
    Dim SFB_i As Integer
    If UCase(PrimaryDoc.Status) <> "CLOSED" And PrimaryDoc.DocNumber <> "" Then
        If NextDoc.Effectivity.Count > 0 Then
            For SFB_i = 1 To NextDoc.Effectivity.Count
                oCL.Add NextDoc.Effectivity.GetKey(SFB_i), NextDoc.Effectivity.GetItem(SFB_i)
            Next
        End If
    Else
        If PrimaryDoc.Effectivity.Count > 0 Then
            For SFB_i = 1 To PrimaryDoc.Effectivity.Count
                oCL.Add PrimaryDoc.Effectivity.GetKey(SFB_i), PrimaryDoc.Effectivity.GetItem(SFB_i)
            Next
        End If
    End If
    Set Effectivity = oCL
End Property


