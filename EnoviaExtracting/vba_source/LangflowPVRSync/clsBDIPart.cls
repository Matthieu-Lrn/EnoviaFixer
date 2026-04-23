Attribute VB_Name = "clsBDIPart"
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
Public PrimaryDoc As clsBDIDoc
Public PrevDoc As clsBDIDoc

Private Sub Class_Initialize()
    Set PrimaryDoc = New clsBDIDoc
    Set PrevDoc = New clsBDIDoc
End Sub
Public Property Get Effectivity() As clsCollection
    Dim oCL As New clsCollection
    Dim i As Integer
    If UCase(PrimaryDoc.Status) <> "CLOSED" And PrimaryDoc.DocNumber <> "" Then
        If PrevDoc.Effectivity.Count > 0 Then
            For i = 1 To PrevDoc.Effectivity.Count
                oCL.Add PrevDoc.Effectivity.GetKey(i), PrevDoc.Effectivity.GetItem(i)
            Next
        End If
    Else
        If PrimaryDoc.Effectivity.Count > 0 Then
            For i = 1 To PrimaryDoc.Effectivity.Count
                oCL.Add PrimaryDoc.Effectivity.GetKey(i), PrimaryDoc.Effectivity.GetItem(i)
            Next
        End If
    End If
    Set Effectivity = oCL
End Property


