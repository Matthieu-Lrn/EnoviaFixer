Attribute VB_Name = "SwapActivityInsertHole"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub Launch()

Dim oProduct As Product

If CATIA.ActiveDocument Is Nothing Then Exit Sub
If Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
    MsgBox ("Active window must be a CATProduct")
    Exit Sub
End If

Set oSelection = CATIA.ActiveDocument.Selection


'Log File
Call AddToLogFile("Swap Hole Activity")



If oSelection.Count > 0 Then

    For i = 1 To oSelection.Count
        
        Set oProduct = oSelection.Item(i).LeafProduct
        oProduct.ApplyWorkMode DESIGN_MODE
        
        Call SwapActivity(oProduct)
    Next
    
Else
    MsgBox ("Preselect before launching the macro")

End If

End Sub

Private Sub SwapActivity(ByVal oProduct As Product)

Dim oPart As Part
Dim oPartBody As Body
Dim oShape As Shape
Dim i As Integer
Dim oAssemble As Assemble

Set oPart = oProduct.GetMasterShapeRepresentation(False).Part
Set oPartBody = oPart.Bodies.Item("PartBody")

For i = 1 To oPartBody.Shapes.Count

    Set oShape = oPartBody.Shapes.Item(i)
    If TypeName(oShape) = "Assemble" Then
        If oShape.Body.Name Like "Inserts_*" Then
        
            If oPart.IsInactive(oShape) Then
                oPart.Activate oShape
            Else
                oPart.Inactivate oShape
            End If
        End If
    End If
    
Next

oPart.Update

End Sub


