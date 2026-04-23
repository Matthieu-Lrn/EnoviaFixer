Attribute VB_Name = "InstanceActivation"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long

'********************************************************************************
'* Name: Activate Components
'* Purpose: Activate a multi selection of instances under a product (not allowed in native Catia). There's 2 modes
'*          Mode 1: User pre-select instances to be activated/deactivated then launch macro
'*          Mode 2: No pre-selection. User launch macro, then select an instance to be activated/deactivated.
'*                  If root product is selected, all instances under it are activated/deactivated
'*
'* Assumption: Any instance for which activity is controled by a relation will not be activated.
'*             Relation must be deactivated before running macro
'*
'* Author: Francois Charette
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub selectInstanceToActivateDeactivate(Optional ByVal bInstanceActivity As Boolean = True)
    
    If TypeName(CATIA.ActiveDocument) = "PartDocument" Then Exit Sub
    
    Set oSelection = CATIA.ActiveDocument.Selection
    
    If oSelection.Count = 0 Then
        Set oSelectedItem = MultiWindowSelectElement("No selection found. Please select instance(s) to activate...")
        If oSelectedItem Is Nothing Then Exit Sub '**** If loop is stopped without selecting anything
        oSelection.Add oSelectedItem.Value
    End If
    
    CATIA.Interactive = False 'Do not allow user to click something
    frmKBEMain.Enabled = False 'Do not allow user to click something
    
    On Error Resume Next '***** If Root product selected, select all instances under it
    If oSelection.Count = 1 And oSelection.Item(1).Value.Parent.Name = CATIA.ActiveDocument.Name Then
        For i = 1 To CATIA.ActiveDocument.Product.Products.Count
            oSelection.Add CATIA.ActiveDocument.Product.Products.Item(i)
        Next i
        oSelection.Remove (1)
    End If
    
    If bInstanceActivity = True Then Call frmProgress.progressBarInitialize("Activate Components")
    If bInstanceActivity = False Then Call frmProgress.progressBarInitialize("Deactivate Components")
    
    For i = 1 To oSelection.Count
   
        If TypeName(oSelection.Item2(i).Value) = "Product" Then
            Set oProduct = oSelection.Item2(i).Value
            If bInstanceActivity = True Then Call frmProgress.progressBarRepaint("Activating " & oSelection.Item(i).Value.Name, CDbl(oSelection.Count), CDbl(i), , , , "(" & CStr(i) & "/" & CStr(oSelection.Count) & ")")
            If bInstanceActivity = False Then Call frmProgress.progressBarRepaint("Deactivating " & oSelection.Item(i).Value.Name, CDbl(oSelection.Count), CDbl(i), , , , "(" & CStr(i) & "/" & CStr(oSelection.Count) & ")")
            Call ActivateOrDeactivate(oProduct, bInstanceActivity)
        End If

        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    Next i
       
endsub:
    On Error GoTo 0
    oSelection.Clear
    frmProgress.Hide
    CATIA.Interactive = True
    frmKBEMain.Enabled = True
    CATIA.StatusBar = "Successfully Done! "
    
End Sub

Public Sub ActivateOrDeactivate(ByVal oProduct As Product, ByVal bInstanceActivity As Boolean)
    
    On Error Resume Next
    Set oParameter = Nothing
    
    If oProduct.Parameters.SubList(oProduct, False).Count = 0 Then
        oProduct.ApplyWorkMode DEFAULT_MODE
    End If
    
    '**** get "component activation State" of the instance
    oParameter = oProduct.Parameters.SubList(oProduct, False).Item("Component Activation State")
    
    If Not oParameter Is Nothing Then
        'Find if the instance activity is controlled by a relation
        Set oRelation = Nothing
        Set oRelation = oParameter.OptionalRelation
        If oRelation Is Nothing Then oParameter.Value = bInstanceActivity
    End If
    
    On Error GoTo 0
    
End Sub


