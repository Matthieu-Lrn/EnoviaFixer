Attribute VB_Name = "clsCollection"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************
'* SCRIPT VERSION: March 28th 2016
'*
'*-Update = Added two functions from New cls collection viz. getitem and setitem
'* When modifying this code please update in all toolbars
'********************************************************

Private oDict
Private Sub Class_Initialize()
    Set oDict = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set oDict = Nothing
End Sub

Public Sub InitializeWithDLLclsColObject(ByVal sClsCollfromDLL, ByRef oClsColl As clsCollection)
    If oClsColl.Count <> 0 Then oClsColl.RemoveAll
    If sClsCollfromDLL.Count > 0 Then
        For i = 1 To sClsCollfromDLL.Count
            If Not (UCase(TypeName(sClsCollfromDLL.GetItemByIndex(i))) Like "*CLSCOLLECTION") Then
                oClsColl.Add sClsCollfromDLL.GetKey(i), sClsCollfromDLL.GetItemByIndex(i)
            Else
                oClsColl.Add sClsCollfromDLL.GetKey(i), New clsCollection
                Call InitializeWithDLLclsColObject(sClsCollfromDLL.GetItemByIndex(i), oClsColl.GetItemByKey(sClsCollfromDLL.GetKey(i)))
            End If
        Next
    End If
End Sub

Public Function GetItems() As Collection
    Dim nColl As New Collection
    If oDict.Count > 0 Then
        Dim itemq: itemq = oDict.items
        For i = 0 To UBound(itemq)
            nColl.Add itemq(i)
        Next
    End If
    Set GetItems = nColl
End Function
Public Function GetKeys() As Collection
    Dim nColl As New Collection
    If oDict.Count > 0 Then
        Dim itemq: itemq = oDict.keys
        For i = 0 To UBound(itemq)
            nColl.Add itemq(i)
        Next
    End If
    Set GetKeys = nColl
End Function
Public Function GetItemByKey(ByVal Key)
    If VarType(oDict.Item(Key)) = vbObject Or TypeName(oDict.Item(Key)) = "Object" Or UCase(TypeName(oDict.Item(Key))) Like "*CLSCOLLECTION" Then
        Set GetItemByKey = oDict.Item(Key)
    Else
        GetItemByKey = oDict.Item(Key)
    End If
End Function
Public Function GetItemByIndex(ByVal index)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(oDict.Item(Key)) = vbObject Or TypeName(oDict.Item(Key)) = "Object" Or UCase(TypeName(oDict.Item(Key))) Like "*CLSCOLLECTION" Then
            Set GetItemByIndex = oDict.Item(Key)
        Else
            GetItemByIndex = oDict.Item(Key)
        End If
    End If
End Function
Public Function SetItemByIndex(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set oDict.Item(Key) = Value
        Else
            oDict.Item(Key) = Value
        End If
    End If
End Function
Public Function SetItemByKey(ByVal Key, ByVal Value)
    If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
        Set oDict.Item(Key) = Value
    Else
        oDict.Item(Key) = Value
    End If
End Function
Public Function GetIndex(ByVal Key)

    Dim keys: keys = oDict.keys
    
    index = 0
    For i = 0 To UBound(keys)
        If keys(i) = Key Then
            index = i + 1
            Exit For
        End If
    Next

    'if key is not found the return index = 0
    GetIndex = index
End Function

Public Function GetKey(ByVal index As Integer)
    Dim items: items = oDict.keys
    GetKey = items(index - 1)
End Function
Public Function Add(ByVal Key, ByVal Item)
    oDict.Add Key, Item
End Function
Public Function Exists(ByVal Key) As Boolean
    Exists = oDict.Exists(Key)
End Function
Public Function Count() As Integer
    Count = oDict.Count
End Function
Public Function RemoveByKey(ByVal Key) As Integer
    On Error GoTo 0
    oDict.Remove Key
End Function
Public Function RemoveByIndex(ByVal index) As Integer
    On Error GoTo 0
    oDict.Remove GetKey(index)
End Function
Public Function RemoveAll()
    Do Until oDict.Count = 0: oDict.Remove GetKey(1): Loop
End Function
Public Function GetItem(ByVal index)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(oDict.Item(Key)) = vbObject Or TypeName(oDict.Item(Key)) = "Object" Or UCase(TypeName(oDict.Item(Key))) Like "*CLSCOLLECTION" Then
            Set GetItem = oDict.Item(Key)
        Else
            GetItem = oDict.Item(Key)
        End If
    Else
        If VarType(oDict.Item(index)) = vbObject Or TypeName(oDict.Item(index)) = "Object" Or UCase(TypeName(oDict.Item(index))) Like "*CLSCOLLECTION" Then
            Set GetItem = oDict.Item(index)
        Else
            GetItem = oDict.Item(index)
        End If
    End If
End Function
Public Function SetItem(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set oDict.Item(Key) = Value
        Else
            oDict.Item(Key) = Value
        End If
    Else
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set oDict.Item(index) = Value
        Else
            oDict.Item(index) = Value
        End If
    End If
End Function
