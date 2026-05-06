Attribute VB_Name = "SFB_clsCollection"
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

Private SFB_oDict
Private Sub Class_Initialize()
    Set SFB_oDict = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set SFB_oDict = Nothing
End Sub

Public Sub InitializeWithDLLclsColObject(ByVal sClsCollfromDLL, ByRef oClsColl As SFB_clsCollection)
    If oClsColl.Count <> 0 Then oClsColl.RemoveAll
    If sClsCollfromDLL.Count > 0 Then
        For SFB_i = 1 To sClsCollfromDLL.Count
            If Not (UCase(TypeName(sClsCollfromDLL.GetItemByIndex(SFB_i))) Like "*CLSCOLLECTION") Then
                oClsColl.Add sClsCollfromDLL.GetKey(SFB_i), sClsCollfromDLL.GetItemByIndex(SFB_i)
            Else
                oClsColl.Add sClsCollfromDLL.GetKey(SFB_i), New SFB_clsCollection
                Call InitializeWithDLLclsColObject(sClsCollfromDLL.GetItemByIndex(SFB_i), oClsColl.GetItemByKey(sClsCollfromDLL.GetKey(SFB_i)))
            End If
        Next
    End If
End Sub

Public Function GetItems() As Collection
    Dim nColl As New Collection
    If SFB_oDict.Count > 0 Then
        Dim itemq: itemq = SFB_oDict.items
        For SFB_i = 0 To UBound(itemq)
            nColl.Add itemq(SFB_i)
        Next
    End If
    Set GetItems = nColl
End Function
Public Function GetKeys() As Collection
    Dim nColl As New Collection
    If SFB_oDict.Count > 0 Then
        Dim itemq: itemq = SFB_oDict.keys
        For SFB_i = 0 To UBound(itemq)
            nColl.Add itemq(SFB_i)
        Next
    End If
    Set GetKeys = nColl
End Function
Public Function GetItemByKey(ByVal Key)
    If VarType(SFB_oDict.Item(Key)) = vbObject Or TypeName(SFB_oDict.Item(Key)) = "Object" Or UCase(TypeName(SFB_oDict.Item(Key))) Like "*CLSCOLLECTION" Then
        Set GetItemByKey = SFB_oDict.Item(Key)
    Else
        GetItemByKey = SFB_oDict.Item(Key)
    End If
End Function
Public Function GetItemByIndex(ByVal index)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(SFB_oDict.Item(Key)) = vbObject Or TypeName(SFB_oDict.Item(Key)) = "Object" Or UCase(TypeName(SFB_oDict.Item(Key))) Like "*CLSCOLLECTION" Then
            Set GetItemByIndex = SFB_oDict.Item(Key)
        Else
            GetItemByIndex = SFB_oDict.Item(Key)
        End If
    End If
End Function
Public Function SetItemByIndex(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set SFB_oDict.Item(Key) = Value
        Else
            SFB_oDict.Item(Key) = Value
        End If
    End If
End Function
Public Function SetItemByKey(ByVal Key, ByVal Value)
    If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
        Set SFB_oDict.Item(Key) = Value
    Else
        SFB_oDict.Item(Key) = Value
    End If
End Function
Public Function GetIndex(ByVal Key)

    Dim keys: keys = SFB_oDict.keys
    
    index = 0
    For SFB_i = 0 To UBound(keys)
        If keys(SFB_i) = Key Then
            index = SFB_i + 1
            Exit For
        End If
    Next

    'if key is not found the return index = 0
    GetIndex = index
End Function

Public Function GetKey(ByVal index As Integer)
    Dim items: items = SFB_oDict.keys
    GetKey = items(index - 1)
End Function
Public Function Add(ByVal Key, ByVal Item)
    SFB_oDict.Add Key, Item
End Function
Public Function Exists(ByVal Key) As Boolean
    Exists = SFB_oDict.Exists(Key)
End Function
Public Function Count() As Integer
    Count = SFB_oDict.Count
End Function
Public Function RemoveByKey(ByVal Key) As Integer
    On Error GoTo 0
    SFB_oDict.Remove Key
End Function
Public Function RemoveByIndex(ByVal index) As Integer
    On Error GoTo 0
    SFB_oDict.Remove GetKey(index)
End Function
Public Function RemoveAll()
    Do Until SFB_oDict.Count = 0: SFB_oDict.Remove GetKey(1): Loop
End Function
Public Function GetItem(ByVal index)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(SFB_oDict.Item(Key)) = vbObject Or TypeName(SFB_oDict.Item(Key)) = "Object" Or UCase(TypeName(SFB_oDict.Item(Key))) Like "*CLSCOLLECTION" Then
            Set GetItem = SFB_oDict.Item(Key)
        Else
            GetItem = SFB_oDict.Item(Key)
        End If
    Else
        If VarType(SFB_oDict.Item(index)) = vbObject Or TypeName(SFB_oDict.Item(index)) = "Object" Or UCase(TypeName(SFB_oDict.Item(index))) Like "*CLSCOLLECTION" Then
            Set GetItem = SFB_oDict.Item(index)
        Else
            GetItem = SFB_oDict.Item(index)
        End If
    End If
End Function
Public Function SetItem(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set SFB_oDict.Item(Key) = Value
        Else
            SFB_oDict.Item(Key) = Value
        End If
    Else
        If VarType(Value) = vbObject Or TypeName(Value) = "Object" Or UCase(TypeName(Value)) Like "*CLSCOLLECTION" Then
            Set SFB_oDict.Item(index) = Value
        Else
            SFB_oDict.Item(index) = Value
        End If
    End If
End Function
