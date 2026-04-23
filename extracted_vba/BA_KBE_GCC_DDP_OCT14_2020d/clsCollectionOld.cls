Attribute VB_Name = "clsCollectionOld"
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
    If VarType(oDict(Key)) = vbObject Then
        Set GetItemByKey = oDict(Key)
    Else
        GetItemByKey = oDict(Key)
    End If
End Function
Public Function GetItemByIndex(ByVal index)
    If IsNumeric(index) Then
        Dim items: items = oDict.items
        If VarType(items(index - 1)) = vbObject Then
            Set GetItemByIndex = items(index - 1)
        Else
            GetItemByIndex = items(index - 1)
        End If
    End If
End Function
Public Function SetItemByIndex(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        oDict.Item(Key) = Value
    End If
End Function
Public Function SetItemByKey(ByVal Key, ByVal Value)
    If VarType(Value) = vbObject Then
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
        Dim items: items = oDict.items
        If VarType(items(index - 1)) = vbObject Then
            Set GetItem = items(index - 1)
        Else
            GetItem = items(index - 1)
        End If
    Else
        If VarType(oDict(index)) = vbObject Then
            Set GetItem = oDict(index)
        Else
            GetItem = oDict(index)
        End If
    End If
End Function
Public Function SetItem(ByVal index, ByVal Value)
    If IsNumeric(index) Then
        Key = GetKey(index)
        oDict.Item(Key) = Value
    Else
        If VarType(Value) = vbObject Then
            Set oDict.Item(index) = Value
        Else
            oDict.Item(index) = Value
        End If
    End If
End Function
