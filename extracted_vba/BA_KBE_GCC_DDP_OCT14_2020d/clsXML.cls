Attribute VB_Name = "clsXML"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oXMLDoc As DOMDocument60

Public Function Load(ByVal sPath As String) As String

Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Check if path exist and load
If oFSO.FileExists(sPath) Then
    Call oXMLDoc.Load(sPath)
    Load = ""
Else
    Load = "File not found"
End If

End Function

Public Function XMLDoc() As DOMDocument60
Set XMLDoc = oXMLDoc
End Function
Public Function Save(ByVal sPath As String, ByVal sName As String) As String

Dim oFSO

Set oFSO = CreateObject("Scripting.FileSystemObject")

'Check if the path name ends with "\"
If Right(sPath, 1) <> "\" Then
    sPath = sPath & "\"
End If

'Check if path exist and save
If oFSO.FolderExists(sPath) Then
    Call oXMLDoc.Save(sPath & sName)
    Save = ""
Else
    Save = "Path not found"
End If

End Function

Public Function AddElement(ByVal sTagName As String, Optional ByVal oParentElem As IXMLDOMElement = Nothing, Optional ByVal oBrotherElem As IXMLDOMElement = Nothing) As IXMLDOMElement

Dim oElement As IXMLDOMElement
Set oElement = Nothing

If Not sTagName = "" Then

    'Add new element
    Set oElement = oXMLDoc.CreateElement(sTagName)

    'Append node
    If oParentElem Is Nothing And oBrotherElem Is Nothing Then
        Call oXMLDoc.appendChild(oElement)
    ElseIf Not oParentElem Is Nothing And oBrotherElem Is Nothing Then
        Call oParentElem.appendChild(oElement)
    Else
        Call oBrotherElem.parentNode.InsertBefore(oElement, oBrotherElem)
    End If

End If

'Return
Set AddElement = oElement

End Function

Public Function SortNodes(ByVal oParentElem As IXMLDOMElement, ByVal sSortAtt As String, Optional ByVal sOrder As String = "ASC")

Dim i As Integer
Dim oClone As IXMLDOMNode, oNode As IXMLDOMNode, oPrvNode As IXMLDOMNode
Dim bLoopAgain As Boolean

Do

    bLoopAgain = False
    For i = 2 To oParentElem.childNodes.Length
    
        Set oNode = oParentElem.childNodes.Item(i - 1)
        Set oPrvNode = oParentElem.childNodes.Item(i - 2)
        
        'Compare
        If sOrder = "ASC" And StrComp(oNode.Attributes.getNamedItem(sSortAtt).nodeValue, oPrvNode.Attributes.getNamedItem(sSortAtt).nodeValue, 1) < 0 Or _
           sOrder = "DESC" And StrComp(oPrvNode.Attributes.getNamedItem(sSortAtt).nodeValue, oNode.Attributes.getNamedItem(sSortAtt).nodeValue, 1) < 0 Then
            Set oClone = oNode.cloneNode(True)
            Call oParentElem.InsertBefore(oClone, oPrvNode)
            Call oNode.parentNode.RemoveChild(oNode)
            
            bLoopAgain = True
        End If
    Next

    If bLoopAgain = False Then Exit Do
Loop

End Function

Public Function Exists(ByVal sQuery As String, Optional ByVal oParent As IXMLDOMNode = Nothing) As Boolean

Dim oNodeList As IXMLDOMNodeList
Set oNodeList = Nothing

If Not sQuery = "" Then
    If oParent Is Nothing Then
        Set oNodeList = oXMLDoc.SelectNodes(sQuery)
    Else
        Set oNodeList = oParent.SelectNodes(sQuery)
    End If
End If

'Return
Exists = oNodeList.Length >= 1
End Function


Public Function AddText(ByVal sValue As String, ByVal oParentElem As IXMLDOMElement) As IXMLDOMText

Dim oText As IXMLDOMText

If Not oParentElem Is Nothing Then
    
    Set oText = oXMLDoc.createTextNode(sValue)
    Call oParentElem.appendChild(oText)

End If

'Return
Set AddText = oText

End Function
Public Sub AddSingleAttribute(ByVal oElement As IXMLDOMElement, ByVal sAttributeName As String, ByVal sAttributeValue As String)

    'Test
    If sAttributeName = "" Then Exit Sub

    'Add element
    Call oElement.setAttribute(sAttributeName, sAttributeValue)
End Sub

Public Sub AddMultipleAttributes(ByVal oElement As IXMLDOMElement, ByVal sAttributeNames As Variant, ByVal sAttributeValues As Variant)

Dim i As Integer

For i = LBound(sAttributeNames) To UBound(sAttributeNames)

    If Trim(sAttributeNames(i)) = "" Then Exit Sub

    Call oElement.setAttribute(sAttributeNames(i), sAttributeValues(i))
Next
End Sub

Public Sub EditAttributeValue(ByVal oElement As IXMLDOMElement, ByVal sAttributeName As String, ByVal sAttributeValue As String)

If Not (oElement Is Nothing Or sAttributeName = "") Then
    oElement.Attributes.getNamedItem(sAttributeName).nodeValue = sAttributeValue
End If

End Sub

Public Function GetAttributeValue(ByVal oElement As IXMLDOMElement, ByVal sAttributeName As String) As String

Dim sAttributeValue As String
sAttributeValue = ""

If Not (oElement Is Nothing Or sAttributeName = "") Then
    sAttributeValue = oElement.Attributes.getNamedItem(sAttributeName).nodeValue
End If

'Return
GetAttributeValue = sAttributeValue
End Function


Public Sub DeleteNode(ByVal oNode As IXMLDOMNode, Optional bKeepChildren = False)

Dim oChildList As IXMLDOMNodeList
Dim oChild As IXMLDOMNode
Dim oClone As IXMLDOMNode

'Don't keep children
If bKeepChildren = False Then

    Call oNode.parentNode.RemoveChild(oNode)
    
'Keep children
Else
    
    'Get all first level children
    Set oChildList = oNode.SelectNodes("./*")
    
    'Clone all the children, if we have any and insert them before the node to be deleted
    If oChildList.Length >= 1 Then
        For Each oChild In oChildList
            Set oClone = oChild.cloneNode(True)
            Call oNode.parentNode.InsertBefore(oClone, oNode)
        Next
    End If
    
    'Delele the node
    Call oNode.parentNode.RemoveChild(oNode)

End If

End Sub

Public Function SelectNodes(ByVal sQuery As String, Optional ByVal oParent As IXMLDOMNode = Nothing) As IXMLDOMNodeList

Dim oNodeList As IXMLDOMNodeList
Set oNodeList = Nothing

If Not sQuery = "" Then
    If oParent Is Nothing Then
        Set oNodeList = oXMLDoc.SelectNodes(sQuery)
    Else
        Set oNodeList = oParent.SelectNodes(sQuery)
    End If
End If

'Return
Set SelectNodes = oNodeList
End Function

Public Function SelectSingleNode(ByVal sQuery As String, Optional ByVal oParent As IXMLDOMNode = Nothing) As IXMLDOMNode

Dim oNode As IXMLDOMNode
Set oNode = Nothing

If Not sQuery = "" Then
    If oParent Is Nothing Then
        Set oNode = oXMLDoc.SelectSingleNode(sQuery)
    Else
        Set oNode = oParent.SelectSingleNode(sQuery)
    End If
End If

'Return
Set SelectSingleNode = oNode
End Function

Public Function RootNode() As IXMLDOMNode
    Set RootNode = oXMLDoc.SelectSingleNode("./*")
End Function

Private Sub Class_Initialize()
    Set oXMLDoc = New DOMDocument60
    oXMLDoc.setProperty "SelectionLanguage", "XPath"
    oXMLDoc.async = True
End Sub

Private Sub Class_Terminate()
    Set oXMLDoc = Nothing
End Sub
