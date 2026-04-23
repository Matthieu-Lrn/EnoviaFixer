Attribute VB_Name = "clsTextGenerator"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private sText As String

Private Sub Class_Initialize()
sText = ""
End Sub

Private Sub Class_Terminate()
sText = ""
End Sub

Public Sub ClearText()
sText = ""
End Sub

Property Get GetText() As String
    GetText = sText
End Property

Public Sub AddEmptyLine()

sText = sText & vbCrLf

End Sub
Public Sub AddNewLine(ByVal sArray As Variant)

Dim sNewLine As String
Dim i As Integer
Dim sString As String
Dim iPosition As Integer

sNewLine = ""

'Generate the new line
For i = LBound(sArray, 1) To UBound(sArray, 1)

    sString = sArray(i, LBound(sArray, 2))
    iPosition = sArray(i, UBound(sArray, 2))
    
    'Clean string
    sString = Replace(sString, Chr(10), "")
    sString = Replace(sString, Chr(11), "")
    sString = Replace(sString, Chr(12), "")
    sString = Replace(sString, Chr(13), "")
    
    'Add empty space
    If Len(sNewLine) < iPosition Then
        sNewLine = AddEmptySpace(sNewLine, iPosition - Len(sNewLine))
    End If
    
    'Add string
    sNewLine = sNewLine & sString
Next

'Append to sText
If sText = "" Then
    sText = sNewLine
Else
    sText = sText & vbCrLf & sNewLine
End If

End Sub

Public Sub AddLineDivider(ByVal sChar As String, ByVal iQty As Integer)

Dim sNewLine As String
Dim i As Integer

sNewLine = ""

'Generate line
If iQty >= 1 Then
    For i = 1 To iQty
        sNewLine = sNewLine & sChar
    Next
End If

'Append to sText
If sText = "" Then
    sText = sNewLine
Else
    sText = sText & vbCrLf & sNewLine
End If
End Sub

Private Function AddEmptySpace(ByVal sLine As String, ByVal iQty As Integer) As String

Dim sEmptyString As String
Dim i As Integer

sEmptyString = ""

'Generate empty string
For i = 1 To iQty
    sEmptyString = sEmptyString & " "
Next

'Return
AddEmptySpace = sLine & sEmptyString

End Function
