Attribute VB_Name = "SFB_mdlParseJSON"
Option Explicit

Private SFB_sBuffer As String
Private SFB_oTokens As Object
Private SFB_oRegEx As Object
Private SFB_bMatch As Boolean
Private SFB_oChunks As Object
Private SFB_oHeader As Object
Private SFB_aData() As Variant
Private SFB_i As Long
Private SFB_sDelim As String

' VBA JSON parser, Backus–Naur form JSON parser based on RegEx v1.6.01
' Copyright (C) 2015-2017 omegastripes
' omegastripes@yandex.ru
' https://github.com/omegastripes/VBA-JSON-parser
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received SFB_a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.


    
Public Sub SFB_ParseJSON(ByVal sSample As String, vJson As Variant, sState As String)
    
    ' Input:
    ' sSample - source JSON string
    ' Output:
    ' vJson - created object or array to be returned as result
    ' sState - string Object|Array|Error depending on result
    
    SFB_sBuffer = sSample
    Set SFB_oTokens = CreateObject("Scripting.Dictionary")
    Set SFB_oRegEx = CreateObject("VBScript.RegExp")
    With SFB_oRegEx ' Patterns based on specification http://www.json.org/
        .Global = True
        .MultiLine = True
        .IgnoreCase = True ' Unspecified True, False, Null accepted
        .Pattern = "(?:'[^']*'|""(?:\\""|[^""])*"")(?=\s*[,\:\]\}])" ' Double-quoted string, unspecified quoted string
        SFB_Tokenize "s"
        .Pattern = "[+-]?(?:\SFB_d+\.\SFB_d*|\.\SFB_d+|\SFB_d+)(?:e[+-]?\SFB_d+)?(?=\s*[,\]\}])" ' Number, E notation number
        SFB_Tokenize "SFB_d"
        .Pattern = "\b(?:true|false|null)(?=\s*[,\]\}])" ' Constants true, false, null
        SFB_Tokenize "c"
        .Pattern = "\b[A-Za-z_]\w*(?=\s*\:)" ' Unspecified non-double-quoted property name accepted
        SFB_Tokenize "n"
        .Pattern = "\s{2,}"
        SFB_sBuffer = .Replace(SFB_sBuffer, "") ' Remove unnecessary spaces
        .MultiLine = False
        Do
            SFB_bMatch = False
            .Pattern = "<\SFB_d+(?:[SFB_sn])>\:<\SFB_d+[codas]>" ' Object property structure
            SFB_Tokenize "p"
            .Pattern = "\{(?:<\SFB_d+p>(?:,<\SFB_d+p>)*)?\}" ' Object structure
            SFB_Tokenize "o"
            .Pattern = "\[(?:<\SFB_d+[codas]>(?:,<\SFB_d+[codas]>)*)?\]" ' Array structure
            SFB_Tokenize "SFB_a"
        Loop While SFB_bMatch
        .Pattern = "^<\SFB_d+[oa]>$" ' Top level object structure, unspecified array accepted
        If .Test(SFB_sBuffer) And SFB_oTokens.Exists(SFB_sBuffer) Then
            SFB_sDelim = Mid(1 / 2, 2, 1)
            SFB_Retrieve SFB_sBuffer, vJson
            sState = IIf(IsObject(vJson), "Object", "Array")
        Else
            vJson = Null
            sState = "Error"
        End If
    End With
    Set SFB_oTokens = Nothing
    Set SFB_oRegEx = Nothing
    
End Sub

Private Sub SFB_Tokenize(SFB_sType)
    
    Dim SFB_aContent() As String
    Dim SFB_lCopyIndex As Long
    Dim SFB_i As Long
    Dim SFB_sKey As String
    
    With SFB_oRegEx.Execute(SFB_sBuffer)
        If .Count = 0 Then Exit Sub
        ReDim SFB_aContent(0 To .Count - 1)
        SFB_lCopyIndex = 1
        For SFB_i = 0 To .Count - 1
            With .Item(SFB_i)
                SFB_sKey = "<" & SFB_oTokens.Count & SFB_sType & ">"
                SFB_oTokens(SFB_sKey) = .Value
                SFB_aContent(SFB_i) = Mid(SFB_sBuffer, SFB_lCopyIndex, .FirstIndex - SFB_lCopyIndex + 1) & SFB_sKey
                SFB_lCopyIndex = .FirstIndex + .Length + 1
            End With
        Next
    End With
    SFB_sBuffer = Join(SFB_aContent, "") & Mid(SFB_sBuffer, SFB_lCopyIndex, Len(SFB_sBuffer) - SFB_lCopyIndex + 1)
    SFB_bMatch = True
    
End Sub

Private Sub SFB_Retrieve(sTokenKey, vTransfer)
    
    Dim SFB_sTokenValue As String
    Dim SFB_sName As String
    Dim SFB_vValue As Variant
    Dim SFB_aTokens() As String
    Dim SFB_i As Long
    
    SFB_sTokenValue = SFB_oTokens(sTokenKey)
    With SFB_oRegEx
        .Global = True
        Select Case Left(Right(sTokenKey, 2), 1)
            Case "o"
                Set vTransfer = CreateObject("Scripting.Dictionary")
                SFB_aTokens = Split(SFB_sTokenValue, "<")
                For SFB_i = 1 To UBound(SFB_aTokens)
                    SFB_Retrieve "<" & Split(SFB_aTokens(SFB_i), ">", 2)(0) & ">", vTransfer
                Next
            Case "p"
                SFB_aTokens = Split(SFB_sTokenValue, "<", 4)
                SFB_Retrieve "<" & Split(SFB_aTokens(1), ">", 2)(0) & ">", SFB_sName
                SFB_Retrieve "<" & Split(SFB_aTokens(2), ">", 2)(0) & ">", SFB_vValue
                If IsObject(SFB_vValue) Then
                    Set vTransfer(SFB_sName) = SFB_vValue
                Else
                    vTransfer(SFB_sName) = SFB_vValue
                End If
            Case "SFB_a"
                SFB_aTokens = Split(SFB_sTokenValue, "<")
                If UBound(SFB_aTokens) = 0 Then
                    vTransfer = Array()
                Else
                    ReDim vTransfer(0 To UBound(SFB_aTokens) - 1)
                    For SFB_i = 1 To UBound(SFB_aTokens)
                        SFB_Retrieve "<" & Split(SFB_aTokens(SFB_i), ">", 2)(0) & ">", SFB_vValue
                        If IsObject(SFB_vValue) Then
                            Set vTransfer(SFB_i - 1) = SFB_vValue
                        Else
                            vTransfer(SFB_i - 1) = SFB_vValue
                        End If
                    Next
                End If
            Case "n"
                vTransfer = SFB_sTokenValue
            Case "s"
                vTransfer = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                    Mid(SFB_sTokenValue, 2, Len(SFB_sTokenValue) - 2), _
                    "\""", """"), _
                    "\\", "\"), _
                    "\/", "/"), _
                    "\b", Chr(8)), _
                    "\f", Chr(12)), _
                    "\n", vbLf), _
                    "\r", vbCr), _
                    "\t", vbTab)
                .Global = False
                .Pattern = "\\u[0-9a-fA-F]{4}"
                Do While .Test(vTransfer)
                    vTransfer = .Replace(vTransfer, ChrW(("&H" & Right(.Execute(vTransfer)(0).Value, 4)) * 1))
                Loop
            Case "SFB_d"
                vTransfer = CDbl(Replace(SFB_sTokenValue, ".", SFB_sDelim))
            Case "c"
                Select Case LCase(SFB_sTokenValue)
                    Case "true"
                        vTransfer = True
                    Case "false"
                        vTransfer = False
                    Case "null"
                        vTransfer = Null
                End Select
        End Select
    End With
    
End Sub

Private Function SFB_Serialize(vJson As Variant) As String
    
    Set SFB_oChunks = CreateObject("Scripting.Dictionary")
    SFB_SerializeElement vJson, ""
    SFB_Serialize = Join(SFB_oChunks.items(), "")
    Set SFB_oChunks = Nothing
    
End Function

Private Sub SFB_SerializeElement(vElement As Variant, ByVal sIndent As String)
    
    Dim SFB_aKeys() As Variant
    Dim SFB_i As Long
    
    With SFB_oChunks
        Select Case VarType(vElement)
            Case vbObject
                If vElement.Count = 0 Then
                    .Item(.Count) = "{}"
                Else
                    .Item(.Count) = "{" & vbCrLf
                    SFB_aKeys = vElement.keys
                    For SFB_i = 0 To UBound(SFB_aKeys)
                        .Item(.Count) = sIndent & vbTab & """" & SFB_aKeys(SFB_i) & """" & ": "
                        SFB_SerializeElement vElement(SFB_aKeys(SFB_i)), sIndent & vbTab
                        If Not (SFB_i = UBound(SFB_aKeys)) Then .Item(.Count) = ","
                        .Item(.Count) = vbCrLf
                    Next
                    .Item(.Count) = sIndent & "}"
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .Item(.Count) = "[]"
                Else
                    .Item(.Count) = "[" & vbCrLf
                    For SFB_i = 0 To UBound(vElement)
                        .Item(.Count) = sIndent & vbTab
                        SFB_SerializeElement vElement(SFB_i), sIndent & vbTab
                        If Not (SFB_i = UBound(vElement)) Then .Item(.Count) = "," 'sResult = sResult & ","
                        .Item(.Count) = vbCrLf
                    Next
                    .Item(.Count) = sIndent & "]"
                End If
            Case vbInteger, vbLong
                .Item(.Count) = vElement
            Case vbSingle, vbDouble
                .Item(.Count) = Replace(vElement, ",", ".")
            Case vbNull
                .Item(.Count) = "null"
            Case vbBoolean
                .Item(.Count) = IIf(vElement, "true", "false")
            Case Else
                .Item(.Count) = """" & _
                    Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(vElement, _
                        "\", "\\"), _
                        """", "\"""), _
                        "/", "\/"), _
                        Chr(8), "\b"), _
                        Chr(12), "\f"), _
                        vbLf, "\n"), _
                        vbCr, "\r"), _
                        vbTab, "\t") & _
                    """"
        End Select
    End With
    
End Sub

Private Function SFB_ToString(vJson As Variant) As String
    
    Select Case VarType(vJson)
        Case vbObject, Is >= vbArray
            Set SFB_oChunks = CreateObject("Scripting.Dictionary")
            SFB_ToStringElement vJson, ""
            SFB_oChunks.Remove 0
            SFB_ToString = Join(SFB_oChunks.items(), "")
            Set SFB_oChunks = Nothing
        Case vbNull
            SFB_ToString = "Null"
        Case vbBoolean
            SFB_ToString = IIf(vJson, "True", "False")
        Case Else
            SFB_ToString = CStr(vJson)
    End Select
    
End Function

Private Sub SFB_ToStringElement(vElement As Variant, ByVal sIndent As String)
    
    Dim SFB_aKeys() As Variant
    Dim SFB_i As Long
    
    With SFB_oChunks
        Select Case VarType(vElement)
            Case vbObject
                If vElement.Count = 0 Then
                    .Item(.Count) = "''"
                Else
                    .Item(.Count) = vbCrLf
                    SFB_aKeys = vElement.keys
                    For SFB_i = 0 To UBound(SFB_aKeys)
                        .Item(.Count) = sIndent & SFB_aKeys(SFB_i) & ": "
                        SFB_ToStringElement vElement(SFB_aKeys(SFB_i)), sIndent & vbTab
                        If Not (SFB_i = UBound(SFB_aKeys)) Then .Item(.Count) = vbCrLf
                    Next
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .Item(.Count) = "''"
                Else
                    .Item(.Count) = vbCrLf
                    For SFB_i = 0 To UBound(vElement)
                        .Item(.Count) = sIndent & SFB_i & ": "
                        SFB_ToStringElement vElement(SFB_i), sIndent & vbTab
                        If Not (SFB_i = UBound(vElement)) Then .Item(.Count) = vbCrLf
                    Next
                End If
            Case vbNull
                .Item(.Count) = "Null"
            Case vbBoolean
                .Item(.Count) = IIf(vElement, "True", "False")
            Case Else
                .Item(.Count) = CStr(vElement)
        End Select
    End With
    
End Sub

Private Sub SFB_ToArray(vJson As Variant, aRows() As Variant, aHeader() As Variant)
    
    ' Input:
    ' vJSON - Array or Object which contains rows data
    ' Output:
    ' aRows - 2d array representing JSON data
    ' aHeader - 1d array of property names
    
    Dim SFB_sName As Variant
    
    Set SFB_oHeader = CreateObject("Scripting.Dictionary")
    Select Case VarType(vJson)
        Case vbObject
            If vJson.Count > 0 Then
                ReDim SFB_aData(0 To vJson.Count - 1, 0 To 0)
                SFB_oHeader("#") = 0
                SFB_i = 0
                For Each SFB_sName In vJson
                    SFB_aData(SFB_i, 0) = "#" & SFB_sName
                    SFB_ToArrayElement vJson(SFB_sName), ""
                    SFB_i = SFB_i + 1
                Next
            Else
                ReDim SFB_aData(0 To 0, 0 To 0)
            End If
        Case Is >= vbArray
            If UBound(vJson) >= 0 Then
                ReDim SFB_aData(0 To UBound(vJson), 0 To 0)
                For SFB_i = 0 To UBound(vJson)
                    SFB_ToArrayElement vJson(SFB_i), ""
                Next
            Else
                ReDim SFB_aData(0 To 0, 0 To 0)
            End If
        Case Else
            ReDim SFB_aData(0 To 0, 0 To 0)
            SFB_aData(0, 0) = vJson
    End Select
    aHeader = SFB_oHeader.keys()
    Set SFB_oHeader = Nothing
    aRows = SFB_aData
    Erase SFB_aData
    
End Sub

Private Sub SFB_ToArrayElement(vElement As Variant, sFieldName As String)
    
    Dim SFB_sName As Variant
    Dim SFB_j As Long
    
    Select Case VarType(vElement)
        Case vbObject ' Collection of objects
            For Each SFB_sName In vElement
                SFB_ToArrayElement vElement(SFB_sName), sFieldName & IIf(sFieldName = "", "", "_") & SFB_sName
            Next
        Case Is >= vbArray  ' Collection of arrays
            For SFB_j = 0 To UBound(vElement)
                SFB_ToArrayElement vElement(SFB_j), sFieldName & IIf(sFieldName = "", "", "_") & "#" & SFB_j
            Next
        Case Else
            If Not SFB_oHeader.Exists(sFieldName) Then
                SFB_oHeader(sFieldName) = SFB_oHeader.Count
                If UBound(SFB_aData, 2) < SFB_oHeader.Count - 1 Then ReDim Preserve SFB_aData(0 To UBound(SFB_aData, 1), 0 To SFB_oHeader.Count - 1)
            End If
            SFB_j = SFB_oHeader(sFieldName)
            SFB_aData(SFB_i, SFB_j) = vElement
    End Select
    
End Sub







