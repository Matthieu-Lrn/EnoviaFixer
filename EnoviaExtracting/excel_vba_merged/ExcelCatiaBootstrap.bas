Attribute VB_Name = "ExcelCatiaBootstrap"
Option Explicit

Public CATIA As Object

Public Sub EnsureCatiaSession()
    If CATIA Is Nothing Then
        On Error Resume Next
        Set CATIA = GetObject(, "CATIA.Application")
        On Error GoTo 0
    End If

    If CATIA Is Nothing Then
        Err.Raise vbObjectError + 4000, , "Could not connect to a running CATIA.Application session."
    End If
End Sub

Public Function BuildServerPath(ByVal rootPath As String, ByVal fileName As String) As String
    Dim normalizedRoot As String

    normalizedRoot = Trim$(rootPath)
    If Right$(normalizedRoot, 1) <> "\" Then
        normalizedRoot = normalizedRoot & "\"
    End If

    BuildServerPath = normalizedRoot & fileName
End Function
