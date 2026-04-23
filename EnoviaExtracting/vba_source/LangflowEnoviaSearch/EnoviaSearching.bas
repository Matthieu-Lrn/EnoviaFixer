Option Explicit

Public Sub OpenTopAssemblyFromEnovia(ByVal topAssy As String, ByVal expectedRevision As String)
    If Len(Trim$(topAssy)) = 0 Then Err.Raise vbObjectError + 1000, , "topAssy is required."

    CATIA.StartCommand "ENOVIA Search..."
    WaitSeconds 2
    AppActivate "CATIA"

    SendKeys EnsureWildcard(topAssy), True
    SendKeys "{TAB}", True
    SendKeys expectedRevision, True
    SendKeys "{ENTER}", True

    WaitSeconds 2
    SendKeys "{DOWN}", True
    SendKeys "{DOWN}", True
    SendKeys "{DOWN}", True
    SendKeys "{ENTER}", True

    WaitSeconds 2
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{ENTER}", True

    WaitUntilSearchDocumentCloses 180

    WaitSeconds 2
    SendKeys "{ENTER}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "{ENTER}", True
End Sub

Private Function EnsureWildcard(ByVal value As String) As String
    value = Trim$(value)
    If Right$(value, 1) = "*" Then
        EnsureWildcard = value
    Else
        EnsureWildcard = value & "*"
    End If
End Function

Private Sub WaitSeconds(ByVal seconds As Long)
    Dim untilTime As Date
    untilTime = DateAdd("s", seconds, Now)
    Do While Now < untilTime
        DoEvents
    Loop
End Sub

Private Sub WaitUntilSearchDocumentCloses(ByVal timeoutSeconds As Long)
    Dim untilTime As Date
    untilTime = DateAdd("s", timeoutSeconds, Now)

    Do While Now < untilTime
        DoEvents
        On Error Resume Next
        If UCase$(CATIA.ActiveDocument.Name) <> UCase$("CATImmSearchDoc") Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    Loop

    Err.Raise vbObjectError + 1001, , "Timed out waiting for ENOVIA search to open the document."
End Sub

