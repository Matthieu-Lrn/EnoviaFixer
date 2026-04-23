Attribute VB_Name = "clsDwgTextFunctions"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************************************
'The following functions are used by:
' - SearchFlagNote
' - SearchFindNo
'********************************************************************************

Public Function FindZone(ByVal sPageSize As String, ByVal dXPos As Double, ByVal dYPos As Double, ByVal sSheetName As String, ByVal bMultiSheet As Boolean, ByVal dPaperWidth As Double, ByVal dPaperHeight As Double, ByVal iNbVerZone As Integer, ByVal iNbHorZone As Integer) As String

Dim sVerticalZone As String
Dim iHorizontalZone As Integer
Dim sHorizontalZone As String
Dim sSheetNb As String

'Not a Bombardier Format
If sPageSize = "Not a Bombardier Format" Then
    FindZone = "Not a Bombardier Format"
    Exit Function
End If

'Outside of drawing format
If dXPos <= 0 Or dXPos >= dPaperWidth Or dYPos <= 0 Or dYPos >= dPaperHeight Then
    FindZone = "Outside of drawing format"
    Exit Function
End If

'Format C and bigger
If Not sPageSize Like "A*" And Not sPageSize = "B" Then

    'On trouve la zone verticale
    If Int(dYPos / (dPaperHeight / iNbVerZone)) = 0 Then
        sVerticalZone = "A"
    ElseIf Int(dYPos / (dPaperHeight / iNbVerZone)) = 1 Then
        sVerticalZone = "B"
    ElseIf Int(dYPos / (dPaperHeight / iNbVerZone)) = 2 Then
        sVerticalZone = "C"
    Else
        sVerticalZone = "D"
    End If
    
    'On trouve la zone horizontale. On commence par la gauche et ca commence par la zone 0
    iHorizontalZone = Int(dXPos / (dPaperWidth / iNbHorZone))
    
    'On transforme la zone qu'on vient de trouver en zone papier qui commence par la droite et avec la zone 1
    sHorizontalZone = CStr(iNbHorZone - iHorizontalZone)

    'Multi-Sheet
    If bMultiSheet = True Then
    
        'On trouve le numéro de la sheet
        sSheetNb = Right(sSheetName, Len(sSheetName) - 2)
        If Left(sSheetNb, 1) = "0" Then
            sSheetNb = Right(sSheetNb, Len(sSheetNb) - 1)
        End If

        'Return value
        FindZone = sHorizontalZone & sVerticalZone & "-" & sSheetNb
    
    'Single Sheet
    Else
        FindZone = sHorizontalZone & sVerticalZone
    End If
    
'Format A ou B
ElseIf bMultiSheet = True And (sPageSize Like "A*" Or sPageSize = "B") Then

    'Multi-Sheet
    If bMultiSheet = True Then
    
        'Return value
        FindZone = sSheetName
    
    'Single Sheet
    Else
        FindZone = ""
    End If

End If

End Function

Public Sub FindPaperSize(ByVal dPaperWidth As Double, ByVal dPaperHeight As Double, ByRef sPageSize As String, ByRef iNbVerZone As Integer, ByRef iNbHorZone As Integer)

If Round(dPaperWidth, 1) = 215.9 And Round(dPaperHeight, 1) = 279.4 Then
    sPageSize = "A Portrait"
    iNbHorZone = 1
    iNbVerZone = 1
ElseIf Round(dPaperWidth, 1) = 279.4 And Round(dPaperHeight, 1) = 215.9 Then
    sPageSize = "A Landscape"
    iNbHorZone = 1
    iNbVerZone = 1
ElseIf Round(dPaperWidth, 1) = 431.8 And Round(dPaperHeight, 1) = 279.4 Then
    sPageSize = "B"
    iNbHorZone = 2
    iNbVerZone = 2
ElseIf Round(dPaperWidth, 1) = 558.8 And Round(dPaperHeight, 1) = 431.8 Then
    sPageSize = "C"
    iNbHorZone = 4
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 863.6 And Round(dPaperHeight, 1) = 558.8 Then
    sPageSize = "D"
    iNbHorZone = 4
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 1117.6 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "E"
    iNbHorZone = 4
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 1676.4 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J66"
    iNbHorZone = 6
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 2235.2 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J88"
    iNbHorZone = 8
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 2794 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J110"
    iNbHorZone = 10
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 3352.8 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J132"
    iNbHorZone = 12
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 3911.6 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J154"
    iNbHorZone = 14
    iNbVerZone = 4
ElseIf Round(dPaperWidth, 1) = 4470.4 And Round(dPaperHeight, 1) = 863.6 Then
    sPageSize = "J176"
    iNbHorZone = 16
    iNbVerZone = 4
Else
    sPageSize = "Not a Bombardier Format"
End If
End Sub

Public Function CheckMultiSheet(ByVal oDrawingDoc As DrawingDocument) As Boolean

Dim oSheet As DrawingSheet
Dim iCount As Integer

iCount = 0
For i = 1 To oDrawingDoc.Sheets.Count
    If Left(oDrawingDoc.Sheets.Item(i).Name, 2) = "SH" Then
        iCount = iCount + 1
    End If
    
    If iCount > 1 Then Exit For
Next

If iCount > 1 Then
    CheckMultiSheet = True
Else
    CheckMultiSheet = False
End If

End Function


Public Sub CalculateFontSize(ByRef sString As String, ByRef dFontsize As Double, ByVal iDwgTableColumn As Integer)

Select Case iDwgTableColumn
    
    Case 1, 2, 3 'Qty LH, Qty RH or Find No
        If Len(sString) <= 3 Then
            dFontsize = 0.095 * 25.4
        Else
            dFontsize = 0.075 * 25.4
        End If
        sString = Left(sString, 4)
    
    Case 4 'Zone Sheet
        If Len(sString) <= 4 Then
            dFontsize = 0.095 * 25.4
        Else
            dFontsize = 0.075 * 25.4
        End If
        sString = Left(sString, 5)
    
    Case 5 'cage Code
        If Len(sString) <= 5 Then
            dFontsize = 0.095 * 25.4
        Else
            dFontsize = 0.075 * 25.4
        End If
        sString = Left(sString, 6)
    
    Case 6 'Part Number
        If Len(sString) <= 25 Then
            dFontsize = 0.095 * 25.4
        ElseIf Len(sString) <= 28 Then
            dFontsize = 0.085 * 25.4
        ElseIf Len(sString) <= 32 Then
            dFontsize = 0.075 * 25.4
        Else
            dFontsize = 0.07 * 25.4
        End If
        sString = Left(sString, 34)
    
    Case 7 'Description
        If Len(sString) <= 29 Then
            dFontsize = 0.095 * 25.4
        ElseIf Len(sString) <= 32 Then
            dFontsize = 0.085 * 25.4
        ElseIf Len(sString) <= 37 Then
            dFontsize = 0.075 * 25.4
        Else
            dFontsize = 0.069 * 25.4
        End If
        sString = Left(sString, 40)
    
     Case 8 'Size
        If Len(sString) <= 6 Then
            dFontsize = 0.095 * 25.4
        ElseIf Len(sString) <= 7 Then
            dFontsize = 0.085 * 25.4
        Else
            dFontsize = 0.075 * 25.4
        End If
        sString = Left(sString, 8)
   
    Case 9 'Matl Spec
        If Len(sString) <= 20 Then
            dFontsize = 0.095 * 25.4
        ElseIf Len(sString) <= 22 Then
            dFontsize = 0.085 * 25.4
        ElseIf Len(sString) <= 25 Then
            dFontsize = 0.075 * 25.4
        Else
            dFontsize = 0.07 * 25.4
        End If
        sString = Left(sString, 27)
   
    Case 10 'Finish Code
        If Len(sString) <= 4 Then
            dFontsize = 0.095 * 25.4
        Else
            dFontsize = 0.075 * 25.4
        End If
        sString = Left(sString, 5)

    Case 11 'PCCN
        dFontsize = 0.095 * 25.4
        sString = Left(sString, 1)

    Case 12 'Note
        If Len(sString) <= 2 Then
            dFontsize = 0.095 * 25.4
        Else
            dFontsize = 0.07 * 25.4
        End If
        sString = Left(sString, 3)

End Select

End Sub

Public Sub DeleteUselessText(ByVal oNewDwgDoc As DrawingDocument, oSheet As DrawingSheet)

Dim oBackgroundView As DrawingView
Dim oDwgText As DrawingText
Dim oToDelete As New Collection

'*** Retrieve background view
Set oBackgroundView = oSheet.Views.Item(2)

'*** Clear Selection
Set oSelection = oNewDwgDoc.Selection
oSelection.Clear

' Retrieve in background view:
' - All text
' - PL.Line* 2D Components
oSelection.Add oBackgroundView
oSelection.Search "(CATDrwSearch.DrwText + CATDrwSearch.DrwDitto.Name_CAP=PL'.'Line*),sel"

'*** We remove from the selection
' - Text when value is "NOTES"
' - Text in the header section
If oSelection.Count > 0 Then
    For j = oSelection.Count To 1 Step -1
    
        If TypeName(oSelection.Item(j).Value) = "DrawingText" Then
            Set oDwgText = oSelection.Item(j).Value
            If Not (oDwgText.Y < 6.71 * 25.4 And oDwgText.X < 10.72 * 25.4 And Not oDwgText.Text Like "NOTES*") Then
                oSelection.Remove (j)
            End If
        Else
        End If
    Next
End If

'*** Delete the selection
If oSelection.Count > 0 Then
    oSelection.Delete
End If

End Sub
