Attribute VB_Name = "SearchFindNo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Sub Launch()

Dim oDrawingDocument As DrawingDocument
Dim oSheet As DrawingSheet
Dim oDrawingView As DrawingView
Dim iViewNumber As Integer
Dim i2DCompNumber As Integer
Dim iSheetNb As Integer
Dim sSheetNb As String

Dim dViewX As Double
Dim dViewY As Double
Dim dViewScale As Double
Dim dViewAngle As Double

Dim oText As DrawingText
Dim o2DComponent As DrawingComponent
Dim dTextX As Double
Dim dTextY As Double

Dim dSheetCoordX As Double
Dim dSheetCoordY As Double

Dim sZone As String
Dim sKey As String
Dim sAnswer As Integer

Dim sString As String
Dim sLocations As String
Dim sSearchFindNb As String

Dim oShearchedText As New clsCollection
Dim oItem As clsCollection

Dim oTextFct As New clsDwgTextFunctions
Dim dPaperWidth As Double
Dim dPaperHeight As Double
Dim iNbHorZone As Integer
Dim iNbVerZone As Integer
Dim sPageSize As String


'Ask user to key in the find number he is looking for
sSearchFindNb = Trim(InputBox("Key in the Find Number you are looking for.", "Find Number Search"))
If sSearchFindNb = "" Then GoTo endsub

'Progress Form
Call frmProgress.progressBarInitialize("Searching for " & sSearchObject)

'On retrouve le drawing document actif
Set oDrawingDocument = CATIA.ActiveDocument

'On loop toutes les sheets (SH) et on extrait la position des texts
For iSheetNb = 1 To oDrawingDocument.Sheets.Count

    'Sheet info
    Set oSheet = oDrawingDocument.Sheets.Item(iSheetNb)
    dPaperWidth = oSheet.GetPaperWidth
    dPaperHeight = oSheet.GetPaperHeight
    Call oTextFct.FindPaperSize(dPaperWidth, dPaperHeight, sPageSize, iNbVerZone, iNbHorZone)

    
    'On veut juste les sheets dont le nom commence par "SH"
    If Left(oSheet.Name, 2) = "SH" Then
    
        'On loop toutes les vues
        For iViewNumber = 1 To oSheet.Views.Count
            
            Set oDrawingView = oSheet.Views.Item(iViewNumber)
 
            'Progress form
            Call frmProgress.progressBarRepaint("Scanning in: " & oSheet.Name & " / " & oDrawingView.Name)
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
                GoTo endsub
            End If
            
            'On retrouve la position de la vue dans la sheet
            dViewX = oDrawingView.xAxisData
            dViewY = oDrawingView.yAxisData
            dViewScale = oDrawingView.Scale
            dViewAngle = oDrawingView.Angle
            
            
            'Check for visible text objects
            Set oSelection = oDrawingDocument.Selection
            oSelection.Clear
            oSelection.Add oDrawingView
            oSelection.Search "(CATDrwSearch.DrwBalloon.Visibility=Visible + CATDrwSearch.DrwText.Visibility=Visible),sel"
            
            'On loop tous les textes dans la sélection
            For iTextNumber = 1 To oSelection.Count
    
                Set oText = oSelection.Item(iTextNumber).Value
                dTextX = oText.X
                dTextY = oText.Y
                
                'On trouve la position des texte:
                '- Si le frame correspond au frame circulaire ou circulaire lock
                '- Si le text est numerique
                If oText.FrameType = 3 Or oText.FrameType = 53 Or oText.GetParameterOnSubString(catBorder, 0, 0) = 3 Then
                
                    'On trouve les coordonnées du texte dans la sheet
                    dSheetCoordX = dViewX + (dTextX * Cos(dViewAngle) - dTextY * Sin(dViewAngle)) * dViewScale
                    dSheetCoordY = dViewY + (dTextX * Sin(dViewAngle) + dTextY * Cos(dViewAngle)) * dViewScale
                    
                    'On trouve la zone du text
                    sZone = oTextFct.FindZone(sPageSize, dSheetCoordX, dSheetCoordY, oSheet.Name, True, dPaperWidth, dPaperHeight, iNbVerZone, iNbHorZone)
                                        
                    'On ajoute le text a oShearchedText
                    Set oItem = New clsCollection
                    Call oItem.Add("Find Number", Trim(oText.Text))
                    Call oItem.Add("Zone", sZone)
                    Call oShearchedText.Add(oShearchedText.Count + 1, oItem)
                End If
            Next
        Next
    End If
Next


frmProgress.Hide

sLocations = ""
sString = ""

'Scan all oShearchedText
If oShearchedText.Count <> 0 Then
    For i = 1 To oShearchedText.Count
    
        'We have a match
        If oShearchedText.GetItemByIndex(i).GetItemByKey("Find Number") = sSearchFindNb Then
            sString = " - " & oShearchedText.GetItemByIndex(i).GetItemByKey("Zone")
            If sLocations = "" Then
                sLocations = sString
            Else
                sLocations = sLocations & vbCrLf & sString
            End If
        End If
    Next
        
    'Display results
    If sLocations <> "" Then
        sAnswer = MsgBox("Find Number " & sSearchFindNb & " was found at the following locations:" & vbCrLf & sLocations & vbCrLf & "Results were transfered to C:\Temp\SearchDraftingObjectsOutout.txt", vbInformation)
        
        'Transfer data to log file
        Call GenerateLogFile("Find Number " & sSearchFindNb & " was found at the following locations:" & vbCrLf & sLocations)

    Else
        sAnswer = MsgBox("Find Number " & sSearchFindNb & " was not found in any of the drawing sheets.", vbExclamation)
    End If
    
Else

    sAnswer = MsgBox("No Find Number were found in any of the drawing sheets.", vbExclamation)
End If
    

endsub:

Call frmKBEMain.resetToolbar
End Sub

Private Sub GenerateLogFile(ByVal sText As String)
    Set oLogFile = CATIA.FileSystem.CreateFile("C:\Temp\SearchDraftingObjectsOutput.txt", True)
    Set oLogStream = oLogFile.OpenAsTextStream("ForWriting")
    oLogStream.Write (sText)
    oLogStream.Close
End Sub




