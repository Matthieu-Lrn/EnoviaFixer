Attribute VB_Name = "FillFindNb"
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
Dim iTextNumber As Integer
Dim iSheetNb As Integer
Dim sSheetNb As String

Dim dViewX As Double
Dim dViewY As Double
Dim dViewScale As Double
Dim dViewAngle As Double

Dim oText As DrawingText
Dim dTextX As Double
Dim dTextY As Double

Dim dSheetCoordX As Double
Dim dSheetCoordY As Double

Dim sZone As String
Dim sKey As String
Dim sAnswer As Integer

Dim sString As String
Dim sWarningNonNumeric As String
Dim sWarningPageSize As String
Dim sWarningOutsideFrame As String

Dim bMultiSheet As Boolean
Dim oDwgTable As DrawingTable
Dim oPLTable As DrawingTable

Dim oDwgSheetsFindNo As New clsCollection
Dim oPLFindNo As New clsCollection
Dim oItem As clsCollection

Dim dFontsize As Double
Dim sCellValue As String
Dim oDwgText As DrawingText

Dim oTextFct As New clsDwgTextFunctions
Dim dPaperWidth As Double
Dim dPaperHeight As Double
Dim iNbHorZone As Integer
Dim iNbVerZone As Integer
Dim sPageSize As String

'Progress Form
Call frmProgress.progressBarInitialize("Searching for " & sSearchObject)

'On retrouve le drawing document actif
Set oDrawingDocument = CATIA.ActiveDocument

'Log File
Call AddToLogFile("Fill Find Number", oDrawingDocument.Name)

'We check if we have more than one SH sheet
bMultiSheet = oTextFct.CheckMultiSheet(oDrawingDocument)

'Initialize
sWarningPageSize = ""
sWarningNonNumeric = ""
sWarningOutsideFrame = ""

'On loop toutes les sheets (SH) et on extrait la position des texts
For iSheetNb = 1 To oDrawingDocument.Sheets.Count

    'Sheet info
    Set oSheet = oDrawingDocument.Sheets.Item(iSheetNb)
    dPaperWidth = oSheet.GetPaperWidth
    dPaperHeight = oSheet.GetPaperHeight
    Call oTextFct.FindPaperSize(dPaperWidth, dPaperHeight, sPageSize, iNbVerZone, iNbHorZone)

    'Check Page size
    If sPageSize = "Not a Bombardier Format" Then
        sString = " - " & oSheet.Name
        If sWarningPageSize = "" Then
            sWarningPageSize = sString
        Else
            sWarningPageSize = sWarningPageSize & vbCrLf & sString
        End If
        
        GoTo TheNextSheet
    End If

    
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
            
            
            'Check for visible text and balllon objects
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
                If oText.FrameType = 3 Or oText.FrameType = 53 Then
                
                    'On trouve les coordonnées du texte dans la sheet
                    dSheetCoordX = dViewX + (dTextX * Cos(dViewAngle) - dTextY * Sin(dViewAngle)) * dViewScale
                    dSheetCoordY = dViewY + (dTextX * Sin(dViewAngle) + dTextY * Cos(dViewAngle)) * dViewScale
                    
                    'On trouve la zone du text
                    sZone = oTextFct.FindZone(sPageSize, dSheetCoordX, dSheetCoordY, oSheet.Name, bMultiSheet, dPaperWidth, dPaperHeight, iNbVerZone, iNbHorZone)
                    
                    'Si le text est numérique
                    If IsNumeric(oText.Text) Then
                                                                    
                             
                        'We add to oDwgSheetsFindNo
                        sKey = "FindNo_" & oText.Text
                        If Not oDwgSheetsFindNo.Exists(sKey) Then
                            Set oItem = New clsCollection
                            Call oItem.Add("Zone", sZone)
                            Call oItem.Add("Sheet Name", oSheet.Name)

                            Call oDwgSheetsFindNo.Add(sKey, oItem)
                        ElseIf oDwgSheetsFindNo.GetItemByKey(sKey).GetItemByKey("Zone") <> sZone Then
                            Call oDwgSheetsFindNo.GetItemByKey(sKey).SetItemByKey("Zone", "VAR")
                        End If


                    'Si le text n'est pas numérique
                    Else
                        'et que le text ne contient pas 'FS', 'BL' ou 'WL'
                        If Not (oText.Text Like "*FS*" Or oText.Text Like "*WL*" Or oText.Text Like "*BL*") Then
                            
                            sString = " - '" & oText.Text & "' was found in zone " & sZone
                            If sWarningNonNumeric = "" Then
                                sWarningNonNumeric = sString
                            Else
                                sWarningNonNumeric = sWarningNonNumeric & vbCrLf & sString
                            End If
                            
                        End If
                    End If
                    
                End If
            Next
        Next
    End If
    
TheNextSheet:
Next


'Scan all sheets and Retrieve all Find No in Part List
For iSheetNb = 1 To oDrawingDocument.Sheets.Count
    
    Set oSheet = oDrawingDocument.Sheets.Item(iSheetNb)
    
    'On veut juste les sheets dont le nom commence par "PL"
    If Left(oSheet.Name, 2) = "PL" Then

        'Get the background view
        Set oSheet = oDrawingDocument.Sheets.Item(iSheetNb)
        Set oDrawingView = oSheet.Views.Item(2)
        
        'Get "PL_Table"
        Set oDwgTable = Nothing
        Set PLTable = Nothing
        If oDrawingView.Tables.Count > 0 Then
            For i = 1 To oDrawingView.Tables.Count
                Set oDwgTable = oDrawingView.Tables.Item(i)
                If oDwgTable.Name = "PL_Table" Then
                    Set PLTable = oDwgTable
                    Exit For
                End If
            Next
        End If
        
        'Extract Find No from "PL_Table"
        If Not PLTable Is Nothing Then
            
            'Scan "PL_Table"
            For i = 1 To PLTable.NumberOfRows
            
                'Get the Find No value from 'PL_Table'
                sString = Trim(PLTable.GetCellString(i, 3))
                               
                'Find No <> ""
                If sString <> "" Then
                
                    Set oItem = New clsCollection
                    Call oItem.Add("Sheet Name", oSheet.Name)
                    Call oItem.Add("PL Table Line", i)
                    sKey = "FindNo_" & sString
                    
                    If Not oPLFindNo.Exists(sKey) Then
                        Call oPLFindNo.Add(sKey, oItem)
                    Else
                        Call MsgBox("Find number " & sString & " can be found twice in the PL. Please correct the PL and run the tool again", vbCritical)
                        GoTo endsub
                    End If
                End If

                'Clear zone sheet in PL_Table
                Call PLTable.SetCellString(i, 4, "")
            Next
        Else
            sAnswer = MsgBox("The 'PL_Table' object was not found in sheet " & oSheet.Name & "." & vbCrLf & "Zone sheet can't be retrieve for this sheet.", vbExclamation)
        End If
    End If
Next

'Transfer Zone to "PL_Table"
For i = oDwgSheetsFindNo.Count To 1 Step -1

    
    If oPLFindNo.Exists(oDwgSheetsFindNo.GetKey(i)) Then
           
        'Get sheet and Background view
        Set oSheet = oDrawingDocument.Sheets.Item(oPLFindNo.GetItemByKey(oDwgSheetsFindNo.GetKey(i)).GetItemByKey("Sheet Name"))
        Set oDrawingView = oSheet.Views.Item(2)
        
        'Get "PL_Table"
        For j = 1 To oDrawingView.Tables.Count
            Set oDwgTable = oDrawingView.Tables.Item(j)
            If oDwgTable.Name = "PL_Table" Then
                Exit For
            End If
        Next

        'Retrieve zone value
        If oDwgSheetsFindNo.GetItemByIndex(i).GetItemByKey("Zone") = "Outside of drawing format" Then
            sCellValue = ""
            
            sString = " - " & Split(oDwgSheetsFindNo.GetKey(i), "_")(1) & " on " & oSheet.Name
            If sWarningOutsideFrame = "" Then
                sWarningOutsideFrame = sString
            Else
                sWarningOutsideFrame = sWarningOutsideFrame & vbCrLf & sString
            End If

        Else
            sCellValue = oDwgSheetsFindNo.GetItemByIndex(i).GetItemByKey("Zone")
        End If

        'Calculate Font Size
        Call oTextFct.CalculateFontSize(sCellValue, dFontsize, 4)
         
        'Add Zone Sheet in "PL_Table"
        iLine = oPLFindNo.GetItemByKey(oDwgSheetsFindNo.GetKey(i)).GetItemByKey("PL Table Line")
        Call oDwgTable.SetCellString(iLine, 4, sCellValue)
       
        
        'Retrieve text in "PL_Table" and apply Font Size
        Set oDwgText = oDwgTable.GetCellObject(iLine, 4)

        oDwgText.SetFontSize 0, 0, dFontsize
        oDwgText.SetFontName 0, 0, "SSS1"

        'When we have a match we remove the items from the collections
        Call oPLFindNo.RemoveByKey(oDwgSheetsFindNo.GetKey(i))
        Call oDwgSheetsFindNo.RemoveByIndex(i)

    End If
Next
    
frmProgress.Hide

'Display sheet dimensions error
If sWarningPageSize <> "" Then
    sAnswer = MsgBox("The dimensions on the following sheets are not as per Bombardier standard:" & vbCrLf & sWarningPageSize, vbExclamation)
End If

'Display non-numeric warning
If sWarningNonNumeric <> "" Then
    sAnswer = MsgBox("The following non-numeric Find Number were found:" & vbCrLf & sWarningNonNumeric, vbExclamation)
End If

'Display outside frame warning
If sWarningOutsideFrame <> "" Then
    sAnswer = MsgBox("The following Find Number were found outside the sheet format:" & vbCrLf & sWarningOutsideFrame, vbExclamation)
End If

'Display Find No not found in PL
If oDwgSheetsFindNo.Count > 0 Then
    sWarning = ""
    For i = 1 To oDwgSheetsFindNo.Count
        sString = " - Find No " & Split(oDwgSheetsFindNo.GetKey(i), "_")(1) & " at location " & oDwgSheetsFindNo.GetItemByIndex(i).GetItemByKey("Zone")
    
        If sWarning = "" Then
            sWarning = sString
        Else
            sWarning = sWarning & vbCrLf & sString
        End If
    Next

    sAnswer = MsgBox("The following Find Number were not found in the Part List:" & vbCrLf & sWarning, vbExclamation)
End If

'Display Find No not found in SH
If oPLFindNo.Count > 0 Then
    sWarning = ""
    For i = 1 To oPLFindNo.Count
          
    
        sString = " - Find No " & Split(oPLFindNo.GetKey(i), "_")(1) & " in " & oPLFindNo.GetItemByIndex(i).GetItemByKey("Sheet Name")
    
        If sWarning = "" Then
            sWarning = sString
        Else
            sWarning = sWarning & vbCrLf & sString
        End If
    Next

    sAnswer = MsgBox("The following Find Number were not found in the drawing sheets:" & vbCrLf & sWarning, vbExclamation)
End If

endsub:


Call frmKBEMain.resetToolbar
End Sub





