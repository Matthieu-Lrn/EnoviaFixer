Attribute VB_Name = "InsertDrawingFormat"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub Launch(ByRef dPageWidth() As Double, ByRef dPageHeight() As Double, ByRef sTemplateSheetName() As String, ByVal oTemplateDoc As DrawingDocument)

Dim iArrayNumber As Integer
Dim oNewDwgDoc As DrawingDocument
Dim oSheet As DrawingSheet
Dim oDetailSheet As DrawingSheet
Dim oTextFct As New clsDwgTextFunctions

CATIA.RefreshDisplay = False

'Retreive active document
Set oNewDwgDoc = CATIA.ActiveDocument

'Retreive the detail sheet in the new drawing document
Set oDetailSheet = oNewDwgDoc.Sheets.Item("DRAWING INFO")

'Loop all the sheets
For Each oSheet In oNewDwgDoc.Sheets

    'Drawing Sheets - SH
    If Left(oSheet.Name, 2) = "SH" Then
    
        'Find array number
        iArrayNumber = CInt(Right(oSheet.Name, 2))
    
        If dPageWidth(iArrayNumber) <> 0# And dPageWidth(iArrayNumber) <> 1 Then
                    
            'Set page size
            oSheet.PaperSize = catPaperUser
            oSheet.SetPaperWidth dPageWidth(iArrayNumber)
            oSheet.SetPaperHeight dPageHeight(iArrayNumber)
            
            'Delete alll components and text in Background view
            Call EmptyBKGBView(oNewDwgDoc, oSheet)
            
            'Copy/Paste drawing format from selected oTemplateDoc
            Call Copy_Paste_Format(oNewDwgDoc, oTemplateDoc, oSheet, sTemplateSheetName(iArrayNumber))
            
            'Expose 2D Components
            Call Expose_Components(oNewDwgDoc, oSheet)
            
        ElseIf dPageWidth(iArrayNumber) = 1 Then
        
            'Delete alll components and text in Background vieq
            Call EmptyBKGBView(oNewDwgDoc, oSheet)

        End If
    End If

    'Part List Sheets - PL
    If Left(oSheet.Name, 2) = "PL" Then
        
        'Find array number
        iArrayNumber = CInt(Right(oSheet.Name, 2)) + 20
    
        If dPageWidth(iArrayNumber) <> 0# And dPageWidth(iArrayNumber) <> 1 Then
        
            'Set page size
            oSheet.PaperSize = catPaperUser
            oSheet.SetPaperWidth (11 * 25.4)
            oSheet.SetPaperHeight (8.5 * 25.4)
            
            'Delete existing components
            Call EmptyBKGBView(oNewDwgDoc, oSheet)
   
            'Copy/Paste drawing format from selected oTemplateDoc
            Call Copy_Paste_Format(oNewDwgDoc, oTemplateDoc, oSheet, sTemplateSheetName(iArrayNumber))
            
            'Delete useless text
            Call oTextFct.DeleteUselessText(oNewDwgDoc, oSheet)
            
            'Expose 2D Components
            Call Expose_Components(oNewDwgDoc, oSheet)
            
        ElseIf dPageWidth(iArrayNumber) = 1 Then
        
            'Delete alll components and text in Background vieq
            Call EmptyBKGBView(oNewDwgDoc, oSheet)
                
        End If
    End If
  
Next


'Position 2D Component in "DRAWING INFO" sheet
Call Position_Component(oNewDwgDoc)

CATIA.RefreshDisplay = True

'Message to user
MsgBox ("Manually delete the unused components in the DRAWING INFO sheet")


End Sub


Private Sub EmptyBKGBView(oNewDwgDoc As DrawingDocument, oSheet As DrawingSheet)

Dim i As Integer
Dim oSelection As Selection
Dim oComponent As DrawingComponent
Dim oBackgroundView As DrawingView

Set oSelection = oNewDwgDoc.Selection
oSelection.Clear

Set oBackgroundView = oSheet.Views.Item(2)
oBackgroundView.Activate

On Error Resume Next
'Select components
For i = 1 To oBackgroundView.Components.Count
    oSelection.Add oBackgroundView.Components.Item(i)
Next

'Select Text
For i = 1 To oBackgroundView.Texts.Count
    oSelection.Add oBackgroundView.Texts.Item(i)
Next

'Select Tables
For i = 1 To oBackgroundView.Tables.Count
    oSelection.Add oBackgroundView.Tables.Item(i)
Next


If oSelection.Count > 0 Then
    oSelection.Delete
End If
On Error GoTo 0

End Sub

Private Sub Copy_Paste_Format(oNewDwgDoc As DrawingDocument, oTemplateDoc As DrawingDocument, oSheet As DrawingSheet, ByVal sTemplateSheetName As String)

Dim oSelection
Dim i As Integer
Dim oMainView As DrawingView
Dim oBackgroundView As DrawingView
Dim oFormatSheet As DrawingSheet
Dim oNewDwgWindow As Window
Dim oFormatWindow As Window
Dim sTempString As String

Set oNewDwgWindow = CATIA.ActiveWindow

'Retrieve window containing formats
sTempString = Left(oTemplateDoc.Name, Len(oTemplateDoc.Name) - 13) 'Remove CATDrawing extension and Revision
For i = 1 To CATIA.Windows.Count
    If CATIA.Windows.Item(i).Name Like "*" & sTempString & "*" Then
        Set oFormatWindow = CATIA.Windows.Item(i)
    End If
Next

'Activate Template window
oFormatWindow.Activate
Set oSelection = oTemplateDoc.Selection
oSelection.Clear

'Retrieve in oTemplateDoc the sheet matching the selection
Set oFormatSheet = oTemplateDoc.Sheets.Item(sTemplateSheetName)
Set oBackgroundView = oFormatSheet.Views.Item(2)

'Select all text
oSelection.Add oFormatSheet
oSelection.Search "CATDrwSearch.DrwText,sel"

'Select all components
For i = 1 To oBackgroundView.Components.Count
    oSelection.Add oBackgroundView.Components.Item(i)
Next

'Copy
oSelection.Copy

'Swap back to drawing
oNewDwgWindow.Activate
Set oSelection = oNewDwgDoc.Selection
Set oMainView = oSheet.Views.Item(1)
Set oBackgroundView = oSheet.Views.Item(2)

'Paste
oSelection.Clear
oSelection.Add oBackgroundView

oSelection.Paste

oMainView.Activate

End Sub

Private Sub Expose_Components(oNewDwgDoc As DrawingDocument, oSheet As DrawingSheet)

Dim oDetailSheet As DrawingSheet
Dim oBackgroundView As DrawingView
Dim oDwgComponent As DrawingComponent
Dim i As Integer

Set oDetailSheet = oNewDwgDoc.Sheets.Item("DRAWING INFO")
Set oBackgroundView = oSheet.Views.Item(2) 'La 2e vue devrait être Background view
    
For i = 1 To oBackgroundView.Components.Count
    Set oDwgComponent = oBackgroundView.Components.Item(i)

    If InStr(oDwgComponent.Name, "Small") <> 0 _
    Or InStr(oDwgComponent.Name, "Main") <> 0 _
    Or InStr(oDwgComponent.Name, "NIEO") <> 0 _
    Or InStr(oDwgComponent.Name, "JBlock") <> 0 Then
    
        On Error Resume Next
        oDwgComponent.ExposeCompRefInSheet oDetailSheet
        On Error GoTo 0
    End If
Next

End Sub

Private Sub Position_Component(oNewDwgDoc As DrawingDocument)
'Position components in "DRAWING INFO" sheet so they are not all superimposed

Dim oDetailSheet As DrawingSheet
Dim o2DComponentView As DrawingView
Dim i As Integer

Set oSelection = oNewDwgDoc.Selection
oSelection.Clear

Set oDetailSheet = oNewDwgDoc.Sheets.Item("DRAWING INFO")

For i = 1 To oDetailSheet.Views.Count

    Set o2DComponentView = oDetailSheet.Views.Item(i)

    If Left(o2DComponentView.Name, 2) = "TB" Then
        o2DComponentView.Y = 50 * i
    End If
Next

End Sub



