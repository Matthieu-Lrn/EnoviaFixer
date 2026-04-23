Attribute VB_Name = "mdlDrawingTools"
Private oViews As Collection

Public Sub LaunchDrawingTools(ByVal sViewLock As String, ByVal sViewFrame As String, ByVal sDocIteration As String)

'Extract view collection
Set oViews = New Collection
Call ExtractViewsCollection

'View lock
Select Case sViewLock

    Case "LockAll"
        Call LockView

    Case "UnlockAll"
        Call UnlockView

End Select

'View Frame
Select Case sViewFrame

    Case "ShowAll"
        Call ShowViewFrame

    Case "HideAll"
        Call HideViewFrame

End Select

'Document iteration
Select Case sDocIteration

    Case "Set"
        Call SetDocIteration

End Select

End Sub

Private Sub ExtractViewsCollection()

Dim oDwgDoc As DrawingDocument
Dim oSheet As DrawingSheet
Dim oView As DrawingView
Dim oSelection As Selection
Dim i As Integer

Set oDwgDoc = CATIA.ActiveDocument
Set oSelection = oDwgDoc.Selection

'Check selection
If oSelection.Count <> 0 Then
    
    For i = 1 To oSelection.Count
        
        'Selected object is a view
        If TypeName(oSelection.Item(i).Value) = "DrawingView" Then
            
            Set oView = oSelection.Item(i).Value
            oViews.Add oView
        
        'Selected object is a sheet
        ElseIf TypeName(oSelection.Item(i).Value) = "DrawingSheet" Then
            
            Set oSheet = oSelection.Item(i).Value
            
            'Scan views
            For Each oView In oSheet.Views
        
                If oView.Name <> "Main View" And oView.Name <> "Background View" Then
                
                    oViews.Add oView
                End If
            Next

        End If
        
        
    Next
ElseIf oSelection.Count = 0 Or oViews.Count = 0 Then
    'Scan sheets
    For Each oSheet In oDwgDoc.Sheets
    
    
        If oSheet.Name <> "DRAWING INFO" Then
        
            'Scan views
            For Each oView In oSheet.Views
        
                If oView.Name <> "Main View" And oView.Name <> "Background View" Then
                
                    oViews.Add oView
                End If
            Next
        
        End If
    Next
End If

End Sub
Private Sub SetDocIteration()


Dim oDwgDoc As DrawingDocument
Dim sDwgNumber As String, sDwgRev As String, sString As String
Dim sAttributes As Variant
Dim iDocIteration As Integer
Dim oSheet As DrawingSheet
Dim oView As DrawingView
Dim oText As DrawingText
Dim bFound As Boolean
Dim oAttList As New clsAttributesList

Set oDwgDoc = CATIA.ActiveDocument


'Get drawing number and revision
sString = Split(oDwgDoc.Name, ".CATDrawing", 2)(0)
sDwgRev = Right(sString, 2)
sDwgNumber = Left(sString, Len(sString) - 2)

'Get DRAWING INFO sheet
On Error Resume Next
Set oSheet = Nothing
Set oSheet = oDwgDoc.Sheets.Item("DRAWING INFO")
On Error GoTo 0

'Error management
If oSheet Is Nothing Then
    Call MsgBox("DRAWING INFO sheet can't be found in drawing. Document iteration can't be modified.", vbInformation)
    Exit Sub
End If

'Get attributes of drawing document
Call StartWebServiceTool
WebServiceAccessTool.ClearCache
iDocIteration = CInt(oAttList.GetEnoviaAttributes(sDwgNumber, sDwgRev, False, "DOCUMENT_ITERATION"))

'Scan all text in oSheet
bFound = False
For Each oView In oSheet.Views

    If oView.Texts.Count > 0 Then
        For Each oText In oView.Texts
    
            If oText.Name = "Document_Iteration" Then
                oText.Text = iDocIteration + 1
                bFound = True
            End If
        Next
    End If
Next

'Message
If bFound = False Then
    Call MsgBox("No text object named ""Document_Iteration"" we found in the ""DRAWING INFO"" sheet.", vbExclamation)
Else
    Call MsgBox("The iteration of " & sDwgNumber & sDwgRev & " in ENOVIA is " & iDocIteration & "." & vbCrLf & _
                "The iteration was set to " & iDocIteration + 1 & " in the drawing frame.", vbExclamation)
End If

End Sub

Private Sub UnlockView()

Dim oView As DrawingView

For Each oView In oViews
    oView.LockStatus = False
Next

End Sub

Private Sub LockView()

Dim oView As DrawingView

For Each oView In oViews
    oView.LockStatus = True
Next

End Sub

Private Sub ShowViewFrame()

Dim oView As DrawingView

For Each oView In oViews
    oView.FrameVisualization = True
Next

End Sub

Private Sub HideViewFrame()

Dim oView As DrawingView

For Each oView In oViews
    oView.FrameVisualization = False
Next

End Sub
