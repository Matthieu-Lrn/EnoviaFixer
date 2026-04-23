Attribute VB_Name = "ResetDraftingColors"
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
Dim oVisProperty As VisPropertySet
Dim oGreen2DCompColl As Collection
Dim oOther2DCompColl As Collection
Dim oTextColl As Collection
Dim sGreen2DCompName(8) As String
Dim oDwgTable As DrawingTable
Dim oText As DrawingText
Dim sLog As String
Dim bGreen2DComp As Boolean
Dim o2DComp As DrawingComponent
Dim iNbSheets As Integer
Dim iSheet As Integer

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim lR As Long
Dim lG As Long
Dim lB As Long



'*** Initial set
Set oDrawingDocument = CATIA.ActiveDocument
Set oSelection = oDrawingDocument.Selection

'Log File
Call AddToLogFile("Reset Drafting Colors", oDrawingDocument.Name)

sGreen2DCompName(0) = "FLAGNOTE*"
sGreen2DCompName(1) = "CONF F/D*"
sGreen2DCompName(2) = "CONF P/L*"
sGreen2DCompName(3) = "CHG BAL*"
sGreen2DCompName(4) = "FS BAL*"
sGreen2DCompName(5) = "DET BAL*"
sGreen2DCompName(6) = "BALL.50*"
sGreen2DCompName(7) = "ASSY BAL*"
sGreen2DCompName(8) = "View Direction*"
sLog = ""
iSheet = 0
Call frmProgress.progressBarInitialize("Resetting Drafting Colors")

'Count number of sheets
On Error Resume Next
Set oSheet = Nothing
Set oSheet = oDrawingDocument.Sheets.Item("DRAWING INFO")
On Error GoTo 0

If oSheet Is Nothing Then
    iNbSheets = oDrawingDocument.Sheets.Count
Else
    iNbSheets = oDrawingDocument.Sheets.Count - 1
End If

'*** Scan all sheets
For Each oSheet In oDrawingDocument.Sheets

    'Cancel on Progress Form
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
        GoTo endsub
    End If


    'Do not check "DRAWING INFO"
    If oSheet.Name <> "DRAWING INFO" Then
        
        oSheet.Activate
        iSheet = iSheet + 1
        Call frmProgress.progressBarRepaint("Resetting Drafting Colors on " & oSheet.Name, iNbSheets, iSheet)
        
        'Select callout and set it black
        oSelection.Clear
        oSelection.Add oSheet
        oSelection.Search "CATDrwSearch.DrwCallout.Visibility=Shown,sel"
        Set oVisProperty = oSelection.VisProperties
        Call oVisProperty.SetRealColor(0, 0, 0, False)
        
        
        'Select
        ' - Text
        ' - Balloon
        ' - Dimension
        ' - GD&T
        ' - Datum Feature
        ' - Roughness Symbol
        ' - Welding Symbol
        'and put it green
        oSelection.Clear
        oSelection.Add oSheet
        oSelection.Search "((Drafting.Text.Visibility=Shown - Drafting.Text.'TEXT STRING'=Bombardier*) + Drafting.Balloon.Visibility=Shown + Drafting.Dimension.Visibility=Shown + Drafting.'Geometrical Tolerance'.Visibility=Shown + Drafting.'Datum Feature'.Visibility=Shown + Drafting.'Roughness Symbol'.Visibility=Shown + Drafting.'Welding Symbol'.Visibility=Shown),sel"
        
        Set oVisProperty = oSelection.VisProperties
        Call oVisProperty.SetRealColor(0, 128, 0, False)
        
                
        'Select all shown 2D Component Instance if name <> "TB.*"
        oSelection.Clear
        oSelection.Add oSheet
        oSelection.Search "CATDrwSearch.DrwDitto.Visibility=Shown,sel"
        oSelection.Search "CATDrwSearch.DrwDitto.Name!=TB'.'*,sel"
        
        'Transfer 2D Component Instance in proper collection using Ref Name
        Set oGreen2DCompColl = New Collection
        Set oOther2DCompColl = New Collection
        If oSelection.Count > 0 Then
            For i = 1 To oSelection.Count
            
                bGreen2DComp = False
                
                For j = 0 To 8
                    'Green 2D Component
                    If oSelection.Item(i).Value.CompRef.Name Like sGreen2DCompName(j) Then
                        oGreen2DCompColl.Add oSelection.Item(i).Value
                        bGreen2DComp = True
                        Exit For
                    End If
                Next
                
                'Other 2D Component
                If bGreen2DComp = False Then
                    oOther2DCompColl.Add oSelection.Item(i).Value
                End If
            Next
        End If
        
        'Transfer all oGreen2DCompColl items to selection and put them green
        'and set "modifiable in instance" text to green
        If oGreen2DCompColl.Count > 0 Then
            oSelection.Clear
            For i = 1 To oGreen2DCompColl.Count
                oSelection.Add oGreen2DCompColl.Item(i)
                
                Set o2DComp = oGreen2DCompColl.Item(i)
                If o2DComp.GetModifiableObjectsCount > 0 Then
                    For j = 1 To o2DComp.GetModifiableObjectsCount
                        Set oText = o2DComp.GetModifiableObject(j)
                        oText.SetParameterOnSubString catColor, 0, 0, 8388863
                    Next
                End If

            Next
            Set oVisProperty = oSelection.VisProperties
            Call oVisProperty.SetRealColor(0, 128, 0, 0)
        End If
        
        'For oOther2DCompColl set "modifiable in instance" text to green
        If oOther2DCompColl.Count > 0 Then
            For i = 1 To oOther2DCompColl.Count
                Set o2DComp = oOther2DCompColl.Item(i)
                
                If o2DComp.GetModifiableObjectsCount > 0 Then
                    For j = 1 To o2DComp.GetModifiableObjectsCount
                        Set oText = o2DComp.GetModifiableObject(j)
                        oText.SetParameterOnSubString catColor, 0, 0, 8388863
                    Next
                End If
            Next
            
            
            Set oVisProperty = oSelection.VisProperties
            Call oVisProperty.SetRealColor(0, 128, 0, False)
        End If
        
        'Select all Table and put the text in green
        oSelection.Clear
        oSelection.Add oSheet
        oSelection.Search "CATDrwSearch.DrwTable.Visibility=Shown,sel"
        
        If oSelection.Count > 0 Then
            For i = 1 To oSelection.Count
                Set oDwgTable = oSelection.Item(i).Value
                
                For j = 1 To oDwgTable.NumberOfRows
                    For k = 1 To oDwgTable.NumberOfColumns
                        Set oDwgText = oDwgTable.GetCellObject(j, k)
                        oDwgText.SetParameterOnSubString catColor, 0, 0, 8388863
                    Next
                Next
            Next
        End If
        
        
        'Check if anything is magenta
        oSelection.Clear
        oSelection.Add oSheet
        oSelection.Search "(Color='(255,0,255)' & Visibility=Shown),sel"
        If oSelection.Count > 0 Then
            If sLog = "" Then
                sLog = " - " & oSheet.Name
            Else
                sLog = sLog & vbCrLf & " - " & oSheet.Name
            End If
        End If
    End If
Next

If sLog <> "" Then
    MsgBox " Magenta elements are still found on sheet:" & vbCrLf & sLog, vbInformation
End If

endsub:

frmProgress.Hide
Call frmKBEMain.resetToolbar

End Sub



