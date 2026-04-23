Attribute VB_Name = "HideElements"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to hide the following elements:
'*  - Axis under the "Axis Systems" node
'*************************************************************************
Public Sub SelectSeedAssy()
        
    On Error Resume Next
    If oSeedAssy Is Nothing Then
        Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    Else
        If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    End If
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If (oSeedAssy Is Nothing Or oSeedAssy.Products.Count = 0) Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical, "No Seed Parts found..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy)
    
    If sSeedPartsList <> "" Then
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else
noSelection:
        Set oSeedAssy = Nothing
        Set sSourceWindow = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set oSeedPartCollection = New Collection
        Exit Sub
    End If
    
    Call HideElem
    
End Sub

Private Sub HideElem()
   
Dim oSeedPartInstance As Product
Dim oPartDocument As PartDocument
Dim oPart As Part
Dim oAxisSystems As AxisSystems
Dim oAxisSystem As AxisSystem
Dim oVisProperties As VisPropertySet
Dim oShow As CatVisPropertyShow

'Set Selection
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""


'We loop the selected Seed Part and Retrieve the axis in the Axis System node
For Each oSeedPartInstance In oSeedPartCollection
    
    Set oPartDocument = oSeedPartInstance.GetMasterShapeRepresentation(False)
    Set oPart = oPartDocument.Part
    Set oAxisSystems = oPart.AxisSystems
    
    For Each oAxisSystem In oAxisSystems
        oSelection.Clear
        oSelection.Add oAxisSystem
    
         Set oVisProperties = oSelection.VisProperties
   
        'Retrieve show state
        Call oVisProperties.GetShow(oShow)
         
        'Only hide the elements in show
        If oShow = catVisPropertyShowAttr Then
            oVisProperties.SetShow (catVisPropertyNoShowAttr)
            sLog = sLog & vbCrLf & " - " & oAxisSystem.Name & " was hidden in " & oSeedPartInstance.Name
        End If
    Next
    
Next

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "The following actions were performed:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_HideElements.txt"
    Call WriteTextFile(sLogFileName, sLog)
    If OpenWithNotepad <> 0 Then
        MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
    End If
Else
    MsgBox ("No actions were performed")
End If

CATIA.DisplayFileAlerts = True
frmKBEMain.resetToolbar

End Sub
