Attribute VB_Name = "PartBodyInWorkObject"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to define in work object the PartBody
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
    
    Call DefineInWork
    
End Sub

Private Sub DefineInWork()
    
    Dim oSeedPartInstance As Product
    Dim sLog As String
    Dim oPartBody As Body
    Dim oActiveObject As AnyObject
    Dim oPart As Part
    Dim oPartDocument As PartDocument
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    
    'We loop the selected Seed Part to Retrieve the KBE Bodies, Boolean Parameters and ATB Parameters and populate the collections
    For Each oSeedPartInstance In oSeedPartCollection
        
        Set oPartDocument = oSeedPartInstance.GetMasterShapeRepresentation(False)
        Set oPart = oPartDocument.Part
        
        Set oPartBody = oPart.Bodies.Item("PartBody")
        Set oActiveObject = oPart.InWorkObject
        If oActiveObject.Name <> "PartBody" Then
            oPart.InWorkObject = oPartBody
            sLog = sLog & vbCrLf & "- In Work Object was modified in : " & oSeedPartInstance.PartNumber
        End If
    Next
    
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following actions were performed:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_PartBodyInWorkObject.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    Else
        MsgBox ("No error were found")
    End If
    
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar
    
End Sub
