Attribute VB_Name = "CheckConstraint"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oCoincidenceConstraintCollection As New Collection
Private oSeedPartCollection As New Collection

'*************************************************************************
'* The goal of this class is to check:
'* - The name of the constraint matches the element it is linked to
'* - The "Instance Description" of the Std Part matches the constraint's reference element

'* This check is required because axis can be renamed after Std Part instantiation
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
    
    Call PopulateConstraintCollection
    Call CheckCoincidenceConstraint
    
End Sub

Private Sub PopulateConstraintCollection()
    
    Dim oConstraints As Constraints
    Dim oSeedPart As Product
    Dim sString As String
    
    'Initial set
    Set oCoincidenceConstraintCollection = Nothing
    Set oConstraints = oSeedAssy.Connections("CATIAConstraints")
    
    'Loop all the constraints
    If Not oConstraints.Count = 0 Then
        For Each oConstraint In oConstraints
        
            'We only check the Coincidence constraints
            If oConstraint.Type = 2 Then
    
                'Loop the Seed Part collection
                For Each oSeedPart In oSeedPartCollection
        
                    'We retrieve the Seed Part name from the constraint
                    sString = Split(oConstraint.Name, "| ")(2)
                    
                    If sString = oSeedPart.PartNumber Then
                        oCoincidenceConstraintCollection.Add oConstraint
                        Exit For
                    End If
                Next
    
            End If
        Next
    End If
    
End Sub

Private Sub CheckCoincidenceConstraint()
    
    Dim sLog As String
    Dim oConstraint As Constraint
    Dim iProgressCount As Integer
    Dim sCstStdPartId As String
    Dim sCstAxisId As String
    Dim sCstSeedPartId As String
    Dim oReference As Reference
    Dim sReferenceName As String
    Dim i As Integer
    Dim oStdPart As Product
    
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    
    sLog = ""
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Check Constraints")
    Call frmProgress.progressBarRepaint("Scanning Seed Assy Coincidence Constraints...")
    
    'Loop the oCoincidenceConstraintCollection
    For Each oConstraint In oCoincidenceConstraintCollection
        
        'Split constraint name in different values
        sCstStdPartId = Split(oConstraint.Name, " | ")(0)
        sCstAxisId = Split(oConstraint.Name, " | ")(1)
        sCstSeedPartId = Split(oConstraint.Name, " | ")(2)
        
        'Put the Std Part associated to the constraint in Default Mode
        Set oStdPart = oSeedAssy.Products.Item(sCstStdPartId)
        oStdPart.ApplyWorkMode DESIGN_MODE
        
        'Retrieve the reference from the constraint
        For i = 1 To 2
        
            Set oReference = oConstraint.GetConstraintElement(i)
            sReferenceName = oReference.DisplayName
        
            'The reference object is in a Seed Part
             If sReferenceName Like "*SEED_PART*" Then
            
                If Not (sReferenceName Like ("*" & sCstAxisId & "*")) Then
                    sLog = sLog & vbCrLf & "- The axis name in the following constraint name is wrong: " & oConstraint.Name
                End If
            
            'The reference object is in a Std Part
            Else
                If Not (oStdPart.DescriptionInst Like ("*" & sCstAxisId & "*")) Then
                    sLog = sLog & vbCrLf & "- The Description of the following Std Part is wrong: " & oStdPart.Name
                End If
            End If
        Next
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckConstraint.txt"
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

