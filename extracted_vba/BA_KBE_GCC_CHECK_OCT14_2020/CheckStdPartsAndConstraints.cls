Attribute VB_Name = "CheckStdPartsAndConstraints"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oCoincidenceConstraintCollection As Collection
Private oBrokenConstraintCollection As Collection
Private oStdPartCollection As Collection

'*************************************************************************
'* The goal of this class is to check the Std Part and constraints:
'*  - The instance name of at least one instance of each Std Part ends with .0 "
'*  - When the instance name ends with ".0", the Description attribute of the instance must be blank
'*  - When the instance name does not end with ".0", the Description attribute of the instance must not be blank
'*  - Check the match between Coincidence Constraint and Std Part
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
    
    If Not oSeedAssy Is Nothing Then Call CheckStandardPartAndConstraints
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub CheckStandardPartAndConstraints()
    
    Dim oProduct As Product
    Dim oConstraint As Constraint
    Dim sLog As String
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    
    'Initial set
    CATIA.DisplayFileAlerts = False
    sLog = ""
    
    'Time set for Progress Bar
    dStartTime = DateTime.Time
    
    'Show progress bar
    Call frmProgress.progressBarInitialize("Check Std Parts and Coincidence Constraints")
    
    'Retrieve Std Part
    Call frmProgress.progressBarRepaint("Retrieve Std Part and put them in Design Mode", 4, 1)
    Call RetrieveStdParts
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
    'Check Std Part nomenclature
    Call frmProgress.progressBarRepaint("Check Std Part Nomenclature", 4, 2)
    Call CheckStdPartNomenclature(sLog)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
    'Retrieve Coincidence Constraints
    Call frmProgress.progressBarRepaint("Retrieve Coincidence Constraints", 4, 3)
    Call PopulateConstraintCollections
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
    'Check for useless Std Part
    Call frmProgress.progressBarRepaint("Check for useless Std Part", 4, 4)
    Call CheckUselessStdPart
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    
endSub:
    'Populate log
    For Each oProduct In oStdPartCollection
        If Right(oProduct.Name, 2) <> ".0" Then
            sLog = sLog & vbCrLf & " - The following Std Part is not linked to any coincidence constraint : " & oProduct.Name
        End If
    Next
           
    For Each oConstraint In oBrokenConstraintCollection
        sLog = sLog & vbCrLf & " - The following Constraint is not linked to any Std Part : " & oConstraint.Name
    Next
    
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following problems were found:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_CheckStdPartAndConstraints.txt"
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

Private Sub PopulateConstraintCollections()
    
    Dim oConstraints As Constraints
    Dim oConstraint As Constraint
    Dim oReference1 As Reference
    Dim oReference2 As Reference
    Dim sCstElemStdPart As String
    Dim sString1 As String
    Dim sString2 As String
    Dim sStdPartID As String
    
    'Initial set
    Set oConstraints = oSeedAssy.Connections("CATIAConstraints")
    
    Set oCoincidenceConstraintCollection = New Collection
    Set oBrokenConstraintCollection = New Collection
    
    'Loop all the constraints
    If Not oConstraints.Count = 0 Then
        For Each oConstraint In oConstraints
        
            'We only check the Coincidence constraints
            If oConstraint.Type = 2 Then
                            
                'Initial set
                sCstElemStdPart = ""
                sString1 = ""
                sString2 = ""
                Set oReference1 = Nothing
                Set oReference2 = Nothing
                
                'Retrieve constraint elements
                On Error Resume Next
                Set oReference1 = oConstraint.GetConstraintElement(1)
                Set oReference2 = oConstraint.GetConstraintElement(2)
                sString1 = oReference1.DisplayName
                sString2 = oReference2.DisplayName
                
                If sString1 Like "*Absolute Axis System*" Then
                    sCstElemStdPart = sString1
                ElseIf sString2 Like "*Absolute Axis System*" Then
                    sCstElemStdPart = sString2
                End If
                
                On Error GoTo 0
                
                If sCstElemStdPart <> "" Then
                    sStdPartID = Split(sCstElemStdPart, "/")(1)
                    oCoincidenceConstraintCollection.Add oConstraint, sStdPartID
                Else
                    oBrokenConstraintCollection.Add oConstraint
                End If
                
            End If
        Next
    End If

End Sub
Private Sub RetrieveStdParts()
'This Sub will retrieve all Std Parts, populate the oStdPartCollection and put in Design Mode
    
    Dim oProduct As Product
    Dim sPartNumber As String
    
    Set oStdPartCollection = New Collection
    
    'Loop instances under the Seed Assy
    For Each oProduct In oSeedAssy.Products
        If Not (oProduct.Name Like "KBE_*" Or oProduct.Name Like "MONUMENTORIGINPART*") Then
                
            oStdPartCollection.Add oProduct, oProduct.Name
            
            'Put product in Design Mode
            sPartNumber = ""
            On Error Resume Next
            sPartNumber = oProduct.PartNumber
            If sPartNumber = "" Then oProduct.ApplyWorkMode DESIGN_MODE
            'Sub Assy in Design Mode
            If TypeName(oProduct.ReferenceProduct.Parent) = "ProductDocument" Then Call SubAssyDesignMode(oProduct)
            On Error GoTo 0
        End If
    Next
    
End Sub

Private Sub SubAssyDesignMode(ByVal oParentProd As Product)
'This Sub will put all the parts of a sub-assemblies in Design Mode
    
    Dim oProduct As Product
    Dim sPartNumber As String
    
    'Loop instances under the Seed Assy
    For Each oProduct In oParentProd.Products
                
        sPartNumber = ""
        On Error Resume Next
        sPartNumber = oProduct.PartNumber
        If sPartNumber = "" Then oProduct.ApplyWorkMode DESIGN_MODE
        
        'Recursive Sub
        If TypeName(oProduct.ReferenceProduct.Parent) = "ProductDocument" Then Call SubAssyDesignMode(oProduct)
    Next
    
End Sub

Private Sub CheckStdPartNomenclature(ByRef sLog As String)
    
    Dim oProduct1 As Product
    Dim oProduct2 As Product
    Dim bFound As Boolean
    
    'We loop oStdPartCollection
    For Each oProduct1 In oStdPartCollection
                                  
    
        'If the instance name ends with ".0"
        If Right(oProduct1.Name, 2) = ".0" Then
            
            If oProduct1.DescriptionInst <> "" Then
                sLog = sLog & vbCrLf & " - The Description attribut of the following instance should be blank: " & oProduct1.Name
            End If
        
        'If the instance name doesn't end with ".0"
        Else
    
            If oProduct1.DescriptionInst = "" Then
                sLog = sLog & vbCrLf & " - The Description attribut of the following instance should not be blank: " & oProduct1.Name
            End If
    
    
            'Loop to find the ".0" instance
            bFound = False
            For Each oProduct2 In oStdPartCollection
                If Right(oProduct2.Name, 2) = ".0" And Split(oProduct2.Name, ".")(0) = Split(oProduct1.Name, ".")(0) Then
                    bFound = True
                    Exit For
                End If
            Next
            
            If bFound = False Then
                sLog = sLog & vbCrLf & " - The following instance is missing: " & Split(oProduct1.Name, ".")(0) & ".0"
            End If
        End If
            
    Next
    
End Sub

Private Sub CheckUselessStdPart()
    
    Dim oStdPart As Product
    Dim oConstraint As Constraint
    Dim i As Integer
    Dim aaa As String
    
    For i = oStdPartCollection.Count To 1 Step -1
    
        Set oStdPart = oStdPartCollection.Item(i)
        
        'Retrieve constraint using name of Std Part
        Set oConstraint = Nothing
        On Error Resume Next
        Set oConstraint = oCoincidenceConstraintCollection.Item(oStdPart.Name)
        On Error GoTo 0
        
        'If we have a match (Constraint - Std Part) we remove the Std Part from the collection
        If Not oConstraint Is Nothing Then
            oStdPartCollection.Remove (i)
        End If
    Next
    
End Sub
