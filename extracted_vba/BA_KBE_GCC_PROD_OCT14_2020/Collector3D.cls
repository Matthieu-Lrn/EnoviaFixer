Attribute VB_Name = "Collector3D"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'Start Create SeedLing Code Copy directly from Module
'COPYRIGHT BOMBARDIER

'************************************************************************************************************
'* Purpose:     Create SeedLing CATPart (Collector) by scanning selected Seed Parts(see launchColl3D)       *
'*              SeedLing is a multi-body CATPart representing each pertinent body of the Seed Parts.        *
'* Purpose(2):  Tool that creates an empty 3D Collector Template (see launchCollTemplate)                   *
'*                                                                                                          *
'* Assumption:  Seed Part(s) be loaded in design mode in CATIA                                              *
'* Prerequisites :                                                                                          *
'* Arguments:   Interactive mode. User has to enter a significant portion of seedling name (any case can be *
'*              used) to help the macro retrieve SeedLing Part in session. Then has to select each Seed Part*
'*              to be synchronized, one at a time.                                                          *
'*                                                                                                          *
'* Script Version : v2.4
'*              V2.4 comment : If InvertSelection is checked then collect everything except bodies in Subset*
'*              V2.3 comment : B0529030 - Clean code, GUI standardization, new options                      *
'*                              clean previous collector, collect material ATBs or other ATBs               *
'*              V2.2 comment : New method to be able to do a selection from many windows                    *
'*              V2.1 comment : new way to indicate script version (user form frame)                         *
'*              + SKLE01 replaced by SKEL01 + code moved to a new class called Collector3D.                 *
'*              V2.0 comment : interface added to select Seed Part and 3D Collector from a CATIA V5         *
'*              multi-windows session. Main treatment changed to clooect only relevant SS parameters        *
'*                                                                                                          *
'* Author: Cedric Begel (cbegel@ca.ibm.com)                                                                 *
'* Updated by: Julien Bigaouette (V2.3 - january 2013)  ,Abhishek (V2.4-March 18,2015)                      *
'* Language: VBA                                                                                            *
'* Initial CATIA version: V5R14 SP6E (PC only) - Last update: V5R21 SP2F                                    *
'************************************************************************************************************
                        
'******************
'* Main Procedure *
'******************
    Private i As Long
    Private j As Long
    Private iI As Long
    Private iJ As Long
    Private iPos As Long
    Private iLastKBEBody As Long
    
    Private oNewBody As Body
    Private oSeedPartCurBody As Body
    Private oSeedLingCurKBEBody As Body
    Private sSeedLingCurKBEBody As String
    Private oSeedLingCurBody As Body
    Private sSeedLingCurBody As String
    Private oSeedLingParameters As Parameters
    Private oCurSeedLingParameter As StrParam
    Private oActivityParam As Parameter
    Private sSeedLingKBEParametersList As String
    Private sSeedLingAxisSystem As String
    Private sKBEBodiesFromOS As String
    Private sCurKBEBodyNum As String
    Private iCurKBEBody As Integer
    Private bSeedPartBooleanParameter As Boolean
    Private oOriginalSolid As Shape
    Private oOriginalSolidParent As Body
    Private sSeedPartBodiesList As String
    Private sSeedPartDatum As String
    Private oCurSeedLingDatum As HybridShape
    Private sCurSeedLingDatum As String
    Private oCurSeedLingOS As HybridShape
    Private sCurSeedLingOS As String
    Private sCurSeedLingOAParamValue As String
    Private oSplit As Split
    
    Private aSeedPartAllATBs As Collection
    Private aSeedPartDesiredATBs As Collection
    Private aSeedPartNotCollectedATBs As Collection
    Private aSeedPartSTRCs As Collection
    Private aSeedPartBooleans As Collection
    Private aActiveBodiesATBs As Collection
    Private aSeedLingATBs As Collection
    Private aSeedLingSTRCs As Collection
    Private aSeedLingActivityParam As Collection
    Private aKBEBodyFromOS As Collection
    Private aSeedPartOS As Collection
    Private aSeedLingOS As Collection
    Private aSeedPartOA As Collection
    Private aSeedPartDatum As Collection
    Private aSeedLingDatum As Collection
    Private aSeedLingRecycleKBEBodiesList As Collection
    Private seedAssySubSetCollection As Collection
    Private seedlingSubSetCollection As Collection
    
    Private oKBEParamSet As ParameterSet
    Private oATBParamSet As ParameterSet
    Private oSTRCParamSet As ParameterSet
    Private SeedPartSSParamSet As ParameterSet
    Private SeedLingSSParamSet As ParameterSet
    Private oDatumSet As HybridBody
    Private oSkeletonSet As HybridBody
    Private oOperationsSet As HybridBody
    Private oEnvironmentSet As HybridBody
    Private oKinematicsSet As HybridBody
    Private oStylingFeatureSet As HybridBody
    Private oBehaviorSet As HybridBody
    
    Private sParameterKBERootName As String
    Private sBodyRootName As String
    Private sSearchString As String
    
Public Sub launchColl3D()
    
    sParameterKBERootName = "ParameterKBE"
    sBodyRootName = "BodyKBE"
   
'-------------------------------------------------------
'     PROGRESS BAR SETTINGS
'-------------------------------------------------------
    timeStart = DateTime.Timer
    iProgressCount1 = 0
    iProgressCount1Max = oSeedPartCollection.Count + 1 '1/4*reset collector + 1/4*initial scan + oSeedPartCollection.Count (7/9*Main treatment + 1/9*collectOS & 1/9*collectOA) + 1/4*SplitOS + 1/4*AddOA
    Call frmProgress.progressBarInitialize("3D Collector")
    Call frmProgress.progressBarRepaint("Initializing 3D Collector...")
    
'--------------------------------------------------------------
'     FORCE ASSEMBLY DESIGN IN BOTH WINDOWS TO ALLOW COPY/PASTE
'--------------------------------------------------------------
    CATIA.ActiveDocument.Selection.Clear
    If Not sTargetWindow Is Nothing And Not oNHAAssy Is Nothing Then
        If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
        If TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
            If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
            If CATIA.GetWorkbenchId <> "Assembly" Then
                CATIA.StartWorkbench "Assembly"
            End If
        End If
    End If
    
    If Not sSourceWindow Is Nothing And Not oSeedAssy Is Nothing Then
        If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
        If TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
            If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
            If CATIA.GetWorkbenchId <> "Assembly" Then
                CATIA.StartWorkbench "Assembly"
            End If
        End If
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     SUBSETS
'-------------------------------------------------------
    
    If frm3DCollector.cbSubSet.Value = True Then
        'Scan Subsets in part
        Call ScanSubsets(oSeedPartCollection)
        
        'Ask user to select a Subset
        Call frmSubsetsSelection.FormInitialize(frm3DCollector.Left + 15, frm3DCollector.Top + frm3DCollector.cmdOK.Top - frmSubsetsSelection.Height)
        If sCurrentSSParameterName = "" Then
            frm3DCollector.Show
            GoTo endsub
        End If
        
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
'-------------------------------------------------------
'     CREATE NEW COLLECTOR
'-------------------------------------------------------
    If frm3DCollector.optCreateNew.Value = True Then
        Call CreateNewCollector
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     RESET COLLECTOR
'-------------------------------------------------------
    'If checkbox selected, clean existing collector.
    iProgressCount1 = iProgressCount1 + 1 / 4
    If frm3DCollector.cbResetCollector.Value = True And frm3DCollector.optCreateNew.Value = False Then
        Call resetCollector
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     INITIALIZE VARIABLES AND COLLECTOR
'-------------------------------------------------------
    Call InitializeCollector
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     INITIAL SCAN
'-------------------------------------------------------
    'Perform initial scan / create missing elements
    iProgressCount1 = iProgressCount1 + 1 / 4
    If frm3DCollector.cbResetCollector.Value = False Or frm3DCollector.optCreateNew.Value = False Then
        Call InitialScan
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     MAIN TREATMENT - COLLECT FROM SEED PART
'-------------------------------------------------------
    For Each iObject In oSeedPartCollection
        
        Set oSeedPart = iObject.ReferenceProduct.Parent.Part
        Call MainTreatment(frmProgress.pbProgressMain.Value, 1 / oSeedPartCollection.Count): If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
        iProgressCount1 = iProgressCount1 + (1 / 10) '9 of 10
        If frm3DCollector.ImgAllSPTrue.Visible = True Or frm3DCollector.ImgSelectedSPTrue.Visible = True Then Call CollectOS
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
        iProgressCount1 = iProgressCount1 + (1 / 10) '10 of 10
        If frm3DCollector.ImgAllSPTrue.Visible = True Or frm3DCollector.ImgSelectedSPTrue.Visible = True Then Call Collect_OA
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
    Next
    
'---------------------------------------------------------------------------
'     FINALIZING COLLECTOR - APPLY OS AND OA OPERATIONS - VERIFY FLAT TREE
'---------------------------------------------------------------------------
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    frmProgress.lblMessageMain.Caption = "Finalizing 3D Collector..."

    iProgressCount1 = iProgressCount1 + 1 / 4
    If frm3DCollector.ImgAllSPTrue.Visible = True Or frm3DCollector.ImgSelectedSPTrue.Visible = True Then Call Split_OS
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
    iProgressCount1 = iProgressCount1 + 1 / 4
    If frm3DCollector.ImgAllSPTrue.Visible = True Or frm3DCollector.ImgSelectedSPTrue.Visible = True Then Call Add_OA
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub

    If Not oNHAAssy Is Nothing Then
        If oNHAAssy.Products.Count > 1 Then CheckKBEBodies_Vs_FlatTreeBodies
    End If
    
'-------------------------------------------------------
'     END
'-------------------------------------------------------
    
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, 1, 1, frmProgress.lblMessage.Caption, 1, 1, sTimer)
    
    On Error Resume Next
    CATIA.StatusBar = "Successfully Done! "
    
    'Different message if
    '   1-User wants to position collector and cannot do it
    '   2-User has to synchronize an existing Flat Tree under collector (sLog variable)
    If (oSeedLingAxisSystem Is Nothing Or oNHAAssy Is Nothing) And frm3DCollector.cbPositionCollector.Value = True Then
        If sLog = "" Then
            MsgBox "3D Collector Synchronized!    " & vbCrLf & vbCrLf & "Warning: Cannot position 3D Collector!     ", vbExclamation, "Warning"
        ElseIf Not sLog = "" Then
            MsgBox "3D Collector Synchronized!    " & vbCrLf & vbCrLf & sLog & vbCrLf & vbCrLf & "Warning: Cannot position 3D Collector!     ", vbExclamation, "Warning"
        End If
    Else
        If sLog = "" Then
            MsgBox "3D Collector Synchronized!    ", vbInformation, "3D Collector"
        ElseIf Not sLog = "" Then
            MsgBox "3D Collector Synchronized!    " & vbCrLf & vbCrLf & sLog, vbExclamation, "Warning"
        End If
    End If
    
endsub:
'-------------------------------------------------------
'     POSITION COLLECTOR
'-------------------------------------------------------
    'If checkbox selected, position existing collector.
    If frm3DCollector.cbPositionCollector.Value = True Then
        
        'Find the last synchronized axis
        On Error Resume Next
        Set oSeedLingAxisSystem = Nothing
        For Each iObject In oSeedLingPart.AxisSystems
            If iObject.Name Like "*DATUM_AXIS_A/C" Then
                Set oSeedLingAxisSystem = iObject
            End If
        Next
        
        If Not oSeedLingAxisSystem Is Nothing And Not oNHAAssy Is Nothing Then
            Call SnapAxisToOrigin(oSeedLingInstance, oNHAAssy, oSeedLingAxisSystem)
            sTargetWindow.ActiveViewer.Reframe
        End If
        On Error GoTo 0
    
    End If
    
    On Error Resume Next
    oSeedLingPart.InWorkObject = oSeedLingPart.MainBody
    oSeedLingPart.Update
    On Error GoTo 0
    
End Sub

Public Sub launchCollTemplate()
    
    sParameterKBERootName = "ParameterDBE"
    sBodyRootName = "BodyDBE"
   
'--------------------------------------------------------------
'     FORCE ASSEMBLY DESIGN IN BOTH WINDOWS TO ALLOW COPY/PASTE
'--------------------------------------------------------------
    CATIA.ActiveDocument.Selection.Clear
    If Not sTargetWindow Is Nothing And Not oNHAAssy Is Nothing Then
        If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
        If TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
            If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
            If CATIA.GetWorkbenchId <> "Assembly" Then
                CATIA.StartWorkbench "Assembly"
            End If
        End If
    End If
    
    If Not sSourceWindow Is Nothing And Not oSeedAssy Is Nothing Then
        If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
        If TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
            If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
            If CATIA.GetWorkbenchId <> "Assembly" Then
                CATIA.StartWorkbench "Assembly"
            End If
        End If
    End If
   
'-------------------------------------------------------
'     PROGRESS BAR SETTINGS
'-------------------------------------------------------
    timeStart = DateTime.Timer
    iProgressCount1 = 0
    iProgressCount1Max = 3 + oSeedPartCollection.Count
    Call frmProgress.progressBarInitialize("Collector Template")

'-------------------------------------------------------
'     CREATE NEW COLLECTOR
'-------------------------------------------------------
    If frm3DCTemplate.optCreateNew.Value = True Then
        Call CreateNewCollector
    End If

'-------------------------------------------------------
'     INITIALIZE VARIABLES AND COLLECTOR
'-------------------------------------------------------
    Call InitializeCollector
    
'-------------------------------------------------------
'     INITIAL SCAN
'-------------------------------------------------------
    'Perform initial scan / create missing elements
    iProgressCount1 = iProgressCount1 + 1
    If frm3DCTemplate.optCreateNew.Value = False Then
        Call InitialScan
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    End If

'-------------------------------------------------------
'     CREATE DBE BODIES
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + 1
    Call CreateDBEBodies
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub

'-------------------------------------------------------
'     DATUM COLLECTION
'-------------------------------------------------------
    If Not oNHAAssy Is Nothing And Not oSeedAssy Is Nothing Then
        For Each iObject In oSeedPartCollection
            iProgressCount1 = iProgressCount1 + 1
            frmProgress.lblMessageMain.Caption = "Scanning " & iObject.Name & " (" & CStr(iProgressCount1) & "/" & CStr(oSeedPartCollection.Count) & ")"
            
            Set oSeedPart = iObject.ReferenceProduct.Parent.Part
            sSeedPartLetter = Split(Split(oSeedPart.Name, "_SEED_PART_")(1), "-")(0)
            Call CopyDatum
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        Next
    End If

'-------------------------------------------------------
'         COPY/PASTE A/C AXIS SYSTEM
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + 1
    iProgressCount2 = 1
    iProgressCount2Max = 1
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Copying A/C Axis..."
    Call frmProgress.progressBarRepaint("Finalizing Collector Template...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    For Each iObject In oSeedPartCollection
        Set oSeedPart = iObject.ReferenceProduct.Parent.Part
        sSeedPartLetter = Split(Split(oSeedPart.Name, "_SEED_PART_")(1), "-")(0)
        Call CopySeedlingAxisSystem
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    Next

'-------------------------------------------------------
'     END
'-------------------------------------------------------
    
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, 1, 1, frmProgress.lblMessage.Caption, 1, 1, sTimer)
    
    On Error Resume Next
    CATIA.StatusBar = "Successfully Done! "
    MsgBox "Collector Template Successfully Done!    ", vbInformation, "Collector Template"
endsub:
    On Error Resume Next
    oSeedLingPart.InWorkObject = oSeedLingPart.MainBody
    oSeedLingPart.Update
    On Error GoTo 0
    
End Sub

Private Sub CreateDBEBodies()
    
    Dim vParameterFunction(5)
    vParameterFunction(0) = "Shell"
    vParameterFunction(1) = "Divider"
    vParameterFunction(2) = "Removable"
    vParameterFunction(3) = "Fixed"
    vParameterFunction(4) = "Function1"
    vParameterFunction(5) = "Function2"
    
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    iProgressCount2 = 0
    iProgressCount2Max = Val(frm3DCTemplate.txtNoOfDBEBodies.Text)
    For i = 1 To Val(frm3DCTemplate.txtNoOfDBEBodies.Text)
        
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Creating / Updating DBE Bodies (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint("Updating Collector Template...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        On Error Resume Next
        sCurKBEBodyNum = Format(CStr(i), "0#")
        
        Set oBodyMain = Nothing
        Set oBodyMain = oSeedLingPart.Bodies.Item(sBodyRootName & sCurKBEBodyNum)
        
        If oBodyMain Is Nothing Then
            
            'Create Body DBE, publish it if it is not published
            Set oBodyMain = oSeedLingPart.Bodies.Add
            oBodyMain.Name = sBodyRootName & sCurKBEBodyNum
            
            Set oPublication = Nothing
            Set oPublication = oPublications.Item(oBodyMain.Name) 'Should return nothing if publication was deleted previously
            If oPublication Is Nothing And frm3DCTemplate.cbPublish.Value Then
                Set oPublication = oPublications.Add(oBodyMain.Name)
            End If
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & oBodyMain.Name)
            oPublications.SetDirect oBodyMain.Name, oReference
            
            
            'Create child body, assemble it to Body DBE
            Set oBody = Nothing
            Set oBody = oSeedLingPart.Bodies.Add
            oBody.Name = "Z_" & sBodyRootName & sCurKBEBodyNum
            
            oSeedLingPart.InWorkObject = oBodyMain
            Set oAssemble = oSeedLingPart.ShapeFactory.AddNewAssemble(oBody)
            oAssemble.Name = "MainAssemble"
            
            
            'Create Parameter Set, move it under oBody
            Set oParameterSet = oSeedLingPart.Parameters.RootParameterSet.ParameterSets.CreateSet("Parameters" & sCurKBEBodyNum)

            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oParameterSet
            oSelectionFromSeedLing.Cut
            
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oBody
            oSelectionFromSeedLing.Paste
            Set oParameterSet = KBESearch("CATKnowledgeSearch.AdvisorParameterSet,sel", oBody, oSelectionFromSeedLing).Item2(1).Value
            
            
            'Create Relations set under oBody
            oSeedLingPart.Relations.CreateSetOfRelations oBody
            Set oRelationSet = KBESearch("CATKnowledgeSearch.AdvisorRelationSet,sel", oBody, oSelectionFromSeedLing).Item2(1).Value
            oSelectionFromSeedLing.Item2(1).Value.Name = "Relations" & sCurKBEBodyNum
            
            
            'Create Parameter DBE, publish it if it is not published
            Set oParameter = oParameterSet.DirectParameters.CreateString(sParameterKBERootName & sCurKBEBodyNum, "Z_" & sBodyRootName & sCurKBEBodyNum & " " & sCollectorID)
            sParameterName = Split(oParameter.Name, "\")(UBound(Split(oParameter.Name, "\")))
            
            Set oPublication = Nothing
            Set oPublication = oPublications.Item(sParameterName) 'Should return nothing if publication was deleted previously
            If oPublication Is Nothing And frm3DCTemplate.cbPublish.Value Then
                Set oPublication = oPublications.Add(sParameterName)
            End If
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oBody.Name & "/" & oParameterSet.Name & "/" & sParameterName)
            oPublications.SetDirect sParameterName, oReference
            
            
            'Create Parameter Activity, publish it if it is not published
            Set oParameter = oParameterSet.DirectParameters.CreateBoolean("Activity_" & oBody.Name, True)
            sParameterName = Split(oParameter.Name, "\")(UBound(Split(oParameter.Name, "\")))
            
            Set oPublication = Nothing
            Set oPublication = oPublications.Item(sParameterName) 'Should return nothing if publication was deleted previously
            If oPublication Is Nothing And frm3DCTemplate.cbPublish.Value Then
                Set oPublication = oPublications.Add(sParameterName)
            End If
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oBody.Name & "/" & oParameterSet.Name & "/" & sParameterName)
            oPublications.SetDirect sParameterName, oReference
            
            
            'Create Parameter Panel Thickness, keep its name for formula
            Set oParameter = oParameterSet.DirectParameters.CreateDimension("PanelThickness", "LENGTH", "0")
            sParameterName = Replace(oParameter.Name, oSeedLingPart.Name & "\", "")
            
            
            'Create ATB Size, link it to Panel Thickness with a formula, move formula under oBody Relation Set
            Set oParameter = oParameterSet.DirectParameters.CreateDimension("ATB_" & sBodyRootName & sCurKBEBodyNum & "_Size", "LENGTH", 0)
            Set oFormula = oSeedLingPart.Relations.CreateFormula("Formula" & sCurKBEBodyNum, "", oParameter, sParameterName)
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oFormula
            oSelectionFromSeedLing.Cut
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oRelationSet
            oSelectionFromSeedLing.Paste
            
            
            'Create ATB Function
            Set oParameter = oParameterSet.DirectParameters.CreateString("ATB_" & sBodyRootName & sCurKBEBodyNum & "_Function", "Shell")
            oParameter.SetEnumerateValues vParameterFunction
            
            
            'Create GeoSet in oBody
            Set oGeometricalSet = oBody.HybridBodies.Add
            oGeometricalSet.Name = "GeoSet"
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
    Next i
    oSelectionFromSeedLing.Clear
    On Error GoTo 0

End Sub

Private Sub CreateNewCollector()
            
    Dim iCol3DIndex As Integer, sCol3DIndex As String
    
    On Error Resume Next
    
    'List all NHA products
    sNHAAssyProducts = ""
    For Each oProduct In oNHAAssy.Products
        If oProduct.Products.Count = 0 Then
            oProduct.ApplyWorkMode DEFAULT_MODE
            If oProduct.PartNumber Like "*-*COL*" Then sNHAAssyProducts = sNHAAssyProducts & oProduct.PartNumber & ";"
        End If
    Next

    'Create new part
    oNHAAssy.Products.AddNewComponent "Part", ""
    Set oSeedLingInstance = oNHAAssy.ReferenceProduct.Products.Item(oNHAAssy.Products.Count)
    Set oSeedLingPart = oSeedLingInstance.ReferenceProduct.Parent.Part

    'Rename new part and instance
    iCol3DIndex = 0
    Do
        'Find Collector index (no duplication)
        Do
            iCol3DIndex = iCol3DIndex + 1
            sCol3DIndex = Format(CStr(iCol3DIndex), "0#")
        Loop Until Not sNHAAssyProducts Like "*" & oNHAAssy.PartNumber & "COL" & sCol3DIndex & "*" And _
                   Not sNHAAssyProducts Like "*" & oNHAAssy.PartNumber & "-COL" & sCol3DIndex & "*"
        
        If InStr(1, oNHAAssy.PartNumber, "-", vbTextCompare) = 0 Then
            oSeedLingInstance.ReferenceProduct.PartNumber = oNHAAssy.PartNumber & "-COL" & sCol3DIndex
            oSeedLingInstance.Name = oNHAAssy.PartNumber & "-COL" & sCol3DIndex & ".1"
        Else
            oSeedLingInstance.ReferenceProduct.PartNumber = oNHAAssy.PartNumber & "COL" & sCol3DIndex
            oSeedLingInstance.Name = oNHAAssy.PartNumber & "COL" & sCol3DIndex & ".1"
        End If
        
    Loop Until oSeedLingInstance.ReferenceProduct.PartNumber Like "*-*COL*" And oSeedLingInstance.Name Like "*-*COL*"
    
    On Error GoTo 0
    
'-------------------------------------------------------
'     APPLY BA STANDARD
'-------------------------------------------------------
    Call oTemplateData.AddGeometricalSets(oSeedLingPart, oSelectionFromSeedLing)
    Call oTemplateData.CreateProperties(oSeedLingInstance, , "KBE")
    Call oTemplateData.HidePlanesAxis(oSeedLingPart, oSelectionFromSeedLing)
    
    sSeedAssyTitle = ""
    If Not oSeedAssy Is Nothing Then sSeedAssyTitle = oSeedAssy.UserRefProperties.Item("Title").ValueAsString
    If Not oSeedAssy Is Nothing Then oSeedLingInstance.ReferenceProduct.UserRefProperties.GetItem("Title").Value = sSeedAssyTitle
    If Not oSeedAssy Is Nothing Then oSeedLingPart.Parent.Product.Definition = "COL_" & Split(oSeedAssy.PartNumber, "_")(1)

End Sub


Private Sub resetCollector()

    iProgressCount2Max = 4
    On Error Resume Next
    
    'Delete Bodies
    iProgressCount2 = 1
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint("Reset 3D Collector...", iProgressCount1Max, iProgressCount1, "Deleting elements...", iProgressCount2Max, iProgressCount2, sTimer)
    
    For Each iObject In oSeedLingPart.Bodies
        If Not iObject.Name = oSeedLingPart.MainBody.Name Then
            oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(iObject)
        End If
    Next
    
    'Delete Parameters
    iProgressCount2 = 2
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, "Deleting elements...", iProgressCount2Max, iProgressCount2, sTimer)
    
    For Each iObject In oSeedLingPart.Parameters.RootParameterSet.ParameterSets
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(iObject)
    Next
    
    'Delete Geoset
    iProgressCount2 = 3
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, "Deleting elements...", iProgressCount2Max, iProgressCount2, sTimer)
    
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("Skeleton"))
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("Environment"))
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("Kinematics"))
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("StylingFeatures"))
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("Behaviors"))
    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingPart.HybridBodies.Item("Operations"))
    
    'Delete Axis
    iProgressCount2 = 4
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, "Deleting elements...", iProgressCount2Max, iProgressCount2, sTimer)
    
    For Each iObject In oSeedLingPart.AxisSystems
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(iObject)
    Next
    
    'Delete broken publications
    Call deleteBrokenPublications(oSeedLingInstance.ReferenceProduct.Parent, , True)
    
    oSeedLingPart.Update
    On Error GoTo 0
    
End Sub


Private Sub InitializeCollector()
    
'-------------------------------------------------------
'     INITIALIZE VARIABLES
'-------------------------------------------------------
    sLog = ""
    iLastKBEBody = 0
    sSeedLingKBEParametersList = ""
    sCollectorID = Mid(oSeedLingPart.Name, InStrRev(oSeedLingPart.Name, "COL"))
    Set oSeedLingParameters = oSeedLingPart.Parameters
    Set oPublications = oSeedLingInstance.ReferenceProduct.Publications
    Set aSeedLingRecycleKBEBodiesList = New Collection
    Set aSeedPartOA = New Collection
    Set aSeedLingATBs = New Collection
    Set aSeedLingSTRCs = New Collection
    Set aSeedLingActivityParam = New Collection
    Set aSeedLingDatum = New Collection
    Set aSeedLingOS = New Collection
    
'-------------------------------------------------------
'     APPLY BA STANDARD
'-------------------------------------------------------
    Call oTemplateData.AddPartBodyPublication(oSeedLingPart, oSelectionFromSeedLing, oSeedLingInstance)
    Call oTemplateData.SetProprietaryNotice(oSeedLingInstance)
        
'-------------------------------------------------------
'     CREATE ABSOLUTE AXIS SYSTEM
'-------------------------------------------------------
    On Error Resume Next
    
    'Create Absolute Axis System
    Set oSeedLingAxisSystem = Nothing
    Set oSeedLingAxisSystem = oSeedLingPart.AxisSystems.Item("Absolute Axis System")
    If oSeedLingAxisSystem Is Nothing Then
        Set oSeedLingAxisSystem = oSeedLingPart.AxisSystems.Add()
        oSeedLingAxisSystem.Name = "Absolute Axis System"
        oSelectionFromSeedLing.Clear
        oSelectionFromSeedLing.Add oSeedLingAxisSystem
        oSelectionFromSeedLing.VisProperties.SetLayer catVisLayerBasic, 1
        oSelectionFromSeedLing.VisProperties.SetShow catVisPropertyNoShowAttr
        oSelectionFromSeedLing.Clear
    End If
    oSeedLingAxisSystem.IsCurrent = True
    
    'Publish Absolute Axis system
    Set oPublication = Nothing
    Set oPublication = oPublications.Item("Absolute Axis System") 'Should return nothing if publication was deleted previously
    If oPublication Is Nothing Then
        Set oPublication = oPublications.Add("Absolute Axis System")
    End If
    Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & oSeedLingAxisSystem.Name)
    oPublications.SetDirect "Absolute Axis System", oReference
    
'-------------------------------------------------------
'     CREATE MANDATORY GEOSETS AND PARAMETER SETS
'-------------------------------------------------------
    
    'Find Param Set "KBE_Parameters" in seedling, create it if does not exit
    Set oKBEParamSet = Nothing
    Set oKBEParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.Item("KBE_Parameters")
    If oKBEParamSet Is Nothing Then
        Set oKBEParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.CreateSet("KBE_Parameters")
    End If
    
    'Find Param Set "ATB_Parameters" in seedling, create it if does not exit
    Set oATBParamSet = Nothing
    Set oATBParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.Item("ATB_Parameters")
    If oATBParamSet Is Nothing Then
        Set oATBParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.CreateSet("ATB_Parameters")
    End If
    
    'Find Param Set "STRC_Parameters" in seedling, create it if does not exit
    Set oSTRCParamSet = Nothing
    Set oSTRCParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.Item("STRC_Parameters")
    If oSTRCParamSet Is Nothing Then
        Set oSTRCParamSet = oSeedLingParameters.RootParameterSet.ParameterSets.CreateSet("STRC_Parameters")
    End If
    
    'Create "Skeleton" Set in SeedLing if it does not exist and hide it
    Set oSkeletonSet = Nothing
    Set oSkeletonSet = oSeedLingPart.HybridBodies.Item("Skeleton")
    If oSkeletonSet Is Nothing Then
        Set oSkeletonSet = oSeedLingPart.HybridBodies.Add
        oSkeletonSet.Name = "Skeleton"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSkeletonSet), 0
    
    'Create "Operations" Set in SeedLing if it does not exist and hide it
    Set oOperationsSet = Nothing
    Set oOperationsSet = oSeedLingPart.HybridBodies.Item("Operations")
    If oOperationsSet Is Nothing Then
        Set oOperationsSet = oSeedLingPart.HybridBodies.Add
        oOperationsSet.Name = "Operations"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oOperationsSet), 0
    
    'Create "Environment" Set in SeedLing if it does not exist and hide it
    Set oEnvironmentSet = Nothing
    Set oEnvironmentSet = oSeedLingPart.HybridBodies.Item("Environment")
    If oEnvironmentSet Is Nothing Then
        Set oEnvironmentSet = oSeedLingPart.HybridBodies.Add
        oEnvironmentSet.Name = "Environment"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oEnvironmentSet), 0
    
    'Create "Kinematics" Set in SeedLing if it does not exist and hide it
    Set oKinematicsSet = Nothing
    Set oKinematicsSet = oSeedLingPart.HybridBodies.Item("Kinematics")
    If oKinematicsSet Is Nothing Then
        Set oKinematicsSet = oSeedLingPart.HybridBodies.Add
        oKinematicsSet.Name = "Kinematics"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oKinematicsSet), 0
    
    'Create "StylingFeatures" Set in SeedLing if it does not exist and hide it
    Set oStylingFeatureSet = Nothing
    Set oStylingFeatureSet = oSeedLingPart.HybridBodies.Item("StylingFeatures")
    If oStylingFeatureSet Is Nothing Then
        Set oStylingFeatureSet = oSeedLingPart.HybridBodies.Add
        oStylingFeatureSet.Name = "StylingFeatures"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oStylingFeatureSet), 1
    
    'Create "Behaviors" Set in SeedLing if it does not exist and hide it
    Set oBehaviorSet = Nothing
    Set oBehaviorSet = oSeedLingPart.HybridBodies.Item("Behaviors")
    If oBehaviorSet Is Nothing Then
        Set oBehaviorSet = oSeedLingPart.HybridBodies.Add
        oBehaviorSet.Name = "Behaviors"
    End If
    oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oBehaviorSet), 1
    
    On Error GoTo 0
    
End Sub
Private Sub InitialScan()
    
'-------------------------------------------------------
'     FIND EXISTING ATBS AND ACTIVITY PARAMETERS
'-------------------------------------------------------
    
'***Find all ATBs in seedling part (collector) and create/reset collection aSeedLingATBs
'***Find all Activity parameters in seedling part (collector) and create/reset collection aSeedLingActivityParam
'***Find all STRCs in seedling part (collector) and create/reset collection aSeedLingSTRCs
    For Each oParameter In oSeedLingParameters
        If Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1) Like "ATB_*" Then
            aSeedLingATBs.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
        ElseIf TypeName(oParameter) = "BoolParam" And Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1) Like "Activity_*" Then
            aSeedLingActivityParam.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
       ElseIf TypeName(oParameter) = "StrParam" And Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1) Like "STRC_*" Then
            aSeedLingSTRCs.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
        End If
    Next
    
'-------------------------------------------------------
'     SCAN BODIES
'-------------------------------------------------------
'Determine available number of BodyKBE bodies and fill initial SeedLing KBEBodies List
'Analyze bodies and associated parameters values & publications, correct them if necessary
    
    On Error Resume Next
    iProgressCount2Max = oSeedLingPart.Bodies.Count
    For iI = 1 To oSeedLingPart.Bodies.Count
        Set oSeedLingCurKBEBody = oSeedLingPart.Bodies.Item(iI)
        sSeedLingCurKBEBody = oSeedLingCurKBEBody.Name
        
        'Progress bar
        iProgressCount2 = iI
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Scanning Bodies (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint("Scanning 3D Collector...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If sSeedLingCurKBEBody Like sBodyRootName & "*" Then
            sCurKBEBodyNum = Right(sSeedLingCurKBEBody, Len(sSeedLingCurKBEBody) - Len(sBodyRootName))
            Err.Clear
            iCurKBEBody = CInt(sCurKBEBodyNum)
            If Err.Number = 0 Then     'Body Name is considered because it is finished by an integer (KBE Body)
                
                If iLastKBEBody < iCurKBEBody Then iLastKBEBody = iCurKBEBody
                
                'Publish body if option "Publish elements" is selected
                Set oPublication = Nothing
                Set oPublication = oPublications.Item(sSeedLingCurKBEBody)
                If oPublication Is Nothing And (frm3DCollector.cbPublish.Value = True Or frm3DCTemplate.cbPublish.Value = True) Then 'If publication does NOT exist, create Publication
                    Set oPublication = oPublications.Add(sSeedLingCurKBEBody)
                End If
                
                If Not oPublication Is Nothing Then
                    Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & sSeedLingCurKBEBody)
                    oPublications.SetDirect sSeedLingCurKBEBody, oReference
                End If
                
                'Find "MainAssemble" in oSeedLingCurKBEBody
                Set oSeedLingCurBody = Nothing
                For Each oCurSeedLingShape In oSeedLingCurKBEBody.Shapes
                    If oCurSeedLingShape.Name = "MainAssemble" Then
                        Set oSeedLingCurBody = oCurSeedLingShape.Body
                        Exit For
                    End If
                Next
                
                'If "MainAssemble" exist
                If Not oSeedLingCurBody Is Nothing Then
                    sSeedLingCurBody = oSeedLingCurBody.Name
                    
                    Set oCurSeedLingParameter = Nothing
                    Set oCurSeedLingParameter = oSeedLingParameters.Item(sParameterKBERootName & sCurKBEBodyNum) 'Recherche du paramètre associé au BodyKBE parent du body Seed traité
                    If oCurSeedLingParameter Is Nothing Then     'Parameter not found in SeedLing => create one
                        oKBEParamSet.DirectParameters.CreateString sParameterKBERootName & sCurKBEBodyNum, sSeedLingCurBody & " " & sCollectorID
                    Else 'Parameter found in SeedLing => update value
                        'If parameter value is RECYCLE, delete what's under corresponding body
                        If UCase(oCurSeedLingParameter.ValueAsString) Like "*RECYCLE*" Then
                            
                            oCurSeedLingParameter.Value = "RECYCLE"
                            
                            'Append aSeedLingRecycleKBEBodiesList collection
                            aSeedLingRecycleKBEBodiesList.Add oSeedLingCurKBEBody
                            
                            'Put Activity Parameter to false before deleting it, then delete its publication
                            aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = False
                            aSeedLingActivityParam.Remove ("Activity_" & sSeedLingCurBody)
                            
                            Set oPublication = Nothing
                            Set oPublication = oPublications.Item("Activity_" & sSeedLingCurBody)
                            If Not oPublication Is Nothing Then oPublications.Remove (oPublication.Name)
                            
                            'Destroy Assemble/Remove/Add/Split operations
                            oSelectionFromSeedLing.Clear
                            For iJ = oSeedLingCurKBEBody.Shapes.Count To 1 Step -1
                                oSelectionFromSeedLing.Add oSeedLingCurKBEBody.Shapes.Item(iJ)
                                oSelectionFromSeedLing.Add oSeedLingCurKBEBody.Shapes.Item(iJ).Body
                            Next
                            oSelectionFromSeedLing.Delete
                            
                        'Else if parameter value is different from 1.BODYNAME COLXX or 2.RESERVED_BODYNAME COLXX, force value to BODYNAME COLXX
                        ElseIf UCase(oCurSeedLingParameter.ValueAsString) <> UCase(sSeedLingCurBody & " " & sCollectorID) And _
                               UCase(oCurSeedLingParameter.ValueAsString) <> "RESERVED_" & UCase(sSeedLingCurBody & " " & sCollectorID) Then
                            oCurSeedLingParameter.Value = sSeedLingCurBody & " " & sCollectorID
                        End If
                    End If
                    
                    'Update the list of KBE parameters found in SeedLing
                    sSeedLingKBEParametersList = sSeedLingKBEParametersList & sSeedLingCurKBEBody & ";"
    
                    'Check if KBE parameter is published. Publish it if it's not published and option "Publish elements" is selected
                    Set oPublication = Nothing
                    Set oPublication = oPublications.Item(sParameterKBERootName & sCurKBEBodyNum)
                    If oPublication Is Nothing And (frm3DCollector.cbPublish.Value = True Or frm3DCTemplate.cbPublish.Value = True) Then
                        Set oPublication = oPublications.Add(sParameterKBERootName & sCurKBEBodyNum)
                    End If
                    
                    If Not oPublication Is Nothing Then
                        'Set Reference for 3D Collector case: ParameterKBE is in parameter set "KBE_Parameters" at the root of the part
                        Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oKBEParamSet.Name & "/" & sParameterKBERootName & sCurKBEBodyNum)
                        'Set Reference for 3D Collector Template case: ParameterDBE is under BodyDBE
                        Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oSeedLingCurKBEBody.Shapes.Item(1).Body.Name & "/" & sParameterKBERootName & sCurKBEBodyNum)
                        oPublications.SetDirect sParameterKBERootName & sCurKBEBodyNum, oReference
                    End If
                    
                    'Check if Activity Parameter is published. Publish it if it's not published and option "Publish elements" is selected
                    Set oActivityParam = Nothing
                    Set oActivityParam = aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody)
                    If Not oActivityParam Is Nothing Then
                        Set oPublication = Nothing
                        Set oPublication = oPublications.Item("Activity_" & sSeedLingCurBody)
                        If oPublication Is Nothing And (frm3DCollector.cbPublish.Value = True Or frm3DCTemplate.cbPublish.Value = True) Then
                            Set oPublication = oPublications.Add("Activity_" & oSeedLingCurKBEBody.Shapes.Item(1).Body.Name)
                        End If
                        
                        If Not oPublication Is Nothing Then
                            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oSeedLingCurKBEBody.Shapes.Item(1).Body.Name & "/" & "Activity_" & oSeedLingCurKBEBody.Shapes.Item(1).Body.Name)
                            oPublications.SetDirect "Activity_" & oSeedLingCurKBEBody.Shapes.Item(1).Body.Name, oReference
                        End If
                    End If
                    
                Else    ' "MainAssemble" was not found under KBE body => KBE Body to be recycled
                    Set oCurSeedLingParameter = Nothing
                    Set oCurSeedLingParameter = oSeedLingParameters.Item(sParameterKBERootName & sCurKBEBodyNum) 'Returns nothing if parameter associated to KBE Body is not found
                    
                    If oCurSeedLingParameter Is Nothing Then     'Parameter not found in SeedLing => create one with NOMAINBODY value
                        oKBEParamSet.DirectParameters.CreateString sParameterKBERootName & sCurKBEBodyNum, "RESERVED_NOMAINBODY"
                    
                    ElseIf UCase(oCurSeedLingParameter.ValueAsString) Like "*RECYCLE*" Then     'Parameter found in SeedLing => update value
                        
                        oCurSeedLingParameter.Value = "RECYCLE"
                        
                        'Append aSeedLingRecycleKBEBodiesList collection
                        aSeedLingRecycleKBEBodiesList.Add oSeedLingCurKBEBody
                        
                        'Put Activity Parameter to false
                        aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = False
                        
                        'Destroy Assemble/Remove/Add/Split operations
                        oSelectionFromSeedLing.Clear
                        For iJ = oSeedLingCurKBEBody.Shapes.Count To 1 Step -1
                            oSelectionFromSeedLing.Add oSeedLingCurKBEBody.Shapes.Item(iJ)
                            oSelectionFromSeedLing.Add oSeedLingCurKBEBody.Shapes.Item(iJ).Body
                        Next
                        oSelectionFromSeedLing.Delete
                        
                    Else
                        oCurSeedLingParameter.Value = "RESERVED_NOMAINBODY"
                    End If
                    
                    'Update the list of KBE parameters found in SeedLing
                    sSeedLingKBEParametersList = sSeedLingKBEParametersList & sSeedLingCurKBEBody & ";"
    
                    'Check if KBE parameter is published. Publish it if it's not published and option "Publish elements" is selected
                    Set oPublication = Nothing
                    Set oPublication = oPublications.Item(sParameterKBERootName & sCurKBEBodyNum) 'Returns nothing if publication not found
                    If oPublication Is Nothing And (frm3DCollector.cbPublish.Value = True Or frm3DCTemplate.cbPublish.Value = True) Then
                        Set oPublication = oPublications.Add(sParameterKBERootName & sCurKBEBodyNum)
                    End If
                    
                    If Not oPublication Is Nothing Then
                        Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oKBEParamSet.Name & "/" & sParameterKBERootName & sCurKBEBodyNum)
                        oPublications.SetDirect sParameterKBERootName & sCurKBEBodyNum, oReference
                    End If
                    
                End If
                oSeedLingPart.Update
            End If
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
'-------------------------------------------------------
'     SCAN KBE PARAMETERS
'-------------------------------------------------------
'Find all Parameters KBE that have no corresponding KBE Bodies, then create a Recycle Body to be used later on
'Publish all Parameters KBE if needed
'Publish all Activity Parameters if needed
    iProgressCount2Max = oKBEParamSet.DirectParameters.Count
    For iI = 1 To oKBEParamSet.DirectParameters.Count
        
        Set oCurSeedLingParameter = Nothing
        Set oCurSeedLingParameter = oKBEParamSet.DirectParameters.Item(iI) 'Returns Nothing is parameter is not a String parameter
        
        'Progress bar
        iProgressCount2 = iI
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Scanning KBE Parameter (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
         
        If Not oCurSeedLingParameter Is Nothing Then
            If oCurSeedLingParameter.Name Like "*" & sParameterKBERootName & "*" Then
                
                iPos = InStrRev(oCurSeedLingParameter.Name, sParameterKBERootName)
                sCurKBEBodyNum = Right(oCurSeedLingParameter.Name, Len(oCurSeedLingParameter.Name) - iPos - Len(sParameterKBERootName) + 1)
                Err.Clear
                iCurKBEBody = CInt(sCurKBEBodyNum)
                If Err.Number = 0 Then  'Parameter name has the nomenclature of a KBE parameter.
                    
                    sSeedLingCurKBEBody = sBodyRootName & sCurKBEBodyNum
                    iPos = InStr(1, sSeedLingKBEParametersList, sSeedLingCurKBEBody, vbBinaryCompare)
                    
                    If iPos = 0 Then    'Parameter is not linked to any KBE Body => set parameter value to RECYCLE and create corresponding KBE body
                        
                        'Put activity parameter to False
                        sSeedLingCurBody = Mid(oCurSeedLingParameter.ValueAsString, 1, InStrRev(oCurSeedLingParameter.ValueAsString, " COL") - 1)
                        aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = False
                        
                        'Modify KBE Parameter value
                        oCurSeedLingParameter.Value = "RECYCLE"
                        
                        'Update the list of KBE parameters found in SeedLing
                        sSeedLingKBEParametersList = sSeedLingKBEParametersList & sSeedLingCurKBEBody & ";"
                        
                        Set oSeedLingCurKBEBody = oSeedLingPart.Bodies.Add
                        oSeedLingCurKBEBody.Name = sSeedLingCurKBEBody
                        
                        'Append aSeedLingRecycleKBEBodiesList collection
                        aSeedLingRecycleKBEBodiesList.Add oSeedLingCurKBEBody
                        
                        'Check if Current SeedLing KBE body has already been published. Publish it if it's not published and option "Publish elements" is selected
                        Set oPublication = Nothing
                        Set oPublication = oPublications.Item(sSeedLingCurKBEBody)
                        If oPublication Is Nothing And frm3DCollector.cbPublish.Value = True Then
                            Set oPublication = oPublications.Add(sSeedLingCurKBEBody)
                        End If
                        
                        If Not oPublication Is Nothing Then
                            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & sSeedLingCurKBEBody)
                            oPublications.SetDirect sSeedLingCurKBEBody, oReference
                        End If
                        
                        'Check if KBE Parameter is published. Publish it if it's not published and option "Publish elements" is selected
                        Set oPublication = Nothing
                        Set oPublication = oPublications.Item(sParameterKBERootName & sCurKBEBodyNum)
                        If oPublication Is Nothing And frm3DCollector.cbPublish.Value = True Then
                            Set oPublication = oPublications.Add(sParameterKBERootName & sCurKBEBodyNum)
                        End If
                        
                        If Not oPublication Is Nothing Then
                            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oKBEParamSet.Name & "/" & sParameterKBERootName & sCurKBEBodyNum)
                            oPublications.SetDirect sParameterKBERootName & sCurKBEBodyNum, oReference
                        End If
                        
                    End If
                End If
            End If
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
'-------------------------------------------------------
'     VERIFY EXISTING DATUM
'-------------------------------------------------------
   
    Call scanExistingDatum(oSkeletonSet)
    Call scanExistingDatum(oOperationsSet)
    Call scanExistingDatum(oEnvironmentSet)
    Call scanExistingDatum(oKinematicsSet)
    Call scanExistingDatum(oStylingFeatureSet)
    Call scanExistingDatum(oBehaviorSet)
    
    
'-------------------------------------------------------
'     VERIFY EXISTING OS
'-------------------------------------------------------
    For Each oCurSeedLingOS In oOperationsSet.HybridShapes
        '*******If surface name contains "_OS_" but does not end with "_OS_"******************
        If oCurSeedLingOS.Name Like "*_OS_*" And Not oCurSeedLingOS.Name Like "*_OS_" Then
            aSeedLingOS.Add oCurSeedLingOS, oCurSeedLingOS.Name 'Error if we have 2 OS with the same name. Do not delete it (can be linked to a split), but do not add it to collection)
        End If
    Next
    
    On Error GoTo 0
    
End Sub

Private Sub scanExistingDatum(ByRef oGeoset As HybridBody)
    
    On Error Resume Next
    For Each oCurSeedLingDatum In oGeoset.HybridShapes
        
        'If we have 2 Datum with the same name > Delete the existing one
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(aSeedLingDatum.Item(oCurSeedLingDatum.Name))
        aSeedLingDatum.Add oCurSeedLingDatum, oCurSeedLingDatum.Name
    
        'Publish Datum if option checked
        Set oPublication = Nothing
        Set oPublication = oPublications.Item(Replace(oCurSeedLingDatum.Name, "EXTRA_", "")) 'Should return nothing if publication was deleted previously
        If oPublication Is Nothing And (frm3DCollector.cbPublish.Value = True Or frm3DCTemplate.cbPublish.Value = True) Then
            Set oPublication = oPublications.Add(Replace(oCurSeedLingDatum.Name, "EXTRA_", ""))
        End If
        
        If Not oPublication Is Nothing Then
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & oCurSeedLingDatum.Name)
            oPublications.SetDirect Replace(oCurSeedLingDatum.Name, "EXTRA_", ""), oReference
        End If
        
    Next
    On Error GoTo 0
    
End Sub

Public Sub MainTreatment(ByVal dMainProgressValue As Double, ByVal dProgressIncrementValue As Double)
'********************************************************************************
'* Name: MainTreatment
'*
'* New: Refer to comments identified by Update 18-March 2015
'* Author:
'* Updated by: Abhishek Kamboj 18 March 2015
'* Language: VBA
'******************************************************************************
    Dim bCollectSSParam As Boolean
    Dim CheckSubset As String
    sSeedPartBodiesList = ""
    
'-------------------------------------------------------
'     SCAN PARAMETERS IN SEED PART AND COLLECTOR
'-------------------------------------------------------
    sSeedPartLetter = Split(Split(oSeedPart.Name, "_SEED_PART_")(1), "-")(0)
    iProgressCount1 = iProgressCount1 + (1 / 10) '1 of 10
    iProgressCount2Max = 1
    iProgressCount2 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Scanning Seed Part..."
    Call frmProgress.progressBarRepaint("Collecting from : " & oSeedPart.Name, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    'Do not collect ZA (Skeleton Part) if other parts were collected
    If sSeedPartLetter = "ZA" Then
        For Each oSeedLingAxisSystem In oSeedLingPart.AxisSystems
            If oSeedLingAxisSystem.Name Like "*DATUM_AXIS*" And Not oSeedLingAxisSystem.Name Like "ZA_*" Then Exit Sub
        Next
    End If
    Call ScanSeedPartParameters
    
'-------------------------------------------------------
'           COLLECTING SEED PART BODIES
'-------------------------------------------------------
    On Error Resume Next
    iProgressCount1 = iProgressCount1 + (1 / 10) '2 of 10
    iProgressCount2Max = oSeedPart.Bodies.Count
    iProgressCount2 = 0
    
    For Each oSeedPartCurBody In oSeedPart.Bodies
        
        'ProgressBarSettings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting KBE body (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ") : " & oSeedPartCurBody.Name
        Call frmProgress.progressBarRepaint("Collecting from : " & oSeedPart.Name, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        '*****If Subset is used, verify if the current body has to be collected
        '*****Update18-March-2015 If InvertSelectin is checked then collect everything except bodies in Subset
        If frm3DCollector.cbSubSet.Value = True Then
            If Not frmSubsetsSelection.cbInvertSelection.Value Then
                bCopySelection = False
                For Each oSSParameterValue In SSParameters.Item(sCurrentSSParameterName).SubsetValues
                    If oSSParameterValue = oSeedPartCurBody.Name Then
                        bCopySelection = True
                        Exit For
                    End If
                Next
            Else
                bCopySelection = True
                For Each oSSParameterValue In SSParameters.Item(sCurrentSSParameterName).SubsetValues
                    If oSSParameterValue = oSeedPartCurBody.Name Then
                        bCopySelection = False
                        Exit For
                    End If
                Next
            End If
        Else
            bCopySelection = True
        End If
        
        'Copy / paste body into collector
        If oSeedPartCurBody.Shapes.Count > 0 And (oSeedPartCurBody.Name Like "KBE_*") And bCopySelection = True Then      'Only non-empty and head bodies are treated
            sSeedPartBodiesList = sSeedPartBodiesList & oSeedPartCurBody.Name & ";"
            
            '******* Looking if body already exist in collector ******
            Set oSeedLingCurBody = Nothing
            Set oSeedLingCurBody = oSeedLingPart.Bodies.Item(sSeedPartLetter & "_" & oSeedPartCurBody.Name) 'Returns nothing if body is not found
            sSeedLingCurBody = sSeedPartLetter & "_" & oSeedPartCurBody.Name
            
            '***** Finding the value of the body boolean parameter. If value is not found, it put it to true (consider Seed Part body as active) >> ICD Parts have no boolean parameters
            bSeedPartBooleanParameter = True
            bSeedPartBooleanParameter = CBool(aSeedPartBooleans.Item(oSeedPartCurBody.Name & "_Boolean").Value)
            
            If bSeedPartBooleanParameter Then    'Copy/paste body only if body is active
                oSelectionFromSeedPart.Clear
                oSelectionFromSeedPart.Add oSeedPartCurBody
                oSelectionFromSeedPart.Copy
                oSelectionFromSeedPart.Clear
                
                oSelectionFromSeedLing.Clear
                oSelectionFromSeedLing.Add oSeedLingPart
                oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
                oSelectionFromSeedLing.Clear
                
                Set oNewBody = oSeedLingPart.Bodies.Item(oSeedLingPart.Bodies.Count)
                
                If Not oSeedLingCurBody Is Nothing Then
                    
                    'Destroy Assemble/Remove/Add/Split (there should be only one dead body)
                    oSelectionFromSeedLing.Clear
                    For iJ = oSeedLingCurBody.Shapes.Count To 1 Step -1
                        oSelectionFromSeedLing.Add oSeedLingCurBody.Shapes.Item(iJ)
                        oSelectionFromSeedLing.Add oSeedLingCurBody.Shapes.Item(iJ).Body
                    Next iJ
                    
                    For iJ = oSeedLingCurBody.Sketches.Count To 1 Step -1
                        oSelectionFromSeedLing.Add oSeedLingCurBody.Sketches.Item(iJ)
                    Next iJ
                    oSelectionFromSeedLing.Delete
                    
                    'Copy geometry inside oSeedLingCurBody
                    oSelectionFromSeedLing.Add oNewBody.Shapes.Item(1)
                    oSelectionFromSeedLing.Copy
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oSeedLingCurBody
                    oSelectionFromSeedLing.Paste
                    oSelectionFromSeedLing.Clear
                    
                    'Destroy shapes under parent body
                    Set oSeedLingCurBody = GetBodyFromObject(oSeedLingCurBody)
                    oSelectionFromSeedLing.Clear
                    For iJ = oSeedLingCurBody.Shapes.Count To 1 Step -1
                        If oSeedLingCurBody.Shapes.Item(iJ).Body.Name <> sSeedLingCurBody Then oSelectionFromSeedLing.Add oSeedLingCurBody.Shapes.Item(iJ)
                    Next iJ
                    
                    For iJ = oSeedLingCurBody.Sketches.Count To 1 Step -1
                        oSelectionFromSeedLing.Add oSeedLingCurBody.Sketches.Item(iJ)
                    Next iJ
                    oSelectionFromSeedLing.Delete
                    
                    'Destroy oNewBody
                    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oNewBody)
                                            
                    Set oSeedLingCurKBEBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, sSeedLingCurBody & " " & sCollectorID, oCurSeedLingParameter)
                    If oSeedLingCurKBEBody Is Nothing Then
                        Set oSeedLingCurKBEBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, "RESERVED_" & sSeedLingCurBody & " " & sCollectorID, oCurSeedLingParameter)
                    End If
                    If Not (oSeedLingCurKBEBody Is Nothing) Then

                        'Modify value of KBE Parameter
                        oCurSeedLingParameter.Value = sSeedLingCurBody & " " & sCollectorID 'Remove "RESERVED_" string in parameter value if it exists
                        
                        'Put activity parameter to True, hide body
                        aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = True
                        oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody), 1 'Set KBE Body to Show mode if required
                    End If
                    
                Else    'Body not found => use a RECYCLE KBE Body if any exists or create a new KBE Body
                    If aSeedLingRecycleKBEBodiesList.Count > 0 Then
                        Set oSeedLingCurKBEBody = aSeedLingRecycleKBEBodiesList.Item(1)
                        aSeedLingRecycleKBEBodiesList.Remove (1)
                        sSeedLingCurKBEBody = oSeedLingCurKBEBody.Name
                        sCurKBEBodyNum = Right(sSeedLingCurKBEBody, Len(sSeedLingCurKBEBody) - Len(sBodyRootName))
                        Set oCurSeedLingParameter = oSeedLingParameters.Item(sParameterKBERootName & sCurKBEBodyNum)
                        
                        'Modify value of KBE Parameter
                        oCurSeedLingParameter.Value = sSeedLingCurBody & " " & sCollectorID 'Recherche du paramètre associé au BodyKBE parent du body Seed traité
                        
                        'Put activity parameter to True, hide body
                        aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = True
                        oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody), 1 'Set KBE Body to Show mode if required
                    Else
                        Do
                            Set oSeedLingCurKBEBody = Nothing
                            iLastKBEBody = iLastKBEBody + 1
                            sCurKBEBodyNum = Format(CStr(iLastKBEBody), "0#")
                            Set oSeedLingCurKBEBody = oSeedLingPart.Bodies.Item(sBodyRootName & sCurKBEBodyNum)
                        Loop Until oSeedLingCurKBEBody Is Nothing
                        
                        Set oSeedLingCurKBEBody = oSeedLingPart.Bodies.Add
                        
                        sSeedLingCurKBEBody = sBodyRootName & sCurKBEBodyNum
                        oSeedLingCurKBEBody.Name = sSeedLingCurKBEBody
                        oKBEParamSet.DirectParameters.CreateString sParameterKBERootName & sCurKBEBodyNum, sSeedLingCurBody & " " & sCollectorID
                            
                        'Update the list of KBE parameters found in SeedLing
                        sSeedLingKBEParametersList = sSeedLingKBEParametersList & sSeedLingCurKBEBody & ";"
                        
                        'Check if Current SeedLing body is published. Publish it if it's not published and option "Publish elements" is selected
                        Set oPublication = Nothing
                        Set oPublication = oPublications.Item(sSeedLingCurKBEBody)
                        If oPublication Is Nothing And frm3DCollector.cbPublish.Value = True Then
                            Set oPublication = oPublications.Add(sSeedLingCurKBEBody)
                        End If
                        
                        If Not oPublication Is Nothing Then
                            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & sSeedLingCurKBEBody)
                            oPublications.SetDirect sSeedLingCurKBEBody, oReference
                        End If
                        
                        'Check if KBE Parameter is published. Publish it if it's not published and option "Publish elements" is selected
                        Set oPublication = Nothing
                        Set oPublication = oPublications.Item(sParameterKBERootName & sCurKBEBodyNum)
                        If oPublication Is Nothing And frm3DCollector.cbPublish.Value = True Then
                            Set oPublication = oPublications.Add(sParameterKBERootName & sCurKBEBodyNum)
                        End If
                        
                        If Not oPublication Is Nothing Then
                            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oKBEParamSet.Name & "/" & sParameterKBERootName & sCurKBEBodyNum)
                            oPublications.SetDirect sParameterKBERootName & sCurKBEBodyNum, oReference
                        End If
                        
                    End If
                    oSeedLingPart.InWorkObject = oSeedLingCurKBEBody
                    Set oAssemble = oSeedLingPart.ShapeFactory.AddNewAssemble(oNewBody)
                    oAssemble.Name = "MainAssemble"
                    sSeedLingCurBody = sSeedPartLetter & "_" & oSeedPartCurBody.Name
                    oNewBody.Name = sSeedLingCurBody
                    
                    Call AddActivityParam
                End If
                
                '*****Look all ATB params, create an array of parameters related to current KBE Body
                For Each oParameter In aSeedPartDesiredATBs
                    Set oTempParameter = Nothing
                    Set oTempParameter = aActiveBodiesATBs.Item(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1))
                
                    If oParameter.Name Like "*" & Replace(oSeedPartCurBody.Name, "KBE", "ATB") & "_*" And oTempParameter Is Nothing Then
                        aActiveBodiesATBs.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
                        aSeedPartNotCollectedATBs.Remove (Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1))
                    End If
                Next

            ElseIf Not oSeedLingCurBody Is Nothing Then     'Body found In SeedLing but it is not active in Seed Part > Reserve it
                'Retrieve KBE Body containing SeedLing current Body by analyzing parameters
                Set oSeedLingCurKBEBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, sSeedLingCurBody & " " & sCollectorID, oCurSeedLingParameter)
                
                'Modify value of KBE Parameter
                oCurSeedLingParameter.Value = "RESERVED_" & sSeedLingCurBody & " " & sCollectorID
                
                'Put activity parameter to False, Hide KBE Body
                aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = False
                oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody), 0 'Set KBE Body to NoShow mode if required
                
            End If
            oSeedLingPart.Update
            
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
'-------------------------------------------------------
'    SCAN SEEDLING BODIES - COMPARE TO CURRENT SEED PART
'-------------------------------------------------------
    'Determine if Current SeedLing body has been found in corresponding Seed Part.
    'If not (i.e. Body was deleted from Seed Part), Update or Create associated parameter with RESERVED_ keyword parameter and publish the KBEbody if not published
    
    iProgressCount1 = iProgressCount1 + (1 / 10) '3 of 10
    iProgressCount2Max = oSeedLingPart.Bodies.Count
    iProgressCount2 = 0
    
    For Each oSeedLingCurKBEBody In oSeedLingPart.Bodies

        sSeedLingCurKBEBody = oSeedLingCurKBEBody.Name
        
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Scanning body (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ") : " & oSeedLingCurKBEBody.Name
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If sSeedLingCurKBEBody Like sBodyRootName & "*" Then
            sCurKBEBodyNum = Right(sSeedLingCurKBEBody, Len(sSeedLingCurKBEBody) - Len(sBodyRootName))
            Err.Clear
            iCurKBEBody = CInt(sCurKBEBodyNum)
            If Err.Number = 0 Then     'Body Name is considered if it finishes by an integer (KBE Body)
                'Look for Body named KBE* in SeedLing KBE Body Main Assemble
                Set oSeedLingCurBody = Nothing
                For Each oCurSeedLingShape In oSeedLingCurKBEBody.Shapes
                    If oCurSeedLingShape.Name Like "MainAssemble" Then
                        Set oSeedLingCurBody = oCurSeedLingShape.Body
                        Exit For
                    End If
                Next
                If Not oSeedLingCurBody Is Nothing Then
                    
                    sSeedLingCurBody = oSeedLingCurBody.Name
                    If Left(sSeedLingCurBody, Len(sSeedPartLetter)) = sSeedPartLetter Then     'Only consider bodies from Selected Seed Part
                        iPos = InStr(1, sSeedPartBodiesList, Mid(sSeedLingCurBody, Len(sSeedPartLetter) + 2), vbBinaryCompare)
                        If iPos = 0 Then    'Body not found in corresponding Seed Part List => body no longer available
                            
                            On Error Resume Next
                            'Modify value of KBE Parameter - Value of parameter corresponding to associated KBE Body set to "RESERVED" status
                            Set oCurSeedLingParameter = oSeedLingParameters.Item(sParameterKBERootName & sCurKBEBodyNum)
                            oCurSeedLingParameter.Value = "RESERVED_" & sSeedLingCurBody & " " & sCollectorID 'oCurSeedLingParameter.Value 'Recherche du paramètre associé au BodyKBE parent du body Seed traité
                            
                            'Put activity parameter to False, Hide body
                            aSeedLingActivityParam.Item("Activity_" & sSeedLingCurBody).Value = False
                            oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody), 0 'Set KBE Body to NoShow mode if required
                            On Error GoTo 0
                            
                        End If
                    End If
                End If
            End If
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
'-------------------------------------------------------
'     UPDATING COLLECTOR
'-------------------------------------------------------
    Err.Clear
    oSeedLingPart.Update
    If Err.Number <> 0 Then
        MsgBox oSeedLingPart.Name + " cannot be updated !   " + vbNewLine + "Use ""Reset Collector"" option, or update it manually.   ", vbExclamation, "Warning"
        bCancelAction = True
    End If
    On Error GoTo 0
    
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
'-------------------------------------------------------
'     DATUM COLLECTION
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + (1 / 10) '4 of 10
    Call CopyDatum
    
'-------------------------------------------------------
'         COPY/PASTE A/C AXIS SYSTEM
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + (1 / 10) '5 of 10
    Call CopySeedlingAxisSystem
    
'-------------------------------------------------------
'     COLLECT SS PARAMETERS
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + (1 / 10) '6 of 10
    
    On Error Resume Next
    Set seedAssySubSetCollection = New Collection
    Set seedlingSubSetCollection = New Collection
    
    ' Determine SP and collector SS parameter sets
    '**** Two names of parameter set to cover: "SubSetParameters" & "ParametersSubSet"
    Set SeedPartSSParamSet = Nothing
    Set SeedPartSSParamSet = oSeedPart.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
    Set SeedPartSSParamSet = oSeedPart.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
    
    Set SeedLingSSParamSet = Nothing
    Set SeedLingSSParamSet = oSeedLingPart.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
    Set SeedLingSSParamSet = oSeedLingPart.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
    
    If Not (SeedPartSSParamSet Is Nothing) Then
        
        '*****Update18-March-2015  if subsets are selected then only selected SS parameters are copied else all SS parameters are copied.
        ' Fill SP SS Parameters collection
        For Each oParameter In SeedPartSSParamSet.DirectParameters
        
            If frm3DCollector.cbSubSet.Value = True Then ' if subsets are selected then only selected SS parameters are copied else all SS parameters are copied.
                bCollectSSParam = False
                Err.Clear
                CheckSubset = SelectedSSParameters.Item(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 4))
                 
                If Err.Number = 0 Then
                    bCollectSSParam = True
                End If
            Else
                bCollectSSParam = True
            End If
            
            If bCollectSSParam Then
               seedAssySubSetCollection.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
            End If
        Next
        
        ' Create collector SubSetParameters parameter set if it does not exist
        If SeedLingSSParamSet Is Nothing Then
            Set SeedLingSSParamSet = oSeedLingPart.Parameters.RootParameterSet.ParameterSets.CreateSet("ParametersSubSet")
        End If
        
        ' Fill collector SS Parameters collection
        For Each oParameter In SeedLingSSParamSet.DirectParameters
            seedlingSubSetCollection.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
        Next
        
        ' Copy SS Parameters from SP to Collector
        iProgressCount2Max = seedAssySubSetCollection.Count
        iProgressCount2 = 0
        For Each oParameter In seedAssySubSetCollection
            
            'Progress Bar Settings
            iProgressCount2 = iProgressCount2 + 1
            timeEnd = DateTime.Timer
            sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
            sProgBarComment = "Collecting SS Parameter (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)

            'Delete existing parameter
            oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(seedlingSubSetCollection(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)))
            
            'Copy from SP
            oSelectionFromSeedPart.Clear
            oSelectionFromSeedPart.Add oParameter
            oSelectionFromSeedPart.Copy
            oSelectionFromSeedPart.Clear
            
            'Paste into collector
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add SeedLingSSParamSet
            oSelectionFromSeedLing.Paste
            oSelectionFromSeedLing.Clear
            
        Next
    End If
    
    
'-------------------------------------------------------
'     COLLECT STRC (CONTROL) PARAMETERS
'-------------------------------------------------------

    iProgressCount1 = iProgressCount1 + (1 / 10) '7 of 10  '*****************************TO MODIFY COUNT DOWN
    On Error Resume Next
    
    ' Copy SS Parameters from SP to Collector
    iProgressCount2Max = aSeedPartSTRCs.Count
    iProgressCount2 = 0
    For Each oParameter In aSeedPartSTRCs

        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting STRC Parameter (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)

        'Delete existing parameter
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(aSeedLingSTRCs(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)))

        'Copy from SP
        oSelectionFromSeedPart.Clear
        oSelectionFromSeedPart.Add oParameter
        oSelectionFromSeedPart.Copy
        oSelectionFromSeedPart.Clear

        'Paste into collector
        oSelectionFromSeedLing.Clear
        oSelectionFromSeedLing.Add oSTRCParamSet
        oSelectionFromSeedLing.Paste
        oSelectionFromSeedLing.Clear

    Next
    
    
'-------------------------------------------------------
'     CLEAN ATB PARAMETERS
'-------------------------------------------------------
    '*****Copy all ATB Parameters related to active KBE bodies
    
    iProgressCount1 = iProgressCount1 + (1 / 10) '8 of 10
    
    On Error Resume Next
    oSelectionFromSeedPart.Clear
    oSelectionFromSeedLing.Clear
    iProgressCount2Max = aActiveBodiesATBs.Count
    iProgressCount2 = 0
    For Each oParameter In aActiveBodiesATBs
        
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Copying ATB Parameter (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        'Delete existing seedling ATB parameters
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(aSeedLingATBs.Item(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)))
        
        'Add existing seedpart ATB parameters to selection to copy/paste them in next step
        oSelectionFromSeedPart.Add oParameter
        
        'Remove collected ATBs from collection aSeedPartNotCollectedATBs
        aSeedPartNotCollectedATBs.Remove (Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1))
    Next
    
    '*****Remove collected ATBs from collection aSeedPartNotCollectedATBs
    If oSelectionFromSeedPart.Count > 0 Then
        oSelectionFromSeedPart.Copy
        oSelectionFromSeedPart.Clear
        
        oSelectionFromSeedLing.Clear
        oSelectionFromSeedLing.Add oATBParamSet
        oSelectionFromSeedLing.Paste
        oSelectionFromSeedLing.Clear
    End If
    
    '*****Delete all ATB that are present in SeedLing, but not collected from seed part*****
    iProgressCount2Max = aSeedPartNotCollectedATBs.Count
    iProgressCount2 = 0
    For Each oParameter In aSeedPartNotCollectedATBs
    
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Cleaning ATB Parameter (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        'Delete existing seedling ATB parameters
        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(aSeedLingATBs.Item(Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)))
        
    Next
    On Error GoTo 0

'-------------------------------------------------------
'     END
'-------------------------------------------------------
    On Error Resume Next
    oSeedLingPart.InWorkObject = oSeedLingPart.MainBody
    oSeedLingPart.Update
    sTargetWindow.ActiveViewer.Reframe
    On Error GoTo 0
    
    '------------ SELECTED SEED PART OPERATIONS ENDING PHASE --------------
    '**********************************************************************
    
End Sub
Private Sub CopySeedlingAxisSystem()
    
    On Error Resume Next
    iProgressCount2 = 0
    iProgressCount2Max = oSeedPart.AxisSystems.Count
    For Each iObject In oSeedPart.AxisSystems
        
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting Axis (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If iObject.Name Like "DATUM_AXIS*" Then
            
            'Delete Original Axis system if it exists
            sSeedLingAxisSystem = sSeedPartLetter & "_" & iObject.Name
            Set oSeedLingAxisSystem = Nothing
            Set oSeedLingAxisSystem = oSeedLingPart.AxisSystems.Item(sSeedLingAxisSystem)
            If Not oSeedLingAxisSystem Is Nothing Then 'Only if A/C Axis System found in SeedLing => delete previous one.
                oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSeedLingAxisSystem)
            End If
            
            'Copy/Paste New Axis system from Seed Part to SeedLing
            oSelectionFromSeedPart.Clear
            oSelectionFromSeedPart.Add iObject
            oSelectionFromSeedPart.Copy
            oSelectionFromSeedPart.Clear
            
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oSeedLingPart
            oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
            oSelectionFromSeedLing.Clear
            
            'Rename copied A/C Axis System without "/" > Impossible to create reference for publication with "/"
            Set oSeedLingAxisSystem = oSeedLingPart.AxisSystems.Item(oSeedLingPart.AxisSystems.Count)
            oSeedLingAxisSystem.Name = Replace(sSeedLingAxisSystem, "/", "")
            
            'Set created Axis No Show mode
            oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oSeedLingAxisSystem), 0
            
            'Delete previous A/C Axis system (so new publication appears at the end)
            For j = oPublications.Count To 1 Step -1
                If oPublications.Item(j).Name = Mid(sSeedLingAxisSystem, InStr(sSeedLingAxisSystem, iObject.Name)) Then
                    oPublications.Remove (oPublications.Item(j).Name)
                    Exit For
                End If
            Next
            
            'Publish A/C Axis system
            Set oPublication = Nothing
            Set oPublication = oPublications.Item(Mid(sSeedLingAxisSystem, InStr(sSeedLingAxisSystem, iObject.Name))) 'Should return nothing if publication was deleted previously
            If oPublication Is Nothing Then
                Set oPublication = oPublications.Add(Mid(sSeedLingAxisSystem, InStr(sSeedLingAxisSystem, iObject.Name)))
            End If
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & oSeedLingAxisSystem.Name)
            oPublications.SetDirect Mid(sSeedLingAxisSystem, InStr(sSeedLingAxisSystem, iObject.Name)), oReference
            
            'Rename copied A/C Axis System
            oSeedLingAxisSystem.Name = sSeedLingAxisSystem
            
         End If
         If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next

    On Error GoTo 0

End Sub

Private Sub CopyDatum()
        
    On Error Resume Next
    
    sSearchString = ""
    If frm3DCollector.cbSK.Value Then sSearchString = "CATGmoSearch.GeometricFeature.Name_CAP=SK_*"
    If frm3DCollector.cbBVR.Value Then
        sSearchString = IIf(sSearchString = "", "", sSearchString & " + "): sSearchString = sSearchString & "CATGmoSearch.GeometricFeature.Name_CAP=BVR_*": End If
    If frm3DCollector.cbSTY.Value Then
        sSearchString = IIf(sSearchString = "", "", sSearchString & " + "): sSearchString = sSearchString & "CATGmoSearch.GeometricFeature.Name_CAP=STY_*": End If
    If frm3DCollector.cbENV Then
        sSearchString = IIf(sSearchString = "", "", sSearchString & " + "): sSearchString = sSearchString & "CATGmoSearch.GeometricFeature.Name_CAP=ENV_*": End If
    If frm3DCollector.cbKIN.Value Then
        sSearchString = IIf(sSearchString = "", "", sSearchString & " + "): sSearchString = sSearchString & "CATGmoSearch.GeometricFeature.Name_CAP=KIN_*": End If
    sSearchString = IIf(sSearchString = "", "", sSearchString & ",sel")
    If sSearchString = "" Then Exit Sub
    
'    sSearchString = _
'    "CATGmoSearch.GeometricFeature.Name_CAP=SK_* + " & _
'    "CATGmoSearch.GeometricFeature.Name_CAP=ENV_* + " & _
'    "CATGmoSearch.GeometricFeature.Name_CAP=KIN_* + " & _
'    "CATGmoSearch.GeometricFeature.Name_CAP=STY_* + " & _
'    "CATGmoSearch.GeometricFeature.Name_CAP=BVR_*,sel"

    'ProgressBarSettings
    sProgBarComment = "Searching Datum elements..."
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, 1, 0.001, sTimer)
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    oSelectionFromSeedPart.Clear
    oSelectionFromSeedPart.Add oSeedPart
    Set oSelectionFromSeedPart = KBESearch(sSearchString, , oSelectionFromSeedPart)
    
    Set aSeedPartDatum = New Collection
    sSeedPartDatum = ""
    For icount = 1 To oSelectionFromSeedPart.Count
        If Not oSelectionFromSeedPart.Item2(icount).Value.Parent.Parent.Name Like "*External References*" Then aSeedPartDatum.Add oSelectionFromSeedPart.Item2(icount).Value 'Exclude Datum elements that belong to ExternalReferences Geometrical Set
        sSeedPartDatum = sSeedPartDatum & aSeedPartDatum.Item(aSeedPartDatum.Count).Name & ";"
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    oSelectionFromSeedPart.Clear
    
    ' Copy all datum from Seed Part (except those in External References)
    iProgressCount2Max = aSeedPartDatum.Count
    For iI = 1 To aSeedPartDatum.Count
        
        'Progress Bar Settings
        iProgressCount2 = iI
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting Datum elements (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If aSeedPartDatum.Item(iI).Name Like "SK_*" Then
            Set oDatumSet = oSkeletonSet
        ElseIf aSeedPartDatum.Item(iI).Name Like "ENV_*" Then Set oDatumSet = oEnvironmentSet
        ElseIf aSeedPartDatum.Item(iI).Name Like "KIN_*" Then Set oDatumSet = oKinematicsSet
        ElseIf aSeedPartDatum.Item(iI).Name Like "STY_*" Then Set oDatumSet = oStylingFeatureSet
        ElseIf aSeedPartDatum.Item(iI).Name Like "BVR_*" Then Set oDatumSet = oBehaviorSet
        Else: Set oDatumSet = oSkeletonSet: End If
        
        sCurSeedLingDatum = sSeedPartLetter & "_" & aSeedPartDatum.Item(iI).Name
        
        'Paste new Surface from Seed Part.
        oSelectionFromSeedPart.Clear
        oSelectionFromSeedPart.Add aSeedPartDatum.Item(iI)
        oSelectionFromSeedPart.Copy
        
        oSelectionFromSeedLing.Clear
        oSelectionFromSeedLing.Add oDatumSet
        oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
        
        Set oCurSeedLingDatum = Nothing
        Set oCurSeedLingDatum = aSeedLingDatum.Item(sCurSeedLingDatum)
        
        If oCurSeedLingDatum Is Nothing Then
            
            'Rename created Surface, add it to Datum Collection
            Set oCurSeedLingDatum = oDatumSet.HybridShapes.Item(oDatumSet.HybridShapes.Count)
            oCurSeedLingDatum.Name = sCurSeedLingDatum
            aSeedLingDatum.Add oCurSeedLingDatum, sCurSeedLingDatum
            
        Else
            Set oFormula = oSeedLingPart.Relations.CreateFormula("oFormula_" & sCurSeedLingDatum, "", oSeedLingPart.Parameters.Item(oCurSeedLingDatum.Name), "`" & oDatumSet.Name & "\" & oDatumSet.HybridShapes.Item(oDatumSet.HybridShapes.Count).Name & "` ")
            oSeedLingPart.Update
            
            'Delete formula and temp surface
            oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oFormula)
            oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oDatumSet.HybridShapes.Item(oDatumSet.HybridShapes.Count))
        End If
        
        'Put surface it in show mode
        oSeedLingPart.HybridShapeFactory.GSMVisibility oSeedLingPart.CreateReferenceFromObject(oCurSeedLingDatum), 1 'Set KBE Body to Show mode if required
        
        'Check if datum is published. Publish it if it's not published and option "Publish elements" is selected
        Set oPublication = Nothing
        Set oPublication = oPublications.Item(sCurSeedLingDatum)
        If oPublication Is Nothing And frm3DCollector.cbPublish.Value Or frm3DCTemplate.cbPublish.Value Then
            Set oPublication = oPublications.Add(sCurSeedLingDatum)
        End If
        
        If Not oPublication Is Nothing Then
            Set oReference = oSeedLingInstance.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/!" & sCurSeedLingDatum)
            oPublications.SetDirect sCurSeedLingDatum, oReference
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    Call manageExtraDatum(oSkeletonSet)
    Call manageExtraDatum(oEnvironmentSet)
    Call manageExtraDatum(oKinematicsSet)
    Call manageExtraDatum(oStylingFeatureSet)
    Call manageExtraDatum(oBehaviorSet)
    oSeedLingPart.Update
    
    On Error GoTo 0
    
End Sub

Private Sub manageExtraDatum(ByRef oGeoset As HybridBody)
    
    'Add prefix EXTRA_ in front of SeedLing datum not found in corresponding Seed Part
    'Delete "EXTRA_" SeedLing datum that are found in corresponding Seed Part (another datum corresponds to it)
    iProgressCount2Max = oGeoset.HybridShapes.Count
    For iI = oGeoset.HybridShapes.Count To 1 Step -1
        
        Set oCurSeedLingDatum = oGeoset.HybridShapes.Item(iI)
        sCurSeedLingDatum = oCurSeedLingDatum.Name
        
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2Max + 1 - iI
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Checking for missing Datum Elements (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If Left(sCurSeedLingDatum, Len(sSeedPartLetter)) = sSeedPartLetter Then     'Only consider datums coming from current Seed Part
            iPos = InStr(1, sSeedPartDatum, Mid(sCurSeedLingDatum, Len(sSeedPartLetter) + 2), vbBinaryCompare)
            If iPos = 0 Then    'Datum originally coming from current Seed Part has been deleted in Seed Part => renamed with "EXTRA_" prefix in SeedLing
                oCurSeedLingDatum.Name = "EXTRA_" & oCurSeedLingDatum.Name
            End If
        ElseIf Left(sCurSeedLingDatum, Len("EXTRA_" & sSeedPartLetter)) = "EXTRA_" & sSeedPartLetter Then
            iPos = InStr(1, sCurSeedLingDatum, aSeedLingDatum.Item(Split(sCurSeedLingDatum, "EXTRA_")(1)).Name, vbBinaryCompare)
            If iPos = 7 Then    'An "EXTRA_" Datum exist with a Datum without "EXTRA_" > delete the Extra one
                oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oCurSeedLingDatum)
            End If
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
End Sub


Private Function GetKBEBodyFromSeedLingParamsScan(ByVal oSeedLingPart As Part, ByVal sBodyRootName As String, ByVal sParamRootName As String, ByVal sSeedLingCurBody As String, ByRef oCurSeedLingParameter As Parameter) As Body
      
'*** Retrieve KBE Body containing SeedLing current Body by analyzing parameters ***
    Set GetKBEBodyFromSeedLingParamsScan = Nothing
    For Each oCurSeedLingParameter In oKBEParamSet.DirectParameters
        If Right(oCurSeedLingParameter.ValueAsString, Len(sSeedLingCurBody)) = sSeedLingCurBody Then
            iPos = InStr(1, oCurSeedLingParameter.Name, sParamRootName, vbBinaryCompare)
            sCurKBEBodyNum = Right(oCurSeedLingParameter.Name, Len(oCurSeedLingParameter.Name) - iPos - Len(sParamRootName) + 1)
            Set GetKBEBodyFromSeedLingParamsScan = oSeedLingPart.Bodies.Item(sBodyRootName & sCurKBEBodyNum)
            Exit For
        End If
    Next

End Function

Private Sub AddActivityParam()
    
    'Add Activity Param
    Set oActivityParam = oSeedLingParameters.CreateBoolean("Activity_" & sSeedLingCurBody, True)
    
    'Copy paste to Parent Body
    oSelectionFromSeedLing.Clear
    oSelectionFromSeedLing.Add oActivityParam
    oSelectionFromSeedLing.Cut
    
    oSelectionFromSeedLing.Clear
    oSelectionFromSeedLing.Add oNewBody
    oSelectionFromSeedLing.Paste
    aSeedLingActivityParam.Add oSelectionFromSeedLing.Item2(1).Value, Mid(oSelectionFromSeedLing.Item2(1).Value.Name, InStrRev(oSelectionFromSeedLing.Item2(1).Value.Name, "\") + 1)
    oSelectionFromSeedLing.Clear
    
    If frm3DCollector.cbPublish.Value Then
        'Check if parameter is published. Publish it if it's not published and option "Publish elements" is selected
        On Error Resume Next
        Set oPublication = Nothing
        Set oPublication = oPublications.Item("Activity_" & sSeedLingCurBody)
        If oPublication Is Nothing And frm3DCollector.cbPublish.Value = True Then
            Set oPublication = oPublications.Add("Activity_" & sSeedLingCurBody)
        End If
        
        If Not oPublication Is Nothing Then
            Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oNewBody.Name & "/" & "Activity_" & sSeedLingCurBody)
            oPublications.SetDirect "Activity_" & sSeedLingCurBody, oReference
        End If
        On Error GoTo 0
    End If
    
End Sub

Private Sub ScanSeedPartParameters()

    On Error Resume Next
    
    Set aSeedPartAllATBs = New Collection
    Set aSeedPartDesiredATBs = New Collection
    Set aActiveBodiesATBs = New Collection
    Set aSeedPartNotCollectedATBs = New Collection
    Set aSeedPartSTRCs = New Collection
    Set aSeedPartBooleans = New Collection
    
    sSearchString = _
    "CATKnowledgeSearch.InternalParameter.Name_CAP=KBE_*_Boolean + " & _
    "CATKnowledgeSearch.InternalParameter.Name_CAP=ATB_* + " & _
    "CATKnowledgeSearch.InternalParameter.Name_CAP=STRC_*,sel"
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    '***Find all boolean parameters in seed part and create a collection aSeedPartBooleans
    Set oSelectionFromSeedPart = KBESearch(sSearchString, oSeedPart, oSelectionFromSeedPart)
    
    For iI = 1 To oSelectionFromSeedPart.Count
        
        sParameterName = Mid(oSelectionFromSeedPart.Item(iI).Value.Name, InStrRev(oSelectionFromSeedPart.Item(iI).Value.Name, "\") + 1)
        
        If sParameterName Like "KBE_*_Boolean" Then
            aSeedPartBooleans.Add oSelectionFromSeedPart.Item(iI).Value, sParameterName
        ElseIf sParameterName Like "ATB_*" Then
            aSeedPartAllATBs.Add oSelectionFromSeedPart.Item(iI).Value, sParameterName
        ElseIf sParameterName Like "STRC_*" Then
            aSeedPartSTRCs.Add oSelectionFromSeedPart.Item(iI).Value, sParameterName
        End If
        
    Next
    oSelectionFromSeedPart.Clear
    
'***Find ATB parameters that need to be collected and create an array aSeedPartDesiredATBs
    If frm3DCollector.cbATBParams.Value = True Then '***Collect only ATB parameters that are related to Material Attributes***
        For Each oParameter In aSeedPartAllATBs
            If oParameter.Name Like "*MType*" Or _
               oParameter.Name Like "*MScale*" Or _
               oParameter.Name Like "*MPosition*" Or _
               oParameter.Name Like "*MOrientation*" Then
                
                aSeedPartDesiredATBs.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
            End If
        Next
    Else '***Collect all ATB parameters, except those related to Material Attributes***
        For Each oParameter In aSeedPartAllATBs
            If Not (oParameter.Name Like "*MType*" Or _
                    oParameter.Name Like "*MScale*" Or _
                    oParameter.Name Like "*MPosition*" Or _
                    oParameter.Name Like "*MOrientation*") Then
                
                aSeedPartDesiredATBs.Add oParameter, Mid(oParameter.Name, InStrRev(oParameter.Name, "\") + 1)
            End If
        Next
    End If
    
'***Reset Not Collected ATBs to all ATBs. Colected ATBs will be removed later from this collection when they will be collected (See sub CollectATBParams)
    Set aSeedPartNotCollectedATBs = aSeedPartAllATBs
    
    On Error GoTo 0

End Sub

Private Sub CollectOS()
        
    '*******Search all surfaces with name containing "_OS_", but does not end with "_OS_"*************************
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    Set oSelectionFromSeedPart = KBESearch("(.Surface.Name_CAP=*_OS_* & .Surface.Name_CAP!=*_OS_),sel", oSeedPart, oSelectionFromSeedPart)
    
    Set aSeedPartOS = New Collection
    For iI = 1 To oSelectionFromSeedPart.Count2
        aSeedPartOS.Add oSelectionFromSeedPart.Item(iI).Value
    Next
    
    If aSeedPartOS.Count > 0 Then
        
        iProgressCount2Max = aSeedPartOS.Count
        For iI = 1 To aSeedPartOS.Count
            
            'Progress Bar Settings
            iProgressCount2 = iI
            timeEnd = DateTime.Timer
            sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
            sProgBarComment = "Collecting & Replacing OS (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
            
            sCurSeedLingOS = sSeedPartLetter & "_" & aSeedPartOS.Item(iI).Name
            
            On Error Resume Next
            Set TempOS = Nothing
            Set TempOS = oOperationsSet.HybridShapes.Item(sCurSeedLingOS)
            On Error GoTo 0
            
            If TempOS Is Nothing Then 'Paste new Surface from Seed Part. & Publish the OS surface
                
                'Copy
                oSelectionFromSeedPart.Clear
                oSelectionFromSeedPart.Add aSeedPartOS.Item(iI)
                oSelectionFromSeedPart.Copy
                
                'Paste
                oSelectionFromSeedLing.Clear
                oSelectionFromSeedLing.Add oOperationsSet
                oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
                oSelectionFromSeedLing.Clear
                
                Set oCurSeedLingOS = oOperationsSet.HybridShapes.Item(oOperationsSet.HybridShapes.Count)
                oCurSeedLingOS.Name = sCurSeedLingOS
                aSeedLingOS.Add oCurSeedLingOS, sCurSeedLingOS
            
            Else    'Paste new Surface from Seed Part with Temp suffix and create rule between old and new surface, and delete the rule and new surface.
                
                On Error Resume Next
                Set oCurSeedLingDatum = Nothing
                Set oCurSeedLingDatum = oOperationsSet.HybridShapes.Item("Temp_" & sCurSeedLingOS)
                On Error GoTo 0

                If oCurSeedLingDatum Is Nothing Then
                    
                    'Copy
                    oSelectionFromSeedPart.Clear
                    oSelectionFromSeedPart.Add aSeedPartOS.Item(iI)
                    oSelectionFromSeedPart.Copy
                    
                    'Paste
                    oSeedLingPart.InWorkObject = oOperationsSet
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oOperationsSet
                    oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
                    
                    'Rename OS with "Temp_"
                    Set oCurSeedLingOSTemp = oOperationsSet.HybridShapes.Item(oOperationsSet.HybridShapes.Count)
                    oCurSeedLingOSTemp.Name = "Temp_" & sCurSeedLingOS
                    
                    'Create formula
                    Set oFormula = oSeedLingPart.Relations.CreateFormula("oFormula_" & sCurSeedLingOS, "", oSeedLingPart.Parameters.Item(sCurSeedLingOS), "`Operations\" & oCurSeedLingOSTemp.Name & "` ")
                    oSeedLingPart.Update
                    
                    'Delete formula and temp surface
                    If Not oCurSeedLingOSTemp Is Nothing Then
                        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oFormula)
                        oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oCurSeedLingOSTemp)
                    End If
                End If
            End If
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
        Next iI
    End If
    
    On Error Resume Next
    Err.Clear
    oSeedLingPart.Update
    If Err.Number <> 0 Then
        MsgBox oSeedLingPart.Name + " cannot be updated !   " + vbNewLine + "Use ""Reset Collector"" option, or update it manually.   ", vbExclamation, "Warning"
        bCancelAction = True
    End If
    On Error GoTo 0
    
End Sub

Private Sub Split_OS()
    
    Dim VolumeBeforeSplit As Double: Dim VolumeAfterSplit As Double
    
    Set wSPAWorkbench = oSeedLingRootDocument.GetWorkbench("SPAWorkbench")
    
    iProgressCount2Max = aSeedLingOS.Count
    iProgressCount2 = 0
    For Each iObject In aSeedLingOS
        
        sCurSeedLingOS = iObject.Name
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Managing OS (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        'If apply OS from all seed parts is selected, OR
        'If apply OS from selected seed parts is selected, and sCurSeedLingOS comes from one of the selected seed parts

        If frm3DCollector.ImgAllSPTrue.Visible = True Or InStr(sSeedPartsList, "SEED_PART_" & Split(sCurSeedLingOS, "_")(0)) > 0 Then
            
            'Read OS impact list, and fill collection "aKBEBodyFromOS" with every impacted body
            sKBEBodiesFromOS = Mid(sCurSeedLingOS, InStr(1, sCurSeedLingOS, "_OS_", vbBinaryCompare) + Len("_OS_"))
            iPos = InStr(1, sKBEBodiesFromOS, "&", vbBinaryCompare)
            
            Set aKBEBodyFromOS = New Collection
            If iPos = 0 Then
                aKBEBodyFromOS.Add sKBEBodiesFromOS
            Else
                Do While iPos > 0
                    'Retrieve the list of panels to be split
                    aKBEBodyFromOS.Add Left(sKBEBodiesFromOS, iPos - 1)
                    sKBEBodiesFromOS = Right(sKBEBodiesFromOS, Len(sKBEBodiesFromOS) - iPos)
                    iPos = InStr(1, sKBEBodiesFromOS, "&", vbBinaryCompare)
                Loop
                aKBEBodyFromOS.Add sKBEBodiesFromOS  'Last value is a valid panel or empty string
            End If
            
            
            'Scan SeedLing to built/update Split operations
            For iJ = 1 To aKBEBodyFromOS.Count
                
                sKBEBodiesFromOS = aKBEBodyFromOS.Item(iJ)
                
                'Retrieve KBE Body containing current SeedLing Body in order to build split operation
                Set oSeedLingCurKBEBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, sKBEBodiesFromOS, oCurSeedLingParameter)
                If oSeedLingCurKBEBody Is Nothing Then
                    Set oSeedLingCurKBEBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, sKBEBodiesFromOS & " " & sCollectorID, oCurSeedLingParameter)
                End If

                'Create Split Operation if an active oSeedLingCurKBEBody body has been found
                
                If Not oSeedLingCurKBEBody Is Nothing Then

                'Check the splitsurface, if it closer enough to do the Split
                    For Each oShape In oSeedLingCurKBEBody.Shapes
                        If oShape.Name = sCurSeedLingOS Then
                            Set oSeedLingCurKBEBody = Nothing 'If split exist, do not create it again
                            Exit For
                        End If
                    Next
                    
                    If Not oSeedLingCurKBEBody Is Nothing Then      'If Split does not exist
                        
                        On Error Resume Next
                        VolumeBeforeSplit = 0: VolumeAfterSplit = 0
                        Set oReference = oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody)
                        VolumeBeforeSplit = wSPAWorkbench.GetMeasurable(oReference).Volume
                        
                        oSeedLingPart.InWorkObject = oSeedLingCurKBEBody.Shapes.Item(1)
                        Set oReference = oSeedLingPart.CreateReferenceFromName("")
                        Set oSplit = oSeedLingPart.ShapeFactory.AddNewSplit(oReference, catPositiveSide)
                        oSplit.Name = sCurSeedLingOS
                        
                        'Assign proper Surface to Split object
                        Set oReference = oSeedLingPart.CreateReferenceFromObject(oSeedLingParameters.Item(sCurSeedLingOS))
                        oSplit.Surface = oReference
                        
                        'Check for true split
                        oSeedLingPart.Update
                        Set oReference = oSeedLingPart.CreateReferenceFromObject(oSeedLingCurKBEBody)
                        VolumeAfterSplit = wSPAWorkbench.GetMeasurable(oReference).Volume
                        
                        If (VolumeBeforeSplit > VolumeAfterSplit) Then
                            iAssemble = 0
                            For Each oShape In oSeedLingCurKBEBody.Shapes
                                If TypeName(oShape) = "Assemble" And oShape.Name = "MainAssemble" Then
                                    iAssemble = iAssemble + 1
                                End If
                            Next
                            
                            If iAssemble > 1 Then
                                oSelectionFromSeedLing.Clear
                                oSelectionFromSeedLing.Add oSplit
                                oSelectionFromSeedLing.Cut
                                oSelectionFromSeedLing.Add oSeedLingCurKBEBody.Shapes.Item("MainAssemble")
                                oSelectionFromSeedLing.Paste
                                oSelectionFromSeedLing.Item(1).Value.Name = sCurSeedLingOS
                            End If
                        Else
                            oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oSplit)
                        End If
                        On Error GoTo 0
                    
                    End If
                End If
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
End Sub

Private Sub Collect_OA()
    
    Dim OAVolume As Double
    
    Set wSPAWorkbench = oSeedLingRootDocument.GetWorkbench("SPAWorkbench")
    
    ' Collects OA bodies in Seedpart
    On Error Resume Next
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    Set oSelectionFromSeedPart = KBESearch("(Name_CAP=*_OA & 'Part Design'.Body),sel", oSeedPart, oSelectionFromSeedPart)
    
    iProgressCount2Max = oSelectionFromSeedPart.Count2
    For iI = 1 To oSelectionFromSeedPart.Count2
            
        'Progress Bar Settings
        iProgressCount2 = iI
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting OA (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
         
        OAVolume = 0: OAVolume = wSPAWorkbench.GetMeasurable(oSeedPart.CreateReferenceFromObject(oSelectionFromSeedPart.Item(iI).Value)).Volume
  
        If OAVolume > 0 Then
            aSeedPartOA.Add New OperationAdd, oSelectionFromSeedPart.Item(iI).Value.Name
            aSeedPartOA.Item(oSelectionFromSeedPart.Item(iI).Value.Name).BodyOA = oSelectionFromSeedPart.Item(iI).Value
            aSeedPartOA.Item(oSelectionFromSeedPart.Item(iI).Value.Name).SeedPartLetter = sSeedPartLetter
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next

    Set oSelectionFromSeedPart = KBESearch("(Name_CAP=*_OA & Knowledgeware.Parameter),sel", oSeedPart, oSelectionFromSeedPart)
    For iI = 1 To oSelectionFromSeedPart.Count2
        sParameterName = Mid(oSelectionFromSeedPart.Item(iI).Value.Name, InStrRev(oSelectionFromSeedPart.Item(iI).Value.Name, "\") + Len("\"))
        aSeedPartOA.Item(sParameterName).ParameterOA = oSelectionFromSeedPart.Item(iI).Value
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    oSelectionFromSeedPart.Clear
    On Error GoTo 0
    
End Sub

Private Sub Add_OA()
    
    iProgressCount2Max = aSeedPartOA.Count
    iProgressCount2 = 0
    For Each iObject In aSeedPartOA
        
        'Progress Bar Settings
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Managing OA (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)

        'For all OA found, find the impacted body in seedling part
        On Error Resume Next
        sCurSeedLingOAParamValue = ""
        sCurSeedLingOAParamValue = iObject.ParameterOA.Value
        
        Set iObject.ImpactedBody = Nothing
        If sCurSeedLingOAParamValue <> "" Then
            iObject.ImpactedBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, CStr(sCurSeedLingOAParamValue), oCurSeedLingParameter)
            If iObject.ImpactedBody Is Nothing Then
                iObject.ImpactedBody = GetKBEBodyFromSeedLingParamsScan(oSeedLingPart, sBodyRootName, sParameterKBERootName, sCurSeedLingOAParamValue & " " & sCollectorID, oCurSeedLingParameter)
            End If
        End If
        
        'If OA impacts an existing body in SeedLing
        'Create/Update Assemble Operation
        If Not iObject.ImpactedBody Is Nothing Then
            oSelectionFromSeedPart.Clear
            oSelectionFromSeedPart.Add iObject.BodyOA
            oSelectionFromSeedPart.Copy
            oSelectionFromSeedPart.Clear
            
            oSelectionFromSeedLing.Clear
            oSelectionFromSeedLing.Add oSeedLingPart
            oSelectionFromSeedLing.PasteSpecial "CATPrtResultWithOutLink"
            Set oNewBody = oSelectionFromSeedLing.Item2(1).Value
            oNewBody.Name = iObject.SeedPartLetter & "_" & iObject.BodyOA.Name
            oSelectionFromSeedLing.Clear
            
            'If the OA already exist in seedling, delete and recreate OA content without deleting body
            On Error Resume Next
            For Each oCurSeedLingShape In iObject.ImpactedBody.Shapes
                If oCurSeedLingShape.Name Like "*" & oNewBody.Name Then
                    
                    'Copy new OA Body content, then delete it
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oNewBody.Shapes.Item(1)
                    oSelectionFromSeedLing.Copy
                    
                    'Delete all shapes in found body
                    oSelectionFromSeedLing.Clear
                    For Each oShape In oCurSeedLingShape.Body.Shapes
                        oSelectionFromSeedLing.Add oShape
                        oSelectionFromSeedLing.Add oShape.Body
                    Next
                    oSelectionFromSeedLing.Delete
                    
                    'Paste new body content in existing OA body
                    oSelectionFromSeedLing.Add oCurSeedLingShape.Body
                    oSelectionFromSeedLing.Paste
                    oSelectionFromSeedLing.Clear
                    
                    oSeedLingPart.HybridShapeFactory.DeleteObjectForDatum oSeedLingPart.CreateReferenceFromObject(oNewBody)
                    Set oNewBody = Nothing
                    Exit For
                End If
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next
            
            'If the OA does not exist in seedling, create a new assemble operation
            
            If Not oNewBody Is Nothing Then
                oSeedLingPart.InWorkObject = iObject.ImpactedBody
                Set oAssemble = oSeedLingPart.ShapeFactory.AddNewAssemble(oNewBody)
                oAssemble.Name = oNewBody.Name
                oSeedLingPart.Update
            End If
            On Error GoTo 0
            
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    On Error Resume Next
    Err.Clear
    oSeedLingPart.Update
    If Err.Number <> 0 Then
        MsgBox oSeedLingPart.Name + " cannot be updated !   " + vbNewLine + "Use ""Reset Collector"" option, or update it manually.   ", vbExclamation, "Warning"
        bCancelAction = True
    End If
    On Error GoTo 0
    
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
End Sub

Private Sub CheckKBEBodies_Vs_FlatTreeBodies()
'Put a warning message if number of active bodies is different from number of Flat Tree parts
    
    Dim NoOfFlatTreeBOdy As Long
    
    'Remove "Activity=False" parameters from collection aSeedLingActivityParam
    For i = aSeedLingActivityParam.Count To 1 Step -1
        If aSeedLingActivityParam.Item(i).Value = False Then
            aSeedLingActivityParam.Remove (i)
        End If
    Next i
    
    'Find all Flat Tree parts under NHA
    NoOfFlatTreeBOdy = 0
    For i = 1 To oNHAAssy.Products.Count
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, 1, 1, "Checking KBE Bodies against Flat Tree Bodies...", 1, 1, sTimer)

        On Error Resume Next
        If TypeName(oNHAAssy.Products.Item(i).ReferenceProduct.Parent) = "PartDocument" Then
            If (oNHAAssy.Products.Item(i).Definition Like "*_KBE_M*" Or oNHAAssy.Products.Item(i).Definition Like "DO NOT PRODUCE*") And Not oNHAAssy.Name Like "*-S*" Then
                For Each oParameter In oKBEParamSet.DirectParameters
                    If Replace(oParameter.Value, "RESERVED_", "") Like Replace(oNHAAssy.Products.Item(i).Definition, "RESERVED_", "") Then
                        oNHAAssy.Products.Item(i).Definition = oParameter.Value
                        Exit For
                    End If
                Next
                NoOfFlatTreeBOdy = NoOfFlatTreeBOdy + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    'Put a warning message if number of active bodies is different from number of Flat Tree parts
    sLog = ""
    If aSeedLingActivityParam.Count < NoOfFlatTreeBOdy And NoOfFlatTreeBOdy > 0 Then
        sLog = "Number of KBE Bodies: " & Str(aSeedLingActivityParam.Count) & vbCrLf & "Number of Flat Tree parts: " & Str(NoOfFlatTreeBOdy) & vbCrLf & vbCrLf & "Please run the Flat Tree - Clean script!   "
    ElseIf aSeedLingActivityParam.Count > NoOfFlatTreeBOdy And NoOfFlatTreeBOdy > 0 Then
        sLog = "Number of KBE Bodies: " & Str(aSeedLingActivityParam.Count) & vbCrLf & "Number of Flat Tree parts: " & Str(NoOfFlatTreeBOdy) & vbCrLf & vbCrLf & "Please run the Flat Tree - Create/Synchronize script!   "
    End If
    
End Sub


