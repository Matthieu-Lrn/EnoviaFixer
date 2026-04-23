Attribute VB_Name = "ActivateExternalObject"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************************************
'* The goal of this class is to activate the following objects:
'*  - All the external references
'*  - All the external parameters
'*  - All axis located under the "Axis Systems"
'*  - for ICD part: all the bodies
'*
'*
'* Author:      Fraocois Charette
'* Updated by:  Abhishek Kamboj
'* Language:    VBA
'********************************************************************************

Private oSeedPartCollection As New Collection
Private oGeosetCollection As Collection
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
    
    Call Activate
    
End Sub

Private Sub Activate()
'This Sub will activate all the External References in all the selected Seed Part
    
    Dim oProduct As Product
    Dim oPartDocument As PartDocument
    Dim oPart As Part
    Dim oSelection As Selection
    Dim oHybridShape As AnyObject
    Dim oGeoset As HybridBody

    Dim sLog As String
    Dim oParamSet As ParameterSet
    Dim oParam As Parameter
    Dim oAxisSystems As AxisSystems
    Dim oAxis As AxisSystem
    Dim oBody As Body
    Dim oShape As Shape
    
    Set oSelection = CATIA.ActiveDocument.Selection
    
    Call frmProgress.progressBarInitialize("Activate External Objects")
    Call frmProgress.progressBarRepaint("Activate: External References / Parameters / Axis .....")
    
    sLog = ""
    For Each oProduct In oSeedPartCollection
        
        'Retrieve Part Object
        Set oPartDocument = oProduct.GetMasterShapeRepresentation(False)
        Set oPart = oPartDocument.Part
    
        'Find "External References" geosets
        On Error Resume Next
        Set oGeoset = Nothing
        Set oGeosetCollection = New Collection
        For Each oGeoset In oPart.HybridBodies
         Call SearchHybridBodies(oGeoset, "*External*Reference*")
        Next
        On Error GoTo 0
    
        'If the Geoset exist we activate all the components
        If oGeosetCollection.Count > 0 Then
            For Each oGeoset In oGeosetCollection
                For Each oHybridShape In oGeoset.HybridShapes
                    If oPart.IsInactive(oHybridShape) Then
                        sLog = sLog & vbCrLf & "- In " & oProduct.PartNumber & ", " & oHybridShape.Name & " was reactivated"
                        oPart.Activate oHybridShape
                    End If
                Next
            Next
        End If
        
        'Find External Parameters set
        On Error Resume Next
        Set oParamSet = Nothing
        Set oParamSet = oPart.FindObjectByName("External Parameters")
        On Error GoTo 0
    
        
        'If the ParamSet exist we activate all the components
        If Not oParamSet Is Nothing Then
            For Each oParam In oParamSet.AllParameters
                If oPart.IsInactive(oParam) Then
                    sLog = sLog & vbCrLf & "- In " & oProduct.PartNumber & ", " & Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\"))) & " was reactivated"
                    oPart.Activate oParam
                End If
            Next
        End If
        
        
        'Find Main Axis System Set
        On Error Resume Next
        Set oAxisSystems = Nothing
        Set oAxisSystems = oPart.AxisSystems
        On Error GoTo 0
    
        
        'If the Main Axis System Set exist we activate all the components
        If Not oAxisSystems Is Nothing Then
            For Each oAxis In oAxisSystems
                If oPart.IsInactive(oAxis) Then
                    sLog = sLog & vbCrLf & "- In " & oProduct.PartNumber & ", " & oAxis.Name & " was reactivated"
                    oPart.Activate oAxis
                End If
            Next
        End If
    
        
        'For the ICD part we check that the solid are activated
        If oProduct.Name Like "*ICD*" Then
            For Each oBody In oPart.Bodies
                For Each oShape In oBody.Shapes
                    If oPart.IsInactive(oShape) Then
                        sLog = sLog & vbCrLf & "- In " & oProduct.PartNumber & ", " & oBody.Name & "/" & oShape.Name & " was reactivated"
                        oPart.Activate oShape
                    End If
                Next
            Next
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "The following actions were performed:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_ActivateExternalReferences.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    Else
        MsgBox ("No External References / External Parameters / Main Axis / Body in ICD part needed to be activated")
    End If
    
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar
    
End Sub

Private Sub SearchHybridBodies(oGeoset As HybridBody, sSearchstring As String)
    '***This sub searches for all geometrical sets inside oGeoset having sSearchstring in their name***
    '*** Geometrical sets are searched recursively to search through all geosets inside selected one***
    Dim oChildGeoset As HybridBody
    
    If oGeoset.Name Like sSearchstring Then
        On Error Resume Next    '*** prevent duplicate entry
        oGeosetCollection.Add oGeoset, oGeoset.Name
        On Error GoTo 0
    End If
    For Each oChildGeoset In oGeoset.HybridBodies
        If oChildGeoset.Name Like sSearchstring Then
            On Error Resume Next  '*** prevent duplicate entry
            oGeosetCollection.Add oChildGeoset, oChildGeoset.Name
            On Error GoTo 0
        End If
    Next
    
    For Each oChildGeoset In oGeoset.HybridBodies
        Call SearchHybridBodies(oChildGeoset, sSearchstring)
    Next
End Sub
