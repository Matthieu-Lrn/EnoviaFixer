Attribute VB_Name = "InstantiateSTDParts"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private oMaintenanceParam As BoolParam
Private oStdPartInstance As Product
Private oBoolParam As Parameter
Private sBoolParamLocalName As String
Private sPartNumberWithUnderscore As String
Private sSeedPartAxisName As String
Private oNewStdPartInstance As Product
Private oSeedPartAxisReference As Reference
Private oStdPartAxisReference As Reference
Private sPreviousRelationText As String

Private oStdPartCollection As Collection
Private oCoincidenceConstraintCollection As Collection
Private oBoolParamCollection As Collection
Private oAxisCollection As Collection

'********************************************************************************
'* Name: Instantiate Standard parts
'* Purpose: Create a new instance for selected standard parts, create a constraint
'*          to put it in position, and if needed create an activity rule at product level
'*
'* Assumption: See below
'*
'* Author:
'* Updated by: François Charette, Julien Bigaouette
'* Language: VBA
'********************************************************************************
'*
'* A constraint name has the following nomenclature:
'* -  Std Part instance name | Axis Name | Seed Part Number
'* Example: "OVALV-1PLC1.23 | SP_M77_Door_BA02_DatumAxisLatch | KBE_M77_SEED_PART_BA-V06"
'*
'******************************************************************************************
'*
'* A Std Part boolean parameter has the following nomenclature:
'*
'* - SP_zzzzzzzzzzz_yyyyyyyyy_Boolean
'* or
'* - SA_zzzzzzzzzzz_yyyyyyyyy_Boolean
'*
'* The corresponding axis system must be named:
'* - SP_zzzzzzzzzzz or SA_zzzzzzzzzzz
'*
'* The corresponding Part Number will be:
'* - "yyyyyyyy"
'*
'******************************************************************************************

Public Sub launchInstantiation()

    Set oConstraints = oSeedAssy.Connections("CATIAConstraints")
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    sLog = ""
    
    'Time set for Progress Bar
    timeStart = DateTime.Timer
    iProgressCount1Max = oSeedPartCollection.Count
    
    'Initialize progress bar
    Call frmProgress.progressBarInitialize("Standard Parts Instantiation")
    
    'Delete existing Constraints, Std Part instances and Activity Rule if required
    If frmStdPartsInstantiation.chkDeleteStdPart.Value = True Then
        Call DeleteConstraintStdPartRule
    Else
        Call frmProgress.progressBarRepaint("Initializing...")
    End If
    
    'Populate oCoincidenceConstraintCollection
    Set oCoincidenceConstraintCollection = New Collection
    If Not oConstraints.Count = 0 Then
        For Each oConstraint In oConstraints
            If oConstraint.Type = 2 Then
                oCoincidenceConstraintCollection.Add oConstraint
            End If
        Next
    End If
    
    'Populate oStdPartCollection
    'Retrieve SEED_OPTION / SEED_CONFIG part and "Maintenance" parameter
    On Error Resume Next
    Set oMaintenanceParam = Nothing
    Set oStdPartCollection = New Collection
    For Each oProduct In oSeedAssy.Products
        If oProduct.Name Like "*_SEED_OPTION*" Or oProduct.Name Like "*_SEED_CONFIG*" Then
            oProduct.ApplyWorkMode DESIGN_MODE
            Set oMaintenanceParam = oProduct.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.DirectParameters.GetItem("Maintenance")
        'A Std Part is any part which doesn't have "*_SEED_PART_*" and "KBE*"
        ElseIf Not oProduct.Name Like "KBE*" Then
            oStdPartCollection.Add oProduct, Replace(Split(oProduct.Name, ".")(0), "-", "_", 1, -1, vbBinaryCompare)
        End If
    Next
    
    'We loop the selected Seed Part
    iProgressCount1 = 0
    For Each oSeedPartInstance In oSeedPartCollection
                   
        iProgressCount1 = iProgressCount1 + 1
        
        'Activity Rule Text
        sRelationText = ""
        
        'Populate the oPopulateBoolParamCollection
        Set oBoolParamCollection = New Collection
        Set oSelection = KBESearch("CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean,sel", oSeedPartInstance)
        If Not oSelection.Count = 0 Then
            For i = 1 To oSelection.Count
                If frmStdPartsInstantiation.optBoolParamTrue.Value = True And oSelection.Item(i).Value.ValueAsString = "true" Then
                    oBoolParamCollection.Add oSelection.Item(i).Value, oSelection.Item(i).Value.Name
                ElseIf frmStdPartsInstantiation.optBoolParamTrue.Value = False Then
                    oBoolParamCollection.Add oSelection.Item(i).Value, oSelection.Item(i).Value.Name
                End If
            Next
        End If
        
        'Populate the oAxisCollection
        Set oAxisCollection = New Collection
        Set oSelection = KBESearch("(Name_CAP=SP_* - Name_CAP=*_Boolean) + (Name_CAP=SA_* - Name_CAP=*_Boolean),sel", oSeedPartInstance)
        If Not oSelection.Count = 0 Then
            For i = 1 To oSelection.Count
                oAxisCollection.Add oSelection.Item(i).Value, oSelection.Item(i).Value.Name
            Next
        End If
        
        'Loop the oBoolParamCollection
        iProgressCount2 = 0
        iProgressCount2Max = oBoolParamCollection.Count
        For Each oBoolParam In oBoolParamCollection
            
            iProgressCount2 = iProgressCount2 + 1
            
            'Progress Bar
            timeEnd = DateTime.Timer
            Call frmProgress.progressBarRepaint("Instantiating from " & oSeedPartInstance.PartNumber & " (" & CStr(iProgressCount1) & "/" & CStr(oSeedPartCollection.Count) & ")", _
            iProgressCount1Max, iProgressCount1, _
            "Creating/Checking constraints (" & CStr(iProgressCount2) & "/" & CStr(oBoolParamCollection.Count) & ")", _
            iProgressCount2Max, iProgressCount2, _
            "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss"))
                
            'Find BoolParam local name
            sBoolParamLocalName = Split(oBoolParam.Name, "\")(UBound(Split(oBoolParam.Name, "\")))
                   
            'Retrieve the matching axis system name
            sSeedPartAxisName = ""
            For Each iObject In oAxisCollection
                If sBoolParamLocalName Like iObject.Name & "_*" Then
                    sSeedPartAxisName = iObject.Name
                    Exit For
                End If
            Next
            
            If sSeedPartAxisName = "" Then
                sLog = sLog & vbCrLf & " - The axis matching parameter " & sBoolParamLocalName & " could not be found"
                GoTo NextBoolParam
            End If
            
            'Find PartNumber
            sPartNumberWithUnderscore = ""
            sPartNumberWithUnderscore = Replace(sBoolParamLocalName, sSeedPartAxisName & "_", "")
            sPartNumberWithUnderscore = Replace(sPartNumberWithUnderscore, "_Boolean", "")
            sPartNumberWithUnderscore = Replace(Split(sPartNumberWithUnderscore, ".")(0), "-", "_", 1, -1, vbBinaryCompare)
            
            'Exit if Std part name is not found in the boolean
            If sPartNumberWithUnderscore = "" Then
                sLog = sLog & vbCrLf & " - The Std Part Number could not be Retrieved from " & sBoolParamLocalName & " and " & sSeedPartAxisName
                GoTo NextBoolParam
            End If
            
            'Find matching Std Part from oStdPartCollection
            Set oStdPartInstance = Nothing
            Set oStdPartInstance = oStdPartCollection.Item(sPartNumberWithUnderscore)
            
            'Exit if Std part was not found. If Std part was found, put it in design mode
            If oStdPartInstance Is Nothing Then
                sLog = sLog & vbCrLf & " - " & sPartNumberWithUnderscore & " was not found in the Seed Assy. Param path is:    " & oBoolParam.Name
                GoTo NextBoolParam
            Else
                oStdPartInstance.ApplyWorkMode DESIGN_MODE
            End If
            
            'Retrieve the existing constraint if it exist (scan only if user did not check "Delete Previous Instances")
            Set oConstraint = Nothing
            If Not frmStdPartsInstantiation.chkDeleteStdPart.Value = True Then
                Set oConstraint = RetrieveConstraint(sPartNumberWithUnderscore, sSeedPartAxisName, oSeedPartInstance.PartNumber)
            End If
            
            'If std part exists and constraint doesn't already exist
            If Not oStdPartInstance Is Nothing And oConstraint Is Nothing Then
                
                'Create a new instance of the std part
                Call oSeedAssy.Products.AddExternalComponent(oStdPartInstance.ReferenceProduct.Parent)
                Set oNewStdPartInstance = oSeedAssy.Products.Item(oSeedAssy.Products.Count)
                oNewStdPartInstance.DescriptionInst = sBoolParamLocalName
                
                'Create a reference of the axis system located in the Seed Part
                Set oSeedPartAxisReference = Nothing
                Set oSeedPartAxisReference = oSeedAssy.CreateReferenceFromName(oSeedAssy.Name & "/" & oSeedPartInstance.Name & "/!" & sSeedPartAxisName)
                
                If oSeedPartAxisReference Is Nothing Then
                    sLog = sLog & vbCrLf & " - Problem creating a reference on " & sSeedPartAxisName & " in " & oSeedPartInstance.PartNumber
                End If
                
                'Create a reference of the axis system located in the Std Part
                'Std Part
                If sBoolParamLocalName Like "SP*" Then
                    
                    Set oStdPartAxisReference = Nothing
                    Set oStdPartAxisReference = oSeedAssy.CreateReferenceFromName(oSeedAssy.Name & "/" & oNewStdPartInstance.Name & "/!Absolute Axis System")
    
                    If oStdPartAxisReference Is Nothing Then
                        sLog = sLog & vbCrLf & " - Absolute Axis System can't be found in " & oNewStdPartInstance.PartNumber
                    End If
                    
                'Std Assy
                Else
                    sOriginPartInstanceName = ""
                    Set oStdPartAxisReference = Nothing
                    
                    For Each oProduct In oNewStdPartInstance.Products
                        If oProduct.Name Like "*ORIGINPART*" Then
                            sOriginPartInstanceName = oProduct.Name
                            Set oStdPartAxisReference = oSeedAssy.CreateReferenceFromName(oSeedAssy.Name & "/" & oNewStdPartInstance.Name & "/" & sOriginPartInstanceName & "/!Absolute Axis System")
                            Exit For
                        End If
                    Next
                    
                    If sOriginPartInstanceName = "" Or oStdPartAxisReference Is Nothing Then
                        sLog = sLog & vbCrLf & " - OriginPart can't be found under " & oNewStdPartInstance.PartNumber & ""
                    End If
                    
                End If
                
                'If we could Retrieve both axis system we create the constraint
                If Not oStdPartAxisReference Is Nothing And Not oSeedPartAxisReference Is Nothing Then
                
                    'Create the constraint
                    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeOn, oStdPartAxisReference, oSeedPartAxisReference)

                    'Activity Rule Text
                    sRelationText = sRelationText & vbCrLf & "`" & oNewStdPartInstance.Name & "\Component Activation State` = `" & oBoolParam.Name & "` == true and `" & oSeedPartInstance.Name & "\Component Activation State` == true"
                    If Not oMaintenanceParam Is Nothing Then sRelationText = sRelationText & " or `" & oMaintenanceParam.Name & "` == true"
                    
                    'Put the constraint in No Show
                    oSelection.Clear
                    oSelection.Add oConstraint
                    oSelection.VisProperties.SetShow (catVisPropertyNoShowAttr)
    
                    'Check if we have a problem with the new constraint
                    If oConstraint.Status = catCstStatusKOBroken Then
                        oSelection.Delete
                        sLog = sLog & vbCrLf & " - Problem creating a constraint between " & sSeedPartAxisName & " and " & oNewStdPartInstance.Name
                    Else
                        oConstraint.Name = oNewStdPartInstance.Name & " | " & sSeedPartAxisName & " | " & oSeedPartInstance.PartNumber
                    End If
                
                'If we don't create a constraint we delete the new std part instance
                Else
                    oSeedAssy.Products.Remove oNewStdPartInstance
                End If
                
            'If we have a constraint
            ElseIf Not oConstraint Is Nothing Then
                'Retrieve the instance name from the constraint
                sRelationText = sRelationText & vbCrLf & "`" & Split(oConstraint.Name, " | ")(0) & "\Component Activation State` = `" & oBoolParam.Name & "` == true and `" & oSeedPartInstance.Name & "\Component Activation State` == true"
                If Not oMaintenanceParam Is Nothing Then sRelationText = sRelationText & " or `" & oMaintenanceParam.Name & "` == true"
            End If
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
            
NextBoolParam:
        Next
        
        'Create/Modify Activity Rule if required
        If frmStdPartsInstantiation.chkActivityRule.Value = True And Not sRelationText = "" Then
            
            timeEnd = DateTime.Timer
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, _
            iProgressCount1Max, iProgressCount1, _
            "Creating/Modifying Activity Rule", _
            iProgressCount2Max, iProgressCount2, _
            "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss"))
            
            Set oRule = Nothing
            For Each oRelation In oSeedAssy.Relations
                If TypeName(oRelation) = "Rule" And oRelation.Name Like oSeedPartInstance.PartNumber & "*" Then
                    Set oRule = oRelation
                    Exit For
                End If
            Next
            
            'Create a new rule
            If oRule Is Nothing Then
                Set oRule = oSeedAssy.Relations.CreateProgram(oSeedPartInstance.PartNumber & "_StdPartsActivityRule", "/* Activity rule for " & oSeedPartInstance.PartNumber & "*/", sRelationText)
            
            'Modify existing rule
            Else
            
                If frmStdPartsInstantiation.chkDeleteStdPart.Value = False Then
                    
                    sPreviousRelationText = oRule.Value
                    
                    'We clean the text of the old rule
                    Call CleanOldRuleText
                    
                    'Before concatenating the sPreviousRelationText to the sRelationText we need to delete line from the sPreviousRelationText to avoid duplicate
                    Call EditRuleText
                    
                    'Concatenate the left over of the old rule to the new rule
                    sRelationText = sRelationText & sPreviousRelationText
                    oRule.Modify (sRelationText)
                Else
                    oRule.Modify (sRelationText)
                End If
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
    Next
    
    oSelection.Clear
    CATIA.StatusBar = "Successfully Done! "
    
    'Update Assy
    If MsgBox("Standard Parts Instantiated!     " & vbCrLf & vbCrLf & "Do you want to update the assembly?    ", vbYesNo + vbQuestion, "Standard Parts Instantiation") = vbYes Then
        Call frmProgress.progressBarRepaint("Updating Constraints...")
        oSeedAssy.Update
        If sLog = "" Then MsgBox "Update Successful!     ", vbInformation, "Standard Parts Instantiation"
    End If

endsub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_KBEStdPartInstantiate.txt"
        sLog = "The following problems were encountered:" & vbCrLf & sLog
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If
    
    On Error GoTo 0
     
End Sub

Private Sub DeleteConstraintStdPartRule()
'This Sub will delete all the constraint and Std Part instances linked to a specific Seed Part
    
    If Not oConstraints.Count = 0 Then
    'Loop all the Seed Parts
        iProgressCount1 = 0
        For Each oSeedPartInstance In oSeedPartCollection
            
            'Retrieve and delete activity rule
            For Each oRelation In oSeedAssy.Relations
                'Progress Bar
                timeEnd = DateTime.Timer
                Call frmProgress.progressBarRepaint("Deleting from " & oSeedPartInstance.PartNumber & " (" & CStr(iProgressCount1) & "/" & CStr(oSeedPartCollection.Count) & ")", _
                iProgressCount1Max, iProgressCount1, "Deleting Activity Rule...", _
                1, 0.001, "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss"))

                If TypeName(oRelation) = "Rule" And oRelation.Name Like oSeedPartInstance.PartNumber & "*" Then
                    oSeedAssy.Relations.Remove (oRelation.Name)
                End If
            Next
            
            'Retrieve and delete constraints and associated Std parts
            iProgressCount1 = iProgressCount1 + 1
            iProgressCount2 = 0
            iProgressCount2Max = oConstraints.Count
            For Each oConstraint In oConstraints
                
                'Progress Bar
                timeEnd = DateTime.Timer
                iProgressCount2 = iProgressCount2 + 1
                Call frmProgress.progressBarRepaint("Deleting from " & oSeedPartInstance.PartNumber & " (" & CStr(iProgressCount1) & "/" & CStr(oSeedPartCollection.Count) & ")", _
                iProgressCount1Max, iProgressCount1, "Scanning previous Standard Parts instances (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")", _
                iProgressCount2Max, iProgressCount2, "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss"))

                'The last portion of the constraint name should be the Seed Part Number
                If oConstraint.Name Like "*" & oSeedPartInstance.PartNumber Then
                    'The Std Part instances name is the first portion of the constraint name
                    On Error Resume Next 'The instance may have been deleted manually
                    
                    'Activate the instance before deleting it, then delete it
                    Call oInstanceActivation.ActivateOrDeactivate(oSeedAssy.Products.Item(Split(oConstraint.Name, " | ")(0)), True) 'Activate the instance before deleting it
                    
                    oSeedAssy.Products.Remove oSeedAssy.Products.Item(Split(oConstraint.Name, " | ")(0))
                    oConstraints.Remove (oConstraint.Name)
                    On Error GoTo 0
                    
                End If
            Next
        Next
    End If
    
End Sub
                
Private Function RetrieveConstraint(ByVal sPartNumber As String, ByVal sAxisName As String, ByVal sSeedPartNumber As String) As Constraint
    
    Dim sConstraintRef1 As String
    Dim sConstraintRef2 As String
    Dim sSourceConstraintName As String 'Concatenation of the Std Part Number + Axis Name + Seed Part Number
    Dim sTargetConstraintName As String 'Name of the constraint found in the oNonFixConstraint collection
    Dim sConstraintName As String
    
    sConstraintName = ""
    sSourceConstraintName = ""
    sTargetConstraintName = ""
    Set RetrieveConstraint = Nothing
    
    sSourceConstraintName = Replace(sPartNumber, "-", "_", 1, -1, vbBinaryCompare)
    sSourceConstraintName = sSourceConstraintName & " | " & sAxisName
    sSourceConstraintName = sSourceConstraintName & " | " & sSeedPartNumber
    
    'We loop all the constraint
    For i = 1 To oCoincidenceConstraintCollection.Count
        Set oConstraint = oCoincidenceConstraintCollection.Item(i)
        
        On Error Resume Next
        sConstraintRef1 = ""
        sConstraintRef2 = ""
        Set oReference = oConstraint.GetConstraintElement(1)
        sConstraintRef1 = oReference.DisplayName
        Set oReference = oConstraint.GetConstraintElement(2)
        sConstraintRef2 = oReference.DisplayName
        
        
        'Find matching constraint
        If sConstraintRef1 Like "*Absolute Axis System" Then
            sTargetConstraintName = Replace(Split(Split(sConstraintRef1, "/")(1), ".")(0), "-", "_")
            sTargetConstraintName = sTargetConstraintName & " | " & Split(sConstraintRef2, "/!")(UBound(Split(sConstraintRef2, "/!")))
            sTargetConstraintName = sTargetConstraintName & " | " & Split(Split(sConstraintRef2, "/")(UBound(Split(sConstraintRef2, "/")) - 1), ".")(0)
            
            sConstraintName = Split(sConstraintRef1, "/")(1)
            sConstraintName = sConstraintName & " | " & Split(sConstraintRef2, "/!")(UBound(Split(sConstraintRef2, "/!")))
            sConstraintName = sConstraintName & " | " & Split(Split(sConstraintRef2, "/")(UBound(Split(sConstraintRef2, "/")) - 1), ".")(0)
            
        ElseIf sConstraintRef2 Like "*Absolute Axis System" Then
            sTargetConstraintName = Replace(Split(Split(sConstraintRef2, "/")(1), ".")(0), "-", "_")
            sTargetConstraintName = sTargetConstraintName & " | " & Split(sConstraintRef1, "/!")(UBound(Split(sConstraintRef1, "/!")))
            sTargetConstraintName = sTargetConstraintName & " | " & Split(Split(sConstraintRef1, "/")(UBound(Split(sConstraintRef1, "/")) - 1), ".")(0)
            
            sConstraintName = Split(sConstraintRef2, "/")(1)
            sConstraintName = sConstraintName & " | " & Split(sConstraintRef1, "/!")(UBound(Split(sConstraintRef1, "/!")))
            sConstraintName = sConstraintName & " | " & Split(Split(sConstraintRef1, "/")(UBound(Split(sConstraintRef1, "/")) - 1), ".")(0)
            
        End If
        On Error GoTo 0
        
        If sTargetConstraintName = sSourceConstraintName Then
            Set RetrieveConstraint = oConstraint
            oCoincidenceConstraintCollection.Remove (i)
            
            If Not RetrieveConstraint.Name = sConstraintName And Not sConstraintName = "" Then
                sLog = sLog & vbCrLf & " - The following constraint was renamed:"
                sLog = sLog & vbCrLf & Chr(9) & "Old name: " & RetrieveConstraint.Name
                sLog = sLog & vbCrLf & Chr(9) & "New name: " & sConstraintName
                RetrieveConstraint.Name = sConstraintName
            End If
            
            Exit For
        End If
    Next i

End Function

Private Sub EditRuleText()

'The goal of this Sub is to remove from the old rule, lines found in the new rule

Dim sOldLine As String
Dim sNewLine As String
Dim sNewStdInstanceName As String
Dim sOldStdInstanceName As String

'We scan all the lines of the new rule
For i = UBound(Split(sRelationText, vbCrLf)) To 0 Step -1

    sNewLine = Split(sRelationText, vbCrLf)(i)
    If sNewLine <> "" Then
    
        'Retrieve the instance name of the Std Part
        sNewStdInstanceName = Split(sNewLine, "`")(1)
        sNewStdInstanceName = Split(sNewStdInstanceName, "\Component Activation State")(0)
        
        'We scan all the lines of the old rule
        For j = UBound(Split(sPreviousRelationText, vbCrLf)) To 0 Step -1
    
            sOldLine = Split(sPreviousRelationText, vbCrLf)(j)
            If sOldLine <> "" Then
                                      
                'Retrieve the instance name of the Std Part
                sOldStdInstanceName = Split(sOldLine, "`")(1)
                sOldStdInstanceName = Split(sOldStdInstanceName, "\Component Activation State")(0)
            
                                      
                If sNewStdInstanceName = sOldStdInstanceName Then
                
                    'Lines from the the old rule found in the new rule will be removed
                    sPreviousRelationText = Replace(sPreviousRelationText, sOldLine, "")
                    
                    Exit For
                End If
            End If
        Next
    End If
Next

'Delete empty lines
sPreviousRelationText = Replace(sPreviousRelationText, vbCrLf & vbCrLf, vbCrLf)

End Sub

Private Sub CleanOldRuleText()

'Sometimes the text retrieved from the old rule is corrupted.
'The name of the Seed Assy is added to the Std Part parameter name.
'I dont't know why it does that but it needs to be corrected.

'`KBE_M55_SEED_ASSY_1-V06\COMBO_OVEN-001KBE.6\Component Activation State` = `KBE_M55......
'instead of
'COMBO_OVEN-001KBE.6\Component Activation State` = `KBE_M55......

Dim sLine As String
Dim sLineCleaned As String

'We scan all the lines of the old rule
For i = 0 To UBound(Split(sPreviousRelationText, vbCrLf))

    sLine = Split(sPreviousRelationText, vbCrLf)(i)
    If sLine <> "" Then
    
        'If the text of the old rule is corrupted
        If sLine Like "`" & oSeedAssy.PartNumber & "*" Then
            sLineCleaned = Replace(sLine, "`" & oSeedAssy.PartNumber & "\", "`")
            sPreviousRelationText = Replace(sPreviousRelationText, sLine, sLineCleaned)
        End If
    End If
        
Next

End Sub
