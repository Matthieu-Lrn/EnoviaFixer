Attribute VB_Name = "StringToModify"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long

'********************************************************************************
'* Name: Find And Replace
'* Purpose: The goal of this macro is to find a string an replace it by another within selection. There's 2 modes:
'*          Mode 1: User pre-select elements to rename. Macro only rename these elements
'*          Mode 2: No pre-selection. User launch macro, then do a selection. The selection and everything under it may be renamed
'*          Checkbox allows to rename part & instances - Not suitable in production (should be used only on local files)
'*
'* Assumption: Parts/Products cannot be renamed by the tool, but everything inside parts
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub FindReplaceOnSelection(Optional ByVal StrToFind As String = "", Optional ByVal StrForReplace As String = "", Optional ByVal bShowProgress As Boolean = True)
    
    Dim sPreviousObjectName As String
    Dim sNewObjectName As String
    Dim sTypeParent As Variant
    Dim oPart As Part
    Dim oReference As Reference
    Dim oPartDocument As PartDocument
    
    Set oSelection = CATIA.ActiveDocument.Selection
    
    'Select an element to scan and rename
    If oSelection.Count2 = 0 Then
        Set oSelectedItem = MultiWindowSelectElement("No active selection found. Please select an item to search...")
        If oSelectedItem Is Nothing Then Exit Sub '**** If loop is stopped without selecting anything
        
        'Called from UserForm
        If StrToFind = "" And StrForReplace = "" Then
            Call frmFindAndReplace.FormInitialize
            StrToFind = frmFindAndReplace.txtStringToFind.Value
            StrForReplace = frmFindAndReplace.txtReplacingString.Value
            If StrToFind = "" Then Exit Sub '***** If user press Cancel
            
            'CATIA.Interactive = False 'Do not allow user to click something
            frmKBEMain.Enabled = False 'Do not allow user to click something
        End If
        
        If bShowProgress = True Then
            Call frmProgress.progressBarInitialize("Find and Replace")
            Call frmProgress.progressBarRepaint("Scanning Selection...")
        End If
        
        Set oSelection = KBESearch("Name_CAP=*" & ModifySpecialCharacters(frmFindAndReplace.txtStringToFind.Value) & "*,sel", oSelectedItem.Value)
        
        If oSelection.Count2 = 0 Then GoTo endsub '***** If search gives no result
        
    Else
        'Called from UserForm
        If StrToFind = "" And StrForReplace = "" Then
            Call frmFindAndReplace.FormInitialize
            StrToFind = frmFindAndReplace.txtStringToFind.Value
            StrForReplace = frmFindAndReplace.txtReplacingString.Value
            
            If StrToFind = "" Then Exit Sub '***** If user press Cancel
            
            'CATIA.Interactive = False 'Do not allow user to click something
            frmKBEMain.Enabled = False 'Do not allow user to click something
        
        End If
        
        If bShowProgress = True Then
            Call frmProgress.progressBarInitialize("Find and Replace")
        End If
        
    End If
    
    'Rename selection
    On Error Resume Next
    For i = 1 To oSelection.Count2
        
        If bShowProgress = True Then
            Call frmProgress.progressBarRepaint("Renaming...", CDbl(oSelection.Count), CDbl(i), , , , "(" & CStr(i) & "/" & CStr(oSelection.Count) & ")")
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        End If
        
        Set iObject = oSelection.Item2(i).Value
        sTypeParent = ""
        sTypeParent = TypeName(iObject.Parent) 'Empty if parent is not found
        
        If TypeName(iObject) = "ParameterSet" Then 'Renaming method specific for parameter sets (part level only)
            sNewObjectName = Replace(iObject.Name, StrToFind, StrForReplace)
            
            If oSelection.Item2(i).Document.Name Like "*CATPart" Then
                Set oPartDocument = oSelection.Item2(i).Document
                Set oPart = oPartDocument.Part
            Else
                Set oPart = oSelection.Item2(i).LeafProduct.ReferenceProduct.Parent.Part
            End If
            Set oReference = oPart.CreateReferenceFromObject(iObject)
            oPart.HybridShapeFactory.ChangeFeatureName oReference, sNewObjectName
        
        ElseIf sTypeParent = "Parameters" And Not (TypeName(iObject) Like "*Shape*") Then 'Renaming method specific for parameters
            sPreviousObjectName = Split(iObject.Name, "\")(UBound(Split(iObject.Name, "\")))
            sNewObjectName = Replace(sPreviousObjectName, StrToFind, StrForReplace)
            If Not UCase(sPreviousObjectName) Like "*TOOLVERSION*" Or UCase(StrToFind) Like "*TOOLVERSION*" Then '*****Do not modify ToolVersion parameter, unless "ToolVersion" is found in the search string
                If InStr(iObject.Name, "\") <= 0 Then iObject.Rename CStr(sNewObjectName)           '*****Modify parameter Global Name (only if path not found in parameter Name, e.g.modified manually by user)
                iObject.Rename "Local:" & CStr(sNewObjectName)                                      '*****Put "Local:" before string to modify parameter Local Name
            End If
        
        ElseIf sTypeParent = "Relations" Then 'Renaming method specific for relations (rule, formula...)
            iObject.Rename CStr(Replace(iObject.Name, StrToFind, StrForReplace))
            
        ElseIf (TypeName(iObject) = "Product" Or TypeName(iObject) = "Part") And frmFindAndReplace.cbRenameInstances.Value = True Then 'Renaming method specific for parts and part instances
            If TypeName(oSelection.Item2(i).Document) = "ProductDocument" Then
                oSelection.Item2(i).LeafProduct.Parent.Parent.ReferenceProduct.Products.Item(oSelection.Item2(i).LeafProduct.Name).Name = Replace(oSelection.Item2(i).LeafProduct.Name, StrToFind, StrForReplace)
                oSelection.Item2(i).LeafProduct.PartNumber = Replace(oSelection.Item2(i).LeafProduct.PartNumber, StrToFind, StrForReplace)
            ElseIf TypeName(oSelection.Item2(i).Document) = "PartDocument" Then
                oSelection.Item2(i).Document.Product.PartNumber = Replace(oSelection.Item2(i).Document.Product.PartNumber, StrToFind, StrForReplace)
            End If
        
        Else 'General method for any other object
            iObject.Name = CStr(Replace(iObject.Name, StrToFind, StrForReplace))
        
        End If
        
    Next i
    
endsub:
    
    If CATIA.Interactive = False Then
        CATIA.Interactive = True
        'Needed to "reactivate" Catia after CATIA.Interactive = True
        CATIA.ActiveWindow.ActivatePrevious
        CATIA.ActiveWindow.ActivateNext
    End If
    
    frmKBEMain.Enabled = True
    CATIA.StatusBar = "Successfully Done! "
    On Error GoTo 0
    
End Sub


'********************************************************************************
'* Name: Modify Special Characters
'* Purpose: Used by Find & Replace > KBE Search to adapt the query on some special
'*          characters (Some characters need to be put between quotes!)
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Function ModifySpecialCharacters(ByVal oString As String) As String
    
    ModifySpecialCharacters = oString
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, "&", "'&'") '   Put & between quotes > Mandatory
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, "(", "'('") '   Put ( between quotes
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, ")", "')'") '   Put ) between quotes
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, "+", "'+'") '   Put + between quotes > Mandatory
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, "-", "'-'") '   Put - between quotes > Mandatory
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, ";", "';'") '   Put ; between quotes
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, ".", "'.'") '   Put . between quotes
    ModifySpecialCharacters = Replace(ModifySpecialCharacters, ",", "','") '   Put , between quotes
    
End Function


'********************************************************************************
'* Name: Remove "Result Of" and "Copy Of"
'* Purpose: The goal of this macro is to remove all "Result Of" and "Copy Of" in features names. There's 2 modes:
'*          Mode 1: User pre-select elements, then the macro scans everything under selection(s).
'*          Mode 2: No pre-selection. User launch macro, then do a selection. The macro scans everything under selection.
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub SearchAndRemoveResultOfCopyOf()
    
    Set oSelection = CATIA.ActiveDocument.Selection

    If oSelection.Count = 0 Then
        Set oSelectedItem = MultiWindowSelectElement("No active selection found. Please select an item to search...")
        If oSelectedItem Is Nothing Then Exit Sub '**** If loop is stopped without selecting anything
    End If

    CATIA.Interactive = False 'Do not allow user to click something
    frmKBEMain.Enabled = False 'Do not allow user to click something
    Call disableOptionNameCheck 'Verify CATSettings (No Name Check)
    
    Call frmProgress.progressBarInitialize("Remove Copy of / Result of")
    Call frmProgress.progressBarRepaint("Scanning Selection...")
    
    Set oSelection = KBESearch("Name_CAP=*Result* + Name_CAP=*Copy*,sel")
    If oSelection.Count2 = 0 Then GoTo endsub '***** If search gives no result
    
    On Error Resume Next
    For i = 1 To oSelection.Count2
        
        Set iObject = Nothing '***** Bug R21: Selection.Item(i) may fail. We loop until it finds a value.
        Do
            Set iObject = CATIA.ActiveDocument.Selection.Item(i).Value
        Loop While iObject Is Nothing
        
        Do Until Not iObject.Name Like "*Result *" And Not iObject.Name Like "*Copy *"
            '*****Result of....
            If iObject.Name Like ("*Result of*") Then
                j = InStr(1, iObject.Name, "Result of", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 10)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Result (#)*") Then
                j = InStr(1, iObject.Name, "Result (", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 14)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Result (##)*") Then
                j = InStr(1, iObject.Name, "Result (", vbTextCompare)
                StrForReplace = Mid(iObject.Name, j, 15)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Result (###)*") Then
                j = InStr(1, iObject.Name, "Result (", vbTextCompare)
                StrForReplace = Mid(iObject.Name, j, 16)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Result (####)*") Then
                j = InStr(1, iObject.Name, "Result (", vbTextCompare)
                StrForReplace = Mid(iObject.Name, j, 17)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            '*****Copy of....
            ElseIf iObject.Name Like ("*Copy of*") Then
                j = InStr(1, iObject.Name, "Copy of", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 8)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Copy (#)*") Then
                j = InStr(1, iObject.Name, "Copy (", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 12)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Copy (##)*") Then
                j = InStr(1, iObject.Name, "Copy (", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 13)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Copy (###)*") Then
                j = InStr(1, iObject.Name, "Copy (", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 14)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            ElseIf iObject.Name Like ("*Copy (####)*") Then
                j = InStr(1, iObject.Name, "Copy (", vbBinaryCompare)
                StrForReplace = Mid(iObject.Name, j, 15)
                iObject.Name = Replace(iObject.Name, StrForReplace, "")
            '***** Else (Copy or Result is found in the name, but not as a Copy of / Result of...)
            Else
                Exit Do
            End If
        Loop
        
        Call frmProgress.progressBarRepaint("Renaming...", CDbl(oSelection.Count), CDbl(i), , , , "(" & CStr(i) & "/" & CStr(oSelection.Count) & ")")
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
        
    Next i
        
endsub:
    On Error GoTo 0
    oSelection.Clear
    CATIA.Interactive = True
    frmKBEMain.Enabled = True
    CATIA.StatusBar = "Successfully Done! "
    
End Sub


'********************************************************************************
'* Name: Rename Assemble Features
'* Purpose: In a 3D Collector, rename all the Assemble features with the name of the body under it
'*          Mode 1: User is in a part, nothing to select, all assembles are renamed
'*          Mode 2: User is in a product. User has to select any item in a part, and all assembles of the selected part are renamed.
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************

Public Sub RenameAssembleFeatures()
    
    Dim oBody As Body
    Dim oShape As Shape
    Dim oSelectedPart As Part
    Dim oSelectedPartInstance As Product
    Dim oSelectedPartRootDocument As Document
    
    On Error Resume Next
    If TypeName(CATIA.ActiveDocument) = "PartDocument" Then
        Set oSelectedPart = CATIA.ActiveDocument.Part
    Else
        Call selectPart(oSelectedPart, oSelectedPartInstance, oSelectedPartRootDocument, , "Please select a Part...  ")
        
        If bStopMultiSelection = False And bToolbarCommandSelected = False Then
            If oSelectedPart Is Nothing Then
                MsgBox "Please select a CATPart!  ", vbCritical, "Wrong Selection..."
                GoTo noSelection
            End If
        Else
            GoTo noSelection
        End If
        
    End If
    
    If oSelectedPart Is Nothing Then
noSelection:
        Set oSelectedPart = Nothing
        Set oSelectedPartInstance = Nothing
        Set oSelectedPartRootDocument = Nothing
        Exit Sub
    End If
    
    For Each oBody In oSelectedPart.Bodies
        If oBody.InBooleanOperation = False Then
            For Each oShape In oBody.Shapes
                If TypeName(oShape) = "Assemble" Then oShape.Name = oShape.Body.Name
            Next
        End If
    Next
        
    CATIA.StatusBar = "Successfully Done! "
    On Error GoTo 0
    
End Sub
