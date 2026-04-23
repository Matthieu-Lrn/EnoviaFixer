Attribute VB_Name = "EvolveParts"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************
'* SCRIPT VERSION: July 19th 2016
'*
'* When modifying this code please update in all toolbars
'********************************************************
Private s_ExistingPartOption As String
Private b_EvolveDefiningPart As String
Private oRootDoc As Document
Private sPartAndDocList As clsCollection

Public Sub LaunchEvolveParts(ByVal oSelectedPart As Product, ByVal sSourceString As String, ByVal sTargetString As String, _
                             ByVal iIncrement As Integer, ByVal bAllLevels As Boolean, ByVal oRootDocument As Document, ByVal sBase As String, _
                             ByVal sExistingPartOption As String, ByVal bEvolveDefiningPart As Boolean)


Dim sSelectedPartType As String
Dim sNewPartNumber As String
Dim oParent As Product
Dim oNewPart As Product
Dim iIndex As Integer
Dim sAnswer As String
Dim bReuseExistingPart As Boolean


'** Initial set
Set oRootDoc = oRootDocument
Set oSelection = CATIA.ActiveDocument.Selection
Set oNewPart = Nothing
Set sPartAndDocList = New clsCollection
bReuseExistingPart = False
s_ExistingPartOption = sExistingPartOption
b_EvolveDefiningPart = bEvolveDefiningPart

'** Retrieve selected part type
sSelectedPartType = RetrievePartType(oSelectedPart)

'** Progress Bar
Call frmProgress.progressBarInitialize("Evolving Parts")

'** Put all children of oSelectedPart in Default Mode
If sSelectedPartType Like "*Product" Then
    oSelectedPart.ApplyWorkMode DEFAULT_MODE
End If

'** Retrieve parent instance of selected part
If sSelectedPartType <> "RootProduct" Then
    Set oParent = oSelectedPart.Parent.Parent
End If

'** Create a list of all loaded parts
Call RetrieveAllPartsAndDoc

'** Find new part number
' Replace string by another one
sNewPartNumber = Replace(oSelectedPart.PartNumber, sSourceString, sTargetString)
' Change dash number
If iIncrement <> 0 And (sBase = "" Or oSelectedPart.PartNumber Like sBase & "*") Then
    sNewPartNumber = RollDashNumber(sNewPartNumber, iIncrement)
End If

'** If the new part number is equal to the original one OR the selected part in the Root Document
If sNewPartNumber = oSelectedPart.PartNumber Or sSelectedPartType = "RootProduct" Then GoTo ScanAllTheLevels

'** Make sure that the new part number is not found anywhere in the parent list
'** Otherwise we will have a loop
sParentList = BuildParentList(oSelectedPart)
If sParentList Like "*" & sNewPartNumber & "*" Then
    sAnswer = MsgBox("The macro will abort the evolution of " & oSelectedPart.PartNumber & " to " & sNewPartNumber & " otherwise a loop will be created.", vbCritical)
    Exit Sub
End If

'** Check if the new part number is already open in CATIA
If sPartAndDocList.Exists(sNewPartNumber) Then
    If s_ExistingPartOption = "Ask Question" Then
        sAnswer = MsgBox(sNewPartNumber & " is already open in the CATIA session." & vbCrLf & _
                         " - Click Yes to skip evolution of this part" & vbCrLf & _
                         " - Click No to reuse with existing reference", 36)
    ElseIf s_ExistingPartOption = "Skip Evolution" Then
        sAnswer = 6
    Else
        sAnswer = 7
    End If
End If


'** Copy/Paste oSelectedPart if sAnswer is different from Yes
Set oNewPart = Nothing
If sAnswer = "7" Or sAnswer = "" Then
    
    '** Progress Bar
    Call frmProgress.progressBarRepaint("Evolving Part: " & oSelectedPart.Name)

    '** Copy / Paste and keep original
    Call Copy_Paste(oSelectedPart, oParent, False)
    
    '** The newly created part is the last one in the Products collection
    Set oNewPart = oParent.Products.Item(oParent.ReferenceProduct.Products.Count)
    
    '** Change part number if we have a new reference and add to list
    If sAnswer = "" Then
        oNewPart.PartNumber = sNewPartNumber
        Call sPartAndDocList.Add(oNewPart.PartNumber, oNewPart.ReferenceProduct.Parent.Name)

    '** Reuse existing reference
    Else
        oParent.Products.ReplaceProduct oNewPart, CATIA.Documents.Item(sPartAndDocList.GetItemByKey(sNewPartNumber)).Product, False
        bReuseExistingPart = True
    End If

    '** Change instance name
    Call oRenameInstance.RenameSelectedInstance(oNewPart)

    '*** Defining Part
    If b_EvolveDefiningPart = True Then
        Call SetDefiningPart(oNewPart)
    End If
    
'** User wants to skip the evolution
Else
    Exit Sub
End If

'** Check if we escape
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub

'** Scan all levels
ScanAllTheLevels:
If bReuseExistingPart = False Then
    If Not oNewPart Is Nothing Then
        If sSelectedPartType Like "*Product" And bAllLevels And Not oNewPart.PartNumber Like "*-*SP##" Then
            Call ScanAllLevels(oNewPart, sSourceString, sTargetString, iIncrement, sBase)
        End If
    Else
        If sSelectedPartType Like "*Product" And bAllLevels And Not oSelectedPart.PartNumber Like "*-*SP##" Then
            Call ScanAllLevels(oSelectedPart, sSourceString, sTargetString, iIncrement, sBase)
        End If
    End If
End If

Call frmProgress.Hide

End Sub

Private Sub ScanAllLevels(ByVal oParent As Product, ByVal sSourceString As String, ByVal sTargetString As String, ByVal iIncrement As Integer, ByVal sBase As String)

Dim oChild As Product
Dim sChildPartType As String
Dim sNewPartNumber As String
Dim oNewPart As Product
Dim iIndex As Integer
Dim sAnswer As String
Dim bReuseExistingPart As Boolean

bReuseExistingPart = False

'** Scan all children of oParent
For Each oChild In oParent.Products

    sAnswer = ""
    
    'Progress Bar
    Call frmProgress.progressBarRepaint("Evolving Part: " & oChild.Name)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub

    '** Retrieve selected part type
    sChildPartType = RetrievePartType(oChild)
    
    '** Find new part number
    ' Replace string by another one
    sNewPartNumber = Replace(oChild.PartNumber, sSourceString, sTargetString)
    ' Change dash number
    If iIncrement <> 0 And (sBase = "" Or oChild.PartNumber Like sBase & "*") Then
        sNewPartNumber = RollDashNumber(sNewPartNumber, iIncrement)
    End If
    
    '** If the new part number is equal to the original one
    If sNewPartNumber = oChild.PartNumber Then GoTo NextPart

    '** Make sure that the new part number is not found anywhere in the parent list
    '** Otherwise we will have a loop
    sParentList = BuildParentList(oChild)
    If sParentList Like "*" & sNewPartNumber & "*" Then
        sAnswer = MsgBox("The macro will skip the evolution of " & oChild.PartNumber & " to " & sNewPartNumber & " otherwise a loop will be created.", vbCritical)
        GoTo NextPart
    End If

    '** Check if the new part number is already open in CATIA
    If sPartAndDocList.Exists(sNewPartNumber) Then
        If s_ExistingPartOption = "Ask Question" Then
            sAnswer = MsgBox(sNewPartNumber & " is already open in the CATIA session." & vbCrLf & _
                             " - Click Yes to skip evolution of this part" & vbCrLf & _
                             " - Click No to reuse with existing reference", 36)
        ElseIf s_ExistingPartOption = "Skip Evolution" Then
            sAnswer = 6
        Else
            sAnswer = 7
        End If
    End If

    '** Copy/Paste oChild if sAnswer is different from vbYes
    Set oNewPart = Nothing
    If sAnswer = "7" Or sAnswer = "" Then
        
        '** Copy / Paste and delete original
        Call Copy_Paste(oChild, oParent, True)
        
        '** The newly created part is the last one in the Products collection
        Set oNewPart = oParent.Products.Item(oParent.ReferenceProduct.Products.Count)
        
        '** Change part number if we have a new reference and add to list
        If sAnswer = "" Then
            oNewPart.PartNumber = sNewPartNumber
            Call sPartAndDocList.Add(oNewPart.PartNumber, oNewPart.ReferenceProduct.Parent.Name)
    
        '** Reuse existing reference
        Else
            oParent.Products.ReplaceProduct oNewPart, CATIA.Documents.Item(sPartAndDocList.GetItemByKey(sNewPartNumber)).Product, False
            bReuseExistingPart = True
        End If
        
        '** Rename instance name
        Call oRenameInstance.RenameSelectedInstance(oNewPart)
        
        '*** Defining Part
        If b_EvolveDefiningPart = True Then
            Call SetDefiningPart(oNewPart)
        End If
        
    '** User wants to skip the evolution
    Else
        GoTo NextPart
    End If

    '** Scan all levels
    If bReuseExistingPart = False Then
        If Not oNewPart Is Nothing Then
            If sChildPartType Like "*Product" And Not oNewPart.PartNumber Like "*-*SP##" Then
                Call ScanAllLevels(oNewPart, sSourceString, sTargetString, iIncrement, sBase)
            End If
        Else
            If sChildPartType Like "*Product" And Not oParent.PartNumber Like "*-*SP##" Then
                Call ScanAllLevels(oChild, sSourceString, sTargetString, iIncrement, sBase)
            End If
        End If
    End If
    
NextPart:

Next

End Sub


Private Sub Copy_Paste(ByVal oProduct As Product, ByVal oParent As Product, ByVal bDelete As Boolean)
'**************************************************************
'* This Sub will Copy / Paste an instance under the same parent
'* It can also delete the original instance
'**************************************************************

'** Copy
oSelection.Clear
oSelection.Add oProduct
oSelection.Copy

'** Paste break link
oSelection.Clear
oSelection.Add oParent
oSelection.PasteSpecial "CATSpecBreakLink"

'** Delete original
If bDelete = True Then
    oSelection.Clear
    oSelection.Add oProduct
    oSelection.Delete
End If

End Sub

Private Sub RetrieveAllPartsAndDoc()
'***************************************************************
'* The Sub will scan all open documents. When it finds a CATPart
'* or a CATProduct it will add it to sPartAndDocList. The key is the Part Number and the value is the document name.
'***************************************************************

Dim oDocument As Document
Dim sPartNumber As String

'** Loop all loaded documents
For Each oDocument In CATIA.Documents

    sPartNumber = ""
       
    'CATPart or CATProduct
    If Right(oDocument.Name, 11) = ".CATProduct" Or Right(oDocument.Name, 8) = ".CATPart" Then
        
        'Keep everything left of the "."
        sPartNumber = Split(oDocument.Name, ".")(0)
    
        'Remove last two character if there is a revision
        If sPartNumber Like "*-[-A-Z]" Then
            sPartNumber = Left(sPartNumber, Len(sPartNumber) - 2)
        End If
    End If


    '** Part Number found, populate sPartAndDocList
    If sPartNumber <> "" And Not sPartAndDocList.Exists(sPartNumber) Then
        Call sPartAndDocList.Add(sPartNumber, oDocument.Name)
    End If
Next

End Sub

Private Function RetrievePartType(oProduct As Product) As String
'*****************************************************************
'The goal of this function is to return the type of an instance:
' - CATPart
' - CATProduct
' - Component
'*****************************************************************

If oProduct.HasAMasterShapeRepresentation Then
    RetrievePartType = "CATPart"
ElseIf oProduct.ReferenceProduct.Parent.Name = oRootDoc.Name Then
    RetrievePartType = "RootProduct"
ElseIf oProduct.ReferenceProduct.Parent.Name = oProduct.Parent.Parent.ReferenceProduct.Parent.Name Then
    RetrievePartType = "Component"
Else
    RetrievePartType = "CATProduct"
End If

End Function

Private Function RollDashNumber(ByVal sPartNumber As String, ByVal iIncrement As Integer) As String
'********************************************************************
' This function will increment the dash number of a given part number
'********************************************************************

Dim sBase As String
Dim iDash As Integer
Dim sDash As String
Dim sComplement As String


'***Find the Dash Number. It is all characters right of the last "-" in the Part Number
'We don't have a "-" in the Part Number
If InStrRev(sPartNumber, "-") = 0 Then
    RollDashNumber = sPartNumber
    Exit Function
End If

'Get Dash Number
sDash = Mid(sPartNumber, InStrRev(sPartNumber, "-"))

'Check if the 3 characters after the "-" in the Dash Number are numeric
If Not sDash Like "-###*" Then
    RollDashNumber = sPartNumber
    Exit Function
End If

'*** Decompose Part Number
sBase = Left(sPartNumber, InStrRev(sPartNumber, "-") - 1)
iDash = CInt(Mid(sDash, 2, 3))
If Len(sDash) > 3 Then
    sComplement = Mid(sDash, 5)
End If

'** Calculate new dash number
iDash = iDash + iIncrement
If iDash < 0 Then
    sDash = "000"
ElseIf iDash < 10 Then
    sDash = "00" & CStr(iDash)
ElseIf iDash < 100 Then
    sDash = "0" & CStr(iDash)
Else
    sDash = CStr(iDash)
End If

'** Transfer new part number
RollDashNumber = sBase & "-" & sDash & sComplement

End Function


Private Sub SetDefiningPart(ByVal oProduct As Product)

Dim oDatasetType As StrParam
Dim oDefiningPart As StrParam
Dim sDefiningPart As String
Dim oParameters As Parameters


'Parameters
Set oParameters = oProduct.ReferenceProduct.UserRefProperties

'**Define Defining Part value
'We have a dash number in the Part Number
If UBound(Split(oProduct.PartNumber, "-")) >= 1 Then
    sDefiningPart = Split(oProduct.PartNumber, "-", 2)(0) & "-" & Left(Split(oProduct.PartNumber, "-", 2)(1), 3)

'No dash number in the Part Number. This case should not exit but I don't want the macro to crash if it happens
Else
    sDefiningPart = sPartNumber
End If


'Get Dataset Type
On Error Resume Next
Set oDatasetType = Nothing
Set oDatasetType = oParameters.GetItem("Dataset Type")
On Error GoTo 0

'Set Defining Part attribute when
' - Dataset Type = SKELETON
' - Dataset Type = CATALOG LIGHT VERSION
' - Dataset Type = FLEXIBLE REPRESENTATION
' - Dataset Type = SPLIT PART
' - Part Number like "*KIN*"
If oDatasetType.Value = "SKELETON" Or oDatasetType.Value = "CATALOG LIGHT VERSION" Or _
    oDatasetType.Value = "FLEXIBLE REPRESENTATION" Or oDatasetType.Value = "SPLIT PART" Or _
    sPartNumber Like "*KIN*" Then

    'Get Defining Part attribute
    On Error Resume Next
    Set oDefiningPart = Nothing
    Set oDefiningPart = oParameters.GetItem("Defining Part")
    On Error GoTo 0

    'Set Defining Part attribute to first 13 characters of Part Number
    If Not oDefiningPart Is Nothing Then

        oDefiningPart.ValuateFromString (sDefiningPart)
    End If
End If

End Sub

