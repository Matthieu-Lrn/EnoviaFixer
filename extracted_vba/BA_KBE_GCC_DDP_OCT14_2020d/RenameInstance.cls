Attribute VB_Name = "RenameInstance"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************
'* SCRIPT VERSION: May 22th 2015
'*
'* When modifying this code please update in all toolbars
'********************************************************
Private sChildNames() As String
Private sChildNameString As String
Private oRootDocument As Document

Public Sub SelectParentInstance()

    Dim oParent As Product

    Dim i As Integer
    Dim oChild As Product
    Dim sParentType As String
    Dim sAnswer As String
    Dim sPartNumber As String
   
    'Select the parent assy
    Call selectAssy(oParent, oRootDocument, , "Select the parent assembly")
    If oParent Is Nothing Then MsgBox "Please select a Product!  ", vbCritical, "Warning": GoTo endsub
    
    'Retrieve the type of the parent assy
    sParentType = FindInstanceType(oParent)
    
    'Retrieve the name of all the child and transfert them in the array sChildNames
    ReDim sChildNames(0)
    For i = 1 To oParent.Products.Count
        Set oChild = oParent.Products.Item(i)
        
        If sChildNames(UBound(sChildNames)) = "" Then
            sChildNames(UBound(sChildNames)) = oChild.Name
        Else
            ReDim Preserve sChildNames(UBound(sChildNames) + 1)
            sChildNames(UBound(sChildNames)) = oChild.Name
        End If
    Next
    
    'Progress bar
    Call frmProgress.progressBarInitialize("Renaming Instances")

    'Rename the first level children only. Do not rename Components
    For i = 1 To oParent.Products.Count
    
        Call frmProgress.progressBarRepaint("Renaming instances", oParent.Products.Count, i)
        
        Set oChild = oParent.Products.Item(i)
        If FindInstanceType(oChild) <> "Component" Then
        
            'Put in default mode if required
            sPartNumber = ""
            On Error Resume Next
            sPartNumber = oChild.PartNumber
            On Error GoTo 0
            
            If sPartNumber = "" Then oChild.ApplyWorkMode DEFAULT_MODE

            If Not (oChild.Name Like oChild.PartNumber & "*") Then
            
                
                'Rename
                Call RenameChildInstance(oChild, oParent, sParentType)
                sChildNames(i - 1) = oChild.Name 'Update the sChildNames with the new instance name
            End If
        End If
    Next
    
    If sParentType = "Component" Then
        sAnswer = MsgBox("Selected parent is a component. You may" & vbCrLf & "need to run the BAReorderStdParts macro.", vbInformation)
    End If

endsub:
    'Log File
    Call AddToLogFile("Rename Instances", , oParent.Name)
    Call frmKBEMain.resetToolbar
    frmProgress.Hide

End Sub


'This is the entry point when the coming from: Replace Multiple Components
Public Sub RenameSelectedInstance(ByVal oChild As Product)

Dim oParent As Product
Dim oItem As Product

'Retrieve parent
Set oParent = oChild.Parent

'Retrieve the type of the parent assy
sParentType = FindInstanceType(oParent)

'Retrieve the name of all the child and transfert them in the array sChildNames
ReDim sChildNames(0)
For i = 1 To oParent.Products.Count
    Set oItem = oParent.Products.Item(i)
    
    If sChildNames(UBound(sChildNames)) = "" Then
        sChildNames(UBound(sChildNames)) = oItem.Name
    Else
        ReDim Preserve sChildNames(UBound(sChildNames) + 1)
        sChildNames(UBound(sChildNames)) = oItem.Name
    End If
Next

'Rename
Call RenameChildInstance(oChild, oParent, sParentType)

End Sub

Private Function FindInstanceType(ByVal oProduct As Product) As String

Dim sType As String
Dim sRootDocumentName As String
Dim sGrandParentName As String


'When coming from the macro Replace Multiple Components oRootDocument is not valuated.
'So I use the Active window to retrieve the root document
If oRootDocument Is Nothing Then
    sRootDocumentName = CATIA.ActiveWindow.Parent.Name
Else
    sRootDocumentName = oRootDocument.Name
End If

'Find Grand Parent Name
sGrandParentName = ""
On Error Resume Next
sGrandParentName = oProduct.Parent.Parent.ReferenceProduct.Parent.Name
On Error GoTo 0


'Find Instance Type
If oProduct.HasAMasterShapeRepresentation Then
    FindInstanceType = "CATPart"
ElseIf oProduct.ReferenceProduct.Parent.Name = sGrandParentName Then
    FindInstanceType = "Component"
ElseIf oProduct.ReferenceProduct.Parent.Name = sRootDocumentName Then
    FindInstanceType = "CATProduct"
Else
    FindInstanceType = "CATProduct"
End If

End Function

Private Sub RenameChildInstance(ByVal oChild As Product, ByVal oParent As Product, ByVal sParentType As String)

Dim sParentDocName As String
Dim sNewInstanceName As String

'We retreive the document name of the parent
sParentDocName = oParent.ReferenceProduct.Parent.Name

'Define the new instance name
sNewInstanceName = DefineNewInstanceName(oChild, oParent)

'We change the instance name
If sParentType = "CATProduct" Then
    CATIA.Documents.Item(sParentDocName).Product.ReferenceProduct.Products.Item(oChild.Name).Name = sNewInstanceName
Else
   CATIA.Documents.Item(sParentDocName).Product.ReferenceProduct.Products.Item(oParent.Name).ReferenceProduct.Products.Item(oChild.Name).Name = sNewInstanceName
End If


End Sub

Private Function DefineNewInstanceName(oChild As Product, oParent As Product) As String

Dim i As Integer
Dim sNewInstanceName As String
Dim iInstanceNumber As Integer
Dim bInstanceFound As Boolean

iInstanceNumber = 0
bInstanceFound = True

'Transfer all the values of the array sChildNames to a string
sChildNameString = "|" & Join(sChildNames, "|") & "|"

Do While bInstanceFound = True

    'We set the new instance name to: Part Number + iInstanceNumber (where iInstanceNumber start at 1)
    iInstanceNumber = iInstanceNumber + 1
    bInstanceFound = False
    sNewInstanceName = oChild.PartNumber & "." & iInstanceNumber
    
    'Check in sChildNameString to see if sNewInstanceName already exist
    If sChildNameString Like "*|" & sNewInstanceName & "|*" Then bInstanceFound = True
        
Loop

DefineNewInstanceName = sNewInstanceName

End Function











