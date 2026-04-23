Attribute VB_Name = "SSParameter"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oCurrentSelection As Collection
Private oSubset As Collection
Private oParameterObject As StrParam
Private sParameterName As String

'********************************************************************************
'* Name: SS Parameter
'* Purpose: This class creates an object "SS Parameter" that has following properties
'*          1- CurrentSelectedItems: Collection of strings that represents the selected objects to add to SS Parameter value
'*          2- SubsetValues: Collection of strings that represents the different values of the SS Parameter
'*          3- ParameterObject: This is the parameter itself. We can add it to selection and highlight it in tree
'*          4- ParameterName: Parameter name without the path, and without the prefix "SS_"
'*
'*          An object SSParameter is used to fill frmSubsetsManagement. The object SSParameter allows to regroup together all information
'*          related to a Subset Parameter. In frmSubsetsManagement, we create a collection "SSParameters", that is a collection of
'*          many different "SSParameter" objects.
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub Class_Initialize()
    Set oCurrentSelection = New Collection
    Set oSubset = New Collection
End Sub


'********************************************************************************
'* Name: Add to selected items
'* Purpose: Add elements to collection SelectedItems, and optionally remove them from collection SubsetValues
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************

Public Sub AddToSelectedItems(ByVal sSelectionName As String, Optional ByVal bRemoveFromSubset As Boolean = True)
    On Error Resume Next
    If bRemoveFromSubset = True Then
        oCurrentSelection.Add sSelectionName, sSelectionName
        oSubset.Remove sSelectionName
    Else 'Add to collection Selected Items only if it does not already exist in Subset
        If oSubset.Item(sSelectionName) = Empty Then oCurrentSelection.Add sSelectionName, sSelectionName
    End If
    On Error GoTo 0
End Sub


'********************************************************************************
'* Name: Add to Subset values
'* Purpose: Add elements to collection SubsetValues, and remove them from collection SelectSelectedItemsionSet
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'********************************************************************************

Public Sub AddToSubsetValues(ByVal sSelectionName As String)
    On Error Resume Next
    oSubset.Add sSelectionName, sSelectionName
    oCurrentSelection.Remove sSelectionName
    On Error GoTo 0
End Sub



'*************************************************************
'************* SS Parameter properties ***********************
'*************************************************************
Public Property Get ParameterName() As String
    ParameterName = sParameterName
End Property
Public Property Let ParameterName(sStrParamName As String)
    sParameterName = sStrParamName
End Property


Public Property Get ParameterObject() As StrParam
    Set ParameterObject = oParameterObject
End Property
Public Property Let ParameterObject(oStrParam As StrParam)
    Set oParameterObject = oStrParam
End Property


Public Property Get CurrentSelectedItems() As Collection
    Set CurrentSelectedItems = oCurrentSelection
End Property
Public Property Let CurrentSelectedItems(oCollection As Collection)
    Set oCurrentSelection = oCollection
End Property


Public Property Get SubsetValues() As Collection
    Set SubsetValues = oSubset
End Property
Public Property Let SubsetValues(oCollection As Collection)
    Set oSubset = oCollection
End Property
