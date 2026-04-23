Attribute VB_Name = "BA_KBE_GCC_CHECK"
'********************************************************************************
'               SCRIPT VERSION
'********************************************************************************
Public Const sScriptVersion As String = "BA_KBE_GCC_CHECK_OCT14_2020"
'********************************************************************************

Public sLogFileName As String
Public sLog As String
Public sParentList As String
Public bCancelAction As Boolean '***** Used to determine if user pressed Cancel in a progress bar
Public oSelection As Selection
Public oSelectedItem As AnyObject
Public oSelectedItems As Collection
Public oSeedAssy As Product
Public sSeedAssyToolVersion As String
Public oSeedAssyToolVersion As StrParam
Public sSeedAssyConfigName As String
Public oSeedAssyConfigName As StrParam
Public sSourceWindow As Window
Public sTargetWindow As Window
Public oSeedPart As Part
Public sSeedPartName As String
Public oSeedPartInstance As Product
Public oSeedPartRootDocument As Document
Public oSelectionFromSeedPart As Selection
Public oControl As Control
Public sSeedPartsList As String
Public sPreviousWorkbench As String

Private i As Long
Private j As Long

'********************************************************************************
'               NEEDED URLS
'********************************************************************************
Public sActiveToolbarModule As String   '***** Active Toolbar Module
Public sBNetPageURL As String           '***** BNet KBE Portal address
Public sInfoPageURL As String           '***** KBE Info page address
Public sHelpPageURL As String           '***** Help file path
Public sToolVersionFile As String       '***** Tool Version Index file path

'********************************************************************************
'               CONSTANT VARIABLES
'********************************************************************************
Public Const sNoSelection As String = "No Selection"

Public Sub CATMain(ByVal sKBEFilePathInput As String)
    
    'Initialize
    sKBEPathFile = sKBEFilePathInput
    sActiveToolbarModule = "BA_KBE_GCC_CHECK"
    
    Call setCommonVariables
    
    '***** Get values from sKBEPathFile
    sToolVersionFile = GetTagValueFromFile(sKBEPathFile, "Tool Version Index")                  '***** Tool Version Index file path
    
    '***** Update CATSettings - Make sure relations update and are synchronous, no name check
    Call enableRelationUpdate
    Call disableOptionNameCheck
    
    '***** Load Toolbar
    frmKBEMain.Show vbModeless
    
End Sub


