Attribute VB_Name = "BA_KBE_GCC_PROD"
'********************************************************************************
'               SCRIPT VERSION
'********************************************************************************
Public Const sScriptVersion As String = "BA_KBE_GCC_PROD_OCT14_2020"
'********************************************************************************

Public sLogFileName As String
Public sLog As String
Public sLastFolderSelected As String
Public sParentList As String
Public bCancelAction As Boolean             '***** Used to determine if user pressed Cancel in a progress bar
Public oSelection As Selection
Public oSelectedItem As AnyObject
Public oSelectedItems As Collection
Public oPart As Part
Public oProduct As Product
Public oDocument As Document
Public oSeedAssy As Product
Public sSeedAssyName As String
Public sSeedAssyTitle As String
Public sSeedAssyToolVersion As String
Public oSeedAssyToolVersion As StrParam
Public sSeedAssyMigratedVersion As String
Public sSeedAssyConfigName As String
Public oSeedAssyConfigName As StrParam
Public oSeedAssyMigrateConfigFile As Boolean
Public sSourceWindow As Window
Public sTargetWindow As Window
Public oSeedConfigPart As Product
Public oMonumentOriginPart As Product
Public oSeedPart As Part
Public sSeedPartName As String
Public sSeedPartLetter As String
Public oSeedPartInstance As Product
Public oSeedPartRootDocument As Document
Public oSelectionFromSeedPart As Selection
Public oSeedLingPart As Part
Public sSeedLingPartName As String
Public sSeedLingLetter As String
Public oSeedLingInstance As Product
Public oSeedLingRootDocument As Document
Public oSelectionFromSeedLing As Selection
Public oSeedLingAxisSystem As AxisSystem
Public sSubSetItem As Variant
Public iMaxRangeKBEBodies As Long
Public iBdyCntIndex As Long
Public iKBEBodyCntIndex As Long
Public oSPCollInstance As Product
Public oSPCollRootDocument As Document
Public oSelectionFromSPColl As Selection
Public oNHAAssy As Product
Public sNHAAssyProducts As String
Public oCurrentSSParameter As StrParam
Public sCurrentSSParameterName As String
Public oReference As Reference
Public oParameter As Parameter
Public sParameterName As String
Public oParameters As Parameters
Public oParameterSet As ParameterSet
Public oRelationSet As KnowledgeObject
Public oTempParameter As Parameter
Public oRelation As Relation
Public sRelationText As String
Public oFormula As Formula
Public oRule As Rule
Public oBody As Body
Public oBodyMain As Body
Public oAssemble As Assemble
Public oExtract As HybridShapeExtract
Public oConstraint As Constraint
Public oConstraints As Constraints
Public oPublication As Publication
Public oPublications As Publications
Public oGeometricalSet As HybridBody
Public oSubGeometricalSet As HybridBody
Public sSeedPartsList As String
Public sStdPartsList As String
Public sXMLSeedPartsList As String
Public sCollectorID As String
Public sProgBarComment As String
Public sProgBarCommentMain As String
Public sTimer As String
Public sPreviousWorkbench As String
Public iObject As Variant
Public timeStart As Double
Public timeEnd As Double
Public iProgressCount1 As Double
Public iProgressCount1Max As Double
Public iProgressCount2 As Double
Public iProgressCount2Max As Double
Public bCopySelection As Boolean
Public oControl As Control
Public oSeedPartCollection As Collection
Public SSParameters As Collection
Public aCopyParams As Collection
Public aPasteParams As Collection
Public oProdPosition As Object
Public wSPAWorkbench As SPAWorkbench
Public ParamFSO As FileSystemObject
Public oTemplateData As New BATemplate
Public oInstanceActivation As New InstanceActivation
Public oReusePanels As New ReusePanels
Public SelectedSSParameters As Collection

Private i As Long
Private j As Long
Private k As Long

'********************************************************************************
'               NEEDED URLS
'********************************************************************************
'Public sActiveToolbarPath As String     '***** Active Toolbar Path
Public sActiveToolbarModule As String   '***** Active Toolbar Module
Public sManufacturedPartsFile As String '***** ManufacturedLCAParts file path

'********************************************************************************
'               CONSTANT VARIABLES
'********************************************************************************
Public Const sNoSelection As String = "No Selection"
Public Const dDeltaPosition As Double = 0.00001

Public Sub CATMain(ByVal sKBEFilePathInput As String)
    
    Dim sUserConnectionLogFile As String
    
    'Initialize
    sKBEPathFile = sKBEFilePathInput
    sActiveToolbarModule = "BA_KBE_GCC_PROD"
    
    Call setCommonVariables
    
    '***** Set all needed URLs (Read txt file)
    sManufacturedPartsFile = GetTagValueFromFile(sKBEPathFile, "Manufactured LCA Parts")        '***** Manufactured LCA Parts
    sUserConnectionLogFile = GetTagValueFromFile(sKBEPathFile, "User Tracking File")
    
    '***** Update CATSettings - Make sure relations update and are synchronous
    Call enableRelationUpdate
    
    '***** Log user connection
    On Error Resume Next
    Dim sUserName As String * 25
    If frmKBEMain.Visible = False Then
        GetUserNameA sUserName, 25
        sLog = UCase(Left(sUserName, InStr(sUserName, Chr(0)) - 1)) & Chr(9) & Format(Date, "Short Date") & Chr(9) & Format(Time, "Long Time") & Chr(9) & "Version: " & sScriptVersion
        Call WriteTextFile(sUserConnectionLogFile, sLog, False)
    End If
    On Error GoTo 0

    
    '***** Load Toolbar
    frmKBEMain.Show vbModeless
    
End Sub


