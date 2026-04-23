Attribute VB_Name = "frmManageData"
Attribute VB_Base = "0{BC969AFA-5352-4A11-A16E-6E65A222CE2F}{CED6F3C0-3192-4B7C-9DB6-1958337844B8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'====================================================================='
'==                    Manage Configuration Data                    =='
'====================================================================='
'== Author:        Christian Isaacs | Rand Worldwide | RPS          =='
'== Author:2       Loganathan | Satyam India                        =='
'== Date:          July 10, 2007                                    =='
'== Updated by:    Julien Bigaouette / François Charette            =='
'== Date:          July 2013                                        =='
'== Version:                                                        =='
'== Application:   Extracts Parameters that contain "DT" & writes   =='
'==                their name and value to an XML File              =='
'==                Imports Parameters from an XML File              =='
'====================================================================='
'====================================================================='

'Private VB Decs
Private i As Long
Private j As Long

Private outputDir As String
Private inputFile As String

Private oSelectedSource As AnyObject
Private oSourceBody As AnyObject
Private oSelectedSourcePart As Part
Private oSelectedTarget As AnyObject
Private oTargetBody As AnyObject
Private oSelectedTargetPart As Part

Private Import_ConfigFile As String
Private Import_ConfigName As String

Private Export_NewXML_ConfigFileDir As String
Private Export_NewXML_ConfigFileName As String
Private Export_NewXML_ConfigName As String
Private Export_NewXML_ConfigDesc As String

Private Export_AddNewConfig_ConfigFileName As String
Private Export_AddNewConfig_ConfigName As String

Private Export_ReleaseConfig_ConfigFileName As String
Private Export_ReleaseConfig_ConfigName As String
Private Export_ReleaseConfig_ConfigFileDir As String
Private Export_ReleaseConfig_NewConfigFileName As String
Private Export_ReleaseConfig_NewConfigName As String
Private Export_ReleaseConfig_NewConfigDesc As String

'Public XML Declares
Public xmlDoc As DOMDocument50
Public xmlRoot As IXMLDOMElement
Public xmlNode As IXMLDOMNode
Public xmlDoc2 As DOMDocument50
Public xmlRoot2 As IXMLDOMElement
Public xmlNode2 As IXMLDOMNode
Public LatestConfig As IXMLDOMNode
Public configData As New XMLConfiguration
Public unrelatedXMLFileMonumentID As String

'Private constants
Private Const sSelectFile As String = "Select File..."
Private Const sSelectDirectory As String = "Select Directory..."
Private Const sEnterConfigName As String = "Enter Configuration Name..."
Private Const sEnterFileName As String = "Enter File Name..."
Private Const sEnterDescription As String = "Enter Description..."
Private Const sSelectSource As String = "Select Source Node..."
Private Const sSelectTarget As String = "Select Target Node..."

Public Sub initializeXMLForm()
    
    '***** Window size according to number of shifters
    With Me
        .StartUpPosition = Manual
        .Height = dTitleBarHeight + fraMain.Height + 31
        .Width = 224
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 190
        .fraExportOptions.Visible = False
        .fraNewConfig.Visible = False
        .fraExistingConfig.Visible = False
        .Caption = "Manage Configuration / Parameters"
        
        '***** Default values
        .txtSelectionSA.Text = sNoSelection
        .optImport.Value = True
        Call SetFormMode
        Call resetTextBoxes
        Call cmdOKActivation
        .Show vbModeless
        
    End With

End Sub

Private Sub cmdSelectSA_Click()
    
    Me.Hide
    Me.txtSelectionSA.Text = sNoSelection
    Set oSeedPartCollection = New Collection
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
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy, Me.Left + 15, Me.Top + Me.cmdSelectSA.Top)

    If sSeedAssyToolVersion <> "" And sSeedAssyConfigName <> "" And sSeedPartsList <> "" Then
        Me.txtSelectionSA.Text = oSeedAssy.PartNumber
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
        
        'Get ConfigFilePath User Defined Property
        Set oTempStrParam = Nothing
        Set oTempStrParam = oSeedAssy.UserRefProperties.Item("ConfigurationFilePath")
        If oTempStrParam Is Nothing Then
            oSeedAssy.UserRefProperties.CreateString "ConfigurationFilePath", ""
        End If
        
        'Get ConfigID User Defined Property, if it does not exist create it!
        Set oTempStrParam = Nothing
        Set oTempStrParam = oSeedAssy.UserRefProperties.Item("ConfigurationID")
        If oTempStrParam Is Nothing Then
            oSeedAssy.UserRefProperties.CreateString "ConfigurationID", ""
        End If
        
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        
    ElseIf oSeedAssyToolVersion Is Nothing And sSeedPartsList <> "" Then
        MsgBox "`ToolVersion_DT` parameter not found!      ", vbCritical, "Warning"
        GoTo noSelection
    ElseIf sSeedAssyToolVersion = "" And sSeedPartsList <> "" Then
        MsgBox "`ToolVersion_DT` parameter not set!      " & vbCrLf & vbCrLf & _
                " Example: `ToolVersion_DT` = VXX      ", vbCritical, "Warning"
        GoTo noSelection
    ElseIf oSeedAssyConfigName Is Nothing And sSeedPartsList <> "" Then
        MsgBox "`ConfigurationName_DT` parameter not found!     ", vbCritical, "Warning"
        GoTo noSelection
    ElseIf sSeedAssyConfigName = "" And sSeedPartsList <> "" Then
        MsgBox "`ConfigurationName_DT` parameter not set!      " & vbCrLf & vbCrLf & _
                " Example: `ConfigurationName_DT` = TX:Config.1      ", vbCritical, "Warning"
        GoTo noSelection
    Else
noSelection:
        Me.txtSelectionSA.Text = sNoSelection
        Set oSeedAssy = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set sSourceWindow = Nothing
    End If
    
    Call cmbConfigName_Change
    If bToolbarCommandSelected = False Then Me.Show vbModeless
    
End Sub

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
    
    CATIA.DisplayFileAlerts = False
'    CATIA.Interactive = False 'Do not allow user to click something
'    frmKBEMain.Enabled = False 'Do not allow user to click something
    Set ParamFSO = CreateObject("Scripting.FileSystemObject")
    
    Import_ConfigFile = ""
    Import_ConfigName = ""
    Import_ConfigFile2 = ""
    Import_ConfigName2 = ""
    
    Export_NewXML_ConfigFileDir = ""
    Export_NewXML_ConfigFileName = ""
    Export_NewXML_ConfigName = ""
    Export_NewXML_ConfigDesc = ""
    
    Export_AddNewConfig_ConfigFileName = ""
    Export_AddNewConfig_ConfigName = ""
    Export_AddNewConfig_ConfigDesc = ""
    
    ReleaseConfig_ConfigFileName = ""
    ReleaseConfig_ConfigName = ""
    ReleaseConfig_ConfigFileDir = ""
    ReleaseConfig_NewConfigFileName = ""
    ReleaseConfig_NewConfigName = ""
    ReleaseConfig_NewConfigDesc = ""
    
    CPSL_Source = ""
    CPSL_Target = ""
    CPSL_ConfigFileDir = ""
    CPSL_NewConfigFileName = ""
    CPSL_NewConfigName = ""
    CPSL_NewConfigDesc = ""
    CPSL_AddNewConfig_ConfigFileName = ""
    CPSL_AddNewConfig_ConfigName = ""
    CPSL_AddNewConfig_ConfigDesc = ""
    CPSL_ConfigFile = ""
    CPSL_ConfigName = ""
    
    If optImport.Value = True Or optCompare.Value = True Then
        Import_ConfigFile = txtConfigFile.Text
        Import_ConfigName = cmbConfigName.Text
        Import_ConfigFile2 = txtConfigFile2.Text
        Import_ConfigName2 = cmbConfigName2.Text
    ElseIf optCPSL.Value = True Then
        If optCopyPaste.Value = True Then
            CPSL_Source = txtSource.Text
            CPSL_Target = txtTarget.Text
        ElseIf optCopySave.Value = True Then
        
            CPSL_Source = txtSource.Text
            
            If optNewXML.Value = True Then
                CPSL_ConfigFileDir = txtConfigFileDir.Text
                CPSL_NewConfigFileName = txtNewConfigFileName.Text
                CPSL_NewConfigName = txtNewConfigName.Text
                CPSL_NewConfigDesc = txtNewConfigDesc.Text
            ElseIf optAddNewConfig.Value = True Then
                CPSL_AddNewConfig_ConfigFileName = txtConfigFile.Text
                CPSL_AddNewConfig_ConfigName = txtConfigName.Text
                CPSL_AddNewConfig_ConfigDesc = txtConfigDesc.Text
            End If
        Else
            CPSL_ConfigFile = txtConfigFile.Text
            CPSL_ConfigName = cmbConfigName.Text
        End If
    ElseIf optReleaseConfig.Value = True Then
        ReleaseConfig_ConfigFileName = txtConfigFile.Text
        ReleaseConfig_ConfigName = cmbConfigName.Text
        ReleaseConfig_ConfigFileDir = txtConfigFileDir.Text
        ReleaseConfig_NewConfigFileName = txtNewConfigFileName.Text
        ReleaseConfig_NewConfigName = txtNewConfigName.Text
        ReleaseConfig_NewConfigDesc = txtNewConfigDesc.Text
    Else
        If optNewXML.Value = True Then
            Export_NewXML_ConfigFileDir = txtConfigFileDir.Text
            Export_NewXML_ConfigFileName = txtNewConfigFileName.Text
            Export_NewXML_ConfigName = txtNewConfigName.Text
            Export_NewXML_ConfigDesc = txtNewConfigDesc.Text
        ElseIf optAddNewConfig.Value = True Then
            Export_AddNewConfig_ConfigFileName = txtConfigFile.Text
            Export_AddNewConfig_ConfigName = txtConfigName.Text
            Export_AddNewConfig_ConfigDesc = txtConfigDesc.Text
        End If
    End If
    
    'Validate the Required Inputs
    If optImport.Value = True Then
        
        If ParamFSO.FileExists(Import_ConfigFile) = False Then
            MsgBox ("Invalid Configuration file Name!  Please Try Again!")
            Exit Sub
        End If
    
        If Import_ConfigName = "" Then
            MsgBox ("Invalid Configuration Name!  Please Try Again!")
            Exit Sub
        End If
    
        Me.Hide
        Call configData.ImportData(Import_ConfigFile, Import_ConfigName)
    
    ElseIf optExport.Value = True Then
        
        If optAddNewConfig.Value = True Then
            If ParamFSO.FileExists(Export_AddNewConfig_ConfigFileName) = False Then
                MsgBox ("Invalid Configuration file Name!  Please Try Again!")
                Exit Sub
            End If
        
            If Export_AddNewConfig_ConfigName = "" Then
                MsgBox ("Invalid Configuration Name!  Please Try Again!")
                Exit Sub
            End If
            
            Me.Hide
            Call configData.ExportData(Export_AddNewConfig_ConfigFileName, Export_AddNewConfig_ConfigName, Export_AddNewConfig_ConfigDesc, sSeedAssyToolVersion, "Add")
            
        ElseIf optNewXML.Value = True Then
            
            If ParamFSO.FolderExists(Export_NewXML_ConfigFileDir) = False Then
                MsgBox ("Invalid Output Directory!  Please Try Again!")
                Exit Sub
            End If
        
            If Export_NewXML_ConfigFileName = "" Then
                MsgBox ("Invalid File Name!  Please Try Again!")
                Exit Sub
            End If
            
            If Export_NewXML_ConfigName = "" Then
                MsgBox ("Invalid Configuration Name!  Please Try Again!")
                Exit Sub
            End If
            
            Me.Hide
            Call configData.ExportData(Export_NewXML_ConfigFileDir & "\" & Export_NewXML_ConfigFileName & ".xml", Export_NewXML_ConfigName, Export_NewXML_ConfigDesc, sSeedAssyToolVersion, "New")
            
        End If
    ElseIf optReleaseConfig.Value = True Then
        
        If ParamFSO.FileExists(ReleaseConfig_ConfigFileName) = False Then
            MsgBox ("Invalid Configuration file Name!  Please Try Again!")
            Exit Sub
        End If
        
        If ReleaseConfig_ConfigName = "" Then
            MsgBox ("Invalid Configuration Name!  Please Try Again!")
            Exit Sub
        End If
        
        If ParamFSO.FolderExists(ReleaseConfig_ConfigFileDir) = False Then
            MsgBox ("Invalid Output Directory!  Please Try Again!")
            Exit Sub
        End If
        
        If ReleaseConfig_NewConfigFileName = "" Then
            MsgBox ("Invalid File Name!  Please Try Again!")
            Exit Sub
        End If
        
        If ReleaseConfig_NewConfigName = "" Then
            MsgBox ("Invalid Configuration Name!  Please Try Again!")
            Exit Sub
        End If
        
        Me.Hide
        Call configData.ReleaseData(ReleaseConfig_ConfigFileName, ReleaseConfig_ConfigName, ReleaseConfig_ConfigFileDir & "\" & ReleaseConfig_NewConfigFileName & ".xml", ReleaseConfig_NewConfigName, ReleaseConfig_NewConfigDesc)
        
    ElseIf optCompare.Value = True Then
        
        If ParamFSO.FileExists(Import_ConfigFile) = False Then
            MsgBox ("Invalid Configuration file Name!  Please Try Again!")
            Exit Sub
        End If
        
        If Import_ConfigName = "" Then
            MsgBox ("Invalid Configuration Name!  Please Try Again!")
            Exit Sub
        End If
        
        If optCompareXML.Value = True Then
            If ParamFSO.FileExists(Import_ConfigFile2) = False Then
                MsgBox ("Invalid Configuration file Name!  Please Try Again!")
                Exit Sub
            End If
            
            Me.Hide
            Call configData.CompareData(Import_ConfigFile, Import_ConfigName, Import_ConfigFile2, Import_ConfigName2)
        Else
            Me.Hide
            Call configData.CompareData(Import_ConfigFile, Import_ConfigName)
        End If
        
    ElseIf optCPSL.Value = True Then
        
        If optCopyPaste.Value = True Then
            
            If CPSL_Source = "" Then
                MsgBox ("Invalid Source!  Please Try Again!")
                Exit Sub
            End If
            
            If CPSL_Target = "" Then
                MsgBox ("Invalid Target!  Please Try Again!")
                Exit Sub
            End If
            Me.Hide
            Call configData.CopyPasteParam(oSelectedSource, oSelectedTarget, oSelectedSourcePart, oSelectedTargetPart)
            
        ElseIf optCopySave.Value = True Then
            
            If CPSL_Source = "" Then
                MsgBox ("Invalid Source!  Please Try Again!")
                Exit Sub
            End If
            
            If optNewXML.Value = True Then
                If ParamFSO.FolderExists(CPSL_ConfigFileDir) = False Then
                    MsgBox ("Invalid Output Directory!  Please Try Again!")
                    Exit Sub
                End If
                
                If CPSL_NewConfigFileName = "" Then
                    MsgBox ("Invalid File Name!  Please Try Again!")
                    Exit Sub
                End If
                
                If CPSL_NewConfigName = "" Then
                    MsgBox ("Invalid Configuration Name!  Please Try Again!")
                    Exit Sub
                End If
            
                Me.Hide
                Call configData.ExportData(CPSL_ConfigFileDir & "\" & CPSL_NewConfigFileName & ".xml", CPSL_NewConfigName, CPSL_NewConfigDesc, sSeedAssyToolVersion, "New", oSelectedSource, oSelectedSourcePart)
            
            ElseIf optAddNewConfig.Value = True Then
                If ParamFSO.FileExists(CPSL_AddNewConfig_ConfigFileName) = False Then
                    MsgBox ("Invalid Configuration file Name!  Please Try Again!")
                    Exit Sub
                End If
            
                If CPSL_AddNewConfig_ConfigName = "" Then
                    MsgBox ("Invalid Configuration Name!  Please Try Again!")
                    Exit Sub
                End If
                
                Me.Hide
                Call configData.ExportData(CPSL_AddNewConfig_ConfigFileName, CPSL_AddNewConfig_ConfigName, CPSL_AddNewConfig_ConfigDesc, sSeedAssyToolVersion, "Add", oSelectedSource, oSelectedSourcePart)
            Else
            End If
        
        Else
            If ParamFSO.FileExists(CPSL_ConfigFile) = False Then
                MsgBox ("Invalid Configuration file Name!  Please Try Again!")
                Exit Sub
            End If
        
            If CPSL_ConfigName = "" And Not txtTarget.Text = sSelectTarget Then
                MsgBox ("Invalid Configuration Name!  Please Try Again!")
                Exit Sub
            End If
        
            Me.Hide
            
            Call configData.LoadPaste(CPSL_ConfigFile, CPSL_ConfigName, oSelectedTarget, oSelectedTargetPart)
        
        End If
    
    End If
    
    frmProgress.Hide
    CATIA.DisplayFileAlerts = True
    CATIA.Interactive = True
    CATIA.RefreshDisplay = True
    
'    'Needed to "reactivate" Catia after CATIA.Interactive = True
'    On Error Resume Next
'    CATIA.ActiveWindow.ActivatePrevious
'    CATIA.ActiveWindow.ActivateNext
'    CATIA.StatusBar = "Successfully Done! "
'    On Error GoTo 0
    
    Call cmdCancel_Click
    
End Sub

Private Sub cmdSelectConfigFile_Click()
        
Dim Result As Variant
Me.txtConfigFile.Text = sSelectFile

'Get User Selected File
inputFile = ""
inputFile = OpenFileDialog("xml files (*.xml)|*.xml", False, "*.xml", , , , FindWindow(0&, Me.Caption))
If inputFile = "" Then
    Me.txtConfigFile.Text = sSelectFile
    cmbConfigName.Clear
    Exit Sub
Else
    Me.txtConfigFile.Text = inputFile
End If
On Error Resume Next

'Load XML file
Set xmlDoc = New DOMDocument50
xmlDoc.Load (inputFile)
Set xmlRoot = Nothing
Set xmlRoot = xmlDoc.documentElement

'Verify if the inputfile contain only data related to current monument.
If Not Me.optCPSL.Value Then
    If configData.IsValidMonumentID(xmlDoc) = False Then
        MsgBox "Configuration file contains data of unrelated monuments : " & unrelatedXMLFileMonumentID & vbCrLf & "Invalid file, please select another file.", vbInformation, "Select a Configuration File..."
        inputFile = ""
        Me.txtConfigFile.Text = sSelectFile
        Exit Sub
    End If
End If
'Verify if the Toolversion matches the Toolversion used to create the inputfile
If configData.IsValidToolVersion(xmlRoot) = False And Not Me.optCPSL.Value Then
    If sSeedAssyToolVersion = "" Then
        MsgBox ("The Current Tool Version is not set! Please Set the ToolVersion Parameter before proceeding! example: ToolVersion_DT = VXX")
        Exit Sub
    Else
        If Me.optExport = True And Me.optAddNewConfig = True Then
            MsgBox "The Version of the Configuration File does not match with the Version of the tool." & vbCrLf & vbCrLf & "Please select a 'New XML' option and proceed!", vbInformation, "Select a Configuration File..."
            inputFile = ""
            Me.txtConfigFile.Text = sSelectFile
            Exit Sub
        ElseIf Me.optReleaseConfig = False Then
            Result = MsgBox("The version of the Configuration File does not match the Version of the Seed Assy." & vbCrLf & "Would you like to migrate a configuration to the current tool level?", vbYesNo + vbInformation, "Select a Configuration File...")
            If Result = vbNo Then
                MsgBox "Please select a New Configuration File!", vbInformation, "Wrong Selection"
                inputFile = ""
                Me.txtConfigFile.Text = sSelectFile
                Exit Sub
            Else
                oSeedAssyMigrateConfigFile = True
            End If
        End If
    End If
Else
    'Load XML file
    oSeedAssyMigrateConfigFile = False
    Set xmlDoc = New DOMDocument50
    xmlDoc.Load (inputFile)
    Set xmlRoot = xmlDoc.documentElement
End If

i = 0
cmbConfigName.Clear
For Each xmlNode In xmlRoot.childNodes
    Call Me.cmbConfigName.AddItem(xmlNode.attributes.Item(0).Text)
    Set LatestConfig = xmlNode
    If Len(xmlNode.attributes.Item(0).Text) > i Then
        i = Len(xmlNode.attributes.Item(0).Text)
        Me.cmbConfigName.ListWidth = i / 0.1
    End If
Next
If Me.cmbConfigName.List(0) <> "" Then Me.cmbConfigName.Value = Me.cmbConfigName.List(0)

If Val(Me.cmbConfigName.ListWidth) < Me.cmbConfigName.Width Then
    Me.cmbConfigName.ListWidth = Me.cmbConfigName.Width
End If

If Me.optAddNewConfig.Value = True And (Not LatestConfig Is Nothing) Then
    If InStr(1, LatestConfig.attributes.Item(0).Text, ":Config.", vbTextCompare) > 0 Then
        Me.txtConfigName.Text = Split(LatestConfig.attributes.Item(0).Text, ":Config.")(0) & ":Config." & Val(Split(LatestConfig.attributes.Item(0).Text, ":Config.")(1)) + 1
    End If
End If

On Error GoTo 0

End Sub

Private Sub cmdSelectConfigFile2_Click()
    
Dim Result As Variant
Me.txtConfigFile2.Text = sSelectFile

'Get User Selected File
inputFile = ""
inputFile = OpenFileDialog("xml files (*.xml)|*.xml", False, "*.xml", , , , FindWindow(0&, Me.Caption))
If inputFile = "" Then
    Me.txtConfigFile2.Text = sSelectFile
    cmbConfigName2.Clear
    Exit Sub
Else
    Me.txtConfigFile2.Text = inputFile
End If

On Error Resume Next

'Load XML file
Set xmlDoc2 = New DOMDocument50
xmlDoc2.Load (inputFile)
Set xmlRoot2 = Nothing
Set xmlRoot2 = xmlDoc2.documentElement

'Verify if the inputfile contain only data related to current monument.
If Not Me.optCPSL.Value Then
    If configData.IsValidMonumentID(xmlDoc2) = False Then
        MsgBox "Configuration file contains data of unrelated monuments : " & unrelatedXMLFileMonumentID & vbCrLf & "Invalid file, please select another file.", vbInformation, "Select a Configuration File..."
        inputFile = ""
        Me.txtConfigFile2.Text = sSelectFile
        Exit Sub
    End If
End If
'Load XML file
oSeedAssyMigrateConfigFile = False

i = 0
cmbConfigName2.Clear
For Each xmlNode In xmlRoot2.childNodes
    Call Me.cmbConfigName2.AddItem(xmlNode.attributes.Item(0).Text)
    Set LatestConfig = xmlNode
    If Len(xmlNode.attributes.Item(0).Text) > i Then
        i = Len(xmlNode.attributes.Item(0).Text)
        Me.cmbConfigName2.ListWidth = i / 0.1
    End If
Next
If Me.cmbConfigName2.List(0) <> "" Then Me.cmbConfigName2.Value = Me.cmbConfigName2.List(0)

If Val(Me.cmbConfigName2.ListWidth) < Me.cmbConfigName2.Width Then
    Me.cmbConfigName2.ListWidth = Me.cmbConfigName2.Width
End If

On Error GoTo 0

End Sub

Private Sub cmdSelectConfigFileDir_Click()
    outputDir = OpenDirectoryDialog("Please select an Output Directory...", sLastFolderSelected)
    If outputDir = "" Then
        Me.txtConfigFileDir.Text = sSelectDirectory
        outputDir = ""
        Exit Sub
    Else
        Me.txtConfigFileDir.Text = outputDir
    End If
    AppActivate (Me.Caption)
End Sub

Private Sub cmdSelectSourceNode_Click()
    
    Me.Hide
    Me.txtSource.Text = sSelectSource
    
    On Error Resume Next
    Set oSelectedSource = Nothing
    Set oSelectedSource = MultiWindowSelectElement("Please select a Source node...")
    If TypeName(oSelectedSource.Document) = "ProductDocument" Then Set oSelectedSourcePart = oSelectedSource.LeafProduct.ReferenceProduct.Parent.Part
    If TypeName(oSelectedSource.Document) = "PartDocument" Then Set oSelectedSourcePart = oSelectedSource.Document.Part
    Set oSelectedSource = oSelectedSource.Value
    Set sSourceWindow = CATIA.ActiveWindow
    
''    To select the body of a tool, when a sub-element is selected
'    Set oBody = GetBodyFromObject(oSelectedSource)
'    If Not oBody Is Nothing Then
'        Set oSelectedSource = oBody
'    End If
    
    If Not oSelectedSource Is Nothing Then
        Me.txtSource.Text = Split(oSelectedSource.Name, "\")(UBound(Split(oSelectedSource.Name, "\")))
    Else
        Me.txtSource.Text = sSelectSource
    End If
    
    CATIA.ActiveDocument.Selection.Clear
    If Not bToolbarCommandSelected Then Me.Show vbModeless
    
End Sub

Private Sub cmdSelectTargetNode_Click()
    
    Me.Hide
    Me.txtTarget.Text = sSelectTarget
    
    On Error Resume Next
    Set oSelectedTarget = Nothing
    Set oSelectedTarget = MultiWindowSelectElement("Please select a Source node...")
    If TypeName(oSelectedTarget.Document) = "ProductDocument" Then Set oSelectedTargetPart = oSelectedTarget.LeafProduct.ReferenceProduct.Parent.Part
    If TypeName(oSelectedTarget.Document) = "PartDocument" Then Set oSelectedTargetPart = oSelectedTarget.Document.Part
    Set oSelectedTarget = oSelectedTarget.Value
    Set sTargetWindow = CATIA.ActiveWindow
    
''    To select the body of a tool, when a sub-element is selected
'    Set oBody = GetBodyFromObject(oSelectedTarget)
'    If Not oBody Is Nothing Then
'        Set oSelectedTarget = oBody
'    End If
    
    If Not oSelectedTarget Is Nothing Then
        Me.txtTarget.Text = Split(oSelectedTarget.Name, "\")(UBound(Split(oSelectedTarget.Name, "\")))
    Else
        Me.txtTarget.Text = sSelectTarget
    End If
    
    CATIA.ActiveDocument.Selection.Clear
    If Not bToolbarCommandSelected Then Me.Show vbModeless

End Sub

'***** Seed Assy Selection *****
Private Sub txtSelectionSA_Change()
    Call SetFormMode
    Call resetTextBoxes
    Call cmdOKActivation
End Sub

'***** Import Button *****
Private Sub optImport_Change()
    If Me.optImport.Value Then
        Call SetFormMode
        Call resetTextBoxes
    End If
End Sub
Private Sub ImgImportFalse_Click()
    Me.optImport.Value = True
End Sub

'***** Export Button *****
Private Sub OptExport_Change()
    If Me.optExport.Value Then
        Me.optAddNewConfig.Value = True 'Default: Add New Config To XML
        Call SetFormMode
        Call resetTextBoxes
    End If
End Sub
Private Sub ImgExportFalse_Click()
    Me.optExport.Value = True
End Sub

'***** Compare Button *****
Private Sub optCompare_Change()
    If Me.optCompare.Value Then
        Me.optCompareSA.Value = True 'Default: Against Seed Assy
        Call SetFormMode
        Call resetTextBoxes
    End If
End Sub
Private Sub ImgCompareFalse_Click()
    Me.optCompare.Value = True
End Sub

'***** Release Button *****
Private Sub optReleaseConfig_Change()
    If Me.optReleaseConfig.Value Then
        Call SetFormMode
        Call resetTextBoxes
    End If
End Sub
Private Sub ImgReleaseConfigFalse_Click()
    Me.optReleaseConfig.Value = True
End Sub

'***** Manage Parameters Button *****
Private Sub OptCPSL_Change()
    If Me.optCPSL.Value Then
        Me.optCopyPaste.Value = True 'Default: Copy/Paste
        Me.optAddNewConfig.Value = True 'Default: Add New Config To XML (2)
        Call SetFormMode
        Call resetTextBoxes
    End If
End Sub
Private Sub ImgCPSLFalse_Click()
    Me.optCPSL.Value = True
End Sub

'***** Create New XML Button *****
Private Sub optNewXML_Change()
    If Me.optNewXML.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgNewXMLFalse_Click()
    Me.optNewXML.Value = True
End Sub

'***** Add New Config to XML Button *****
Private Sub optAddNewConfig_Change()
    If Me.optAddNewConfig.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgAddNewConfigFalse_Click()
    Me.optAddNewConfig.Value = True
End Sub

'***** Against Seed Assy Button *****
Private Sub optCompareSA_Change()
    If Me.optCompareSA.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgCompareSAFalse_Click()
    Me.optCompareSA.Value = True
End Sub

'***** Against XML [Options] Button *****
Private Sub optCompareXML_Change()
    If Me.optCompareXML.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgCompareXMLFalse_Click()
    Me.optCompareXML.Value = True
End Sub

'***** Copy/Paste Button *****
Private Sub OptCopyPaste_Change()
    If Me.optCopyPaste.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgCopyPasteFalse_Click()
    Me.optCopyPaste.Value = True
End Sub

'***** Copy/Save Button *****
Private Sub OptCopySave_Change()
    If Me.optCopySave.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgCopySaveFalse_Click()
    Me.optCopySave.Value = True
End Sub

'***** Load/Paste Button *****
Private Sub OptLoadPaste_Change()
    If Me.optLoadPaste.Value Then
        Call SetFormMode
        Call cmdOKActivation
    End If
End Sub
Private Sub ImgLoadPasteFalse_Click()
    Me.optLoadPaste.Value = True
End Sub

Private Sub SetFormMode()
    
    With Me
        
        'Initializing
        .fraSeedAssy.Visible = False
        .fraExportOptions.Visible = False
        .fraExistingConfig.Visible = False
        .fraCompare.Visible = False
        .fraNewConfig.Visible = False
        .fraCPSLOptions.Visible = False
        .fraSourceTarget.Visible = False
        
        'Buttons Visibility
        .ImgImportTrue.Visible = .optImport.Value
        .ImgExportTrue.Visible = .optExport.Value
        .ImgCompareTrue.Visible = .optCompare.Value
        .ImgReleaseConfigTrue.Visible = .optReleaseConfig.Value
        .ImgCPSLTrue.Visible = .optCPSL.Value
        .ImgNewXMLTrue.Visible = .optNewXML.Value
        .ImgAddNewConfigTrue.Visible = .optAddNewConfig.Value
        .ImgCopyPasteTrue.Visible = .optCopyPaste.Value
        .ImgCopySaveTrue.Visible = .optCopySave.Value
        .ImgLoadPasteTrue.Visible = .optLoadPaste.Value
        .ImgCompareSATrue.Visible = .optCompareSA.Value
        .ImgCompareXMLTrue.Visible = .optCompareXML.Value
        
        .txtConfigDesc.BackColor = &H80000005
        .txtConfigDesc.Locked = False
        .fraMain.Top = 3.75
        
        .Show vbModeless
        
        '*** Import and Compare modes ***
        If .optImport.Value Or .optCompare.Value Then
            
            .fraSeedAssy.Visible = True
            
            If .txtSelectionSA.Text = sNoSelection Then
                .fraSeedAssy.Top = .fraMain.Top + .fraMain.Height
                .Height = dTitleBarHeight + .fraSeedAssy.Top + .fraSeedAssy.Height + 33
                GoTo endsub
            End If
            
            .fraExistingConfig.Visible = True
            .txtConfigName.Visible = False: .cmbConfigName.Visible = True
            
            If .optImport.Value Then
                .fraCompare.Visible = False
                .fraExistingConfig.Height = 70
                .txtConfigDesc.BackColor = &H80000016
                .txtConfigDesc.Locked = True
                
            ElseIf .optCompare.Value Then
                .fraCompare.Visible = True
                .fraExistingConfig.Height = 114
    
                If .optCompareSA.Value Then
                    .cmdSelectConfigFile2.Enabled = False
                    .lblConfigFileName2.Enabled = False
                    .lblConfigFile2.Enabled = False
                    .txtConfigFile2.Enabled = False
                    .cmbConfigName2.Enabled = False
                ElseIf .optCompareXML.Value Then
                    .cmdSelectConfigFile2.Enabled = True
                    .lblConfigFileName2.Enabled = True
                    .lblConfigFile2.Enabled = True
                    .txtConfigFile2.Enabled = True
                    .cmbConfigName2.Enabled = True
                End If
    
            End If
    
            .fraExistingConfig.Top = .fraSeedAssy.Top + .fraSeedAssy.Height + 4
            .Height = dTitleBarHeight + .fraExistingConfig.Top + .fraExistingConfig.Height + 33
            
            
        '*** Export mode ***
        ElseIf .optExport.Value Then
            
            .fraSeedAssy.Visible = True
            
            If .txtSelectionSA.Text = sNoSelection Then
            .fraSeedAssy.Top = .fraMain.Top + .fraMain.Height
            .Height = dTitleBarHeight + .fraSeedAssy.Top + .fraSeedAssy.Height + 33
            GoTo endsub
            End If
            
            .fraExportOptions.Visible = True
            .fraExportOptions.Top = .fraSeedAssy.Top + .fraSeedAssy.Height + 4
            
            If .optNewXML.Value Then
                .fraNewConfig.Visible = True
                .fraNewConfig.Height = 90
                .fraNewConfig.Top = .fraExportOptions.Height + .fraExportOptions.Top + 4
                .Height = dTitleBarHeight + .fraNewConfig.Top + .fraNewConfig.Height + 33
            ElseIf .optAddNewConfig.Value Then
                .fraExistingConfig.Visible = True
                .txtConfigName.Visible = True: .cmbConfigName.Visible = False
                .fraExistingConfig.Height = 70
                .fraExistingConfig.Top = fraExportOptions.Height + fraExportOptions.Top + 4
                .Height = dTitleBarHeight + .fraExistingConfig.Top + .fraExistingConfig.Height + 33
            End If
            
            
        '*** Release mode ***
        ElseIf .optReleaseConfig.Value Then
            
            .fraSeedAssy.Visible = True
            
            If .txtSelectionSA.Text = sNoSelection Then
            .fraSeedAssy.Top = .fraMain.Top + .fraMain.Height
            .Height = dTitleBarHeight + .fraSeedAssy.Top + .fraSeedAssy.Height + 33
            GoTo endsub
            End If
            
            .fraExistingConfig.Visible = True
            .fraExistingConfig.Top = .fraSeedAssy.Height + .fraSeedAssy.Top + 4
            .txtConfigName.Visible = False: .cmbConfigName.Visible = True: .fraCompare.Visible = False
            .fraExistingConfig.Height = 70
            .txtConfigDesc.BackColor = &H80000016
            .txtConfigDesc.Locked = True
            
            .fraNewConfig.Visible = True
            .fraNewConfig.Top = .fraExistingConfig.Height + .fraExistingConfig.Top + 4
            .fraNewConfig.Height = 90
            
            .Height = dTitleBarHeight + .fraNewConfig.Top + .fraNewConfig.Height + 33
            
        
        '*** Manage mode
        ElseIf .optCPSL.Value Then
            
            .fraSourceTarget.Visible = True
            .fraCPSLOptions.Visible = True
            .fraCPSLOptions.Left = .fraExportOptions.Left
            .fraCPSLOptions.Top = .fraMain.Top + .fraMain.Height + 4

            If .optCopyPaste.Value Then
                
                .txtSource.Visible = True: .cmdSelectSourceNode.Visible = True: .lblSource.Visible = True
                .txtTarget.Visible = True: .cmdSelectTargetNode.Visible = True: .lblTarget.Visible = True
                .txtTarget.Top = .txtSource.Top + 20: .cmdSelectTargetNode.Top = .cmdSelectSourceNode.Top + 20: .lblTarget.Top = .lblSource.Top + 20
                .fraSourceTarget.Top = .fraCPSLOptions.Top + .fraCPSLOptions.Height + 4
                .fraSourceTarget.Height = 48
                .fraNewConfig.Visible = False
                .fraNewConfig.Height = 100
                .Height = dTitleBarHeight + .fraSourceTarget.Top + .fraSourceTarget.Height + 33
                
            ElseIf .optCopySave.Value Then
                .fraExportOptions.Visible = True
                .fraExportOptions.Top = .fraCPSLOptions.Top + .fraCPSLOptions.Height + 4
                .fraSourceTarget.Visible = True
                .fraSourceTarget.Top = .fraExportOptions.Top + .fraExportOptions.Height + 4
                .fraSourceTarget.Height = 28
                .txtSource.Visible = True: .cmdSelectSourceNode.Visible = True: .lblSource.Visible = True
                .txtTarget.Visible = False: .cmdSelectTargetNode.Visible = False: .lblTarget.Visible = False
                
                If .optNewXML.Value Then
                    .fraNewConfig.Visible = True
                    .fraNewConfig.Height = 68
                    .fraNewConfig.Top = .fraSourceTarget.Top + .fraSourceTarget.Height + 4
                    .Height = dTitleBarHeight + .fraNewConfig.Top + .fraNewConfig.Height + 33
                    
                    
                ElseIf .optAddNewConfig.Value Then
                    .fraExistingConfig.Visible = True
                    .fraExistingConfig.Height = 48
                    .fraExistingConfig.Top = .fraSourceTarget.Top + .fraSourceTarget.Height + 4
                    .txtConfigName.Visible = True: .cmbConfigName.Visible = False
                    .Height = dTitleBarHeight + .fraExistingConfig.Top + .fraExistingConfig.Height + 33
                    
                End If
                
            ElseIf .optLoadPaste.Value Then
                
                .txtSource.Visible = False: .cmdSelectSourceNode.Visible = False: .lblSource.Visible = False
                .txtTarget.Visible = True: .cmdSelectTargetNode.Visible = True: .lblTarget.Visible = True
                .txtTarget.Top = .txtSource.Top: .cmdSelectTargetNode.Top = .cmdSelectSourceNode.Top: .lblTarget.Top = .lblSource.Top
                .fraSourceTarget.Top = .fraCPSLOptions.Top + .fraCPSLOptions.Height + 4
                .fraSourceTarget.Height = 28
                .fraExistingConfig.Visible = True
                .fraExistingConfig.Top = .fraSourceTarget.Top + .fraSourceTarget.Height + 4
                .fraExistingConfig.Height = 48
                .txtConfigName.Visible = False: .cmbConfigName.Visible = True
                .Height = dTitleBarHeight + .fraExistingConfig.Top + .fraExistingConfig.Height + 33
                
            End If
        
        End If
   
endsub:
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3.75
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        
        If .Top > dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height Then .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 10
        If .Top < 0 Then .Top = 10
        
    End With

End Sub

Private Sub resetTextBoxes()
    With Me
    
        .txtConfigFile.Text = sSelectFile
        .txtConfigName.Text = ""
        If .txtConfigDesc.Locked = False Then .txtConfigDesc.Text = sEnterDescription
        If .txtConfigDesc.Locked = True Then .txtConfigDesc.Text = ""
        .cmbConfigName.Clear
        
        .txtConfigFile2.Text = sSelectFile
        .cmbConfigName2.Clear
        
        .txtConfigFileDir.Text = sSelectDirectory
        .txtNewConfigFileName.Text = sEnterFileName
        .txtNewConfigName.Text = ""
        .txtNewConfigDesc.Text = sEnterDescription
        
        .txtSource.Text = sSelectSource
        .txtTarget.Text = sSelectTarget
        
    End With
End Sub


'***********************
'Frame fraExistingConfig
Private Sub txtConfigFile_Change()
    If Me.txtConfigFile.Text = "" Or Me.txtConfigFile.Text = sSelectFile Then
        Me.txtConfigName = ""
    Else
        Me.txtConfigName.Text = Me.txtConfigFile.Text & ":Config.1"
    End If
    Call cmdOKActivation
End Sub
Private Sub txtConfigDesc_Change()
    Call cmdOKActivation
End Sub
Private Sub txtConfigDesc_Enter()
    If Me.txtConfigDesc.Text = sEnterDescription Then Me.txtConfigDesc.Text = ""
End Sub
Private Sub txtConfigDesc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtConfigDesc.Text = "" And Me.txtConfigDesc.Locked = False Then Me.txtConfigDesc.Text = sEnterDescription
End Sub
Private Sub cmbConfigName_Change()
    
    On Error Resume Next
    If cmbConfigName.Value <> "" And cmbConfigName.Visible = True Then
        
        If Not oSeedAssy Is Nothing Then
            ' Sortlisting the common Seed Parts in Seed Assy and XML
            sXMLSeedPartsList = ""
            For Each xmlNode In xmlRoot.childNodes
                If xmlNode.attributes.Item(0).Text = cmbConfigName.Text Then
                    For j = 0 To xmlNode.childNodes.Length - 1
                        If InStr(1, sSeedPartsList, "_" & Split(xmlNode.childNodes.Item(j).attributes.Item(0).nodeValue, "SEED_PART_")(1), vbTextCompare) > 0 Then
                            sXMLSeedPartsList = sXMLSeedPartsList & "SEED_PART_" & Split(xmlNode.childNodes.Item(j).attributes.Item(0).nodeValue, "SEED_PART_")(1) & ";"
                        End If
                    Next j
                End If
            Next
            
            If sXMLSeedPartsList = "" Then
                MsgBox "No Seed Part corresponding to this configuration was selected!   " & vbCrLf & _
                        "Make sure to select all Seed Parts when selecting Seed Assy!   " & vbCrLf & vbCrLf & _
                        "Note: Seed Parts must be in Design mode!    ", vbExclamation, "Warning"
            End If
        End If
        
        Set xmlNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & cmbConfigName.Value & Chr(34) & "]")
        txtConfigDesc.Text = xmlNode.childNodes.Item(0).childNodes.Item(2).Text
        
        If txtConfigDesc.Text = "" Then txtConfigDesc.Text = "No description found..."
    
    Else
        If txtConfigDesc.Locked = True Then txtConfigDesc.Text = ""
        If txtConfigDesc.Locked = False Then txtConfigDesc.Text = sEnterDescription
    End If
    On Error GoTo 0
        
    Me.Show vbModeless
    Call cmdOKActivation

End Sub
Private Sub txtConfigFile2_Change()
    Call cmdOKActivation
End Sub
Private Sub cmbConfigName2_Change()
    Call cmdOKActivation
End Sub

'******************
'Frame fraNewConfig
Private Sub txtConfigFileDir_Change()
    Call cmdOKActivation
End Sub

Private Sub txtNewConfigDesc_Change()
    Call cmdOKActivation
End Sub
Private Sub txtNewConfigFileName_Change()
    If Me.txtNewConfigFileName.Text = "" Or Me.txtNewConfigFileName.Text = sEnterFileName Then
        Me.txtNewConfigName = ""
    Else
        Me.txtNewConfigName.Text = Me.txtNewConfigFileName.Text & ":Config.1"
    End If
    Call cmdOKActivation
End Sub
Private Sub txtNewConfigFileName_Enter()
    If Me.txtNewConfigFileName.Text = sEnterFileName Then Me.txtNewConfigFileName.Text = ""
End Sub
Private Sub txtNewConfigFileName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtNewConfigFileName.Text = "" Then Me.txtNewConfigFileName.Text = sEnterFileName
End Sub
Private Sub txtNewConfigDesc_Enter()
    If Me.txtNewConfigDesc.Text = sEnterDescription Then Me.txtNewConfigDesc.Text = ""
End Sub
Private Sub txtNewConfigDesc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtNewConfigDesc.Text = "" Then Me.txtNewConfigDesc.Text = sEnterDescription
End Sub

'**********************
'Select Source & Target
Private Sub txtSource_Change()
    Call cmdOKActivation
End Sub
Private Sub txtTarget_Change()
    Call cmdOKActivation
End Sub

'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if form is not complete
'*          Otherwise, set the focus on OK button
'********************************************************************************
Private Sub cmdOKActivation()
    
    On Error Resume Next
    With Me
        
        '*** No Seed Assy Selected
        If .txtSelectionSA.Text = sNoSelection And (.optImport.Value Or .optCompare.Value Or .optExport.Value Or .optReleaseConfig.Value) Then
            GoTo activateCANCEL
            
        '*** Import and Compare modes ***
        ElseIf (.optImport.Value Or .optCompare.Value) And txtConfigFile.Text <> "" And txtConfigFile.Text <> sSelectFile And cmbConfigName.Text <> "" Then
            
            If .optImport.Value And txtConfigDesc.Text <> "" And txtConfigDesc.Text <> sEnterDescription Then
                GoTo activateOK
            ElseIf .optCompare.Value Then
                If .optCompareSA.Value Or (.optCompareXML.Value And txtConfigFile2.Text <> sSelectFile And cmbConfigName2.Text <> "") Then
                    GoTo activateOK
                Else
                    GoTo activateCANCEL
                End If
            Else
                GoTo activateCANCEL
            End If
            
        '*** Export mode ***
        ElseIf .optExport.Value Then
            If .optNewXML.Value And txtConfigFileDir.Text <> "" And txtConfigFileDir.Text <> sSelectDirectory And txtNewConfigFileName.Text <> "" And txtNewConfigFileName.Text <> sEnterFileName And txtNewConfigDesc.Text <> "" And txtNewConfigDesc.Text <> sEnterDescription Then
                GoTo activateOK
            ElseIf .optAddNewConfig.Value And txtConfigFile.Text <> sSelectFile And txtConfigName.Text <> "" And txtConfigName.Text <> sEnterConfigName And txtConfigDesc.Text <> "" And txtConfigDesc.Text <> sEnterDescription Then
                GoTo activateOK
            Else
                GoTo activateCANCEL
            End If
            
        '*** Release mode ***
        ElseIf .optReleaseConfig.Value And txtConfigFile.Text <> "" And txtConfigFile.Text <> sSelectFile And cmbConfigName.Text <> "" Then
            If txtConfigFileDir.Text <> "" And txtNewConfigFileName.Text <> "" And txtNewConfigFileName.Text <> sEnterFileName And txtNewConfigDesc.Text <> "" And txtNewConfigDesc.Text <> sEnterDescription Then
                GoTo activateOK
            Else
                GoTo activateCANCEL
            End If
        
        '*** Manage mode
        ElseIf .optCPSL.Value Then

            If .optCopyPaste.Value And txtSource.Text <> "" And txtSource.Text <> sSelectSource And txtTarget.Text <> "" And txtTarget.Text <> sSelectTarget Then
                GoTo activateOK
            ElseIf .optCopySave.Value And txtSource.Text <> "" And txtSource.Text <> sSelectSource Then
                If .optNewXML.Value And txtConfigFileDir.Text <> "" And txtConfigFileDir.Text <> sSelectDirectory And txtNewConfigFileName.Text <> "" And txtNewConfigFileName.Text <> sEnterFileName Then
                    GoTo activateOK
                ElseIf .optAddNewConfig.Value And txtConfigFile.Text <> sSelectFile And txtConfigName.Text <> "" And txtConfigName.Text <> sEnterConfigName Then
                    GoTo activateOK
                Else
                    GoTo activateCANCEL
                End If
            ElseIf .optLoadPaste.Value And txtTarget.Text <> "" And txtTarget.Text <> sSelectTarget And txtConfigFile.Text <> sSelectFile And cmbConfigName.Text <> "" Then
                GoTo activateOK
            Else
                GoTo activateCANCEL
            End If
        
        Else
            GoTo activateCANCEL
        End If
        
activateOK:
        .cmdOK.ForeColor = -2147483630
        .cmdOK.Locked = False
        .cmdOK.Default = True
        On Error GoTo 0
        Exit Sub
    
activateCANCEL:
        .cmdOK.ForeColor = -2147483631
        .cmdOK.Locked = True
        .cmdCancel.Default = True
        On Error GoTo 0
    
    End With
    
End Sub
