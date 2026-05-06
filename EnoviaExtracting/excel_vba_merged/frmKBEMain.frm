Attribute VB_Name = "frmKBEMain"
Attribute VB_Base = "0{25E00E79-53E5-4170-B6F0-128A315E6618}{3EA86231-4BD4-4C8F-BA92-71ED254A4EAA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private iButtonCount As Integer
Private iButtonTop As Integer
Private iButtonLeft As Integer
Private iSpacing As Integer
Private iLines As Integer
Private iToolbarTop As Integer
Private iToolbarLeft As Integer
Private oToolbarWindow As Integer
Private myViewPoint As Viewpoint3D
Private params() As Variant


'********************************************************************************
'* Name: Prod Toolbar
'* Purpose: Creates a toolbar. Size is automatically defined by the number of buttons.
'*          There's a main section for the toolbar itself (e.g. size, positioning of all buttons, etc..)
'*          and a specific section for each button (e.g. what happens when you press a button)
'*
'* Assumption: Each button must be composed of two images:
'*             A first image represents the active button (default), and must be named "ImgXXX" where XXX is the command name
'*             A second image represents the inactive button (orange), and must be named "ImgXXX2" where XXX is the command name
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'******************************************************************************
'******************************************************************************
'******************************************************************************
'******************************************************************************
'                          USERFORM MAIN SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************

Private Sub UserForm_Initialize()
   
    'Check resolution
    If GetSystemMetrics32(78) < 1280 Then
        Call MsgBox("Low resolution (< 1280) may cause display problem for the DDP Toolbar", 48)
    End If
   
    On Error Resume Next
    '*****Catia status bar value
    CATIA.StatusBar = sScriptVersion & " - Welcome to KBE DDP Tools"
    
    '*****Toolbar title bar modification (toolbar look)*****
    Me.Caption = sScriptVersion & " - DDP Tools"
    oToolbarWindow = FindWindow(0&, Me.Caption)
    SetWindowLong oToolbarWindow, -20, 384 '256 for regular window
    DrawMenuBar oToolbarWindow
    
    '*****Position the toolbar
    Me.StartUpPosition = Manual
    
    'Position setting not found
    If GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".TOP") = "" Or GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".LEFT") = "" Then
        Call PositionFormInCATIAMiddle
        iToolbarTop = Me.Top
        iToolbarLeft = Me.Left
    
    'Setting were found
    Else
        iToolbarTop = GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".TOP")
        iToolbarLeft = GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".LEFT")
        iLines = GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".LINES") '*****Number of lines
        Call resetToolbar

        
        'Checking if the form is in the virtual screen
        If FormInVirtualScreen(iToolbarLeft, iToolbarTop) = True Then
            Me.Top = iToolbarTop '***** Bug: Need to do it twice
            Me.Top = iToolbarTop
            Me.Left = iToolbarLeft
        Else
            Call PositionFormInCATIAMiddle
        End If
    End If
    
    'Number of line not determined
    If iLines = 0 Then iLines = 1
    
    On Error GoTo 0
    
End Sub
Private Sub PositionFormInCATIAMiddle()
    With Me
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width / 2) - .Width / 2
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height / 2) - .Height / 2
    End With
    
    iToolbarLeft = Me.Left
    iToolbarTop = Me.Top
End Sub

Private Function FormInVirtualScreen(ByVal dLeft As Double, ByVal dTop As Double) As Boolean

    '*****GetSystemMetrics32(xx) > value is in pixel
    '     xx = 76 : The coordinates for the left side of the virtual screen
    '     xx = 77 : The coordinates for the top of the virtual screen
    '     xx = 78 : The width of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors
    '     xx = 79 : The height of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors
    
    If dLeft < GetSystemMetrics32(76) * dPointsToPixelRatioH Then
        FormInVirtualScreen = False
        Exit Function
    End If
    
    If dLeft > (GetSystemMetrics32(76) + GetSystemMetrics32(78)) * dPointsToPixelRatioH - Me.Width Then
        FormInVirtualScreen = False
        Exit Function
    End If
    
    If dTop < GetSystemMetrics32(77) Then
        FormInVirtualScreen = False
        Exit Function
    End If
    
    If Me.Top > (GetSystemMetrics32(77) + GetSystemMetrics32(79)) * dPointsToPixelRatioV - Me.Height Then
        FormInVirtualScreen = False
        Exit Function
    End If
    
    FormInVirtualScreen = True

End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call userFormFormat
End Sub

Private Sub UserForm_Click()
    Call resetToolbar
    
End Sub

Public Sub resetToolbar()
    
    Dim oExcelApp
    
'*****makes Active buttons (default images) visible
'*****and Inactive buttons (orange images) not visible
    Me.Enabled = True
    
    '***** Reset Public variable
    bToolbarCommandSelected = True 'Boolean used to stop any procedure when clicking on toolbar
    bCancelAction = True 'Used to determine if user pressed Cancel in a progress bar
    
    '***** Unload all userforms
    For Each iObject In UserForms
        If Not iObject.Name Like "*KBEMain*" Then Unload iObject
    Next
    
    '***** Show all active controls on toolbar
    For Each Control In Me.Controls
        If Right(Control.Name, 1) = 2 Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Next
    
    Call userFormFormat
    
End Sub

Private Sub userFormFormat(Optional ByVal bToolbarHorizontal As Boolean = True)

'****************************
'Position all buttons on toolbar (calling sub setButtonPosition)
'Create an index number for each button
'Toolbar width is function of number of buttons
'By default, if no arguments, toolbar is horizontal
'****************************
    
    iButtonCount = 0
    iButtonTop = 5 '*****Top position of buttons in toolbar
    iButtonLeft = 12 '*****Starting left position of first button in toolbar
    iSpacing = 21 '*****Spacing between buttons
        
        '***** 1 - Insert Part Template
        setButtonPosition Me.ImgInsertPartTemplate, Me.ImgInsertPartTemplate2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 2 - Create Flat Panel
        setButtonPosition Me.ImgCreateFlatPanel, Me.ImgCreateFlatPanel2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 3 - Add Finish on panel
        setButtonPosition Me.ImgFinish, Me.ImgFinish2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 4 - Insert Hole
        setButtonPosition Me.ImgInsertHole, Me.ImgInsertHole2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 5 - Fill FS, WL, BL Ballons
        setButtonPosition Me.ImgFillFSNumber, Me.ImgFillFSNumber2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 6 - Create Capture
        setButtonPosition Me.ImgCreateCapture, Me.ImgCreateCapture2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 7 - Evolve Parts
        setButtonPosition Me.ImgEvolveParts, Me.ImgEvolveParts2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 8 - Replace Multiple Components
        setButtonPosition Me.ImgReplaceMultiComponents, Me.ImgReplaceMultiComponents2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 9 - Rename Instances
        setButtonPosition Me.ImgRenameInstance, Me.ImgRenameInstance2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 10 - Instantiate Hardware
        setButtonPosition Me.ImgInstantiateHardware, Me.ImgInstantiateHardware2, iButtonCount
        iButtonCount = iButtonCount + 1
                
        '***** 11 - Insert Drawing Format
        setButtonPosition Me.ImgInsertDrawingFormat, Me.ImgInsertDrawingFormat2, iButtonCount
        iButtonCount = iButtonCount + 1

        '***** 12 - Search Drafting Objects
        setButtonPosition Me.ImgSearchDraftingObjects, Me.ImgSearchDraftingObjects2, iButtonCount
        iButtonCount = iButtonCount + 1

        '***** 13 - Fill Find Number
        setButtonPosition Me.ImgFillFindNumber, Me.ImgFillFindNumber2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 14 - Update Location Info
        setButtonPosition Me.ImgLocationInfo, Me.ImgLocationInfo2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 15 - Get Data for EDRN Box 21
        setButtonPosition Me.ImgEDRNBox21, Me.ImgEDRNBox212, iButtonCount
        iButtonCount = iButtonCount + 1
                        
        '***** 16 - Drawing Tools
        setButtonPosition Me.ImgDrawingTools, Me.ImgDrawingTools2, iButtonCount
        iButtonCount = iButtonCount + 1
                        
        '***** 17 - Extract PL
        setButtonPosition Me.ImgPLExtract, Me.ImgPLExtract2, iButtonCount
        iButtonCount = iButtonCount + 1
                        
        '***** 18 - Import PL
        setButtonPosition Me.ImgPLImport, Me.ImgPLImport2, iButtonCount
        iButtonCount = iButtonCount + 1
                                 
        '***** 19 - Transfer PL to Excel
        setButtonPosition Me.ImgPLTransferToExcel, Me.ImgPLTransferToExcel2, iButtonCount
        iButtonCount = iButtonCount + 1
                        
        '***** 20 - Reset Drafting Colors
        setButtonPosition Me.ImgResetDraftingColors, Me.ImgResetDraftingColors2, iButtonCount
        iButtonCount = iButtonCount + 1
                        
        '***** 21 - Import SDD
        setButtonPosition Me.ImgSDDImport, Me.ImgSDDImport2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 22 - Transfer SDD to Excel
        setButtonPosition Me.ImgSDDTransferToExcel, Me.ImgSDDTransferToExcel2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 23 - Attribute Analysis
        setButtonPosition Me.ImgAttributeAnalysis, Me.ImgAttributeAnalysis2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 24 - Extract IDL
        setButtonPosition Me.ImgExtractIDL, Me.ImgExtractIDL2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 25 - Compare IDL
        setButtonPosition Me.ImgIDLCompare, Me.ImgIDLCompare2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 26 - Import Color Code Table
        setButtonPosition Me.ImgImportColorCodeTable, Me.ImgImportColorCodeTable2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 27 - Check Fastener Compatibility
        setButtonPosition Me.ImgChkFastenerCompatibility, Me.ImgChkFastenerCompatibility2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 28 - Check Panel Assembly
        setButtonPosition Me.ImgPanelAssemblyCheck, Me.ImgPanelAssemblyCheck2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 29 - Clash Analysis
        setButtonPosition Me.ImgClashAnalysis, Me.ImgClashAnalysis2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 30 - PVR Sync Tool
        setButtonPosition Me.ImgPVRSync, Me.ImgPVRSync2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 31 - Resize Toolbar
        setButtonPosition Me.ImgResizeToolbar, Me.ImgResizeToolbar2, iButtonCount
        iButtonCount = iButtonCount + 1
                                
        '***** ii - XXXXXXXXXXXXXXXXXXX ADD NEW BUTTON
        'setButtonPosition Me.ImgMYNEWBUTTON, Me.ImgMYNEWBUTTON2, iButtonCount
        'iButtonCount = iButtonCount + 1
        
    '***** Toolbar size *****
    If Not iButtonCount / iLines - Int(iButtonCount / iLines) = 0 Then iButtonCount = (iLines * Int(iButtonCount / iLines)) + iLines
    Me.Width = iButtonLeft + iButtonCount / iLines * iSpacing + 5
    Me.Height = iButtonTop + iLines * iSpacing + 18
        
    On Error Resume Next
    CATIA.StatusBar = ""
    On Error GoTo 0
    
End Sub

Private Sub setButtonPosition(buttonActive As Image, buttonInactive As Image, ByVal iButtonIndex As Integer)
'****************************
'Position an active button (default image)
'Position an inactive button (orange image) at Top+1 , Left+1
'****************************
    
    With buttonActive
        .Left = iButtonLeft + Int(iButtonIndex / iLines) * iSpacing
        .Top = iButtonTop + iSpacing * (iButtonIndex / iLines - Int(iButtonIndex / iLines)) * iLines
    End With
    
    With buttonInactive
        .Left = iButtonLeft + Int(iButtonIndex / iLines) * iSpacing + 1
        .Top = iButtonTop + iSpacing * (iButtonIndex / iLines - Int(iButtonIndex / iLines)) * iLines + 1
    End With
    
End Sub

Private Sub UserForm_Layout()
    
'****************************
'Every time userform is moved, make sure it does not get out of the screen!
'****************************
    
    '*****GetSystemMetrics32(xx) > value is in pixel
    '     xx = 0 : Width of principal screen
    '     xx = 1 : Heigth of principal screen
    '     xx = 76 : The coordinates for the left side of the virtual screen
    '     xx = 77 : The coordinates for the top of the virtual screen
    '     xx = 78 : The width of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors
    '     xx = 79 : The height of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors
    
    If Me.Left < GetSystemMetrics32(76) * dPointsToPixelRatioH Then
        Me.Left = GetSystemMetrics32(76) * dPointsToPixelRatioH
    End If
    
    If Me.Left > (GetSystemMetrics32(76) + GetSystemMetrics32(78)) * dPointsToPixelRatioH - Me.Width Then
        Me.Left = (GetSystemMetrics32(76) + GetSystemMetrics32(78)) * dPointsToPixelRatioH - Me.Width
    End If
    
    If Me.Top < GetSystemMetrics32(77) Then
        Me.Top = GetSystemMetrics32(77)
    End If
    
    If Me.Top > (GetSystemMetrics32(77) + GetSystemMetrics32(79)) * dPointsToPixelRatioV - Me.Height Then
        Me.Top = (GetSystemMetrics32(77) + GetSystemMetrics32(79)) * dPointsToPixelRatioV - Me.Height
    End If
    
    

    
'****************************
'Every time userform is moved by the user we keep the position
'****************************
    If Me.Visible = True Then
        iToolbarTop = Me.Top
        iToolbarLeft = Me.Left
    End If
    
End Sub

Private Sub UserForm_Terminate()
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".TOP", Str(iToolbarTop))
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".LEFT", Str(iToolbarLeft))
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".LINES", Str(iLines))
    Call GetTagValueFromFile("", "", True)
    Call resetToolbar
    Call WebServiceExecute("Stop")
    Unload Me
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Insert New Part Template
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInsertPartTemplate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInsertPartTemplate.Top = Me.ImgInsertPartTemplate2.Top
    Me.ImgInsertPartTemplate.Left = Me.ImgInsertPartTemplate2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Insert a new Part Template"
    On Error GoTo 0
    
End Sub

Private Sub ImgInsertPartTemplate_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgInsertPartTemplate.Visible = False: Me.ImgInsertPartTemplate2.Visible = True
    frmNewPartTemplate.InitializeForm
End Sub

Private Sub ImgInsertPartTemplate2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Create Flat Panel
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCreateFlatPanel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCreateFlatPanel.Top = Me.ImgCreateFlatPanel2.Top
    Me.ImgCreateFlatPanel.Left = Me.ImgCreateFlatPanel2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create a Flat Panel"
    On Error GoTo 0
    
End Sub

Private Sub ImgCreateFlatPanel_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCreateFlatPanel.Visible = False: Me.ImgCreateFlatPanel2.Visible = True
    frmCreateFlatPanel.InitializeForm
    
End Sub

Private Sub ImgCreateFlatPanel2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Insert Hole
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInsertHole_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInsertHole.Top = Me.ImgInsertHole2.Top
    Me.ImgInsertHole.Left = Me.ImgInsertHole2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create insert holes in panel"
    On Error GoTo 0
    
End Sub

Private Sub ImgInsertHole_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgInsertHole.Visible = False: Me.ImgInsertHole2.Visible = True
    frmInsertHole.InitializeForm
    
End Sub

Private Sub ImgInsertHole2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Create Capture
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCreateCapture_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCreateCapture.Top = Me.ImgCreateCapture2.Top
    Me.ImgCreateCapture.Left = Me.ImgCreateCapture2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create Captures in CATPart"
    On Error GoTo 0
    
End Sub

Private Sub ImgCreateCapture_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCreateCapture.Visible = False: Me.ImgCreateCapture2.Visible = True
    Call oCreateCapture.Create
    
End Sub

Private Sub ImgCreateCapture2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Open Attribute Analysis
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgAttributeAnalysis_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgAttributeAnalysis.Top = Me.ImgAttributeAnalysis2.Top
    Me.ImgAttributeAnalysis.Left = Me.ImgAttributeAnalysis2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Open the Attribute Analysis tool"
    On Error GoTo 0
    
End Sub

Private Sub ImgAttributeAnalysis_Click()
    
    Dim oExcel As clsExcel
    
    Call resetToolbar
    
    On Error Resume Next
'    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgAttributeAnalysis.Visible = False: Me.ImgAttributeAnalysis2.Visible = True
'    Set oExcel = New clsExcel
'    Call oExcel.GetExcel
'    Call oExcel.OpenExcelFile(sAttributesTemplateFile)
'    Call oExcel.ShowExcelWindow
    frmAttributeAnalysis.InitializeForm
    'Call resetToolbar
End Sub

Private Sub ImgAttributeAnalysis2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Extract IDL
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgExtractIDL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgExtractIDL.Top = Me.ImgExtractIDL2.Top
    Me.ImgExtractIDL.Left = Me.ImgExtractIDL2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "IDL Extract Tool"
    On Error GoTo 0
    
End Sub

Private Sub ImgExtractIDL_Click()
    
    Call resetToolbar
        
    'Active document must be a CATProduct
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
        sAnswer = MsgBox("Active document must be a CATProduct", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgExtractIDL.Visible = False: Me.ImgExtractIDL2.Visible = True
    
    Set oIDLExtract = New IDLExtract
    Call oIDLExtract.oMain

End Sub

Private Sub ImgExtractIDL2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Compare IDL
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgIDLCompare_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgIDLCompare.Top = Me.ImgIDLCompare2.Top
    Me.ImgIDLCompare.Left = Me.ImgIDLCompare2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "IDL Compare Tool"
    On Error GoTo 0
    
End Sub

Private Sub ImgIDLCompare_Click()
    
    Dim oExcel As clsExcel
    
    Call resetToolbar
    
    '*** Get Excel file
    Set oExcel = New clsExcel
    Call oExcel.GetExcel

    '***Open Excel file
    Call oExcel.OpenExcelFile(sIDLCompareFile)
    oExcel.ShowExcelWindow

    Call frmKBEMain.resetToolbar

End Sub

Private Sub ImgIDLCompare2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'                       Import Color Code Table
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgImportColorCodeTable_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgImportColorCodeTable.Top = Me.ImgImportColorCodeTable2.Top
    Me.ImgImportColorCodeTable.Left = Me.ImgImportColorCodeTable2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Import Color Coding Table"
    On Error GoTo 0
    
End Sub

Private Sub ImgImportColorCodeTable_Click()
    
    Dim sAnswer As String
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        sAnswer = MsgBox("Active document must be a CATDrawing", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgImportColorCodeTable.Visible = False: Me.ImgImportColorCodeTable2.Visible = True
    Call mdlImportColorCodeTable.ColorCodeMain
    
End Sub

Private Sub ImgImportColorCodeTable2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Replace Multiple Components
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgReplaceMultiComponents_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgReplaceMultiComponents.Top = Me.ImgReplaceMultiComponents2.Top
    Me.ImgReplaceMultiComponents.Left = Me.ImgReplaceMultiComponents2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Replace Multiple Components"
    On Error GoTo 0
    
End Sub

Private Sub ImgReplaceMultiComponents_Click()
    
    Dim i As Integer
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgReplaceMultiComponents.Visible = False: Me.ImgReplaceMultiComponents2.Visible = True
    
    'Check preselection
    Set oSelection = CATIA.ActiveDocument.Selection

    If oSelection.Count = 0 Then
        MsgBox ("You must preselect the parts to be replaced before launching the macro")
        Call resetToolbar
        Exit Sub
    End If

    For i = 1 To oSelection.Count
        If TypeName(oSelection.Item(i).Value) <> "Product" Then
            MsgBox ("Only select instances from the graph")
            Exit Sub
        End If
    Next
    
    Call frmReplaceMulti.InitializeForm
    
End Sub

Private Sub ImgReplaceMultiComponents2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Rename Instances
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgRenameInstance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgRenameInstance.Top = Me.ImgRenameInstance2.Top
    Me.ImgRenameInstance.Left = Me.ImgRenameInstance2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Rename Instances"
    On Error GoTo 0
    
End Sub

Private Sub ImgRenameInstance_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgRenameInstance.Visible = False: Me.ImgRenameInstance2.Visible = True
    Call oRenameInstance.SelectParentInstance
    
End Sub

Private Sub ImgRenameInstance2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Instantiate Hardware
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInstantiateHardware_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInstantiateHardware.Top = Me.ImgInstantiateHardware2.Top
    Me.ImgInstantiateHardware.Left = Me.ImgInstantiateHardware2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Instantiate Hardware"
    On Error GoTo 0
    
End Sub

Private Sub ImgInstantiateHardware_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgInstantiateHardware.Visible = False: Me.ImgInstantiateHardware2.Visible = True
    Call frmInstantiateHardware.InitializeForm
    
End Sub

Private Sub ImgInstantiateHardware2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Add Finish on panel
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgFinish_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgFinish.Top = Me.ImgFinish2.Top
    Me.ImgFinish.Left = Me.ImgFinish2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Add Finish on Panel"
    On Error GoTo 0
    
End Sub

Private Sub ImgFinish_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgFinish.Visible = False: Me.ImgFinish2.Visible = True
    Call frmFinish.InitializeForm
    
End Sub

Private Sub ImgFinish2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Fill FS Number
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgFillFSNumber_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgFillFSNumber.Top = Me.ImgFillFSNumber2.Top
    Me.ImgFillFSNumber.Left = Me.ImgFillFSNumber2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Fill FS,WL,BL"
    On Error GoTo 0
    
End Sub

Private Sub ImgFillFSNumber_Click()
    
    Call resetToolbar

    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgFillFSNumber.Visible = False: Me.ImgFillFSNumber2.Visible = True
    frmBalloons.InitializeForm
End Sub

Private Sub ImgFillFSNumber2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Inser Drawing Format
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInsertDrawingFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInsertDrawingFormat.Top = Me.ImgInsertDrawingFormat2.Top
    Me.ImgInsertDrawingFormat.Left = Me.ImgInsertDrawingFormat2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Insert Drawing Format"
    On Error GoTo 0
    
End Sub

Private Sub ImgInsertDrawingFormat_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgInsertDrawingFormat.Visible = False: Me.ImgInsertDrawingFormat2.Visible = True
    Call frmInsertDrawingFormat.InitializeForm
    
End Sub

Private Sub ImgInsertDrawingFormat2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Drawing Tools
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgDrawingTools_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgDrawingTools.Top = Me.ImgDrawingTools2.Top
    Me.ImgDrawingTools.Left = Me.ImgDrawingTools2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Drawing Tools"
    On Error GoTo 0
    
End Sub

Private Sub ImgDrawingTools_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgDrawingTools.Visible = False: Me.ImgDrawingTools2.Visible = True
    Call frmDrawingTools.InitializeForm
    
End Sub

Private Sub ImgDrawingTools2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Search Drafting Objects
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSearchDraftingObjects_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSearchDraftingObjects.Top = Me.ImgSearchDraftingObjects2.Top
    Me.ImgSearchDraftingObjects.Left = Me.ImgSearchDraftingObjects2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Search Drafting Objects"
    On Error GoTo 0
    
End Sub

Private Sub ImgSearchDraftingObjects_Click()
    
    Call resetToolbar
    Dim sAnswer As String
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Call resetToolbar
        Exit Sub
    End If
    On Error GoTo 0
    
    sAnswer = InputBox("Key in the number of the objects you are looking for:" & vbCrLf & _
                       " 1) Flag Notes" & vbCrLf & _
                       " 2) Find Numbers")
                       
    '*****Log File
    Dim sDocumentName As String, sDocRev As String
    On Error Resume Next
        sDocumentName = Left(Split(CATIA.ActiveDocument, ".CATDrawing")(0), Len(Split(CATIA.ActiveDocument, ".CATDrawing")(0)) - 2)
        sDocRev = Right(Split(CATIA.ActiveDocument, ".CATDrawing")(0), 2)
    On Error GoTo 0
    Call AddToLogFile("Search Drafting Objects", sDocumentName, sDocRev)
    '*****
    Me.ImgSearchDraftingObjects.Visible = False: Me.ImgSearchDraftingObjects2.Visible = True
    
    If Left(sAnswer, 1) = "1" Then
        Call oSearchFlagNote.Launch
    ElseIf Left(sAnswer, 1) = "2" Then
        Call oSearchFindNo.Launch
    Else
        Call resetToolbar
    End If
End Sub

Private Sub ImgSearchDraftingObjects2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'                       Fill Find Number
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgFillFindNumber_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgFillFindNumber.Top = Me.ImgFillFindNumber2.Top
    Me.ImgFillFindNumber.Left = Me.ImgFillFindNumber2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Fill Find Numbers in PL"
    On Error GoTo 0
    
End Sub

Private Sub ImgFillFindNumber_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgFillFindNumber.Visible = False: Me.ImgFillFindNumber2.Visible = True
    Call oFillFindNumber.Launch
    
End Sub

Private Sub ImgFillFindNumber2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Location Info Text
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgLocationInfo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgLocationInfo.Top = Me.ImgLocationInfo2.Top
    Me.ImgLocationInfo.Left = Me.ImgLocationInfo2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Find and Update Location Info Box"
    On Error GoTo 0
    
End Sub

Private Sub ImgLocationInfo_Click()
    Dim sdummy As String
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgLocationInfo.Visible = False: Me.ImgLocationInfo2.Visible = True
    Call mdlVueFindNumber.LaunchVueFindNumber(sdummy)
    
End Sub

Private Sub ImgLocationInfo2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Get Data for EDRN Box 21
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgEDRNBox21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgEDRNBox21.Top = Me.ImgEDRNBox212.Top
    Me.ImgEDRNBox21.Left = Me.ImgEDRNBox212.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Get Data for EDRN Box 21 in Clipboard"
    On Error GoTo 0
    
End Sub

Private Sub ImgEDRNBox21_Click()
    Dim sdummy As String
    Call resetToolbar
    
    Me.ImgEDRNBox21.Visible = False: Me.ImgEDRNBox212.Visible = True
    Call frmEDRNBox21.InitializeForm
    
End Sub

Private Sub ImgEDRNBox212_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Extract Part List
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgPLExtract_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgPLExtract.Top = Me.ImgPLExtract2.Top
    Me.ImgPLExtract.Left = Me.ImgPLExtract2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Extract Part List"
    On Error GoTo 0
    
End Sub

Private Sub ImgPLExtract_Click()
    
    Call resetToolbar
    
    Me.ImgPLExtract.Visible = False: Me.ImgPLExtract2.Visible = True
    Call frmPLExtract.InitializeForm
    
End Sub

Private Sub ImgPLExtract2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Import PL in CATDrawing
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgPLImport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgPLImport.Top = Me.ImgPLImport2.Top
    Me.ImgPLImport.Left = Me.ImgPLImport2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Import PL in CATDrawing"
    On Error GoTo 0
    
End Sub

Private Sub ImgPLImport_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgPLImport.Visible = False: Me.ImgPLImport2.Visible = True
    Call oPlImport.Launch
    
End Sub

Private Sub ImgPLImport2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Transfer PL to Excel
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgPLTransferToExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgPLTransferToExcel.Top = Me.ImgPLTransferToExcel2.Top
    Me.ImgPLTransferToExcel.Left = Me.ImgPLTransferToExcel2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Transfer PL to Excel"
    On Error GoTo 0
    
End Sub

Private Sub ImgPLTransferToExcel_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgPLTransferToExcel.Visible = False: Me.ImgPLTransferToExcel2.Visible = True
    Call oPlTransferToExcel.Launch
    
End Sub

Private Sub ImgPLTransferToExcel2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Reset Drafting Colors
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgResetDraftingColors_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgResetDraftingColors.Top = Me.ImgResetDraftingColors2.Top
    Me.ImgResetDraftingColors.Left = Me.ImgResetDraftingColors2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Reset Drafting Colors"
    On Error GoTo 0
    
End Sub

Private Sub ImgResetDraftingColors_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgResetDraftingColors.Visible = False: Me.ImgResetDraftingColors2.Visible = True
    Call oResetDraftingColors.Launch
    
End Sub

Private Sub ImgResetDraftingColors2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'                       Import SDD in CATDrawing
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSDDImport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSDDImport.Top = Me.ImgSDDImport2.Top
    Me.ImgSDDImport.Left = Me.ImgSDDImport2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Import WP4 SDD in CATDrawing"
    On Error GoTo 0
    
End Sub

Private Sub ImgSDDImport_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgSDDImport.Visible = False: Me.ImgSDDImport2.Visible = True
    Call oSDDImport.Launch
    
End Sub

Private Sub ImgSDDImport2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Transfer SDD to Excel
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSDDTransferToExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSDDTransferToExcel.Top = Me.ImgSDDTransferToExcel2.Top
    Me.ImgSDDTransferToExcel.Left = Me.ImgSDDTransferToExcel2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Transfer WP4 SDD to Excel"
    On Error GoTo 0
    
End Sub

Private Sub ImgSDDTransferToExcel_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "DrawingDocument" Then
        Call MsgBox("Active document must be a CATDrawing.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Me.ImgSDDTransferToExcel.Visible = False: Me.ImgSDDTransferToExcel2.Visible = True
    Call oSDDTransferToExcel.Launch
    
End Sub

Private Sub ImgSDDTransferToExcel2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Check fastener compatibility
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgChkFastenerCompatibility_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgChkFastenerCompatibility.Top = Me.ImgChkFastenerCompatibility2.Top
    Me.ImgChkFastenerCompatibility.Left = Me.ImgChkFastenerCompatibility2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Fastener Compatibility"
    On Error GoTo 0
    
End Sub

Private Sub ImgChkFastenerCompatibility_Click()
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
        MsgBox ("Active document must be an assembly")
        Call resetToolbar
        GoTo EndCheckassy
    End If
    On Error GoTo 0
    
    Me.ImgChkFastenerCompatibility.Visible = False: Me.ImgChkFastenerCompatibility2.Visible = True
    Call frmCheckFastenerCompatibility.LaunchAssemblyAnalysis
EndCheckassy:
    If frmCheckFastenerCompatibility.Visible = False Then
        Dim cClashes As Clashes
        Dim oClash As Clash
        On Error Resume Next
            Set cClashes = CATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")
            Set cGroups = CATIA.ActiveDocument.Product.GetTechnologicalObject("Groups")
            cClashes.Remove ("AutomationClash_001")
            cGroups.Remove ("AutomationGroup_001")
        On Error GoTo 0
        If Not bToolbarCommandSelected Then Call resetToolbar
    End If
End Sub

Private Sub ImgChkFastenerCompatibility2_Click()
    Dim cClashes As Clashes
    Dim oClash As Clash
    On Error Resume Next
        Set cClashes = CATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")
        Set cGroups = CATIA.ActiveDocument.Product.GetTechnologicalObject("Groups")
        cClashes.Remove ("AutomationClash_001")
        cGroups.Remove ("AutomationGroup_001")
    On Error GoTo 0
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'                       Check Check Panel Assy
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgPanelAssemblyCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgPanelAssemblyCheck.Top = Me.ImgPanelAssemblyCheck2.Top
    Me.ImgPanelAssemblyCheck.Left = Me.ImgPanelAssemblyCheck2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Panel Assembly"
    On Error GoTo 0
    
End Sub

Private Sub ImgPanelAssemblyCheck_Click()
    Call resetToolbar
    Dim sdummy As String
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
        MsgBox ("Active document must be an assembly")
        Call resetToolbar
        GoTo EndCheckassy
    End If
    On Error GoTo 0
    If MsgBox("Do you Want to run a new analysis ?" & vbCrLf & _
              "- Press Yes to run a new analysis" & vbCrLf & _
              "- Press No to link an existing report with Catia" _
              , vbYesNo + vbDefaultButton1 + vbQuestion, "Select An Action") = vbYes Then
        Me.ImgPanelAssemblyCheck.Visible = False: Me.ImgPanelAssemblyCheck2.Visible = True
        Call ScanPanelAssies(sdummy)
    Else
        MsgBox "Make sure that the active window in Catia is of the assembly you " & vbCrLf & _
               "want to analyse and the active document in excel is the analysis " & vbCrLf & _
               "report you want to connect with Catia.", vbOKOnly + vbInformation
        Call frmDisplayInCatia.LaunchDisplayInCatia(CATIA.ActiveWindow)
    End If
EndCheckassy:
End Sub

Private Sub ImgPanelAssemblyCheck2_Click()
    Call resetToolbar
End Sub
'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Clash Analysis
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgClashAnalysis_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgClashAnalysis.Top = Me.ImgClashAnalysis2.Top
    Me.ImgClashAnalysis.Left = Me.ImgClashAnalysis2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Clash Analysis"
    On Error GoTo 0
    
End Sub

Private Sub ImgClashAnalysis_Click()

    Call resetToolbar

    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub

    Me.ImgClashAnalysis.Visible = False: Me.ImgClashAnalysis2.Visible = True

   
    Call frmClashAnalysis.initializeClashAnalysisForm
        
    
    On Error GoTo 0

End Sub

Private Sub ImgClashAnalysis2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       RESIZE TOOLBAR
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgResizeToolbar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgResizeToolbar.Top = Me.ImgResizeToolbar2.Top
    Me.ImgResizeToolbar.Left = Me.ImgResizeToolbar2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Resize Toolbar"
    On Error GoTo 0
    
End Sub

Private Sub ImgResizeToolbar_Click()
    
    Call resetToolbar
    Me.ImgResizeToolbar.Visible = False: Me.ImgResizeToolbar2.Visible = True
    iLines = iLines + 1
    If iLines > 3 Then iLines = 1
    Call resetToolbar
    
End Sub

Private Sub ImgResizeToolbar2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       Evolve Parts
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgEvolveParts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgEvolveParts.Top = Me.ImgEvolveParts2.Top
    Me.ImgEvolveParts.Left = Me.ImgEvolveParts2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Evolve Parts"
    On Error GoTo 0
    
End Sub

Private Sub ImgEvolveParts_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgEvolveParts.Visible = False: Me.ImgEvolveParts2.Visible = True
    frmEvolveParts.InitializeForm
    
End Sub

Private Sub ImgEvolveParts2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'                       PVR Sync Tool
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgPVRSync_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgPVRSync.Top = Me.ImgPVRSync2.Top
    Me.ImgPVRSync.Left = Me.ImgPVRSync2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "PVR Sync Tool"
    On Error GoTo 0
    
End Sub

Private Sub ImgPVRSync_Click()
    
    Dim params(0)
    Dim SysService
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
        
    Me.ImgPVRSync.Visible = False: Me.ImgPVRSync2.Visible = True
    
'    params(0) = sKBEPathFile
'    Set SysService = CATIA.SystemService
'    Call SysService.ExecuteScript(sPVRSYNCToolbarPath, catScriptLibraryTypeVBAProject, sPVRSYNCToolbarModule, "CATMain", params)
    Call mdlPVRSync.PVRSyncMain
    
    Call resetToolbar

End Sub

Private Sub ImgPVRSync2_Click()
    Call resetToolbar
End Sub

