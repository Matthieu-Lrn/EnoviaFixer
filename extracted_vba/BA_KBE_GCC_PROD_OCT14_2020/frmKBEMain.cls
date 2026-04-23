Attribute VB_Name = "frmKBEMain"
Attribute VB_Base = "0{8514B98A-BCAF-4DA8-BC18-ED7BB2347CCD}{37134D27-E7A9-46B1-80EB-0EE97756A894}"
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
Private iToolbarTop As Integer
Private iToolbarLeft As Integer
Private oToolbarWindow As Integer

'********************************************************************************
'               ONE NEW OBJECT PER CLASS MODULE
'********************************************************************************
Private oSelectionRename As New StringToModify
Private oCollectorPosition As New CollectorPosition

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
    
    '*****Set controls position and visibility
    Call resetToolbar
    
    On Error Resume Next
    '*****Catia status bar value
    CATIA.StatusBar = sScriptVersion & " - Welcome to KBE Production Tools"
    
    '*****Toolbar title bar modification (toolbar look)*****
    Me.Caption = sScriptVersion & " - Production Tools"
    oToolbarWindow = FindWindow(0&, Me.Caption)
    SetWindowLong oToolbarWindow, -20, 384 '256 for regular window
    DrawMenuBar oToolbarWindow
        
    '*****Position the toolbar at the same place it was open last time*****
    '*****Verify value in Settings File
    '*****If variable is empty
    If GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".TOP") = "" Or GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".LEFT") = "" Then
        Me.StartUpPosition = WindowsDefaultLocation
        iToolbarTop = Me.Top
        iToolbarLeft = Me.Left
    '*****Else, put toolbar at previously saved position
    Else
        Me.StartUpPosition = Manual
        iToolbarTop = GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".TOP")
        iToolbarLeft = GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".LEFT")
        Me.Top = iToolbarTop
        Me.Left = iToolbarLeft
    End If
    
    If GetTagValueFromFile(sSettingsPath, "LastSelected.FOLDER") = "" Then
        sLastFolderSelected = sTempDirectory
    Else
        sLastFolderSelected = GetTagValueFromFile(sSettingsPath, "LastSelected.FOLDER")
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call userFormFormat
End Sub

Private Sub UserForm_Click()
    Call resetToolbar
End Sub

Public Sub resetToolbar()
    
'*****makes Active buttons (default images) visible
'*****and Inactive buttons (orange images) not visible
    Me.Enabled = True
    
    '***** Reset Public variable
    bToolbarCommandSelected = True 'Boolean used to stop any procedure when clicking on toolbar
    bCancelAction = True 'Used to determine if user pressed Cancel in a progress bar
    sSeedPartsList = ""
    Set oSeedAssy = Nothing
    Set oNHAAssy = Nothing
    Set oSeedPart = Nothing
    Set oSeedLingPart = Nothing
    
    '***** Unload all userforms
    For Each iObject In UserForms
        If Not iObject.Name Like "*KBEMain*" Then Unload iObject
    Next
    
    '***** Show all active controls on toolbar
    For Each oControl In Me.Controls
        If Right(oControl.Name, 1) = 2 Then
            oControl.Visible = False
        Else
            oControl.Visible = True
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
        
        '***** 1 - Manage DT Parameters (XML)
        setButtonPosition Me.ImgXML, Me.ImgXML2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 2 - Cabinet Project
        setButtonPosition Me.ImgCabinetProject, Me.ImgCabinetProject2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 3 - Instantiate STD Parts
        setButtonPosition Me.ImgInstantiateSTDParts, Me.ImgInstantiateSTDParts2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 4 - 3D Collector
        setButtonPosition Me.Img3DCollector, Me.Img3DCollector2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 5 - 3D Collector Position check
        setButtonPosition Me.ImgCollPosition, Me.ImgCollPosition2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 6 - 3D Collector Template
        setButtonPosition Me.Img3DCollTemplate, Me.Img3DCollTemplate2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 7 - SP Collector
        setButtonPosition Me.ImgSPCollector, Me.ImgSPCollector2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 8 - Flat Tree
        setButtonPosition Me.ImgFlatTree, Me.ImgFlatTree2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 9 - Identify / Reuse existing parts
        'setButtonPosition Me.ImgReuse, Me.ImgReuse2, iButtonCount
        'iButtonCount = iButtonCount + 1
        
        '***** 10 - Subsets creation tool
        setButtonPosition Me.ImgSubset, Me.ImgSubset2, iButtonCount
        iButtonCount = iButtonCount + 1

        '***** 11 - Change Render Style
        setButtonPosition Me.ImgRenderStyle, Me.ImgRenderStyle2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 12 - Find and replace
        setButtonPosition Me.ImgFindReplace, Me.ImgFindReplace2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** ii - XXXXXXXXXXXXXXXXXXX ADD NEW BUTTON
        'setButtonPosition Me.ImgMYNEWBUTTON, Me.ImgMYNEWBUTTON2, iButtonCount
        'iButtonCount = iButtonCount + 1
        
    '***** Toolbar size *****
    Me.Width = iButtonLeft + iButtonCount * iSpacing + 5
    Me.Height = iButtonTop + iSpacing + 18
    
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
        .Left = iButtonLeft + iButtonIndex * iSpacing
        .Top = iButtonTop
    End With
        
    With buttonInactive
        .Left = iButtonLeft + iButtonIndex * iSpacing + 1
        .Top = iButtonTop + 1
    End With
    
End Sub

Private Sub UserForm_Layout()
    
'****************************
'Every time userform is moved, make sure it does not get out of the screen!
'****************************
    
    '*****GetSystemMetrics32(xx) > value is in pixel
    '     xx = 0 : Width of principal screen
    '     xx = 1 : Heigth of principal screen
    '     xx = 78 : Width of complete area (idem as 0 if only 1 screen)
    '     xx = 79 : Max Heigth of complete area (idem as 1 if only 1 screen)
    
    If Me.Left < 0 Then
        Me.Left = 0
    End If
    
    If Me.Left > GetSystemMetrics32(78) * dPointsToPixelRatioH - Me.Width Then
        Me.Left = GetSystemMetrics32(78) * dPointsToPixelRatioH - Me.Width
    End If
    
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    
    If Me.Top + Me.Height > GetSystemMetrics32(79) * dPointsToPixelRatioV Then
        Me.Top = GetSystemMetrics32(79) * dPointsToPixelRatioV - Me.Height
    End If
    
'****************************
'Every time userform is moved, save its last position
'****************************
    iToolbarTop = Me.Top
    iToolbarLeft = Me.Left
    
End Sub

Private Sub UserForm_Terminate()
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".TOP", Str(iToolbarTop))
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".LEFT", Str(iToolbarLeft))
    Call SetTagValueToFile(sSettingsPath, "LastSelected.FOLDER", sLastFolderSelected)
    Call resetToolbar
    Unload Me
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                     MANAGE DT PARAMETERS (XML) SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgXML_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgXML.Top = Me.ImgXML2.Top
    Me.ImgXML.Left = Me.ImgXML2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Manage DT Parameters (xml)"
    On Error GoTo 0
    
End Sub

Private Sub ImgXML_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgXML.Visible = False: Me.ImgXML2.Visible = True
    frmManageData.initializeXMLForm
    On Error GoTo 0
    
End Sub

Private Sub ImgXML2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                     CABINET PROJECT SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCabinetProject_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCabinetProject.Top = Me.ImgCabinetProject2.Top
    Me.ImgCabinetProject.Left = Me.ImgCabinetProject2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create Cabinet Project"
    On Error GoTo 0
    
End Sub

Private Sub ImgCabinetProject_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgCabinetProject.Visible = False: Me.ImgCabinetProject2.Visible = True
    frmProject.initializeProjectForm
    On Error GoTo 0
    
End Sub

Private Sub ImgCabinetProject2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                          INSTANTIATE STANDARD PARTS SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInstantiateSTDParts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInstantiateSTDParts.Top = Me.ImgInstantiateSTDParts2.Top
    Me.ImgInstantiateSTDParts.Left = Me.ImgInstantiateSTDParts2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Instantiate Standard Parts"
    On Error GoTo 0
    
End Sub

Private Sub ImgInstantiateSTDParts_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgInstantiateSTDParts.Visible = False: Me.ImgInstantiateSTDParts2.Visible = True
    frmStdPartsInstantiation.initializeStdPartsForm
    On Error GoTo 0
    
End Sub

Private Sub ImgInstantiateSTDParts2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                           3D COLLECTOR SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub Img3DCollector_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.Img3DCollector.Top = Me.Img3DCollector2.Top
    Me.Img3DCollector.Left = Me.Img3DCollector2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create/Synchronize 3D Collector"
    On Error GoTo 0
    
End Sub

Private Sub Img3DCollector_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.Img3DCollector.Visible = False: Me.Img3DCollector2.Visible = True
    frm3DCollector.initialize3DCollectorForm
    On Error GoTo 0
    
End Sub

Private Sub Img3DCollector2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        CHECK COLLECTORS POSITION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCollPosition_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCollPosition.Top = Me.ImgCollPosition2.Top
    Me.ImgCollPosition.Left = Me.ImgCollPosition2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check 3D Collectors position"
    On Error GoTo 0
    
End Sub

Private Sub ImgCollPosition_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgCollPosition.Visible = False: Me.ImgCollPosition2.Visible = True
    Call oCollectorPosition.CheckPosition
    On Error GoTo 0
    
End Sub

Private Sub ImgCollPosition2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                     COLLECTOR TEMPLATE SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub Img3DCollTemplate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.Img3DCollTemplate.Top = Me.Img3DCollTemplate2.Top
    Me.Img3DCollTemplate.Left = Me.Img3DCollTemplate2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create/Synchronize Collector Template"
    On Error GoTo 0
    
End Sub

Private Sub Img3DCollTemplate_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.Img3DCollTemplate.Visible = False: Me.Img3DCollTemplate2.Visible = True
    frm3DCTemplate.initialize3DCTemplateForm
    On Error GoTo 0
    
End Sub

Private Sub Img3DCollTemplate2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                           SP COLLECTOR SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSPCollector_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSPCollector.Top = Me.ImgSPCollector2.Top
    Me.ImgSPCollector.Left = Me.ImgSPCollector2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create/Synchronize Standard Parts Collector"
    On Error GoTo 0
    
End Sub

Private Sub ImgSPCollector_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgSPCollector.Visible = False: Me.ImgSPCollector2.Visible = True
    frmSPCollector.initializeSPCollectorForm
    On Error GoTo 0
    
End Sub

Private Sub ImgSPCollector2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                           FLAT TREE SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgFlatTree_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgFlatTree.Top = Me.ImgFlatTree2.Top
    Me.ImgFlatTree.Left = Me.ImgFlatTree2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Create/Synchronize Flat Tree"
    On Error GoTo 0
    
End Sub

Private Sub ImgFlatTree_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgFlatTree.Visible = False: Me.ImgFlatTree2.Visible = True
    frmFlatTree.initializeFlatTreeForm
    On Error GoTo 0
    
End Sub

Private Sub ImgFlatTree2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                           REUSE PARTS SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgReuse_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgReuse.Top = Me.ImgReuse2.Top
    Me.ImgReuse.Left = Me.ImgReuse2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Identify/Reuse existing parts"
    On Error GoTo 0
    
End Sub

Private Sub ImgReuse_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgReuse.Visible = False: Me.ImgReuse2.Visible = True
    frmReuse.initializeReusePartsForm
    On Error GoTo 0
    
End Sub

Private Sub ImgReuse2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       MANAGE SUBSETS SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSubset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSubset.Top = Me.ImgSubset2.Top
    Me.ImgSubset.Left = Me.ImgSubset2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Subsets Management"
    On Error GoTo 0
    
End Sub

Private Sub ImgSubset_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    
    Me.ImgSubset.Visible = False: Me.ImgSubset2.Visible = True
    frmSubsetsManagement.initializeSubSetsForm
    On Error GoTo 0
    
End Sub

Private Sub ImgSubset2_Click()
    Call resetToolbar
End Sub



'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       CHANGE RENDER STYLE SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Name: Change Render Style
'* Purpose: Switch between render styles "Parallel" and "Perspective"
'*
'* Assumption:
'*
'* Author: Julien Bigaouette
'* Updated by:
'* Language: VBA
'******************************************************************************
Private Sub ImgRenderStyle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgRenderStyle.Top = Me.ImgRenderStyle2.Top
    Me.ImgRenderStyle.Left = Me.ImgRenderStyle2.Left
    
    On Error Resume Next
    If Not TypeName(CATIA.ActiveDocument) = "PartDocument" And Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then
        CATIA.StatusBar = "Change Render Style"
        Exit Sub
    Else
        If CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.ProjectionMode = catProjectionConic Then
            CATIA.StatusBar = "Switch to Parallel"
            Me.ImgRenderStyle.ControlTipText = "Switch to Parallel"
            Me.ImgRenderStyle2.ControlTipText = "Switch to Parallel"
        Else
            CATIA.StatusBar = "Switch to Perspective"
            Me.ImgRenderStyle.ControlTipText = "Switch to Perspective"
            Me.ImgRenderStyle2.ControlTipText = "Switch to Perspective"
        End If
    End If
    On Error GoTo 0
    
End Sub

Private Sub ImgRenderStyle_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    If Not TypeName(CATIA.ActiveDocument) = "PartDocument" And Not TypeName(CATIA.ActiveDocument) = "ProductDocument" Then Exit Sub
    On Error GoTo 0
    
    Me.ImgRenderStyle.Visible = False: Me.ImgRenderStyle2.Visible = True
    
    If CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.ProjectionMode = catProjectionConic Then
        CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.ProjectionMode = catProjectionCylindric
        CATIA.StatusBar = "Switch to Perspective"
        Me.ImgRenderStyle.ControlTipText = "Switch to Perspective"
        Me.ImgRenderStyle2.ControlTipText = "Switch to Perspective"
    Else
        CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.ProjectionMode = catProjectionConic
        CATIA.StatusBar = "Switch to Parallel"
        Me.ImgRenderStyle.ControlTipText = "Switch to Parallel"
        Me.ImgRenderStyle2.ControlTipText = "Switch to Parallel"
    End If
    
    Me.Repaint
    Sleep 200
    Call resetToolbar
    
End Sub

Private Sub ImgRenderStyle2_Click()
    Call resetToolbar
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
'                          FIND AND REPLACE SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgFindReplace_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgFindReplace.Top = Me.ImgFindReplace2.Top
    Me.ImgFindReplace.Left = Me.ImgFindReplace2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Find And Replace"
    On Error GoTo 0
    
End Sub

Private Sub ImgFindReplace_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgFindReplace.Visible = False: Me.ImgFindReplace2.Visible = True
    oSelectionRename.FindReplaceOnSelection
    Call resetToolbar
    
End Sub

Private Sub ImgFindReplace2_Click()
    Call resetToolbar
End Sub



