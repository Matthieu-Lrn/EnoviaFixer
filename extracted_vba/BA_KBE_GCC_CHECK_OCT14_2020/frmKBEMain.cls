Attribute VB_Name = "frmKBEMain"
Attribute VB_Base = "0{934412CD-52FD-4B69-B569-770A3AF70570}{FDC9742E-823E-4134-9F35-B29BEB2FEB74}"
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
Private oPartBodyInWorkObject As New PartBodyInWorkObject
Private oActivateExtObject As New ActivateExternalObject
Private oBooleanAxisNomenclature As New BooleanAxisNomenclature
Private oStdPartLink As New StdPartLink
Private oBoolVsKBEBodyVsATBParam As New BooleanVsKBEBodyVsATBParam
Private oCheckStdPartsAndConstraints As New CheckStdPartsAndConstraints
Private oSeedPartPosition As New SeedPartPosition
Private oHideElement As New HideElements
Private oCheckAttribut As New CheckAttribute
Private oCheckParameter As New CheckParameter
Private oCheckRevision As New CheckRevision
Private oBodyColor As New BodyColor
Private oObjectToBeDeleted As New ObjectToBeDeleted
Private oParamPath As New CheckParamPath
Private oReportTV As New ReportToolVersion
Private oCheckRules As New CheckRules
Private oCheckVolume As New CheckVolume


'********************************************************************************
'* Name: Check Toolbar
'* Purpose: Creates a toolbar. Size is automatically defined by the number of buttons.
'*          There's a main section for the toolbar itself (e.g. size, positioning of all buttons, etc..)
'*          and a specific section for each button (e.g. what happens when you press a button)
'*
'* Assumption: Each button must be composed of two images:
'*             A first image represents the active button (default), and must be named "ImgXXX" where XXX is the command name
'*             A second image represents the inactive button (orange), and must be named "ImgXXX2" where XXX is the command name
'*
'* Author: François Charette
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
    
    On Error Resume Next
    '*****Catia status bar value
    CATIA.StatusBar = sScriptVersion & " - Welcome to KBE Check Tools"
    
    '*****Toolbar title bar modification (toolbar look)*****
    Me.Caption = sScriptVersion & " - Check Tools"
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
        Me.Top = iToolbarTop '***** Bug: Need to do it twice
        Me.Top = iToolbarTop
        Me.Left = iToolbarLeft
    End If
    
    '*****Set controls position and visibility
    Call resetToolbar
    
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
        
        '***** 1 - Define PartBody as In Work Object
        setButtonPosition Me.ImgInWorkObject, Me.ImgInWorkObject2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 2 - Activate External Objects
        setButtonPosition Me.ImgActivateExtObj, Me.ImgActivateExtObj2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 3 - Check SP/SA Boolean Parameter and Std Part axis nomenclature
        setButtonPosition Me.ImgBooleanAxisNomenclature, Me.ImgBooleanAxisNomenclature2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 4 - Check Std Part linked objects
        setButtonPosition Me.ImgStdPartLink, Me.ImgStdPartLink2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 5 - Check Std Parts and Coincidence Constraints
        setButtonPosition Me.ImgSTD, Me.ImgSTD2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 6 - Compare Boolean Parameters, KBE Bodies and ATB Parameter
        setButtonPosition Me.ImgCompareParamBody, Me.ImgCompareParamBody2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 7 - Hide Elements
        setButtonPosition Me.ImgHide, Me.ImgHide2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 8 - Check the Seed Part Position in the Seed Assy
        setButtonPosition Me.ImgSeedPosition, Me.ImgSeedPosition2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 9 - Check attributs
        setButtonPosition Me.ImgCheckAttribute, Me.ImgCheckAttribute2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 10 - Check parameters
        setButtonPosition Me.ImgCheckParameter, Me.ImgCheckParameter2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 11 - Check document revision
        setButtonPosition Me.ImgCheckRevision, Me.ImgCheckRevision2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 12 - Body Color
        setButtonPosition Me.ImgBodyColor, Me.ImgBodyColor2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 13 - Object to be deleted
        setButtonPosition Me.ImgObjectToBeDeleted, Me.ImgObjectToBeDeleted2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 14 - Seed Option Parameter Path
        setButtonPosition Me.ImgParamPath, Me.ImgParamPath2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 15 - Report Tools Version
        setButtonPosition Me.ImgReportTV, Me.ImgReportTV2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 16 - Check Seed Assy Rules
        setButtonPosition Me.ImgCheckRules, Me.ImgCheckRules2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 17 - Check KBE Body Volume
        setButtonPosition Me.ImgCheckVolume, Me.ImgCheckVolume2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** 18 - Verify Subsets
        setButtonPosition Me.ImgVerifySS, Me.ImgVerifySS2, iButtonCount
        iButtonCount = iButtonCount + 1
        
        '***** ii - XXXXXXXXXXXXXXXXXXX ADD NEW BUTTON
        'setButtonPosition Me.ImgMYNEWBUTTON, Me.ImgMYNEWBUTTON2, iButtonCount
        'iButtonCount = iButtonCount + 1
        
    '***** Toolbar size *****
    Me.Width = iButtonLeft + iButtonCount * iSpacing + 5
    Me.Height = iButtonTop + iSpacing + 18 '***** Catia toolbar heigth
    
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
    Call resetToolbar
    Unload Me
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Define PartBody as In Work Object
'
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgInWorkObject_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgInWorkObject.Top = Me.ImgInWorkObject2.Top
    Me.ImgInWorkObject.Left = Me.ImgInWorkObject2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Define PartBody as In Work Object"
    On Error GoTo 0
    
End Sub

Private Sub ImgInWorkObject_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgInWorkObject.Visible = False: Me.ImgInWorkObject2.Visible = True
    oPartBodyInWorkObject.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgInWorkObject2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                          Activate External Objects
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgActivateExtObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgActivateExtObj.Top = Me.ImgActivateExtObj2.Top
    Me.ImgActivateExtObj.Left = Me.ImgActivateExtObj2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Activate External References"
    On Error GoTo 0
    
End Sub

Private Sub ImgActivateExtObj_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgActivateExtObj.Visible = False: Me.ImgActivateExtObj2.Visible = True
    oActivateExtObject.SelectSeedAssy
    Call resetToolbar
End Sub

Private Sub ImgActivateExtObj2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'     Check SP/SA Boolean Parameters and Std Part Axis nomenclature
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgBooleanAxisNomenclature_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgBooleanAxisNomenclature.Top = Me.ImgBooleanAxisNomenclature2.Top
    Me.ImgBooleanAxisNomenclature.Left = Me.ImgBooleanAxisNomenclature2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Compare Boolean Parameters, Std Part Axis and Coincidence Constraints"
    On Error GoTo 0
    
End Sub

Private Sub ImgBooleanAxisNomenclature_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgBooleanAxisNomenclature.Visible = False: Me.ImgBooleanAxisNomenclature2.Visible = True
    oBooleanAxisNomenclature.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgBooleanAxisNomenclature2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                  Check Std Part linked objects
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgStdPartLink_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgStdPartLink.Top = Me.ImgStdPartLink2.Top
    Me.ImgStdPartLink.Left = Me.ImgStdPartLink2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Std Part linked objects"
    On Error GoTo 0
    
End Sub

Private Sub ImgStdPartLink_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    If MsgBox("Before running this check all the Std Parts should be active." & vbCrLf & "Do you want to continue ?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Using the Task Manager kill all Excel process before running this macro." & vbCrLf & "Do you want to continue ?", vbYesNo) = vbNo Then Exit Sub
    
    Me.ImgStdPartLink.Visible = False: Me.ImgStdPartLink2.Visible = True
    oStdPartLink.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgStdPartLink2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       Check Std Parts and Coincidence Constraints
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSTD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSTD.Top = Me.ImgSTD2.Top
    Me.ImgSTD.Left = Me.ImgSTD2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Std Parts and Coincidence Constraints"
    On Error GoTo 0
    
End Sub

Private Sub ImgSTD_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    If MsgBox("Before running this check all the Std Parts should be active" & vbCrLf & "Do you want to continue ?", vbYesNo) = vbNo Then Exit Sub
    
    Me.ImgSTD.Visible = False: Me.ImgSTD2.Visible = True
    oCheckStdPartsAndConstraints.SelectSeedAssy
    Call resetToolbar
   
End Sub

Private Sub ImgSTD2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'           Compare Boolean Parameter, KBE Bodies and ATB Param
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCompareParamBody_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCompareParamBody.Top = Me.ImgCompareParamBody2.Top
    Me.ImgCompareParamBody.Left = Me.ImgCompareParamBody2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Compare Boolean Parameter, KBE Bodies and ATB Param"
    On Error GoTo 0
    
End Sub

Private Sub ImgCompareParamBody_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCompareParamBody.Visible = False: Me.ImgCompareParamBody2.Visible = True
    oBoolVsKBEBodyVsATBParam.SelectSeedAssy
    Call resetToolbar
   
End Sub

Private Sub ImgCompareParamBody2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                              Hide Elements
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgHide_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgHide.Top = Me.ImgHide2.Top
    Me.ImgHide.Left = Me.ImgHide2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Hide Elements"
    On Error GoTo 0
    
End Sub

Private Sub ImgHide_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgHide.Visible = False: Me.ImgHide2.Visible = True
    oHideElement.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgHide2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'          Check the position of the Seed Part in the Seed Assy
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgSeedPosition_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgSeedPosition.Top = Me.ImgSeedPosition2.Top
    Me.ImgSeedPosition.Left = Me.ImgSeedPosition2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check the Seed Part Position in the Seed Assy"
    On Error GoTo 0
    
End Sub

Private Sub ImgSeedPosition_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgSeedPosition.Visible = False: Me.ImgSeedPosition2.Visible = True
    oSeedPartPosition.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgSeedPosition2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                  Check the value of different attributes
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCheckAttribute_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCheckAttribute.Top = Me.ImgCheckAttribute2.Top
    Me.ImgCheckAttribute.Left = Me.ImgCheckAttribute2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check the value of Seed Parts attributs"
    On Error GoTo 0
    
End Sub

Private Sub ImgCheckAttribute_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCheckAttribute.Visible = False: Me.ImgCheckAttribute2.Visible = True
    oCheckAttribut.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgCheckAttribute2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'          Check the value of different parameter
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCheckParameter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCheckParameter.Top = Me.ImgCheckParameter2.Top
    Me.ImgCheckParameter.Left = Me.ImgCheckParameter2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check the value of Seed Option or Seed Config parameters"
    On Error GoTo 0
    
End Sub

Private Sub ImgCheckParameter_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCheckParameter.Visible = False: Me.ImgCheckParameter2.Visible = True
    oCheckParameter.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgCheckParameter2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'          Check the Seed Part documents revision
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCheckRevision_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCheckRevision.Top = Me.ImgCheckRevision2.Top
    Me.ImgCheckRevision.Left = Me.ImgCheckRevision2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Document Revision"
    On Error GoTo 0
    
End Sub

Private Sub ImgCheckRevision_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCheckRevision.Visible = False: Me.ImgCheckRevision2.Visible = True
    oCheckRevision.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgCheckRevision2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Check the color of all KBE Bodies
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgBodyColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgBodyColor.Top = Me.ImgBodyColor2.Top
    Me.ImgBodyColor.Left = Me.ImgBodyColor2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check the color of KBE Bodies"
    On Error GoTo 0
    
End Sub

Private Sub ImgBodyColor_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgBodyColor.Visible = False: Me.ImgBodyColor2.Visible = True
    oBodyColor.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgBodyColor2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Object to be deleted
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgObjectToBeDeleted_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgObjectToBeDeleted.Top = Me.ImgObjectToBeDeleted2.Top
    Me.ImgObjectToBeDeleted.Left = Me.ImgObjectToBeDeleted2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check object that should be deleted"
    On Error GoTo 0
    
End Sub

Private Sub ImgObjectToBeDeleted_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgObjectToBeDeleted.Visible = False: Me.ImgObjectToBeDeleted2.Visible = True
    oObjectToBeDeleted.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgObjectToBeDeleted2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Seed Parameter Path
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgParamPath_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgParamPath.Top = Me.ImgParamPath2.Top
    Me.ImgParamPath.Left = Me.ImgParamPath2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check parameter path"
    On Error GoTo 0
    
End Sub

Private Sub ImgParamPath_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgParamPath.Visible = False: Me.ImgParamPath2.Visible = True
    oParamPath.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgParamPath2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Report Tools Version
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgReportTV_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgReportTV.Top = Me.ImgReportTV2.Top
    Me.ImgReportTV.Left = Me.ImgReportTV2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Report Tools Version"
    On Error GoTo 0
    
End Sub

Private Sub ImgReportTV_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgReportTV.Visible = False: Me.ImgReportTV2.Visible = True
    oReportTV.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgReportTV2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Check Seed Assy Rules
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCheckRules_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCheckRules.Top = Me.ImgCheckRules2.Top
    Me.ImgCheckRules.Left = Me.ImgCheckRules2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check Seed Assy Rules"
    On Error GoTo 0
    
End Sub

Private Sub ImgCheckRules_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCheckRules.Visible = False: Me.ImgCheckRules2.Visible = True
    oCheckRules.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgCheckRules2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                        Check KBE Body Volume
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgCheckVolume_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgCheckVolume.Top = Me.ImgCheckVolume2.Top
    Me.ImgCheckVolume.Left = Me.ImgCheckVolume2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Check KBE Body Volume"
    On Error GoTo 0
    
End Sub

Private Sub ImgCheckVolume_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgCheckVolume.Visible = False: Me.ImgCheckVolume2.Visible = True
    oCheckVolume.SelectSeedAssy
    Call resetToolbar
    
End Sub

Private Sub ImgCheckVolume2_Click()
    Call resetToolbar
End Sub


'******************************************************************************
'******************************************************************************
'******************************************************************************
'                       VERIFY SUBSETS SECTION
'******************************************************************************
'******************************************************************************
'******************************************************************************
Private Sub ImgVerifySS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.ImgVerifySS.Top = Me.ImgVerifySS2.Top
    Me.ImgVerifySS.Left = Me.ImgVerifySS2.Left
    
    On Error Resume Next
    CATIA.StatusBar = "Subsets Verification"
    On Error GoTo 0
    
End Sub

Private Sub ImgVerifySS_Click()
    
    Call resetToolbar
    
    On Error Resume Next
    If CATIA.ActiveDocument Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Me.ImgVerifySS.Visible = False: Me.ImgVerifySS2.Visible = True
    frmSubsetsVerification.LaunchVerification
    Call resetToolbar
    
End Sub

Private Sub ImgVerifySS2_Click()
    Call resetToolbar
End Sub



