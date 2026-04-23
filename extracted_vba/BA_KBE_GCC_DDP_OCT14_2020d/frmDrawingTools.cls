Attribute VB_Name = "frmDrawingTools"
Attribute VB_Base = "0{0143CD56-D2CF-45B9-8658-92DCD0C4A137}{3F5C741F-2E2D-4382-B3D2-EC2EE01F746C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub InitializeForm()


With Me

    .lblTitle.Top = 9
    .lblTitle.Left = 6
    
    .fraViewLock.Left = 6
    .fraViewLock.Top = .lblTitle.Top + .lblTitle.Height + 5
    
    .fraViewFrame.Left = .fraViewLock.Left + .fraViewLock.Width + 5
    .fraViewFrame.Top = .fraViewLock.Top
    
    .lblNote.Top = .fraViewLock.Top + .fraViewLock.Height + 5
    .lblNote.Left = .fraViewLock.Left
    
    .fraDocIteration.Left = .lblNote.Left
    .fraDocIteration.Top = .lblNote.Top + .lblNote.Height + 5
    
    .Height = dTitleBarHeight + .fraDocIteration.Top + .fraDocIteration.Height + 30
    .Width = .fraViewFrame.Left + .fraViewFrame.Width + 10

    .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
    .imgDSGlobe.Left = -3
    .cmdCancel.Top = .Height - dTitleBarHeight - 24
    .cmdCancel.Left = .Width - .cmdCancel.Width - 12
    .cmdOK.Top = .Height - dTitleBarHeight - 24
    .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
    
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
          
    .optViewLockDoNothing.Value = True
    .optDocIterationNothing.Value = True
    .optViewFrameNothing.Value = True

End With

Me.Show vbModless
End Sub


Private Sub cmdCancel_Click()

    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOK_Click()

Dim sViewLock As String
Dim sViewFrame As String
Dim sDocIteration As String
Dim sLogDump As String

    
'Lock
If Me.optViewLockAll.Value = True Then
    sViewLock = "LockAll"
ElseIf Me.optViewUnlockAll.Value = True Then
    sViewLock = "UnlockAll"
Else
    sViewLock = "DoNothing"
End If

'Frame
If Me.optViewFrameShow.Value = True Then
    sViewFrame = "ShowAll"
ElseIf Me.optViewFrameHide.Value = True Then
    sViewFrame = "HideAll"
Else
    sViewFrame = "DoNothing"
End If


'Doc Iteration
If Me.optDocIterationSet = True Then
    sDocIteration = "Set"
Else
    sDocIteration = "DoNothing"
End If

'Log File
sLogDump = "Lock View = " & sViewLock
sLogDump = sLogDump & " | View Frame = " & sViewFrame
sLogDump = sLogDump & " | Doc Iteration = " & sDocIteration
Call AddToLogFile("Drawing Tools", CATIA.ActiveDocument.Name, , , , , , sLogDump)


Call LaunchDrawingTools(sViewLock, sViewFrame, sDocIteration)

Call frmKBEMain.resetToolbar
End Sub




















