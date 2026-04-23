Attribute VB_Name = "CreateCapture"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************************************
'* Name: Create_Capture
'* Purpose: The goal of this macro is to create the mandatory captures and assign the respective cameras:
'*          (ALL, PART INFORMATION, ISO, DESIGN NOTES)
'*          Also creates other captures
'*
'* Assumption: None
'*
'* Script Current Version : 1.0
'* Version 1.0 (V5R18 SP6) Comment : None
'*
'* Author:  Francois Charette
'* Language: VBA
'********************************************************************************


Public Sub Create()

Dim oSpecsAndGeomWindow As SpecsAndGeomWindow
Dim oPartDocument As PartDocument
Dim oAnnotationSet As AnnotationSet
Dim oCaptureFactory As CaptureFactory
Dim oCapture As Capture
Dim oCaptures As Captures
Dim oCameras As Cameras
Dim oCamera As Camera3D
Dim oViewer3D As Viewer3D
Dim sNames As String
Dim sName As String
Dim sAnswer As String

'Initial set
On Error Resume Next
If CATIA.ActiveDocument Is Nothing Then GoTo endsub
If Not TypeName(CATIA.ActiveDocument) = "PartDocument" Then
    MsgBox ("Active document must be a CATPart")
    GoTo endsub
End If
On Error GoTo 0

Set oSpecsAndGeomWindow = CATIA.ActiveWindow
Set oViewer3D = oSpecsAndGeomWindow.ActiveViewer
Set oPartDocument = CATIA.ActiveDocument
Set oPart = oPartDocument.Part
Set oPart = CATIA.ActiveDocument.Part

'Annotation set doesn't exist
If oPart.AnnotationSets.Count = 0 Then
    sAnswer = MsgBox("Annotation set doesn't exist. Create it manually and rename." & vbCrLf & "Name should be in line with BM10014.07.03", vbExclamation)
    GoTo endsub
End If
Set oAnnotationSet = oPart.AnnotationSets.Item(1)
Set oCaptureFactory = oAnnotationSet.CaptureFactory
Set oCaptures = oPart.AnnotationSets.Item(1).Captures
Set oCameras = oPartDocument.Cameras


'Mandatory Captures
sAnswer = MsgBox("Do you want to create the mandatory captures ?", vbYesNo)
If sAnswer = vbYes Then
    
    'Captures for DESIGN NOTES
    sNames = "ALL|PART ATTRIBUTES|DESIGN NOTES|ISO"


    For i = LBound(Split(sNames, "|")) To UBound(Split(sNames, "|"))
    
        'Retrieve name
        sName = Split(sNames, "|")(i)
        
        'Check if capture exist
        If FindCapture(oCaptures, sName) = False Then
            Set oCapture = oCaptureFactory.CreateCapture
            oCapture.Name = sName
        Else
            Set oCapture = oCaptures.Item(RetreiveCaptureNumber(oCaptures, sName))
        End If
        
        oCapture.ManageHideShowBody = True
        
        'Check if camera exist
        If FindCamera(oCameras, sName) = False Then
            Set oCamera = oViewer3D.NewCamera()
            oCamera.Name = sName
        Else
            Set oCamera = oCameras.Item(RetreiveCameraNumber(oCameras, sName))
        End If
    
        'Assign camera to capture
        oCapture.Camera = oCamera
    
        iRow = iRow + 1
    Next
Else

End If

On Error Resume Next
oPart.Update
If Err.Number <> 0 Then
    sString = MsgBox("Error when updating Part." & vbCrLf & "Macro aborted", vbCritical)
    GoTo endsub
End If
On Error GoTo 0
    
'Other Captures
Do
    sAnswer = MsgBox("Do you want to create another capture ?", vbYesNo)
    If sAnswer = vbYes Then
        
        'Retrieve capture name
        sName = InputBox("Key in the capture name", "Capture Name")
    
        If sName = "" Then GoTo endsub
        
        'Check if capture exist
        If FindCapture(oCaptures, sName) = False Then
            Set oCapture = oCaptureFactory.CreateCapture
            oCapture.Name = sName
        Else
            Set oCapture = oCaptures.Item(RetreiveCaptureNumber(oCaptures, sName))
        End If
        
        'Check if camera exist
        If FindCamera(oCameras, sName) = False Then
            Set oCamera = oViewer3D.NewCamera()
            oCamera.Name = sName
        Else
            Set oCamera = oCameras.Item(RetreiveCameraNumber(oCameras, sName))
        End If
        
        'Assign camera to capture
        oCapture.Camera = oCamera
        
        oPart.Update
    
        On Error Resume Next
        oPart.Update
        If Err.Number <> 0 Then
            sString = MsgBox("Error when updating Part." & vbCrLf & "Macro aborted", vbCritical)
            GoTo endsub
        End If
        On Error GoTo 0

    Else: GoTo endsub
    End If
Loop

endsub:

'*****Log File
Call AddToLogFile("Create Capture", oPartDocument.Name)

Call frmKBEMain.resetToolbar

End Sub

Private Function FindCapture(oCaptures As Captures, sName As String) As Boolean

Dim i As Integer
Dim bCaptureExist As Boolean
Dim oCapture As Capture

bCaptureExist = False
For i = 1 To oCaptures.Count
    Set oCapture = oCaptures.Item(i)
    If oCapture.Name = sName Then
        bCaptureExist = True
        Exit For
    End If

Next

FindCapture = bCaptureExist

End Function

Private Function FindCamera(oCameras As Cameras, sName As String) As Boolean

Dim i As Integer
Dim bCameraExist As Boolean

bCameraExist = False
For i = 1 To oCameras.Count
    If oCameras.Item(i).Name = sName Then
        bCameraExist = True
        Exit For
    End If
Next

FindCamera = bCameraExist

End Function


Private Function RetreiveCameraNumber(oCameras As Cameras, sName As String) As Integer

Dim i As Integer

For i = 1 To oCameras.Count
    If oCameras.Item(i).Name = sName Then
    Exit For
End If
Next

RetreiveCameraNumber = i

End Function

Private Function RetreiveCaptureNumber(oCaptures As Captures, sName As String) As Integer

Dim i As Integer
Dim oCapture As Capture

For i = 1 To oCaptures.Count
    Set oCapture = oCaptures.Item(i)
    If oCapture.Name = sName Then
        Exit For
    End If
Next

RetreiveCaptureNumber = i

End Function








