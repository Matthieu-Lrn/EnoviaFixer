Attribute VB_Name = "clsExcel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************
'* SCRIPT VERSION: August 12th 2016
'*
'* When modifying this code please update in all toolbars
'********************************************************

Private oExcelApp As Object
Private bExcelIsOpen As Boolean
Private bExcelFileWasAlreadyOpen As Boolean
Private iNbProcess As Integer

Public Sub GetExcel()
    
'Note: pour que ce Sub fonctionne dans le menu Tools + Reference de l'éditeur VBA il faut que:
'Microsoft Office x.x Object Library
'Microsoft Excel x.x Object Library
'soient sélectionnés

On Error Resume Next

'Si Excel est déjà ouvert on le retrouve
Set oExcelApp = Nothing
Set oExcelApp = GetObject(, "Excel.Application")

'Si Excel n'est pas ouvert on lance l'application
If oExcelApp.Value <> "Microsoft Excel" Then
    Err.Clear
    Set oExcelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        MsgBox "Impossible to open Excel, please open manualy."
        End
    End If
End If
End Sub
Public Sub GetNewExcel()
    On Error Resume Next
    
    'Create a new Excel
    Set oExcelApp = Nothing
    Set oExcelApp = CreateObject("Excel.Application")
    
    'Si Excel n'est pas ouvert on lance l'application
    If oExcelApp.Value <> "Microsoft Excel" Then
        MsgBox "Impossible to open Excel, please open manualy."
    End If
    On Error GoTo 0
End Sub
Public Sub GetExistingExcel()
    On Error Resume Next
    
    'Get Existing Excel
    Set oExcelApp = Nothing
    Set oExcelApp = GetObject(, "Excel.Application")

    On Error GoTo 0

End Sub

Public Function GetNbExcelProcesses() As Integer

Dim objProcesses As Object
Set objProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2").ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & "excel.exe" & "'")

GetNbExcelProcesses = objProcesses.Count

End Function
Public Sub OpenExcelFile(Optional sFileName As String, Optional sMessage As String)

Dim sFileToOpen As String
Dim oWorkbook 'As Workbook
Dim iNbOpenWorkbook As Integer

'Check if file is already open
On Error Resume Next
iNbOpenWorkbook = 0
iNbOpenWorkbook = oExcelApp.Workbooks.Count
If Err.Number = 0 And iNbOpenWorkbook <> 0 Then
    For Each oWorkbook In oExcelApp.Workbooks
        If sFileName Like "*" & oWorkbook.Name Then
            oWorkbook.Activate
            oExcelApp.DisplayAlerts = False
            bExcelFileWasAlreadyOpen = True
            bExcelIsOpen = True
            Exit Sub
        End If
    Next
End If

'Open file
If sFileName = "" Then
    sFileToOpen = oExcelApp.GetOpenFileName()
    If sFileToOpen = "False" Then
        bExcelIsOpen = False
        Exit Sub
    End If
    oExcelApp.Workbooks.Open Filename:=sFileToOpen
    bExcelFileWasAlreadyOpen = False
Else
    oExcelApp.Workbooks.Open Filename:=sFileName
    bExcelFileWasAlreadyOpen = False
End If
bExcelIsOpen = True

End Sub

Public Function IsOneWkbOpen() As Boolean

If oExcelApp.Workbooks.Count = 0 Then
    IsOneWkbOpen = False
Else
    IsOneWkbOpen = True
End If

End Function

Public Sub NewExcelFile()

oExcelApp.Workbooks.Add

End Sub

Public Sub SaveWorkbook(ByVal sFileName As String)

oExcelApp.Application.DisplayAlerts = False
oExcelApp.ActiveWorkbook.SaveAs sFileName

End Sub
Public Sub ShowExcelWindow()

oExcelApp.Application.Visible = True

End Sub
Public Sub CloseExcelFile()

oExcelApp.DisplayAlerts = False
oExcelApp.ActiveWorkbook.Close
DoEvents
oExcelApp.DisplayAlerts = True

End Sub

Public Property Get App() As Object
Set App = oExcelApp
End Property

Public Property Get IsOpen() As Boolean
IsOpen = bExcelIsOpen
End Property

Public Property Get WasOpen() As Boolean
WasOpen = bExcelFileWasAlreadyOpen
End Property

Public Sub RunMacro(sMacroName As String)

oExcelApp.Application.Run (sMacroName)

End Sub

Public Sub CloseAlreadyOpenWorkbook(ByVal sFileName As String)

oExcelApp.DisplayAlerts = False

If oExcelApp.Workbooks.Count <> 0 Then
    For Each oWorkbook In oExcelApp.Workbooks
        If sFileName Like oWorkbook.FullName Then
            oWorkbook.Close
            Exit For
        End If
    Next
End If

oExcelApp.DisplayAlerts = True

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not oExcelApp Is Nothing Then
        If oExcelApp.Workbooks.Count > 0 Then Exit Sub
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
        DoEvents
    End If
    On Error GoTo 0

End Sub
