Attribute VB_Name = "SeedPartPosition"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private oFixConstraintCollection As New Collection
Private Const dDeltaPosition As Double = 0.00001

'*************************************************************************
'* The goal of this class is to check that Seed Part are:
'*  - Located at 0,0,0,
'*  - Have a Fix constraint
'*************************************************************************
Public Sub SelectSeedAssy()
    
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
    
    If Not oSeedAssy Is Nothing Then Call CheckSeedPartPosition
    
noSelection:
    Set oSeedAssy = Nothing
    
End Sub

Private Sub PopulateFixConstraintCollection()
'This Sub will populate the oFixConstraintCollection with all the fix constraints

Dim oConstraint As Constraint
Dim oConstraints As Constraints

Set oFixConstraintCollection = Nothing
Set oConstraints = oSeedAssy.Connections("CATIAConstraints")

'Make sure there is at least one constraint
If oConstraints.Count = 0 Then
    Exit Sub
End If

'Loop all the constraints
For Each oConstraint In oSeedAssy.Connections("CATIAConstraints")
    If oConstraint.Type = catCstTypeReference Then
        oFixConstraintCollection.Add oConstraint
    End If
Next

End Sub

Private Function RetrieveConstraint(ByVal sProductName As String) As Constraint

Dim oConstraint As Constraint
Dim sString As String
Dim oReference As Reference

Set RetrieveConstraint = Nothing

'We loop all the constraint
For Each oConstraint In oFixConstraintCollection

    sString = ""
    On Error Resume Next
    Set oReference = oConstraint.GetConstraintElement(1)
    sString = oReference.DisplayName
    On Error GoTo 0
    
    If sString Like "*" & sProductName & "*" Then
        Set RetrieveConstraint = oConstraint
        Exit For
    End If
        
Next

End Function

Private Sub CheckSeedPartPosition()

Dim oProduct As Product
Dim sLog As String
Dim dPositionMatrix(11)
Dim oProdPosition 'As Position
Dim oReference As Reference
Dim i As Integer
Dim sTargetValue As Double

'Initial set
CATIA.DisplayFileAlerts = False

sLog = ""

Call PopulateFixConstraintCollection

'We loop the Seed Assy
For Each oProduct In oSeedAssy.Products

    'We only check Seed Part or MONUMENTORIGINPART-KBE
    If oProduct.Name Like "*SEED_PART*" Or oProduct.Name Like "MONUMENTORIGINPART*" Then
        'Check for a Fix constraint
        Set oConstraint = RetrieveConstraint(oProduct.Name)
        If oConstraint Is Nothing Then
            sLog = sLog & vbCrLf & " - Fix constraint is missing for: " & oProduct.Name
        End If
        
        'Check position
        Set oProdPosition = oProduct.Position
        oProdPosition.GetComponents dPositionMatrix
    
        For i = 0 To UBound(dPositionMatrix)
    
            If i = 0 Or i = 4 Or i = 8 Then
                sTargetValue = 1
            Else
                sTargetValue = 0
            End If
    
            If dPositionMatrix(i) > sTargetValue + dDeltaPosition Then
                sLog = sLog & vbCrLf & " - Wrong position matrix for: " & oProduct.Name
            ElseIf dPositionMatrix(i) < sTargetValue - dDeltaPosition Then
                sLog = sLog & vbCrLf & " - Wrong position matrix for: " & oProduct.Name
            End If
        Next
    End If
Next

'Display Log file in Notepad
If sLog <> "" Then
    sLog = "The following problems were found:" & vbCrLf & sLog
    sLogFileName = sTempDirectory & "LogFile_KBECheck_SeedPartPosition.txt"
    Call WriteTextFile(sLogFileName, sLog)
    If OpenWithNotepad <> 0 Then
        MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
    End If
Else
    MsgBox ("No error were found")
End If

CATIA.DisplayFileAlerts = True
frmKBEMain.resetToolbar

End Sub
