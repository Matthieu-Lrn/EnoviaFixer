Attribute VB_Name = "ReportToolVersion"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private oSeedPartCollection As New Collection
Private sParamList() As String
Private sToolVersionList As String
Private oToolVersionList As Collection

'*************************************************************************
'* The goal of this class is to report the ToolVersion parameter of all
'* tools used in the selected Seed Part
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
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy)
    
    If sSeedPartsList <> "" Then
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else
noSelection:
        Set oSeedAssy = Nothing
        Set sSourceWindow = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set oSeedPartCollection = New Collection
        Exit Sub
    End If
    
    Call ScanParameters
    
End Sub

Private Sub ScanParameters()
    
    Dim oSeedPart As Product
    Dim sLog As String
    Dim i, j, iMaxIndex As Integer
    Dim sString As String
    Dim sMaxString As String
    Dim oParam As Parameter
    Dim sToolNameList As String
    Dim iProgress As Integer
    
    'Set Selection
    Set oSelection = CATIA.ActiveDocument.Selection
    
    'Initial set
    sLog = ""
    
    'Read sToolVersionFile and populate collection oToolVersionList
    'oToolVersionList is a collection of strings of each line in sToolVersionFile
    Set oToolVersionList = New Collection
    sToolVersionList = ReadTextFile(sToolVersionFile, , True)
    For i = LBound(Split(sToolVersionList, vbCrLf)) To UBound(Split(sToolVersionList, vbCrLf))
        oToolVersionList.Add Split(sToolVersionList, vbCrLf)(i)
    Next
    
    Call frmProgress.progressBarInitialize("Report Tools Version")
    iProgress = 0
    
    'We loop all the seed part in the collection
    For Each oSeedPart In oSeedPartCollection
        
        iProgress = iProgress + 1
        Call frmProgress.progressBarRepaint("Scanning " & oSeedPart.PartNumber, oSeedPartCollection.Count, iProgress, "Searching for ToolVersion parameters", 3, 1)
        
        'We scan all the ToolVersion* parameters
        oSelection.Clear
        Set oSelection = KBESearch("(Name=ToolVersion* & CATKnowledgeSearch.InternalParameter),sel", oSeedPart, oSelection)
        
        'Transfert the parameter name and value in the sParamList
        If oSelection.Count > 0 Then
            
            ReDim sParamList(oSelection.Count)
    
            For i = 1 To oSelection.Count
                Set oParam = oSelection.Item(i).Value
                
                'Param local name
                sString = Split(oParam.Name, "\")(UBound(Split(oParam.Name, "\")))
                
                'Add Param value
                sString = sString & "=" & oParam.Value & "|"
                
                'Add Param Full Path
                sString = sString & oParam.Name
                
                'Transfer to sParamList
                sParamList(i) = sString
            Next
        Else
            GoTo NextSeedPart
        End If
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        
        'Sort the sParamList
        Call frmProgress.progressBarRepaint("Scanning " & oSeedPart.PartNumber, oSeedPartCollection.Count, iProgress, "Sorting the list", 3, 2)
        
        For i = UBound(sParamList) To 2 Step -1
            sMaxString = sParamList(i)
            iMaxIndex = i
        
            'Find at largest string
            For j = i - 1 To 1 Step -1
                If sMaxString < sParamList(j) Then
                    sMaxString = sParamList(j)
                    iMaxIndex = j
                End If
            Next
            
            'Swap string
            If iMaxIndex <> i Then
                sParamList(iMaxIndex) = sParamList(i)
                sParamList(i) = sMaxString
            End If
        Next
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
        
        'Transfert to log file
        Call frmProgress.progressBarRepaint("Scanning " & oSeedPart.PartNumber, oSeedPartCollection.Count, iProgress, "Generating report", 3, 3)
        
        sToolNameList = ""
        sLog = sLog & vbCrLf & oSeedPart.PartNumber
        i = 1
        Do While i <= UBound(sParamList)
                  
            If i = 1 Then
                'Local Parameter name with value
                sString = Split(sParamList(i), "|")(0)
                sLog = sLog & vbCrLf & " - " & sString & CountNbInstance(sString, i) & CompareWithToolVersionFile(sString)
                sToolNameList = sToolNameList & "|" & Split(sParamList(i), "|")(0)
                
                'Global Parameter name
                sString = Replace(sParamList(i), Split(sParamList(i), "|")(0), "")
                sString = Replace(sString, "|", "")
                sLog = sLog & vbCrLf & "    - " & sString
            Else
                If Not Split(sParamList(i), "|")(0) = Split(sParamList(i - 1), "|")(0) Then
                    'Local Parameter name with value
                    sString = Split(sParamList(i), "|")(0)
                    sLog = sLog & vbCrLf & " - " & sString & CountNbInstance(sString, i) & CompareWithToolVersionFile(sString)
                    sToolNameList = sToolNameList & "|" & Split(sParamList(i), "|")(0)
                End If
                
                'Global Parameter name
                sString = Replace(sParamList(i), Split(sParamList(i), "|")(0), "")
                sString = Replace(sString, "|", "")
                sLog = sLog & vbCrLf & "    - " & sString
                
            End If
            
            i = i + 1
        Loop
        
        'Add a summary of all the tools in the Seed Part
        sLog = sLog & vbCrLf
        sLog = sLog & vbCrLf & "Summary of all tools in " & oSeedPart.PartNumber
        For i = 1 To UBound(Split(sToolNameList, "|"))
            sLog = sLog & vbCrLf & Split(sToolNameList, "|")(i)
        Next
        sLog = sLog & vbCrLf
        
NextSeedPart:
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endSub
    Next
    
endSub:
    'Display Log file in Notepad
    If sLog <> "" Then
        sLog = "Tool Version Report:" & vbCrLf & sLog
        sLogFileName = sTempDirectory & "LogFile_KBECheck_ToolVersion.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    Else
        MsgBox ("No actions were performed")
    End If
    
    CATIA.DisplayFileAlerts = True
    frmKBEMain.resetToolbar
    
End Sub

Private Function CompareWithToolVersionFile(ByVal sString As String) As String
    
    Dim sParamName As String
    Dim sParamValue As String
    Dim i As Integer
    Dim sMessage As String
    
    sParamName = Split(sString, "=")(0)
    sParamValue = Split(sString, "=")(1)
    
    sMessage = "  -  Tool was not found in file"
    For i = 1 To oToolVersionList.Count
        If oToolVersionList.Item(i) Like sParamName & "*" Then
            If oToolVersionList.Item(i) Like "*" & sParamValue & "*" Then
                sMessage = "  -  Up to date"
            Else
                sMessage = "  -  WARNING: Tool version is not as per file"
            End If
            
            Exit For
        End If
    Next
    
    CompareWithToolVersionFile = sMessage
    
End Function

Private Function CountNbInstance(ByVal sString As String, ByVal iIndex As Integer) As String
    
    'Counts the number of time a specific tool is used in a given Seed Part
    Dim iCount As Integer
    
    iCount = 0
    Do While sParamList(iIndex) Like sString & "*"
        iCount = iCount + 1
        iIndex = iIndex + 1
        
        If iIndex > UBound(sParamList) Then
            Exit Do
        End If
    Loop
    
    CountNbInstance = "  -  " & iCount & " instances"
    
End Function
