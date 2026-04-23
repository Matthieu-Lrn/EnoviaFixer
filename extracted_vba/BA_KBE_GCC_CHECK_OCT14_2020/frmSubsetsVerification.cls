Attribute VB_Name = "frmSubsetsVerification"
Attribute VB_Base = "0{4B938C1F-FD03-4EDA-9176-6512D76D21B7}{FC4ACA18-7D7A-44B9-85F3-6B48B06E3A1F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************************************
'* Name:        Subsets Verification check
'* Purpose:     To verify Subset parameters & list broken ones.
'*              This check will also list duplicate entries in the subsets
'*              It will check all the seed parts selected by user from Seed part List.
'*
'* Assumption:  Seed parts are identified from Subset parameters in following fashion:
'*              KBE_MXX_XXX*_AA*
'*              SP_MXX_XXX*_AA*
'*
'*              AA in above examples are SEED part identifiers, it is identified
'*              as fourth item seperated by an underscore in the string.
'*              frmSeedPartsList is used by this procedure.
'*
'* Author:      Abhishek Kamboj
'* Updated by:
'* Language:    VBA
'********************************************************************************
Private i As Long
Private j As Long
Private SSParameters As Collection           '**** collection having COLLECTIONS of elements of subsets
Private SSName  As Collection                '**** collection having only name of Subsets
Private Booleanlist As Collection            '**** list of all boolean parameters ending with "*_Boolean"
Private missingList As Collection            '**** list of broken parameters in each subset
Private SSListOfDuplicateNames As Collection '**** list of duplicate parameters in each subset
Private sLog As String                       '**** string for log file
Private sLogDup As String                    '**** string for duplicate entries in log file
Private oSeedPartList As Collection          '**** list of seed parts involved in Subsets
Private oDoNotReScanSeedParts As Collection  '**** list of seed parts that are already scanned for boolean parameters
Private oPartSeedinstance As Product
Private oPartSeed As Part
Private iNumberOfSubSets As Integer
Private iMax As Integer
Private isecondmax As Integer
Private iprogresscntr As Integer
Private sDummylist() As String
Private sSeedPartKeyList As Collection
Private sDummyname As String
Private bCheck As Boolean
Private bExitcheck As Boolean                '**** Trigger to exit CheckSubSetOfseedpart sub
Public Sub LaunchVerification()
    bExitcheck = False
    Set oSeedAssy = Nothing
    '**** Asking user to select seed assy****
    Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSeedAssy Is Nothing Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical
            GoTo noSelect
        ElseIf oSeedAssy.Products.Count = 0 Then
            MsgBox "No Seed Parts found!  ", vbCritical
            GoTo noSelect
        End If
    Else
        GoTo noSelect
    End If
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy)

    If frmSeedPartsList.Visible = False And frmSeedPartsList.seedPartsListBox.ListCount > 0 Then '*** if OK was pressed in SeedPartList form
        Set oSeedPartCollection = New Collection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else '*** if Cancel was pressed in SeedPartList form then exit
        GoTo noSelect
    End If
    
    sLog = ""                                   '**** final log string
    Set Booleanlist = New Collection            '**** list of all boolean parameters ending with "*_Boolean"
    Set oSeedPartList = New Collection          '**** list of seed parts involved in Subsets
    Set oDoNotReScanSeedParts = New Collection  '**** list of seed parts that are already scanned for boolean parameters
    If oSeedPartCollection.Count > 0 Then
        For j = 1 To oSeedPartCollection.Count
            Call CheckSubSetOfseedpart(oSeedPartCollection.Item(j).ReferenceProduct.Parent.Part)
            If bExitcheck = True Then '***if user presses cancel on progress toolbar
                 GoTo noSelect
            End If
        Next j
    End If
    
    '******** Display Log file in Notepad *********
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_SSCheck.txt"
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    Else
        MsgBox ("No error was found")
    End If
noSelect:
    Set oSeedAssy = Nothing
    Set oSeedPartInstance = Nothing
    Set oSeedPartRootDocument = Nothing
    CATIA.DisplayFileAlerts = True
End Sub

Private Sub CheckSubSetOfseedpart(oCurrentSeedPart As Part)
    On Error Resume Next
    
    '**** Scan the reference Part for the SubSetparameters with the prefix of SS_
    '**** Two names of parameter set to cover: "SubSetParameters" & "ParametersSubSet"
    iNumberOfSubSets = 0
    Set SSParameters = New Collection
    Set sSeedPartKeyList = New Collection
    Set SSListOfDuplicateNames = New Collection
    Set SSName = New Collection
    Set oParameterSet = Nothing
    Set oParameterSet = oCurrentSeedPart.Parameters.RootParameterSet.ParameterSets.Item("SubSetParameters")
    Set oParameterSet = oCurrentSeedPart.Parameters.RootParameterSet.ParameterSets.Item("ParametersSubSet")
    iNumberOfSubSets = oParameterSet.DirectParameters.Count
    If Not (oParameterSet Is Nothing) Then
        For i = 1 To oParameterSet.DirectParameters.Count
            Set oCurrentSSParameter = oParameterSet.DirectParameters.Item(i)
            sCurrentSSParameterName = Mid(oCurrentSSParameter.Name, InStrRev(oCurrentSSParameter.Name, "\SS_") + Len("\SS_"))
            SSName.Add sCurrentSSParameterName, sCurrentSSParameterName  '**** collection having names of Subsets
            SSParameters.Add New Collection, sCurrentSSParameterName     '**** collection having COLLECTION of elements of subsets
            SSListOfDuplicateNames.Add New Collection, sCurrentSSParameterName
            For Each iObject In ConvertStringToCollection(oCurrentSSParameter.ValueAsString, "|")
                If InStr(iObject, "_Boolean") = 0 Then
                    iObject = iObject + "_Boolean"
                End If
                Err.Clear
                SSParameters.Item(sCurrentSSParameterName).Add iObject, iObject
                If Err.Number <> 0 Then                                 '**** to verify double entries
                    SSListOfDuplicateNames.Item(sCurrentSSParameterName).Add iObject
                End If
                '**** populating possible list of involved seed parts****
                sDummylist = Split(iObject, "_")
                If UBound(sDummylist) >= 3 Then
                    sDummyname = Left(sDummylist(3), 2)         '**** extracts identification key of Seed Part Name
                    sSeedPartKeyList.Add sDummyname, sDummyname '**** duplicate entries will be ignored
                End If
            Next
        Next
    End If
    
    '******** if no subsets are found then exit sub*********
    If iNumberOfSubSets = 0 Then
        GoTo noSelection
    End If
    
    '******** searching for different seedparts involved in SS*********
    For Each oPartSeedinstance In oSeedAssy.Products
         sDummylist = Split(oPartSeedinstance.Name, "_")
         bCheck = False
         For Each iObject In sSeedPartKeyList
            If UBound(sDummylist) >= 4 Then
                If Left(sDummylist(4), 2) = iObject Then
                 bCheck = True
                End If
            End If
         Next
        If bCheck = True Then
            Err.Clear
            oPartSeedinstance.ApplyWorkMode DESIGN_MODE '**** ensuring that part is in design mode
            oSeedPartList.Add oPartSeedinstance.ReferenceProduct.Parent.Part, oPartSeedinstance.ReferenceProduct.Parent.Part.Name
        End If
     Next
    '******** searching parameters in seed part in which SubSet exist*********
    Call frmProgress.progressBarInitialize("SubSetVerification")
    frmProgress.Show vbModeless
    iMax = 0
    For Each oPartSeed In oSeedPartList
        iMax = iMax + oPartSeed.Parameters.Count        '**** Max value for main progress bar progress bar
    Next
    Set oParameter = Nothing
    iprogresscntr = 1                                   '**** counter for progress bar
    '***check Do Not rescan List
            Err.Clear
            bCheck = False
            dummyname = oDoNotReScanSeedParts.Item(oCurrentSeedPart.Name).Name
            If Err.Number = 0 Then
                bCheck = True
            End If
    '***end check Do Not rescan List
    If bCheck = False Then '***seedpart is not yet scanned
        For Each oParameter In oCurrentSeedPart.Parameters
            If TypeName(oParameter) = "BoolParam" And oParameter.Name Like "*_Boolean" Then
                sDummylist = Split(oParameter.Name, "\")
                sDummyname = sDummylist(UBound(sDummylist))
                Booleanlist.Add sDummyname, sDummyname
            End If
            If iprogresscntr / 100 = Int(iprogresscntr / 100) Then
                Call frmProgress.progressBarRepaint("Scanning " + oCurrentSeedPart.Name, iMax, iprogresscntr)
                If bCancelAction = True Then '***if user presses cancel on progress toolbar
                    bExitcheck = True
                    GoTo noSelection
                End If
            End If
             iprogresscntr = iprogresscntr + 1
        Next
        oDoNotReScanSeedParts.Add oCurrentSeedPart, oCurrentSeedPart.Name '*** add seed part to Do Not rescan list
    End If
    On Error GoTo 0
    
    '******** searching for broken automation boolean parameters in seed part of subsets *********
    
    Set missingList = New Collection
    For i = 1 To iNumberOfSubSets
       sDummyname = vbCrLf & "  " + "SubSet_" + SSName.Item(i)
       missingList.Add sDummyname
       For Each iObject In SSParameters.Item(i)
           On Error Resume Next
           Err.Clear
           sDummyname = Booleanlist.Item(iObject)
           If Err.Number <> 0 Then
               missingList.Add iObject
           End If
           On Error GoTo 0
       Next
    Next
    
    '******** If all parameters are not found in parent seedpart then *********
    If missingList.Count > iNumberOfSubSets Then
        On Error Resume Next
        For Each oPartSeed In oSeedPartList             '*** searching for boolean parameters in all seed parts except from parent seed part
            '*** check Do Not rescan List
            Err.Clear
            bCheck = False
            dummyname = oDoNotReScanSeedParts.Item(oPartSeed.Name).Name
            If Err.Number = 0 Then
                bCheck = True
            End If
            '*** end check Do Not rescan List
            If bCheck = False Then '*** subset seedpart is not yet scanned
                isecondmax = oPartSeed.Parameters.Count
                i = 1                                   '**** counter for secondary progress bar
                For Each oParameter In oPartSeed.Parameters
                    If TypeName(oParameter) = "BoolParam" And oParameter.Name Like "*_Boolean" Then
                        sDummylist = Split(oParameter.Name, "\")
                        sDummyname = sDummylist(UBound(sDummylist))
                        Booleanlist.Add sDummyname, sDummyname
                    End If
                    If iprogresscntr / 100 = Int(iprogresscntr / 100) Then
                        Call frmProgress.progressBarRepaint("Scanning for  Automation Parameters ", iMax, iprogresscntr, "Searching in " + oPartSeed.Name, isecondmax, i)
                        If bCancelAction = True Then '***if user presses cancel on progress toolbar
                            bExitcheck = True
                            GoTo noSelection
                        End If
                    End If
                    iprogresscntr = iprogresscntr + 1
                    i = i + 1
                Next
             End If
             oDoNotReScanSeedParts.Add oPartSeed, oPartSeed.Name '*** add seed part to Do Not rescan list
        Next
        On Error GoTo 0
    End If
    
    '******** searching for broken automation boolean parameters in all seedparts used by subsets *********
    On Error Resume Next
    Set missingList = New Collection '**** re initialising missing list collection
    For i = 1 To iNumberOfSubSets
       sDummyname = vbCrLf & "  " + "SubSet_" + SSName.Item(i)
       missingList.Add sDummyname
       For Each iObject In SSParameters.Item(i)
           Err.Clear
           sDummyname = Booleanlist.Item(iObject)
           If Err.Number <> 0 Then
               missingList.Add iObject
           End If
       Next
    Next
    On Error GoTo 0
    frmProgress.Hide
    '******** populating Log string *********
    
    sLogDup = ""    '*** string having duplicate parameters list
    For i = 1 To iNumberOfSubSets
        If SSListOfDuplicateNames.Item(SSName.Item(i)).Count > 0 Then
            sLogDup = sLogDup + vbCrLf + "  SubSet_" + SSName.Item(i) + vbCrLf
            For Each iObject In SSListOfDuplicateNames.Item(SSName.Item(i))
                sLogDup = sLogDup + "   " + iObject + vbCrLf
            Next
        End If
    Next

    If missingList.Count > iNumberOfSubSets Or sLogDup <> "" Then
        sLog = sLog + "Seed Part: " + oCurrentSeedPart.Name + vbCrLf
        If missingList.Count > iNumberOfSubSets Then '*** if missing parameters exist
            sLog = sLog + " Broken Parameter(s) in Subset(s):" + vbCrLf
            For Each iObject In missingList
             sLog = sLog + "   " + iObject & vbCrLf
            Next
        End If
        If sLogDup <> "" Then '*** if duplicate parameters exist
            sLog = sLog + vbCrLf + " Duplicate Parameter(s) in SubSet(s):" + vbCrLf & sLogDup
        End If
        sLog = sLog + vbCrLf + "-------------------------------------------------------" + vbCrLf
    End If
noSelection:
End Sub


