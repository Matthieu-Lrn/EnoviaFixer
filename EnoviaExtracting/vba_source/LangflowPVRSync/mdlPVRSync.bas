Attribute VB_Name = "mdlPVRSync"
Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Option Explicit
Private oConfXML As DOMDocument60
Private oPVRXML As DOMDocument60

Private oEffDocReport As DOMDocument60
Private oRelatedPVRReport As DOMDocument60
Private oLogReport As DOMDocument60
Private oAttList As clsAttributesList
Private oRefProductList As clsCollection 'This is a collection of all the Reference Product found in the open PVR.
Private oLoadErrDoc As clsCollection
Private oPartList As clsCollection
Private iRelationID As Long
Private o3DMarkerText As clsTextGenerator
Private sDestinationFolder As String

Private bSaveXML As Boolean
Private sXMLPath As String

Private sErroLogFilePath As String
Private sEffectiveDocReportPath As String

Private oTopNode As clsPartInfo

Public bLangflowAutomated As Boolean

Public sLangflowProject As String

Public bLangflowSyncFromBSF As Boolean

Public sLangflowCIOption As String

Public sLangflowNonCIOption As String


Private Function CheckFileExist(ByVal sFileName As String) As Boolean

Dim oFSO

Set oFSO = CreateObject("Scripting.FileSystemObject")

If oFSO.FileExists(sFileName) = True Then
    CheckFileExist = True
Else
    CheckFileExist = False
End If

End Function


Private Function ExtractProductStructureFromProduct(ByVal oParentPart As clsPartInfo) As clsCollection

Dim oTopProduct As Product
Dim oParentProduct, oChild As Product
Dim sPN As String
Dim dPositionMatrix(11) As Variant
Dim oProdPosition 'As Position
Dim oChildren As clsCollection
Dim oItem As clsCollection
Dim sType As String
Dim sRev As String

'Initialize
Set oChildren = New clsCollection
Set oTopProduct = CATIA.ActiveDocument.Product

'Find the parent product
For Each oParentProduct In oTopProduct.Products

    'Get the Part Number
    On Error Resume Next
    sPN = ""
    sPN = oParentProduct.PartNumber
    If sPN = "" Then
        oParentProduct.ApplyWorkMode DEFAULT_MODE
    End If
    sPN = oParentProduct.PartNumber
    On Error GoTo 0

    If sPN = oParentPart.OriginalPartNumber Then Exit For
    
Next
    
'Extract the product structure from oParentProduct
For Each oChild In oParentProduct.Products

    'Get the Part Number
    On Error Resume Next
    sPN = ""
    sPN = oChild.PartNumber
    If sPN = "" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    sPN = oChild.PartNumber
    On Error GoTo 0

    If sPN <> "" Then
        
        'Get Type
        sType = GetType(oChild)
       
        'Override Part Number with the one found in the file name
        If sType = "CATPart" Or sType = "CATProduct" Then
            sPN = oChild.ReferenceProduct.Parent.Name
            sPN = Left(Split(sPN, ".")(0), Trim(Len(Split(sPN, ".")(0)) - 2))
        End If

        'Position Matrix
        Set oProdPosition = oChild.Position
        oProdPosition.GetComponents dPositionMatrix
    
        'Edit position matrix
        dPositionMatrix(9) = dPositionMatrix(9) / 1000
        dPositionMatrix(10) = dPositionMatrix(10) / 1000
        dPositionMatrix(11) = dPositionMatrix(11) / 1000
        
        'Special case for all parts with same base number as NHA (SP, Skel, MBP....)
        If sPN Like oParentPart.OriginalPartNumber & "*" Or sPN Like oParentPart.DefiningPartAtt & "*" Then
            sRev = oChild.ReferenceProduct.Parent.Name
            sRev = Right(Split(sRev, ".")(0), 2)
        Else
            sRev = "N/A"
        End If
        
        'Package data in oChildren
        Set oItem = New clsCollection
        Call oItem.Add("Part Number", sPN)
        Call oItem.Add("Revision", sRev)
        Call oItem.Add("Position Matrix", dPositionMatrix)
        
        Call oChildren.Add("Instance" & oChildren.Count + 1, oItem)
    
    Else
        Call oLoadErrDoc.Add(oChild.Name, " : Part Number of can't be retrieved.")
    End If
    
Next

'Return
Set ExtractProductStructureFromProduct = oChildren
End Function



'***************************************************************************
'*
'*                                  MAIN
'*
'***************************************************************************
Public Sub PVRSyncMain()

Dim oPVRDoc As ProductDocument
Dim oSelection As Selection
Dim bExitByUser As Boolean, oBool As Boolean
Dim sAnswer As String, sMessage As String, sIter As String, sArray() As String
Dim oWindow As Window
Dim oBDIPartColl As clsCollection
Dim oSettingCtrl As SettingControllers
Dim oGeneralSettingCtrl As GeneralSessionSettingAtt

'Tracking
Set oTracking = New clsTracking
Call oTracking.CreateEmptyTrackFile

'Error Log General Settings
sErroLogFilePath = Environ("temp") & "\PVRSync_ErrorLog.txt"

'XML General Settings
bSaveXML = True
sXMLPath = Environ("temp") & "\"

'Find PVR Report path
sEffectiveDocReportPath = Environ("temp") & "\"

'Start web service access tool
Call StartWebServiceTool

'A window must be open
If CATIA.Windows.Count = 0 Then
    Call MsgBox("Active window must be a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'Check the active window
Set oWindow = CATIA.ActiveWindow
If Not oWindow.Name Like "ENOVIA5\*PVRREF*.CATProduct" Then
    Call MsgBox("Active window must be a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'######################################################
'It seems that this option is not working well on many computers

'Check Load Reference Document option
'Set oSettingCtrl = CATIA.SettingControllers
'Set oGeneralSettingCtrl = oSettingCtrl.Item("CATCafGeneralSessionSettingCtrl")
'oBool = CBool(oGeneralSettingCtrl.RefDoc)
'If oBool = False Then
'    sMessage = "It looks like Tools + Options + General + Load referenced documents is not selected."
'    sMessage = sMessage & vbCrLf & "You may want to select the option, close the loaded PVR and load it again."
'    sMessage = sMessage & vbCrLf & " - Click YES to stop the process and make the changes."
'    sMessage = sMessage & vbCrLf & " - Click NO to continue the process."
'    sAnswer = MsgBox(sMessage, 52, "PVR Sync Tool")
'    If sAnswer = vbYes Then
'        GoTo endsub
'    End If
'End If


'Initialize objects
Call InitializeObjects

'Track
Call oTracking.AddHeader("PVRSyncMain")

'Get the PVR document and retrieve information on the document
Set oPVRDoc = CATIA.ActiveDocument
oTopNode.OriginalPartNumber = oPVRDoc.Product.PartNumber
oTopNode.RefPartNumber = oTopNode.OriginalPartNumber
oTopNode.UsedPartNumber = oTopNode.OriginalPartNumber
oTopNode.PVRDocNb = Left(Split(oPVRDoc.Name, ".")(0), Len(Split(oPVRDoc.Name, ".")(0)) - 2)
oTopNode.PVRDocRev = Right(Split(oPVRDoc.Name, ".")(0), 2)
oTopNode.PVRDocIteration = oAttList.GetEnoviaAttributes(oTopNode.PVRDocNb, oTopNode.PVRDocRev, False, "DOCUMENT_ITERATION")
oTopNode.PVRDocStatus = oAttList.GetEnoviaAttributes(oTopNode.PVRDocNb, oTopNode.PVRDocRev, False, "Revision Status")
oTopNode.PartType = "PVRREF"
oTopNode.Comment = "No comment"
Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Got info from the PVR REF.", oTopNode.PVRDocNb & oTopNode.PVRDocRev)

'Check if the part linked to the PVRREF is a CI
Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Cheking if PVR REF is CI.")
If oPasswordMgr.GetEncryptedPassword = "" Then GoTo endsub



'*****************
'Set oBDIPartColl = GetBDIPartCollection("G25046079-001", "", bExitByUser, True)

'*******************
Set oBDIPartColl = GetBDIPartCollection(oPVRDoc.Product.PartNumber, "", bExitByUser, True)
If oBDIPartColl.Count >= 1 Then
    oTopNode.PartIsCI = oBDIPartColl.GetItem(1).IsCI
    oTopNode.PartTitle = oBDIPartColl.GetItem(1).PartTitle
Else
    oTopNode.PartIsCI = False
End If

'Add text to o3DMarkerText
Call o3DMarkerText.AddEmptyLine
ReDim sArray(0, 1)
sArray(0, 0) = " - Process performed on iteration " & oTopNode.PVRDocIteration + 1 & " of " & oTopNode.PVRDocNb & " " & oTopNode.PVRDocRev
sArray(0, 1) = 0
Call o3DMarkerText.AddNewLine(sArray)

'Progress bar
Call frmProgress.progressBarInitialize("PVR Sync Tool")
Call frmProgress.progressBarRepaint("Starting Process", 7, 0)

'The part linked to the PVRREF is a CI we need to ask user for project number
If oTopNode.PartIsCI = True Then

    Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Part is a CI, asking user for project.")
    
    'Add text to o3DMarkerText
    ReDim sArray(0, 1)
    sArray(0, 0) = " - " & oTopNode.OriginalPartNumber & " is a CI"
    sArray(0, 1) = 0
    Call o3DMarkerText.AddNewLine(sArray)

    If bLangflowAutomated = True Then

        If bLangflowSyncFromBSF = True Then

            Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Langflow selected to update as per BSF.")
            oTopNode.SyncFromBSF = True
            oTopNode.CIOption = "N/A"
            oTopNode.NonCIOption = "N/A"

            ReDim sArray(0, 1)
            sArray(0, 0) = " - Langflow selected to update PVR using best so far structure from ENOVIA"
            sArray(0, 1) = 0
            Call o3DMarkerText.AddNewLine(sArray)

        Else

            oTopNode.Project = UCase(Trim$(sLangflowProject))
            oTopNode.SyncFromBSF = False
            oTopNode.CIOption = sLangflowCIOption
            oTopNode.NonCIOption = sLangflowNonCIOption

            Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Langflow selected project is " & oTopNode.Project)
            Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "CI Option is " & oTopNode.CIOption)
            Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "NonCI Option is " & oTopNode.NonCIOption)

            If Not oTopNode.Project Like "S####" Then
                Err.Raise vbObjectError + 2100, "PVRSyncMain", "The project number format is wrong: " & oTopNode.Project
            End If

            oTopNode.Tail = WebServiceAccessTool.JSON_TailNumberFromProjectNumber(oTopNode.Project, bExitByUser, , oPasswordMgr.GetEncryptedPassword)
            Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Tail Number is " & oTopNode.Tail)
            Call oTracking.Save(True)

            If bExitByUser = True Then
                Err.Raise vbObjectError + 2101, "PVRSyncMain", "Process aborted while retrieving tail number."
            End If

            If Not oTopNode.Tail Like "#####" Then
                Err.Raise vbObjectError + 2102, "PVRSyncMain", "For " & oTopNode.Project & " the tail is " & oTopNode.Tail & "."
            End If

            ReDim sArray(0, 1)
            sArray(0, 0) = " - Langflow selected to update PVR as per " & oTopNode.Project & "/" & oTopNode.Tail
            sArray(0, 1) = 0
            Call o3DMarkerText.AddNewLine(sArray)

        End If

        GoTo LangflowPVRSyncOptionsDone

    End If

    Call frmPVRSync2.InitializeForm(oTopNode.OriginalPartNumber)
    If frmPVRSync2.Cancel = True Then
        Call MsgBox("Process aborted.", vbCritical, "PVR Sync Tool")
        GoTo endsub
    End If

    If frmPVRSync2.SyncWithBSF = True Then
        Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "User selected to update as per BSF.")
        oTopNode.SyncFromBSF = True
        oTopNode.CIOption = "N/A"
        oTopNode.NonCIOption = "N/A"
        Unload frmPVRSync2
                
        'Add text to o3DMarkerText
        ReDim sArray(0, 1)
        sArray(0, 0) = " - User selected to update PVR using best so far structure from ENOVIA"
        sArray(0, 1) = 0
        Call o3DMarkerText.AddNewLine(sArray)

    Else
        
        'Get the project number
        oTopNode.Project = UCase(Trim(frmPVRSync2.Project))
        oTopNode.SyncFromBSF = False
        oTopNode.CIOption = frmPVRSync2.CIOption
        oTopNode.NonCIOption = frmPVRSync2.NonCIOption
        Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Selected project is " & oTopNode.Project)
        Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "CI Option is " & oTopNode.CIOption)
        Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "NonCI Option is " & oTopNode.NonCIOption)
        
        Unload frmPVRSync2
        If Not oTopNode.Project Like "S####" Then
            Call MsgBox("The project number format is wrong. Process aborted.", vbCritical, "PVR Sync Tool")
            GoTo endsub
        End If
        
        'Get the tail number
        oTopNode.Tail = WebServiceAccessTool.JSON_TailNumberFromProjectNumber(oTopNode.Project, bExitByUser, , oPasswordMgr.GetEncryptedPassword)
        Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Tail Number is " & oTopNode.Tail)
        Call oTracking.Save(True)
        
        If bExitByUser = True Then
            Call MsgBox("Process aborted.", vbCritical, "PVR Sync Tool")
            GoTo endsub
        End If
        
        If Not oTopNode.Tail Like "#####" Then
            Call MsgBox("For " & oTopNode.Project & " the tail is " & oTopNode.Tail & ". This doesn't seem be to right. Process aborted.", vbCritical, "PVR Sync Tool")
            GoTo endsub
        Else
            sAnswer = MsgBox("The correponding tail for " & oTopNode.Project & " is " & oTopNode.Tail & ". Do you want to continue ?", 36, "PVR Sync Tool")
            If sAnswer = vbNo Then GoTo endsub
        End If
        
        'Add text to o3DMarkerText
        ReDim sArray(0, 1)
        sArray(0, 0) = " - User selected to update PVR as per " & oTopNode.Project & "/" & oTopNode.Tail
        sArray(0, 1) = 0
        Call o3DMarkerText.AddNewLine(sArray)
        
    End If
Else

    Call oTracking.AddInformationLine("PVRSyncMain", "EXECUTING", "Part is not a CI, extracting from BSF.")
    oTopNode.SyncFromBSF = True
    
    'Add text to o3DMarkerText
    ReDim sArray(0, 1)
    sArray(0, 0) = " - " & oTopNode.OriginalPartNumber & " is not a CI"
    sArray(0, 1) = 0
    Call o3DMarkerText.AddNewLine(sArray)
    sArray(0, 0) = " - PVR updated using best so far structure from ENOVIA "
    sArray(0, 1) = 0
    Call o3DMarkerText.AddNewLine(sArray)
End If

LangflowPVRSyncOptionsDone:

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    Call oTracking.Save(True)
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 1 of 8 - Retrieving Configured Product Structure", 8, 1)

'Get the Configured Product Structure
Call GetConfProductStructure

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    Call oTracking.Save(True)
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 2 of 8 - Saving Configured Structure Report", 8, 2)

'Get the Configured Product Structure
Call oTracking.AddHeader("SAVING LOG REPORT")
Call GenerateLogReport
Call oTracking.Save(True)

'Add text to o3DMarkerText
ReDim sArray(0, 1)
Call o3DMarkerText.AddEmptyLine
sArray(0, 0) = "Process run on " & Format(Now(), "yyyy mm dd") & " at " & Format(Time, "hh:mm:ss")
sArray(0, 1) = 0
Call o3DMarkerText.AddNewLine(sArray)
Environ ("temp") & "\"
sArray(0, 0) = "Refer to " & Environ("temp") & "\PVRSync_ConfiguredStructure_Report_" & oTopNode.OriginalPartNumber & ".xml for full report"
sArray(0, 1) = 0
Call o3DMarkerText.AddNewLine(sArray)

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 3 of 8 - Accessing Part Number", 8, 3)

'Load PVR. Un load global est plus rapide qu'un load à chaque instance sur le TA G25015068-001PVRREF
oPVRDoc.Product.ApplyWorkMode DEFAULT_MODE
  
'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 4 of 8 - Retrieving PVR Product Structure", 8, 4)

'Get the PVR Product Structure
Call GetPVRProductStructure(oPVRDoc.Product)
  
'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 5 of 8 - Renaming components", 8, 5)

'Check if we can rename components
Call RenameComponent

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 6 of 8 - Moving instances", 8, 6)

'Move instances in the PVR structure that are not at the right position
Call MovePVRInstances

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 7 of 8 - Deleting non-required instances in PVR", 8, 7)

'Delete instances in PVR structure that should not be there
Call DeletePVRInstances
Call CleanRefProductList

'Update Progress Bar
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 8 of 8 - Adding missing instances in PVR", 8, 8)

'Add missing instances to PVR structure
Call AddPVRInstances
Call oTracking.Save(True)

'Generate Error log
Call GenerateErrorLog

'Add 3D Marker and log to database
Call Add3DMarker

'Log File
Call AddToLogFile("PVR Sync", , oTopNode.PVRDocNb, oTopNode.PVRDocRev, oTopNode.PVRDocIteration + 1, "OK")

'Ask to save PVR if requires
If oTopNode.PVRDocStatus = "WIP" Then
    Call MsgBox("PVR has been synchronize. Please save the PVR immediately.", 64, "PVR Sync Tool")
End If

'Exit
endsub:
Call frmKBEMain.resetToolbar
End Sub

Sub InitializeObjects()

Dim sArray(0, 1) As String

'Initialize attribute class
Set oAttList = New clsAttributesList

'Initialize oConfXML
Set oConfXML = New DOMDocument60
oConfXML.setProperty "SelectionLanguage", "XPath"
oConfXML.async = True

'Initialize oPVRXML
Set oPVRXML = New DOMDocument60
oPVRXML.setProperty "SelectionLanguage", "XPath"
oPVRXML.async = True

'Initialize oRefProductList
Set oRefProductList = New clsCollection

'Initialize oPartList
Set oPartList = New clsCollection

'Initialize iRelationID
iRelationID = 0

'Load Error Document list
Set oLoadErrDoc = New clsCollection

'Initialize oTopNode
Set oTopNode = New clsPartInfo

'Initialize o3DMarkerText
Set o3DMarkerText = New clsTextGenerator
sArray(0, 0) = "PVR Sync Tool"
sArray(0, 1) = "0"
Call o3DMarkerText.AddNewLine(sArray)

End Sub
Private Sub RenameComponent()

Dim oPVRNodeList As IXMLDOMNodeList
Dim oPVRNode As IXMLDOMNode
Dim oPVRNodeList2 As IXMLDOMNodeList
Dim oPVRNode2 As IXMLDOMNode
Dim oConfNodeList As IXMLDOMNodeList
Dim oConfNode As IXMLDOMNode
Dim oConfNodeList2 As IXMLDOMNodeList
Dim oConfNode2 As IXMLDOMNode
Dim oPVRPartInfo As clsPartNumber
Dim oConfPartInfo As clsPartNumber
Dim bRenamed As Boolean
Dim sNewInstanceName As String, sOldInstanceName As String, sRev As String, sPN As String
Dim oPVRNodeList3 As IXMLDOMNodeList
Dim oPVRNode3 As IXMLDOMNode

'We first loop all components in the PVR and try to find a match in Conf.
'When I find a match I change the ComponentRename attribute to "MatchFound"
Set oPVRNodeList = oPVRXML.SelectNodes("//Instance[@DocType='Component' and @ComponentRename='']")
For Each oPVRNode In oPVRNodeList

    If oPVRNode.Attributes.getNamedItem("ComponentRename").nodeValue = "" Then
    
        'Get info from active oPVRNode
        sPN = oPVRNode.Attributes.getNamedItem("PartNumber").nodeValue
    
        'Check if there is a instance with the same part number in Conf.
        Set oConfNodeList = oConfXML.SelectNodes("//Instance[@DocType='Component' and @PartNumber='" & sPN & "' and @ComponentRename='']")
        
        'If the part number is found in Conf we tag all the instance in PVR and Conf with Part Number = sPN to "MatchFound"
        If oConfNodeList.Length >= 1 Then
            
            Set oPVRNodeList2 = oPVRXML.SelectNodes("//Instance[@DocType='Component' and @PartNumber='" & sPN & "']")
            For Each oPVRNode2 In oPVRNodeList2
                oPVRNode2.Attributes.getNamedItem("ComponentRename").nodeValue = "MatchFound"
            Next
            Set oConfNodeList = oConfXML.SelectNodes("//Instance[@DocType='Component' and @PartNumber='" & sPN & "']")
            For Each oConfNode In oConfNodeList
                oConfNode.Attributes.getNamedItem("ComponentRename").nodeValue = "MatchFound"
            Next
            
        End If
    End If
Next

'Save
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If

'All the instance in PVR with ComponentRename = "" could be renamed. Let's scan them and see.
Do

    'Seach for component where ComponentRename = ""
    Set oPVRNodeList = oPVRXML.SelectNodes("//Instance[@DocType='Component' and @ComponentRename='']")
    If oPVRNodeList.Length = 0 Then Exit Do
    
    'Get the first node in the list
    Set oPVRNode = oPVRNodeList.Item(0)
 
    'Get info from active oPVRNode
    sPN = oPVRNode.Attributes.getNamedItem("PartNumber").nodeValue
    
    'Initialize oPVRPartInfo
    Set oPVRPartInfo = New clsPartNumber
    Set oPVRPartInfo = oPVRPartInfo.Initialize(sPN)
             
    'Only try to find a part with a different dash number when we have a numeric dash
    If oPVRPartInfo.NumericDash <> "" Then
    
            'Initial set
            bRenamed = False

            'Scan all the components in Conf where ComponentRename = ""
            Set oConfNodeList = oConfXML.SelectNodes("//Instance[@DocType='Component' and @ComponentRename='']")
            For Each oConfNode In oConfNodeList
                
                'Initialize oConfPartInfo
                Set oConfPartInfo = New clsPartNumber
                Set oConfPartInfo = oConfPartInfo.Initialize(oConfNode.Attributes.getNamedItem("PartNumber").nodeValue)
                
                'Compare
                If oPVRPartInfo.BaseNumber = oConfPartInfo.BaseNumber And _
                   oConfPartInfo.NumericDash <> "" And _
                   oPVRPartInfo.NumericDashParity = oConfPartInfo.NumericDashParity And _
                   oPVRPartInfo.Suffix = oConfPartInfo.Suffix Then
                                       
                   bRenamed = True
                    
                    'Change Part Number in PVR
                    Call RenameProduct(oPVRPartInfo.PartNumber, oConfPartInfo.PartNumber)
                   
                    'Change Part Number in oPVRXML
                    Set oPVRNodeList2 = oPVRXML.SelectNodes("//Instance[@PartNumber='" & sPN & "']")
                    For Each oPVRNode2 In oPVRNodeList2
                        oPVRNode2.Attributes.getNamedItem("PartNumber").nodeValue = oConfPartInfo.PartNumber
                    Next

                    'Edit key in oRefProductList
                    sRev = oConfNode.Attributes.getNamedItem("DocRev").nodeValue
                    Call oRefProductList.Add(oConfPartInfo.PartNumber & sRev, oRefProductList.GetItemByKey(sPN & "N/A"))
                    Call oRefProductList.RemoveByKey(sPN & "N/A")
                    
                    'Change rev on all instance of oConfPartInfo.PartNumber
                    Set oPVRNodeList2 = oPVRXML.SelectNodes("//Instance[@PartNumber='" & oConfPartInfo.PartNumber & "']")
                    For Each oPVRNode2 In oPVRNodeList2
                        oPVRNode2.Attributes.getNamedItem("DocRev").nodeValue = sRev
                    Next
                    
                    'Change status of all Conf instance with same Part Number to "MatchFound"
                    Set oConfNodeList2 = oConfXML.SelectNodes("//Instance[@PartNumber='" & oConfPartInfo.PartNumber & "']")
                    For Each oConfNode2 In oConfNodeList2
                        oConfNode2.Attributes.getNamedItem("ComponentRename").nodeValue = "MatchFound"
                    Next

                    'Rename instances and change status in PVR XML
                    Do
                        
                        Set oPVRNodeList2 = oPVRXML.SelectNodes("//Instance[@DocType='Component' and @PartNumber='" & oConfPartInfo.PartNumber & "' and @ComponentRename='']")
                        If oPVRNodeList2.Length = 0 Then Exit Do
                    
                        'Get the first node in the list
                        Set oPVRNode2 = oPVRNodeList2.Item(0)
                    
                        'Get old instance name
                        sOldInstanceName = oPVRNode2.Attributes.getNamedItem("InstanceName").nodeValue
                        
                        'Rename instance
                        sNewInstanceName = ChangeInstanceName(oPVRNode2)
                        
                        'Change instance name in XML
                        oPVRNode2.Attributes.getNamedItem("InstanceName").nodeValue = sNewInstanceName
                        
                        'Change status in XML
                        oPVRNode2.Attributes.getNamedItem("ComponentRename").nodeValue = "Renamed"
                        
                        'Update XLM for all other instance with same parent Part Number
                        Set oPVRNodeList3 = oPVRXML.SelectNodes("//Instance[@PartNumber='" & oPVRNode2.parentNode.Attributes.getNamedItem("PartNumber").nodeValue & "']/Instance[@InstanceName='" & sOldInstanceName & "']")
                        For Each oPVRNode3 In oPVRNodeList3
                            oPVRNode3.Attributes.getNamedItem("InstanceName").nodeValue = sNewInstanceName
                            oPVRNode3.Attributes.getNamedItem("ComponentRename").nodeValue = "Renamed"
                        Next
                    Loop
                   
                   
                    'Edit attribut of component
                    Call AddAttributesToComponent(oRefProductList.GetItemByKey(oConfPartInfo.PartNumber & sRev), sRev)
                    Exit For
                End If
            Next
            
            'No match was found. Change the status
            If bRenamed = False Then
                Set oPVRNodeList2 = oPVRXML.SelectNodes("//Instance[@PartNumber='" & sPN & "']")
                For Each oPVRNode2 In oPVRNodeList2
                    oPVRNode2.Attributes.getNamedItem("ComponentRename").nodeValue = "NoMatchFound"
                Next
                
            End If
            
    'Weird part number. We don't rename
    Else
    
        Set oPVRNodeList = oPVRXML.SelectNodes("//Instance[@DocType='Component' and @PartNumber='" & sPN & "']")
        For Each oPVRNode In oPVRNodeList
            oPVRNode.Attributes.getNamedItem("ComponentRename").nodeValue = "NoMatchFound"
        Next
        
    End If
    
Loop

'Save
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If
End Sub
Private Function ChangeInstanceName(ByVal oPVRNode As IXMLDOMElement) As String

Dim sPN As String, sInstanceName As String, sParentPN As String, sParentRev As String
Dim sNewInstanceName As String, sInstanceNameList As String
Dim oParentRefProduct As Product
Dim iInstanceNumber As Integer, i As Integer


'Get info from oPVRNode
sPN = oPVRNode.Attributes.getNamedItem("PartNumber").nodeValue
sInstanceName = oPVRNode.Attributes.getNamedItem("InstanceName").nodeValue
sParentPN = oPVRNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
sParentRev = oPVRNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue

'Get the parent reference product
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

'Compile the list of all the child instance name
sInstanceNameList = "|"
For i = 1 To oParentRefProduct.Products.Count
    sInstanceNameList = sInstanceNameList & oParentRefProduct.Products.Item(i).Name & "|"
Next

'Find the new instance name
iInstanceNumber = 0
Do
    iInstanceNumber = iInstanceNumber + 1
    sNewInstanceName = sPN & "." & iInstanceNumber
    If Not sInstanceNameList Like "*|" & sNewInstanceName & "|*" Then Exit Do
Loop

'Change instance name
oParentRefProduct.Products.Item(sInstanceName).Name = sNewInstanceName

'Return Value
ChangeInstanceName = sNewInstanceName

End Function

Private Sub RenameProduct(ByVal sPN As String, sNewPN As String)

Dim oProductToRename As Product

Set oProductToRename = oRefProductList.GetItemByKey(sPN & "N/A")

oProductToRename.PartNumber = sNewPN

End Sub


Private Sub AddPVRInstances()

Dim oConfNodeList As IXMLDOMNodeList
Dim oConfNode As IXMLDOMNode
Dim sExtension As String, sPN, sRev, sParentPN, sParentRev As String
Dim iTotalQty As Integer, iQty As Integer, iLoopNb As Integer

CATIA.DisplayFileAlerts = False

'Initial set
iTotalQty = 0

'Track
Call oTracking.AddHeader("Add PVR Instances")

'Search for all instances in the Conf that should be added
Do
    iLoopNb = iLoopNb + 1

    'Search all nodes where Status = ""
    Set oConfNodeList = oConfXML.SelectNodes("/Instance//Instance[@SyncStatus='']")
    If oConfNodeList.Length = 0 Then Exit Do

    'Get qty for progesss bar
    If iTotalQty = 0 Then iTotalQty = oConfNodeList.Length
    
    'Update Progress Bar
    If bCancelAction = True Then Exit Sub
    Call frmProgress.progressBarRepaint("Step 8 of 8 - Adding missing instances in PVR", 8, 8, "Adding instance " & iTotalQty - oConfNodeList.Length & " of " & iTotalQty, iTotalQty, iTotalQty - oConfNodeList.Length)

    'Get the first node in the list
    Set oConfNode = oConfNodeList.Item(0)

    'Get info from active oConfNode
    sPN = oConfNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oConfNode.Attributes.getNamedItem("DocRev").nodeValue
    sParentPN = oConfNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
    sParentRev = oConfNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Track
    Call oTracking.AddInformationLine("AddPVRInstances", "EXECUTING", "Loop " & iLoopNb, "Part is " & sPN & " " & sRev)
    
    'Get the extension
    sExtension = oAttList.GetEnoviaAttributes(sPN, sRev, False, "EXTENSION")

    'Create the new object
    If sExtension = "CATPart" Or sExtension = "CATProduct" Then
        Call AddComponentFromFile(sParentPN, sParentRev, sPN, sRev, sExtension, oConfNode)
    Else
        Call AddComponent(sParentPN, sParentRev, sPN, oConfNode, sRev)

    End If
    
Loop

CATIA.DisplayFileAlerts = True
Call oTracking.Save(True)

'Save
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If

End Sub

Private Sub CleanRefProductList()

Dim i As Integer
Dim sKey As String
Dim sPN, sRev As String
Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode

'Scan the list of all Reference Product
For i = oRefProductList.Count To 1 Step -1

    'Get the key
    sKey = oRefProductList.GetKey(i)

    'Get the PN and Rev
    If sKey Like "*N/A" Then
        sRev = "N/A"
        sPN = Left(sKey, Len(sKey) - 3)
    Else
        sPN = Left(sKey, Len(sKey) - 2)
        sRev = Right(sKey, 2)
    End If
    
    'Check if we have a non deleted instance in the PVR structure. If not we delete it from the list.
    Set oNodeList = oPVRXML.SelectNodes("//Instance[@PartNumber='" & sPN & "' and @DocRev='" & sRev & "' and @SyncStatus!='Deleted']")
    If oNodeList.Length = 0 Then
        
        oRefProductList.RemoveByIndex (i)
    
    End If
Next

End Sub
Private Sub DeletePVRInstances()

Dim oPVRNodeList As IXMLDOMNodeList
Dim oPVRNode As IXMLDOMNode
Dim oPVRChildNodeList As IXMLDOMNodeList
Dim oPVRChildNode As IXMLDOMNode
Dim sInstanceName As String, sRev As String, sParentPN As String, sParentRev As String

'Search all instances (except top node) in PVR Product Structure when Status = "". We need to delete these instances.
Do

    'Search all nodes to be deleted
    Set oPVRNodeList = oPVRXML.SelectNodes("/Instance//Instance[@SyncStatus='']")
    If oPVRNodeList.Length = 0 Then Exit Do

    'Get the first node from the list
    Set oPVRNode = oPVRNodeList.Item(0)

    'Get info from active oPVRNode
    sInstanceName = oPVRNode.Attributes.getNamedItem("InstanceName").nodeValue
    sRev = oPVRNode.Attributes.getNamedItem("DocRev").nodeValue
    sParentPN = oPVRNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
    sParentRev = oPVRNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue

    'Remove the instance
    Call RemoveInstance(sParentPN, sParentRev, sInstanceName)

    'Set the status
    Call SetInstancesStatusDelete(sParentPN, sInstanceName)

Loop

'Save
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If

End Sub

Private Sub SetInstancesStatusOKorMoved(ByVal oDoc As DOMDocument60, ByVal sParentPN As String, ByVal sInstanceName As String, ByVal sStatus As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with sStatus
Set oNodeList = oDoc.SelectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "' and @SyncStatus='']")
For Each oNode In oNodeList
    oNode.Attributes.getNamedItem("SyncStatus").nodeValue = sStatus
Next

End Sub

Private Sub SetInstancesStatusKO(ByVal oDoc As DOMDocument60, ByVal sPN As String, ByVal sRev As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode


'All instances with sPN should be set to "KO"
Set oNodeList = oDoc.SelectNodes("//Instance[@PartNumber='" & sPN & "' and @DocRev='" & sRev & "']")
For Each oNode In oNodeList
    oNode.Attributes.getNamedItem("Status").nodeValue = "KO"
Next

End Sub

Private Sub SetInstancesStatusDelete(ByVal sParentPN As String, ByVal sInstanceName As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oChildNodeList As IXMLDOMNodeList
Dim oChildNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with "Deleted"
Set oNodeList = oPVRXML.SelectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "']")
For Each oNode In oNodeList

    'Set Status of active node
    oNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    
    'Set Status to "Deleted" for all children (all levels) of the active node
    Set oChildNodeList = oNode.SelectNodes(".//Instance")
    For Each oChildNode In oChildNodeList
        oChildNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    Next
Next

End Sub

Private Function GetClosestInstance(ByVal oConfNodeList As IXMLDOMNodeList, ByVal oPVRNode As IXMLDOMNode) As IXMLDOMNode

Dim i As Integer
Dim oConfNode As IXMLDOMNode
Dim dClosestDistance As Double
Dim dDistance As Double
Dim iIndex As Integer
Dim dPVRCoord(1 To 3) As Double
Dim dConfCoord(1 To 3) As Double

'Initial set
dClosestDistance = 99999999

'We only have one node in the list
If oConfNodeList.Length = 1 Then
    Set GetClosestInstance = oConfNodeList.Item(0)
    Exit Function
End If

'Get coordinates of PVR instance
dPVRCoord(1) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
dPVRCoord(2) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
dPVRCoord(3) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

'Find the instance which is closest to PVR one
For i = 1 To oConfNodeList.Length
    
    Set oConfNode = oConfNodeList.Item(i - 1)
    
    'Get coordinates of Conf instance
    dConfCoord(1) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
    dConfCoord(2) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
    dConfCoord(3) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

    'Calculate distance
    dDistance = Sqr((dPVRCoord(1) - dConfCoord(1)) ^ 2 + (dPVRCoord(2) - dConfCoord(2)) ^ 2 + (dPVRCoord(3) - dConfCoord(3)) ^ 2)

    'Compare distance
    If dDistance < dClosestDistance Then
        dClosestDistance = dDistance
        iIndex = i - 1
    End If
Next

'Return
Set GetClosestInstance = oConfNodeList.Item(iIndex)

End Function

Private Sub MovePVRInstances()

Dim oPVRNodeList As IXMLDOMNodeList
Dim oPVRNode As IXMLDOMNode
Dim oConfNodeList As IXMLDOMNodeList
Dim oConfNode As IXMLDOMNode
Dim oChildNodeList As IXMLDOMNodeList
Dim oChildNode As IXMLDOMNode
Dim sPN, sRev, sInstanceName, sParentPN, sParentRev, sDocType As String, sRelationID As String

'Search all instances (except top node) in PVR Product Structure and compare the position with Conf instances
Set oPVRNodeList = oPVRXML.SelectNodes("Instance//Instance[@SyncStatus='']")
For Each oPVRNode In oPVRNodeList
        
    'Only the instance that have a Status = ""
    If oPVRNode.Attributes.getNamedItem("SyncStatus").nodeValue = "" Then
    
        'Get info from active oPVRNode
        sPN = oPVRNode.Attributes.getNamedItem("PartNumber").nodeValue
        sRev = oPVRNode.Attributes.getNamedItem("DocRev").nodeValue
        sDocType = oPVRNode.Attributes.getNamedItem("DocType").nodeValue
        sInstanceName = oPVRNode.Attributes.getNamedItem("InstanceName").nodeValue
        sParentPN = oPVRNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
        sParentRev = oPVRNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue
        
        'Search in Conf
        If sDocType = "Component" Then
            Set oConfNodeList = oConfXML.SelectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@PartNumber='" & sPN & "' and @DocType='" & sDocType & "' and @SyncStatus='']")
        Else
            Set oConfNodeList = oConfXML.SelectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@PartNumber='" & sPN & "' and @DocType='" & sDocType & "' and @DocRev ='" & sRev & "' and @SyncStatus='']")
        End If

        'We found an instance that matches the search criteria
        If oConfNodeList.Length >= 1 Then
             
            'Get the closest Conf node
            Set oConfNode = GetClosestInstance(oConfNodeList, oPVRNode)
    
            'Get the RelationID
            sRelationID = oConfNode.Attributes.getNamedItem("RelationID").nodeValue
            
            'Transfer the instance name to all Conf instance with same RelationID
            Call TransferInstanceName(sRelationID, sInstanceName)
    
            'Compare the position
            If ComparePosition(oPVRNode, oConfNode) <> "Same Position" Then
                Call MoveInstance(sParentPN, sParentRev, sInstanceName, oConfNode)
                Call SetInstancesStatusOKorMoved(oPVRXML, sParentPN, sInstanceName, "Moved")
                Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, sInstanceName, "OK")
            Else
                Call SetInstancesStatusOKorMoved(oPVRXML, sParentPN, sInstanceName, "OK")
                Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, sInstanceName, "OK")
            End If
            
        'We didn't find anything. This instance (along with all it's children) will the deleted later. For now I set all it's children to "Deleted".
        Else
            
            Set oChildNodeList = oPVRNode.SelectNodes(".//Instance")
            For Each oChildNode In oChildNodeList
                oChildNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
            Next
        End If
    End If
Next

'Save
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If

End Sub

Private Sub TransferInstanceName(ByVal sRelationID As Long, ByVal sInstanceName As String)

Dim oConfNodeList As IXMLDOMNodeList
Dim oConfNode As IXMLDOMNode

Set oConfNodeList = oConfXML.SelectNodes("//Instance[@RelationID='" & sRelationID & "']")
For Each oConfNode In oConfNodeList
    oConfNode.Attributes.getNamedItem("InstanceName").nodeValue = sInstanceName
Next

End Sub


Private Sub AddComponent(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sNewPN As String, oConfElem As IXMLDOMElement, ByVal sRev As String)

Dim oRefProduct As Product
Dim oParentRefProduct As Product
Dim oNewProduct As Product
Dim sRelationID As String

'Track
Call oTracking.AddInformationLine("AddComponent", "EXECUTING", "Adding component")

'Get the parent reference product
On Error Resume Next
Err.Clear
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("AddComponent", "CODEERROR", "Can't get oParentRefProduct")
End If
On Error GoTo 0

'Check if we can copy an existing component
Set oRefProduct = Nothing
If oRefProductList.Exists(sNewPN & sRev) Then
    Set oRefProduct = oRefProductList.GetItemByKey(sNewPN & sRev)
End If

'Create a new component
If oRefProduct Is Nothing Then

    On Error Resume Next
    Err.Clear

    'Create a new component
    Set oNewProduct = oParentRefProduct.Products.AddNewProduct(sNewPN)
    
    'Add attributes to components
    Call AddAttributesToComponent(oNewProduct, sRev)
    
    'Add reference product to oRefProductList
    Call oRefProductList.Add(sNewPN & sRev, oNewProduct.ReferenceProduct)
    
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Can't create a new component " & sNewPN)
    End If
    On Error GoTo 0
    
'Create a component from an existing one
Else
    
    On Error Resume Next
    Err.Clear

    'Create new component
    Set oNewProduct = oParentRefProduct.Products.AddComponent(oRefProduct)
    
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Can't create a component from an existing one: " & sNewPN)
    End If
    On Error GoTo 0
End If

'Get the RelationID of oConfElem
sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue

'Transfer the instance name to all Conf instance with same RelationID
Call TransferInstanceName(sRelationID, oNewProduct.Name)

'Set status
Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, oNewProduct.Name, "OK")

'Move instance
Call MoveInstance(sParentPN, sParentRev, oNewProduct.Name, oConfElem)

End Sub

Private Sub AddAttributesToComponent(ByVal oProduct As Product, ByVal sRev As String)

Dim oProperty As StrParam
Static oAttToCreate As clsCollection
Dim i As Integer
Dim sParamName As String, sParamNameInList As String, sParamValue As String, sCheckString As String

If oAttToCreate Is Nothing Then

    Set oAttToCreate = New clsCollection

    'The Key is the name of the attribute to be added and the value is the name of the attribute in clsAttributesList
    Call oAttToCreate.Add("Title", "Title")
    Call oAttToCreate.Add("Dataset Type", "Dataset Type")
    Call oAttToCreate.Add("Design Authority Program", "Design Authority Program")
    Call oAttToCreate.Add("Supplier Name And CAGE Code", "Supplier Name And CAGE Code")
    Call oAttToCreate.Add("Major Supplier Code", "Major Supplier Code")
    Call oAttToCreate.Add("3D Only", "3D Only")
    Call oAttToCreate.Add("Color Coded", "Color Coded")
    Call oAttToCreate.Add("PCCN", "PCCN")
    Call oAttToCreate.Add("Material Specifications", "Material Specifications")
    Call oAttToCreate.Add("Material Description", "Material Description")
    Call oAttToCreate.Add("Material Type", "Material Type")
    Call oAttToCreate.Add("Material Form", "Material Form")
    Call oAttToCreate.Add("Size", "Size")
    Call oAttToCreate.Add("Thickness", "Thickness")
    Call oAttToCreate.Add("Inside Diameter", "Inside Diameter")
    Call oAttToCreate.Add("Outside Diameter", "Outside Diameter")
    Call oAttToCreate.Add("Length", "Length")
    Call oAttToCreate.Add("Width", "Width")
    Call oAttToCreate.Add("Wall", "Wall")
    Call oAttToCreate.Add("Alloy", "Alloy")
    Call oAttToCreate.Add("Final Condition", "Final Condition")
    Call oAttToCreate.Add("Density", "Density")
    Call oAttToCreate.Add("Eng. Make From", "Eng. Make From")
    Call oAttToCreate.Add("Form", "Form")
    Call oAttToCreate.Add("Grade/Composition", "Grade/Composition")
    Call oAttToCreate.Add("Material Class", "Material Class")
    Call oAttToCreate.Add("Mesh Cell Size", "Mesh Cell Size")
    Call oAttToCreate.Add("Standard Spec Die", "Standard Spec Die")
    Call oAttToCreate.Add("Type", "Type")
    Call oAttToCreate.Add("TD Material Code", "TD Material Code")
    Call oAttToCreate.Add("Finish Code", "Finish Code")
    Call oAttToCreate.Add("MFG Process", "MFG Process")
    Call oAttToCreate.Add("Defining Part", "Defining Part")
    Call oAttToCreate.Add("Material Specification Production", "Material Specification Production")
    Call oAttToCreate.Add("Material Description Production", "Material Description Production")
    Call oAttToCreate.Add("Organization", "Revision Organization")
    Call oAttToCreate.Add("Project", "Revision Project")

End If

For i = 1 To oAttToCreate.Count
    
    'Get the parameter name
    sParamName = oAttToCreate.GetKey(i)
    sParamNameInList = oAttToCreate.GetItemByIndex(i)
    
    'Check if parameter exits
    On Error Resume Next
    sCheckString = ""
    sCheckString = oProduct.ReferenceProduct.UserRefProperties.Item(sParamName).Name
    On Error GoTo 0
    
    'Parameter doesn't exist we need to create it
    If sCheckString = "" Then
        sParamValue = oAttList.GetEnoviaAttributes(oProduct.PartNumber, sRev, False, sParamNameInList)
        Set oProperty = oProduct.ReferenceProduct.UserRefProperties.CreateString(sParamName, sParamValue)
        
    'Parameter exist, just change the value
    Else
        sParamValue = oAttList.GetEnoviaAttributes(oProduct.PartNumber, sRev, False, sParamNameInList)
        Call oProduct.UserRefProperties.Item(sParamName).ValuateFromString(sParamValue)
    End If
Next

End Sub

Private Sub AddComponentFromFile(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sNewPN As String, ByVal sNewRev As String, sNewExtension As String, oConfElem As IXMLDOMElement)

Dim oPVRWindow As Window
Dim oNewDoc As Document
Dim oParentRefProduct As Product
Dim oNewProduct As Product
Dim sRelationID As String
Dim dLoadTimer As Double
Dim bTimeout As Boolean
Dim EnoviaDoc As EnoviaDocument
Dim EV5product As Product
Dim sLoadStatus As String

'Initialize
Set oPVRWindow = CATIA.ActiveWindow
sLoadStatus = "Already Loaded"

'Connect to API
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call ConnectToAPI
End If
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Can't connect to CATIA API")
End If
On Error GoTo 0

'Get the parent reference product
On Error Resume Next
Err.Clear
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Can't get oParentRefProduct")
End If
On Error GoTo 0

'Check if document already loaded in session.
Set oNewDoc = Nothing
On Error Resume Next
Set oNewDoc = CATIA.Documents.Item(sNewPN & sNewRev & "." & sNewExtension)
On Error GoTo 0

'Check if document is in oLoadErrDoc
If oLoadErrDoc.Exists(sNewPN & sNewRev) Then sLoadStatus = "Load Error"

'Open document from ENOVIA
If oNewDoc Is Nothing And sLoadStatus <> "Load Error" Then

    'Track
    Call oTracking.AddInformationLine("AddComponentFromFile", "EXECUTING", sNewPN & sNewRev & "." & sNewExtension & " will be loaded from ENOVIA.")
    
    'Set Status
    sLoadStatus = "New Load"

    'Loading document
    On Error Resume Next
    Err.Clear
    Set EV5product = EnoviaDoc.OpenPartDocument(sNewPN, sNewRev)
    If Err.Description <> "" Then Call oTracking.AddInformationLine("AddComponentFromFile", "WARNING", "After load the error description is: " & Err.Description)
    
    'Make sure the document is loaded
    dLoadTimer = Timer
    bTimeout = False
    Do
        Set oNewDoc = EV5product.ReferenceProduct.Parent

        If Not oNewDoc Is Nothing Then Exit Do

        DoEvents
        Sleep 150
        Call oTracking.AddInformationLine("AddComponentFromFile", "WARNING", "Sleeping while waiting for document to load")
        If Abs(Timer - dLoadTimer) > 5 Then
            bTimeout = True
            Exit Do
        End If
    Loop
    On Error GoTo 0

    'Timeout reached, add document to oLoadErrDoc
    If bTimeout = True Then
        Call oLoadErrDoc.Add(sNewPN & sNewRev, "AddComponentFromFile")
        Call SetInstancesStatusKO(oConfXML, sNewPN, sNewRev)
        sLoadStatus = "Load Error"
        Call oTracking.AddInformationLine("AddComponentFromFile", "WARNING", "The document could not load")
    End If

    'Swap windows
    oPVRWindow.Activate
Else
    'Track
    Call oTracking.AddInformationLine("AddComponentFromFile", "EXECUTING", sNewPN & sNewRev & "." & sNewExtension & " already loaded in session.")
End If

'We have a document, let's add an instance to the PVR structure
If sLoadStatus = "Already Loaded" Or sLoadStatus = "New Load" Then

    Call oTracking.AddInformationLine("AddComponentFromFile", "EXECUTING", "Adding an new instance to structure")
    
    'Add instance in PVR
    On Error Resume Next
    Err.Clear
    Set oNewProduct = oParentRefProduct.Products.AddExternalComponent(oNewDoc)
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Could not add a new instance to structure")
    End If
    On Error GoTo 0
    
    'Add to oRefProductList
    If Not oRefProductList.Exists(sNewPN & sNewRev) Then
        On Error Resume Next
        Err.Clear
        Call oRefProductList.Add(sNewPN & sNewRev, oNewProduct.ReferenceProduct)
        If Err.Number <> 0 Then
            Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Could not add oNewProduct to oRefProductList")
        End If
        On Error GoTo 0
    End If
    
    'Get the RelationID of oConfElem
    On Error Resume Next
    Err.Clear
    sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Could not retrieve 'Relation ID' attribut")
    End If
    On Error GoTo 0
    
    'Transfer the instance name to all Conf instance with same RelationID
    Call TransferInstanceName(sRelationID, oNewProduct.Name)
    
    'Set status
    Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, oNewProduct.Name, "OK")
    
    'Move instance
    Call MoveInstance(sParentPN, sParentRev, oNewProduct.Name, oConfElem)
        
End If

'We need to close a newly open document
If sLoadStatus = "New Load" Then

    'Close document
    On Error Resume Next
    Err.Clear
    oNewDoc.Close
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("AddComponentFromFile", "CODEERROR", "Could not close the document")
    End If
    On Error GoTo 0
End If

End Sub

Private Sub RemoveInstance(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sInstanceName As String)

Dim oParentRefProduct As Product

'Get the Parent Reference Product
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

'Remove the product
Call oParentRefProduct.Products.Remove(sInstanceName)

End Sub

Private Sub MoveInstance(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sInstanceName As String, ByVal oConfElem As IXMLDOMElement)

Dim oParentRefProduct As Product
Dim oChildProduct As Product
Dim oPosition
Dim oMatrix(11) As Variant

'Track
Call oTracking.AddInformationLine("MoveInstance", "EXECUTING", "Parent: " & sParentPN & sParentRev & ", Child: " & sInstanceName)

'Get the Parent Reference Product
On Error Resume Next
Err.Clear
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("MoveInstance", "CODEERROR", "Could not get parent " & sParentPN & sParentRev)
End If
On Error GoTo 0

'Get child
On Error Resume Next
Err.Clear
Set oChildProduct = oParentRefProduct.Products.Item(sInstanceName)
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("MoveInstance", "CODEERROR", "Could not get child " & sInstanceName)
End If
On Error GoTo 0

'Retrieve position matrix
On Error Resume Next
Err.Clear
oMatrix(0) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue)
oMatrix(1) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue)
oMatrix(2) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue)
oMatrix(3) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue)
oMatrix(4) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue)
oMatrix(5) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue)
oMatrix(6) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue)
oMatrix(7) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue)
oMatrix(8) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue)
oMatrix(9) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue)
oMatrix(10) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue)
oMatrix(11) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue)
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("MoveInstance", "CODEERROR", "Could not retrieve position matrix")
End If
On Error GoTo 0

'Position instance
On Error Resume Next
Err.Clear
Set oPosition = oChildProduct.Position
oPosition.SetComponents oMatrix
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("MoveInstance", "CODEERROR", "Could not position the instance")
End If
On Error GoTo 0

End Sub

Private Function ComparePosition(ByVal oPVRElem As IXMLDOMElement, ByVal oConfElem As IXMLDOMElement) As String

Dim sReturnValue As String

'Initial set
sReturnValue = "Same Position"

'Check all position
If oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
End If

ComparePosition = sReturnValue


End Function
Private Sub GetPVRProductStructure(ByVal oProduct As Product)

Dim oElem As IXMLDOMElement

'Create PVRRootNode
Set oElem = oPVRXML.CreateElement("Instance")
oPVRXML.appendChild oElem
Call Add_Attribute(oPVRXML, oElem, "PartNumber", oTopNode.OriginalPartNumber)
Call Add_Attribute(oPVRXML, oElem, "InstanceName", "")
Call Add_Attribute(oPVRXML, oElem, "DocRev", oTopNode.OriginalPartRev)
Call Add_Attribute(oPVRXML, oElem, "SyncStatus", "")
Call Add_Attribute(oPVRXML, oElem, "Structure", "PVR")

'Add to oRefProductList
Call oRefProductList.Add(oTopNode.OriginalPartNumber & oTopNode.OriginalPartRev, oProduct.ReferenceProduct)

'Scan structure
Call ScanPVRProductStructure(oProduct, oElem)

'Save
If bSaveXML = True Then oPVRXML.Save sXMLPath & "PVR.xml"

End Sub

Private Sub ScanPVRProductStructure(ByVal oDMUParent As Product, ByVal oXMLParent As IXMLDOMElement)

Dim oChild As Product
Dim oInstanceElem As IXMLDOMElement
Dim oPosElem As IXMLDOMElement
Dim oSourceElem As IXMLDOMElement
Dim oClonedElem As IXMLDOMElement
Dim dPositionMatrix(11) As Variant
Dim oProdPosition 'As Position
Dim oDoc As Document
Dim oNode As IXMLDOMNode

Dim oChildPart As clsPartInfo

'Scan all children products
For Each oChild In oDMUParent.Products

    Set oChildPart = New clsPartInfo
    
    'Retrieve Part Number

    On Error Resume Next
    oChildPart.OriginalPartNumber = ""
    oChildPart.OriginalPartNumber = oChild.PartNumber
    If oChildPart.OriginalPartNumber = "" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    oChildPart.OriginalPartNumber = oChild.PartNumber
    On Error GoTo 0
    
    If oChildPart.OriginalPartNumber <> "" Then
        'Get Type
        oChildPart.PartType = GetType(oChild)
        
        'Get the document revision from the file name for CATPart and CATProduct
        If oChildPart.PartType = "CATPart" Or oChildPart.PartType = "CATProduct" Then
            
            'Get document
            Set oDoc = oChild.ReferenceProduct.Parent
    
            'Get the revision
            oChildPart.OriginalPartRev = Right(Split(oDoc.Name, ".")(0), 2)
        
        'For components we can't extract the revision from CATIA. I'll copy the one from oConfXML if it exist
        Else
            Set oNode = Nothing
            Set oNode = oConfXML.SelectSingleNode("//Instance[@PartNumber='" & oChildPart.OriginalPartNumber & "']")
            If oNode Is Nothing Then
                oChildPart.OriginalPartRev = "N/A"
            Else
                oChildPart.OriginalPartRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
            End If
        End If
    Else
        oChildPart.OriginalPartNumber = "G25XXXXXX-XXX" 'I set the part to a dummy part number. This way it will get deleted in a futur step
        oChildPart.PartType = "CATPart"
        oChildPart.OriginalPartRev = "--"
    End If
    
    'Get child position matrix
    Set oProdPosition = oChild.Position
    oProdPosition.GetComponents dPositionMatrix
    
    'We first check if the structure was already extracted
    Set oSourceElem = GetExistingNode(oPVRXML, oChildPart.OriginalPartNumber)

    'The structure already exist. We clone the structure and append it
    If Not oSourceElem Is Nothing Then
        
        'Clone
        Set oClonedElem = oSourceElem.cloneNode(True)
        oXMLParent.appendChild oClonedElem
        Set oInstanceElem = oXMLParent.lastChild
        
        'Edit instance name
        oInstanceElem.Attributes.getNamedItem("InstanceName").nodeValue = oChild.Name
    
        'Edit position
        Set oPosElem = oInstanceElem.SelectSingleNode("./Position")
        
        oPosElem.Attributes.getNamedItem("Position0").nodeValue = CDbl(dPositionMatrix(0))
        oPosElem.Attributes.getNamedItem("Position1").nodeValue = CDbl(dPositionMatrix(1))
        oPosElem.Attributes.getNamedItem("Position2").nodeValue = CDbl(dPositionMatrix(2))
        oPosElem.Attributes.getNamedItem("Position3").nodeValue = CDbl(dPositionMatrix(3))
        oPosElem.Attributes.getNamedItem("Position4").nodeValue = CDbl(dPositionMatrix(4))
        oPosElem.Attributes.getNamedItem("Position5").nodeValue = CDbl(dPositionMatrix(5))
        oPosElem.Attributes.getNamedItem("Position6").nodeValue = CDbl(dPositionMatrix(6))
        oPosElem.Attributes.getNamedItem("Position7").nodeValue = CDbl(dPositionMatrix(7))
        oPosElem.Attributes.getNamedItem("Position8").nodeValue = CDbl(dPositionMatrix(8))
        oPosElem.Attributes.getNamedItem("Position9").nodeValue = CDbl(dPositionMatrix(9))
        oPosElem.Attributes.getNamedItem("Position10").nodeValue = CDbl(dPositionMatrix(10))
        oPosElem.Attributes.getNamedItem("Position11").nodeValue = CDbl(dPositionMatrix(11))
                
    'We add a new child to the structure
    Else
    
        'Add to oRefProductList
        If oChildPart.OriginalPartNumber <> "G25XXXXXX-XXX" Then
            Call oRefProductList.Add(oChildPart.OriginalPartNumber & oChildPart.OriginalPartRev, oChild.ReferenceProduct)
        End If
                
        'Add instance to structure
        Set oInstanceElem = Add_Element(oPVRXML, "Instance", oXMLParent)
        Call Add_Attribute(oPVRXML, oInstanceElem, "PartNumber", oChildPart.OriginalPartNumber)
        Call Add_Attribute(oPVRXML, oInstanceElem, "InstanceName", oChild.Name)
        Call Add_Attribute(oPVRXML, oInstanceElem, "DocRev", oChildPart.OriginalPartRev)
        Call Add_Attribute(oPVRXML, oInstanceElem, "DocType", oChildPart.PartType)
        Call Add_Attribute(oPVRXML, oInstanceElem, "SyncStatus", "")
        If oChildPart.PartType = "Component" Then
            Call Add_Attribute(oPVRXML, oInstanceElem, "ComponentRename", "")
        End If
    
        'Add Position
        Set oPosElem = Add_Element(oPVRXML, "Position", oInstanceElem)
        Call Add_Attribute(oPVRXML, oPosElem, "Position0", CDbl(dPositionMatrix(0)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position1", CDbl(dPositionMatrix(1)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position2", CDbl(dPositionMatrix(2)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position3", CDbl(dPositionMatrix(3)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position4", CDbl(dPositionMatrix(4)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position5", CDbl(dPositionMatrix(5)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position6", CDbl(dPositionMatrix(6)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position7", CDbl(dPositionMatrix(7)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position8", CDbl(dPositionMatrix(8)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position9", CDbl(dPositionMatrix(9)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position10", CDbl(dPositionMatrix(10)))
        Call Add_Attribute(oPVRXML, oPosElem, "Position11", CDbl(dPositionMatrix(11)))

        'Recursive call on Components
        If oChildPart.PartType = "Component" Then
            Call ScanPVRProductStructure(oChild, oInstanceElem)
        End If
    End If
Next
End Sub

Private Sub GetConfProductStructure()

Dim oElem As IXMLDOMElement
Dim sVariant As Variant

'Create BSFRootNode
Set oElem = oConfXML.CreateElement("Instance")
oConfXML.appendChild oElem
Call Add_Attribute(oConfXML, oElem, "PartNumber", oTopNode.OriginalPartNumber)
Call Add_Attribute(oConfXML, oElem, "InstanceName", "")
Call Add_Attribute(oConfXML, oElem, "DocRev", "")
Call Add_Attribute(oConfXML, oElem, "DocStatus", "")
Call Add_Attribute(oConfXML, oElem, "DocType", "PVRREF")
Call Add_Attribute(oConfXML, oElem, "SyncStatus", "")
iRelationID = iRelationID + 1
Call Add_Attribute(oConfXML, oElem, "RelationID", CStr(iRelationID))
Call Add_Attribute(oConfXML, oElem, "Structure", "Configured")

'Scan structure
Call oTracking.AddHeader("ScanConfProductStructure")
Call ScanConfProductStructure(oTopNode, oElem, oTopNode.PartIsCI, sVariant)

'Save
If bSaveXML = True Then oConfXML.Save sXMLPath & "Conf.xml"

End Sub

Private Function PopulateChildrenFromWebService(ByVal oData As Variant, ByVal oParentPart As clsPartInfo) As clsCollection

Dim i, j As Integer
Dim sPN As String
Dim sRev As String
Dim oChildren As clsCollection
Dim oItem As clsCollection

'Initialize
Set oChildren = New clsCollection

'Scan oData and transfer value to oChildren
For i = 1 To oData.Count
    
    For j = 2 To oData.Item(i).Count
    
        'Get Part Number
        sPN = oData.Item(i).Item(1)

        'Special case for all parts with same base number as NHA (SP, Skel, MBP....)
        If oParentPart.PartIsCI = True And sPN Like oParentPart.UsedPartNumber & "*" Then
            sRev = oAttList.GetEnoviaAttributes(sPN, "LatestRev", False, "BA Document Revision")
        Else
            sRev = "N/A"
        End If

        Set oItem = New clsCollection
        Call oItem.Add("Part Number", sPN)
        Call oItem.Add("Revision", sRev)
        Call oItem.Add("Position Matrix", oData.Item(i).Item(j).Item(2))
        
        Call oChildren.Add("Instance" & oChildren.Count + 1, oItem)
    Next
Next

'Return
Set PopulateChildrenFromWebService = oChildren

End Function

Private Sub ConnectToAPI()

Dim session As EnoviaV5Session
On Error Resume Next
Set session = CATIA.GetItem("BAGExtEnoviaV5Session")
If Err.Number > 0 Then
    Dim SysSrv As SystemService
    Dim App As Application
    Set App = CATIA
    Set SysSrv = App.SystemService
    SysSrv.ExecuteProcessus "CNEXT.exe /regserver"
    Set session = CATIA.GetItem("BAGExtEnoviaV5Session")
End If
On Error GoTo 0

End Sub


Private Function ExtractProductStructureCIPVR(ByVal sPVRNumber As String, ByVal sPVRRev As String, ByVal sPartNumber As String) As clsCollection

Dim oPVRWindow As Window
Dim oPVRDoc As Document
Dim oProduct As Product
Dim oChild As Product
Dim EV5product As Product
Dim bTimeout As Boolean
Dim dPositionMatrix(11) As Variant
Dim oProdPosition 'As Position
Dim sPN As String
Dim sDocPN As String
Dim oChildren As clsCollection
Dim oItem As clsCollection
Dim EnoviaDoc As EnoviaDocument
Dim dLoadTimer As Double
Dim bPVRAlreadyOpen As Boolean
Dim sString As String
Dim sRev As String
Dim sType As String

'Initialize
Set oPVRWindow = CATIA.ActiveWindow
Set oChildren = New clsCollection

'Connect to API
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call ConnectToAPI
End If
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "CODEERROR", "Can't connect to CATIA API")
End If
On Error GoTo 0

'Check if PVR already open
bPVRAlreadyOpen = True
Set oPVRDoc = Nothing
On Error Resume Next
Set oPVRDoc = CATIA.Documents.Item(sPVRNumber & sPVRRev & ".CATProduct")
On Error GoTo 0

'Load PVR from Enovia
If oPVRDoc Is Nothing Then

    'Track
    Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "EXECUTING", sPVRNumber & sPVRRev & ".CATProduct will be loaded from ENOVIA.")

    On Error Resume Next
    Err.Clear
    bPVRAlreadyOpen = False
    Set EV5product = EnoviaDoc.OpenPartDocument(sPVRNumber, sPVRRev)
    If Err.Description <> "" Then Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "WARNING", "After load the error description is: " & Err.Description)
    
    'Make sure the document is loaded
    dLoadTimer = Timer
    bTimeout = False
    Do
    
        Set oPVRDoc = EV5product.ReferenceProduct.Parent
        
        If Not oPVRDoc Is Nothing Then Exit Do
        
        DoEvents
        Sleep 150
        Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "WARNING", "Sleeping while waiting for document to load")
        If Abs(Timer - dLoadTimer) > 5 Then
            bTimeout = True
            Exit Do
        End If
    Loop
    On Error GoTo 0
    
    'Timeout reached
    If bTimeout = True Then
        Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "CODEERROR", sPVRNumber & sPVRRev & " can't be loaded")
    End If
Else
    Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "EXECUTING", sPVRNumber & sPVRRev & ".CATProduct is already loaded in session.")
End If

'No children in PVR
Set oProduct = oPVRDoc.Product
If oProduct.Products.Count = 0 Then
    Call MsgBox(sPVRNumber & sPVRRev & " doesn't have any children. Process aborted.", vbCritical)
    Set ExtractProductStructureCIPVR = Nothing
    Exit Function
End If


'Get the top node
Set oProduct = oPVRDoc.Product

'***SPECIAL CASE***
'Sometimes users will cheat when creating a PVR, for example G25027879-001PVRREF. When you open this PVR the top node
'is not G25027879-001, it is L7EEM-1540021-E101-Z06-C01-IN.
'For this case I will check if I can find and instance at level 2 with a P/N = sPartNumber. If I find one I will use it
'as my top node. Otherwise I stop the process.
'The other special case is that sometimes the part number of the top node includes "PVRREF" or ".1" , refer to G25040790-001PVRREF-A and G25041376-001PVRREF--
If (oProduct.PartNumber <> sPartNumber) And (Not oProduct.PartNumber Like sPartNumber & "*") Then


    'Look for oChild with P/N = sPartNumber at level 2
    For Each oChild In oProduct.Products
    
        'Get the Part Number
        On Error Resume Next
        sPN = ""
        sPN = oChild.PartNumber
        If sPN = "" Then
            oChild.ApplyWorkMode DEFAULT_MODE
        End If
        sPN = oChild.PartNumber
        On Error GoTo 0
    
        If sPN = sPartNumber Then
            Set oProduct = oChild
            Exit For
        End If
    Next
    
    'oChild was not found
    If sPN <> sPartNumber Then
        Call MsgBox("In " & sPVRNumber & sPVRRev & " " & sPartNumber & " can't be found at levele 2. Process aborted.", vbCritical, "PVR Sync")
        Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "CODEERROR", "In " & sPVRNumber & sPVRRev & " " & sPartNumber & " can't be found at levele 2. Process aborted.")
        bCancelAction = True
        Exit Function
    End If
End If

'Extract first level children
'dTimer = Timer
For Each oChild In oProduct.Products

    'Get the Part Number
    On Error Resume Next
    sPN = ""
    sPN = oChild.PartNumber
    If sPN = "" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    sPN = oChild.PartNumber
    On Error GoTo 0

    If sPN <> "" Then
        
        'Get Type
        sType = GetType(oChild)
       
        'Override Part Number with the one found in the file name
        If sType = "CATPart" Or sType = "CATProduct" Then
            sPN = oChild.ReferenceProduct.Parent.Name
            sPN = Left(Split(sPN, ".")(0), Trim(Len(Split(sPN, ".")(0)) - 2))
        End If
        
        'Position Matrix
        Set oProdPosition = oChild.Position
        oProdPosition.GetComponents dPositionMatrix
    
        'Edit position matrix
        dPositionMatrix(9) = dPositionMatrix(9) / 1000
        dPositionMatrix(10) = dPositionMatrix(10) / 1000
        dPositionMatrix(11) = dPositionMatrix(11) / 1000
        
        'Special case for all parts with same base number as NHA (SP, Skel, MBP....)
        If sPN Like sPartNumber & "*" And (sType = "CATPart" Or sType = "CATProduct") Then
            sRev = oChild.ReferenceProduct.Parent.Name
            sRev = Right(Split(sRev, ".")(0), 2)
        Else
            sRev = "N/A"
        End If
        
        'Package data in oChildren
        Set oItem = New clsCollection
        Call oItem.Add("Part Number", sPN)
        Call oItem.Add("Revision", sRev)
        Call oItem.Add("Position Matrix", dPositionMatrix)
        
        Call oChildren.Add("Instance" & oChildren.Count + 1, oItem)
    
    Else
        Call oLoadErrDoc.Add(oChild.Name, " : Part Number of can't be retrieve. Instance found in " & sPVRNumber & sPVRRev)
    End If
    
Next

'Swap windows
oPVRWindow.Activate
    
'Close document
If bPVRAlreadyOpen = False Then

    'Close document
    On Error Resume Next
    Err.Clear
    oPVRDoc.Close
    If Err.Number <> 0 Then
        Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "CODEERROR", "Could not close the document " & oPVRDoc.Name)
    End If
    On Error GoTo 0
    
End If

'Return
Set ExtractProductStructureCIPVR = oChildren

End Function



Private Sub ScanConfProductStructure(ByRef oParentPart As clsPartInfo, ByVal oXMLParent As IXMLDOMElement, ByVal bGParentCI As Boolean, ByVal dPosMatrix As Variant)

Dim oData
Dim i, j As Integer
Dim oInstanceElem As IXMLDOMElement
Dim oChildren As clsCollection, oBDIPartColl As clsCollection
Dim bExitByUser As Boolean, bAlreadyExist As Boolean
Dim sString As String, sArray() As String, sReqdRev As String, sBomCompare As String, sAnswer As String
Dim oChildPart As clsPartInfo

'Track
Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "---Parent part number is " & oParentPart.OriginalPartNumber)

If bCancelAction = True Then Exit Sub

'Check if parent is a CI only when:
' - Grand-parent is a CI
'-  User doesn't want to update with BSF
If bGParentCI = True And oTopNode.SyncFromBSF = False Then
    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Cheking if part is a CI")
    
    'Get the info of the BDI Part
    Set oBDIPartColl = GetBDIPartCollection(oParentPart.RefPartNumber, "", bExitByUser, True)

    'Extract info from latest BDI Part rev
    If oBDIPartColl.Count >= 1 Then
        oParentPart.PartIsCI = oBDIPartColl.GetItem(1).IsCI
    Else
        oParentPart.PartIsCI = False
    End If
End If

'Activate CATIA
If oTopNode.CIOption <> "DisplayNever" Then
    AppActivate CATIA.Caption
    DoEvents
End If

'Part is the PVRREF, PartIsCI = True, PVR is not "REL" and oTopNode.SyncFromBSF = False
If oParentPart.OriginalPartNumber = oTopNode.OriginalPartNumber And oParentPart.PartIsCI = True And oTopNode.PVRDocStatus <> "REL" And oTopNode.SyncFromBSF = False Then
    
    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part is PVR + CI + Non Released")
    oParentPart.Mode = "BSF"
    oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
    oParentPart.OriginalPartStatus = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "Revision Status")
    oParentPart.RefPartRev = oParentPart.OriginalPartRev
    oParentPart.RefPartStatus = oParentPart.OriginalPartStatus
    oParentPart.UsedPartRev = oParentPart.OriginalPartRev
    oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
    oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
    oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
    oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
    oParentPart.SelectedSourceRev = oParentPart.OriginalPartRev
    oParentPart.EffectiveDwgNb = "N/A"
    oParentPart.EffectiveDwgRev = "N/A"
    oParentPart.SelectedDwgNb = "N/A"
    oParentPart.SelectedDwgRev = "N/A"
    
'Part is the PVRREF, PartIsCI = True, PVR is "REL" and oTopNode.SyncFromBSF = False
ElseIf oParentPart.OriginalPartNumber = oTopNode.OriginalPartNumber And oParentPart.PartIsCI = True And oTopNode.PVRDocStatus = "REL" And oTopNode.SyncFromBSF = False Then
   
    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part is PVR + CI + Released")
    
    oParentPart.Mode = "PVR"
    
    'Find the BDI document effective for the given tail
    Set oEffDocReport = Nothing
    sString = GetEffectiveBDIDoc(oBDIPartColl, oTopNode.Tail)
    oParentPart.EffectiveDwgRev = Right(sString, 2)
    oParentPart.EffectiveDwgNb = Left(sString, Len(sString) - 2)

    'Generate and save the report
    Call GenerateEffectiveDocReport(oParentPart.EffectiveDwgNb, oParentPart.EffectiveDwgRev, oParentPart.OriginalPartNumber)

    'Get the related PVRREF for a given document revision
    Set oRelatedPVRReport = Nothing
    Call GetRelatedPVRandNone(oParentPart, oParentPart.EffectiveDwgNb, oParentPart.EffectiveDwgRev, "NONE|ProposedPVR")

    'Compare with the loaded document
    If oParentPart.ProposedSourceNb <> oTopNode.PVRDocNb Or oParentPart.ProposedSourceRev <> oTopNode.PVRDocRev Then
        Call frmPVRSync3.InitializeForm(oParentPart.EffectiveDwgNb & oParentPart.EffectiveDwgRev, _
                                        oParentPart.ProposedSourceNb & oParentPart.ProposedSourceRev, _
                                        oTopNode.PVRDocNb & oTopNode.PVRDocRev, _
                                        sEffectiveDocReportPath & "PVRSync_GetEffectiveDocument_" & oParentPart.OriginalPartNumber & ".txt", _
                                        oParentPart.Tail)

        sAnswer = frmPVRSync3.Answer
        Unload frmPVRSync3

        If sAnswer = "Yes" Then

            oParentPart.SelectedSourceNb = oTopNode.PVRDocNb
            oParentPart.SelectedSourceRev = oTopNode.PVRDocRev
            oParentPart.SelectedDwgRev = "N/A"
            oParentPart.SelectedDwgNb = "N/A"

            'Get the NONE related to the selected PVR
            Call GetRelatedPVRandNone(oParentPart, oParentPart.SelectedSourceNb, oParentPart.SelectedSourceRev, "NONE")
            oParentPart.UsedPartRev = oParentPart.RefPartRev
            oParentPart.UsedPartStatus = oParentPart.RefPartStatus
        Else
            bCancelAction = True
            Exit Sub
        End If
    Else
        oParentPart.SelectedSourceNb = oParentPart.ProposedSourceNb
        oParentPart.SelectedSourceRev = oParentPart.ProposedSourceRev
        oParentPart.SelectedDwgNb = oParentPart.EffectiveDwgNb
        oParentPart.SelectedDwgRev = oParentPart.EffectiveDwgRev
        oParentPart.UsedPartRev = oParentPart.OriginalPartRev
        oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
    End If


'Part is a CI and oTopNode.SyncFromBSF = False
ElseIf oParentPart.PartIsCI = True And oTopNode.SyncFromBSF = False Then

    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part is CI")
    oParentPart.Mode = "PVR"
    
    'Find the BDI document effective for the given tail
    Set oEffDocReport = Nothing
    sString = GetEffectiveBDIDoc(oBDIPartColl, oTopNode.Tail)
    
    'An effective document was found
    If sString <> "" Then
        oParentPart.EffectiveDwgRev = Right(sString, 2)
        oParentPart.EffectiveDwgNb = Left(sString, Len(sString) - 2)

        'Generate and save the report
        Call GenerateEffectiveDocReport(oParentPart.EffectiveDwgNb, oParentPart.EffectiveDwgRev, oParentPart.RefPartNumber)

        'Get the related PVRREF for the proposed drawing
        Set oRelatedPVRReport = Nothing
        Call GetRelatedPVRandNone(oParentPart, oParentPart.EffectiveDwgNb, oParentPart.EffectiveDwgRev, "NONE|ProposedPVR")
        
        'No PVR was found, use BSF
        If oParentPart.ProposedSourceNb = "" Then
            oParentPart.Mode = "BSF"
            oParentPart.Comment = "No PVR found, using BSF"
            
        'A PVR was found. Let's find the status of the latest rev of this part
        Else
            oParentPart.PartLatestRevStatus = oAttList.GetEnoviaAttributes(oParentPart.RefPartNumber, "LatestRev", False, "Revision Status")
        End If

    'No effective drawing was found, use latest released
    Else
        oParentPart.Mode = "BSF"
        oParentPart.Comment = "No effective drawing could be found, using BSF"
    End If

    'User selection only when the mode is PVR
    If oParentPart.Mode = "PVR" And (oTopNode.CIOption = "DisplayAlways" Or (oTopNode.CIOption = "DisplayNonRel" And oParentPart.PartLatestRevStatus <> "REL")) Then

        AppActivate CATIA.Caption
        DoEvents

        Call frmPVRSync.InitializeForm(oParentPart.RefPartNumber, oParentPart.PartTitle, oParentPart.EffectiveDwgNb & oParentPart.EffectiveDwgRev, sEffectiveDocReportPath & "PVRSync_GetEffectiveDocument_" & oParentPart.RefPartNumber & ".txt", oEffDocReport, oParentPart.Tail)
        sString = frmPVRSync.SelectedDoc
        Unload frmPVRSync

        'Exit when user selected "Cancel" in userform
        If sString = "" Then
            bCancelAction = True
            Exit Sub
        End If

        If sString <> "BSF" Then
            oParentPart.SelectedDwgRev = Split(sString, " ")(1)
            oParentPart.SelectedDwgNb = Split(sString, " ")(0)

            If oParentPart.SelectedDwgRev = oParentPart.EffectiveDwgRev Then
                oParentPart.SelectedSourceRev = oParentPart.ProposedSourceRev
                oParentPart.SelectedSourceNb = oParentPart.ProposedSourceNb

            Else
                Set oRelatedPVRReport = Nothing
                Call GetRelatedPVRandNone(oParentPart, oParentPart.SelectedDwgNb, oParentPart.SelectedDwgRev, "NONE|SelectedPVR")
            End If


            'Flexible CIs
            If oParentPart.OriginalPartNumber <> oParentPart.RefPartNumber Then
                
                Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part is a flexible CI, comparing BOMs")
                
                'Compare the BOMs
                sBomCompare = CompareBOM(oParentPart.OriginalPartNumber, oParentPart.ProposedSourceNb, oParentPart.ProposedSourceRev)
                
                If sBomCompare = "Same BOM" Then
                    oParentPart.Mode = "Loaded PVR"
                    oParentPart.Comment = "BOM of flexible and non-flexible assemby are equal. Flexible assembly was used."
                    oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
                    If oParentPart.OriginalPartRev <> "--" Then
                        oParentPart.OriginalPartRev = "NA"
                    End If
                    oParentPart.UsedPartNumber = oParentPart.OriginalPartNumber
                    oParentPart.UsedPartRev = oParentPart.OriginalPartRev
                    oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
                ElseIf sBomCompare = "Different BOM" Then
                    oParentPart.UsedPartNumber = oParentPart.RefPartNumber
                    oParentPart.UsedPartRev = oParentPart.RefPartRev
                    oParentPart.UsedPartStatus = oParentPart.RefPartStatus
                    oParentPart.PartTitle = oAttList.GetEnoviaAttributes(oParentPart.UsedPartNumber, oParentPart.UsedPartRev, False, "Title")
                    oParentPart.Comment = "BOM of flexible and non-flexible assemby are different. Non-Flexible assembly was used."
                Else
                    bCancelAction = True
                    Exit Sub
                End If
            
            'Non Flexible
            Else
                oParentPart.UsedPartRev = oParentPart.RefPartRev
                oParentPart.UsedPartStatus = oParentPart.RefPartStatus
            End If


        Else
            oParentPart.Mode = "BSF"
            oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
            oParentPart.OriginalPartStatus = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "Revision Status")
            oParentPart.UsedPartRev = oParentPart.OriginalPartRev
            oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
            oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
            oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
            oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
            oParentPart.SelectedSourceRev = oParentPart.OriginalPartRev
            oParentPart.EffectiveDwgNb = "N/A"
            oParentPart.EffectiveDwgRev = "N/A"
            oParentPart.SelectedDwgNb = "N/A"
            oParentPart.SelectedDwgRev = "N/A"
        End If

    ElseIf oParentPart.Mode = "PVR" And (oTopNode.CIOption = "DisplayNever" Or (oTopNode.CIOption = "DisplayNonRel" And oParentPart.PartLatestRevStatus = "REL")) Then
        oParentPart.SelectedSourceRev = oParentPart.ProposedSourceRev
        oParentPart.SelectedSourceNb = oParentPart.ProposedSourceNb
        oParentPart.SelectedDwgNb = oParentPart.EffectiveDwgNb
        oParentPart.SelectedDwgRev = oParentPart.EffectiveDwgRev
        
        'Flexible CIs
        If oParentPart.OriginalPartNumber <> oParentPart.RefPartNumber Then
            
            'Compare the BOMs
            sBomCompare = CompareBOM(oParentPart.OriginalPartNumber, oParentPart.ProposedSourceNb, oParentPart.ProposedSourceRev)
            
            If sBomCompare = "Same BOM" Then
                oParentPart.Mode = "Loaded PVR"
                oParentPart.Comment = "BOM of flexible and non-flexible assemby are equal. Flexible assembly was used."
                oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
                If oParentPart.OriginalPartRev <> "--" Then
                    oParentPart.OriginalPartRev = "NA"
                End If
                oParentPart.UsedPartRev = oParentPart.OriginalPartRev
                oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
            ElseIf sBomCompare = "Different BOM" Then
                oParentPart.UsedPartNumber = oParentPart.RefPartNumber
                oParentPart.PartTitle = oAttList.GetEnoviaAttributes(oParentPart.RefPartNumber, oParentPart.RefPartRev, False, "Title")
                oParentPart.Comment = "BOM of flexible and non-flexible assemby are different. Non-Flexible assembly was used."
                oParentPart.UsedPartRev = oParentPart.RefPartRev
                oParentPart.UsedPartStatus = oParentPart.RefPartStatus
            Else
                bCancelAction = True
                Exit Sub
            End If
        
            'Non Flexible
            Else
                oParentPart.UsedPartRev = oParentPart.RefPartRev
                oParentPart.UsedPartStatus = oParentPart.RefPartStatus
            End If

    ElseIf oParentPart.Mode = "BSF" Then
        oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
        oParentPart.OriginalPartStatus = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "Revision Status")
        oParentPart.UsedPartNumber = oParentPart.OriginalPartNumber
        oParentPart.UsedPartRev = oParentPart.OriginalPartRev
        oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
        oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
        oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
        oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
        oParentPart.SelectedSourceRev = oParentPart.OriginalPartRev
        oParentPart.EffectiveDwgNb = "N/A"
        oParentPart.EffectiveDwgRev = "N/A"
        oParentPart.SelectedDwgNb = "N/A"
        oParentPart.SelectedDwgRev = "N/A"
    End If


'Part is not a CI and oTopNode.SyncFromBSF = False
ElseIf oParentPart.PartIsCI = False And oTopNode.SyncFromBSF = False Then

    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part is not a CI")

    oParentPart.EffectiveDwgNb = "N/A"
    oParentPart.EffectiveDwgRev = "N/A"
    oParentPart.SelectedDwgNb = "N/A"
    oParentPart.SelectedDwgRev = "N/A"

    'We want the BSF or the BSF is released
    If oTopNode.NonCIOption = "BSF" Or (oTopNode.NonCIOption = "LatestRel" And oParentPart.PartLatestRevStatus = "REL") Then
        oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
        oParentPart.SelectedSourceRev = oParentPart.OriginalPartRev
        oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
        oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
        oParentPart.UsedPartNumber = oParentPart.OriginalPartNumber
        oParentPart.UsedPartRev = oParentPart.OriginalPartRev
        oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
        oParentPart.Mode = "BSF"
        
   'We want latest released and the BSF is not released. We need to get the latest released PVR
   Else
        'Looking for the latest released PVR
        oParentPart.Mode = "PVR"
        oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber & "PVRREF"
        oParentPart.SelectedSourceRev = oAttList.GetEnoviaAttributes(oParentPart.SelectedSourceNb, "LatestRelRev", False, "BA Document Revision")
        oParentPart.ProposedSourceNb = oParentPart.SelectedSourceNb
        oParentPart.ProposedSourceRev = oParentPart.SelectedSourceRev
        oParentPart.UsedPartNumber = oParentPart.OriginalPartNumber
        oParentPart.UsedPartRev = oParentPart.OriginalPartRev
        oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
   
        'PVR not found, using the BSF
        If oParentPart.SelectedSourceRev = "Part or Attribute doesn't exist" Then
            oParentPart.Comment = "BSF if not REL and no released PVR could not be found, using BSF."
            oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
            oParentPart.SelectedSourceRev = oAttList.GetEnoviaAttributes(oParentPart.SelectedSourceNb, "LatestRev", False, "BA Document Revision")
            oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
            oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
            oParentPart.Mode = "BSF"
        End If
   
   End If

'We update using the BSF
Else

    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Part should be sync with BSF")
    
    oParentPart.Mode = "BSF"
    oParentPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "BA Document Revision")
    oParentPart.OriginalPartStatus = oAttList.GetEnoviaAttributes(oParentPart.OriginalPartNumber, "LatestRev", False, "Revision Status")
    oParentPart.ProposedSourceNb = oParentPart.OriginalPartNumber
    oParentPart.ProposedSourceRev = oParentPart.OriginalPartRev
    oParentPart.SelectedSourceNb = oParentPart.OriginalPartNumber
    oParentPart.SelectedSourceRev = oParentPart.OriginalPartRev
    oParentPart.UsedPartNumber = oParentPart.OriginalPartNumber
    oParentPart.UsedPartRev = oParentPart.OriginalPartRev
    oParentPart.UsedPartStatus = oParentPart.OriginalPartStatus
    oParentPart.EffectiveDwgNb = "N/A"
    oParentPart.EffectiveDwgRev = "N/A"
    oParentPart.SelectedDwgNb = "N/A"
    oParentPart.SelectedDwgRev = "N/A"
End If

'Add to oPartList
If Not oPartList.Exists(oParentPart.UsedPartNumber & oParentPart.UsedPartRev) Then
    Call oPartList.Add(oParentPart.UsedPartNumber & oParentPart.UsedPartRev, oParentPart)
End If

'Edit XML structure except for top node
If oTopNode.OriginalPartNumber <> oParentPart.OriginalPartNumber Then
    Set oInstanceElem = Nothing
    bAlreadyExist = False
    Call EditXML(oInstanceElem, bAlreadyExist, oXMLParent, oParentPart, dPosMatrix)
    
    If bAlreadyExist Then
        Exit Sub
    Else
        Set oXMLParent = oInstanceElem
    End If
Else
    oXMLParent.Attributes.getNamedItem("DocRev").nodeValue = oParentPart.UsedPartRev
End If

'Activate CATIA
If oTopNode.CIOption <> "DisplayNever" Then
    AppActivate CATIA.Caption
    DoEvents
End If

'Extract Product structure from selected PVR
If oParentPart.Mode = "PVR" Then

    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Product structure will be extracted from " & oParentPart.SelectedSourceNb & " " & oParentPart.SelectedSourceRev)

    Set oChildren = ExtractProductStructureCIPVR(oParentPart.SelectedSourceNb, oParentPart.SelectedSourceRev, oParentPart.UsedPartNumber)

    'Activate CATIA
    If oTopNode.CIOption <> "DisplayNever" Then
        AppActivate CATIA.Caption
        DoEvents
    End If

    'PVR could not be loaded or there is no children in it.
    If oChildren Is Nothing Then Exit Sub

'Extract Product Structure from BSF
ElseIf oParentPart.Mode = "BSF" Then

    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Product structure will be extracted from BSF")
    
    'Retrieve children and position matrix from Web Service
    Set oData = WebServiceAccessTool.GetAssyPositionMatricies(oParentPart.UsedPartNumber)
    Call oTracking.WebServiceResultAnalysis("GetAssyPositionMatricies", oData, oParentPart.UsedPartNumber)
    
    'Transfer children info to oChildren
    Set oChildren = PopulateChildrenFromWebService(oData, oParentPart)

'Extract Product Structure from the loaded PVR (for flexible CIs)
ElseIf oParentPart.Mode = "Loaded PVR" Then
    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "Product structure will be extracted from Loaded PVR (Flexible CI)")
    Set oChildren = ExtractProductStructureFromProduct(oParentPart)
End If

'Add all children to XML structure
For i = 1 To oChildren.Count

    '**Required Revision
    'When we update on a specific tail we use the NonCIOption to know if we use:
    '  The BSF
    '  The latest released revision.
    'However if the part is a CI the required revision may change when we will retrieve the effective one
    If oTopNode.SyncFromBSF = False Then

        'Special case for SP collectors of CI
        If oChildren.GetItemByIndex(i).GetItemByKey("Revision") <> "N/A" Then
            sReqdRev = oChildren.GetItemByIndex(i).GetItemByKey("Revision")
        ElseIf oTopNode.NonCIOption = "LatestRel" Then
            sReqdRev = "LatestRelRevIfExist"
        Else
            sReqdRev = "LatestRev"
        End If
    Else
        sReqdRev = "LatestRev"
    End If

    'Part Number
    Set oChildPart = New clsPartInfo
    oChildPart.OriginalPartNumber = oChildren.GetItemByIndex(i).GetItemByKey("Part Number")
    
    Call oTracking.AddInformationLine("ScanConfProductStructure", "EXECUTING", "*Children is " & oParentPart.OriginalPartNumber & "/" & oChildPart.OriginalPartNumber & " " & sReqdRev)
    
    'Get other attributes of the document
    oChildPart.PartExtension = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "EXTENSION")
    oChildPart.PartType = IIf(oChildPart.PartExtension = "", "Component", oChildPart.PartExtension)
    oChildPart.OriginalPartRev = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "BA Document Revision")
    oChildPart.OriginalPartStatus = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "Revision Status")
    oChildPart.UsedPartNumber = oChildPart.OriginalPartNumber
    oChildPart.UsedPartRev = oChildPart.OriginalPartRev
    oChildPart.UsedPartStatus = oChildPart.OriginalPartStatus
    oChildPart.PartLatestRevStatus = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, "LatestRev", False, "Revision Status")
    oChildPart.PartTitle = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "Title")
    oChildPart.DefiningPartAtt = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "Defining Part")
    oChildPart.DatasetTypeAtt = oAttList.GetEnoviaAttributes(oChildPart.OriginalPartNumber, sReqdRev, False, "Dataset Type")

    'Defining Part
    If (oChildPart.DatasetTypeAtt = "CATALOG LIGHT VERSION" Or oChildPart.DatasetTypeAtt = "FLEXIBLE REPRESENTATION") And oChildPart.DefiningPartAtt <> "" Then
        oChildPart.RefPartNumber = UCase(oChildPart.DefiningPartAtt)
        oChildPart.RefPartNumber = Trim(Replace(oChildPart.RefPartNumber, "(DON'T USE THIS PART)", ""))
        oChildPart.RefPartNumber = Trim(Replace(oChildPart.RefPartNumber, "(CANCELLED)", ""))
        oChildPart.RefPartRev = oAttList.GetEnoviaAttributes(oChildPart.RefPartNumber, sReqdRev, False, "BA Document Revision")
        oChildPart.RefPartStatus = oAttList.GetEnoviaAttributes(oChildPart.RefPartNumber, sReqdRev, False, "Revision Status")
    Else
        oChildPart.RefPartNumber = oChildPart.OriginalPartNumber
        oChildPart.RefPartRev = oAttList.GetEnoviaAttributes(oChildPart.RefPartNumber, sReqdRev, False, "BA Document Revision")
        oChildPart.RefPartStatus = oAttList.GetEnoviaAttributes(oChildPart.RefPartNumber, sReqdRev, False, "Revision Status")
    End If
    
    'Initialize
    oChildPart.EffectiveDwgNb = "N/A"
    oChildPart.EffectiveDwgRev = "N/A"
    oChildPart.SelectedDwgNb = "N/A"
    oChildPart.SelectedDwgRev = "N/A"
    oChildPart.Comment = "No comment"
    
    'Get position matrix
    dPosMatrix = oChildren.GetItemByIndex(i).GetItemByKey("Position Matrix")

    'Recursive call
    If oChildPart.PartType <> "Component" Then
    
        'Add to oPartList if not already in the collection
        If Not oPartList.Exists(oChildPart.OriginalPartNumber & oChildPart.OriginalPartRev) Then
            Call oPartList.Add(oChildPart.OriginalPartNumber & oChildPart.OriginalPartRev, oChildPart)
        End If
    
        'Edit XML structure
        Call EditXML(oInstanceElem, bAlreadyExist, oXMLParent, oChildPart, dPosMatrix)
    Else
        If bCancelAction = True Then Exit Sub
        Call ScanConfProductStructure(oChildPart, oXMLParent, oParentPart.PartIsCI, dPosMatrix)
    End If
    
Next

End Sub

Sub EditXML(ByRef oInstanceElem As IXMLDOMElement, ByRef bAlreadyExist As Boolean, ByVal oXMLParent As IXMLDOMElement, ByVal oChildPart As clsPartInfo, ByVal dPosMatrix As Variant)


Dim oSourceElem As IXMLDOMElement
Dim oClonedElem As IXMLDOMElement
Dim oPosElem As IXMLDOMElement

'We first check if the structure was already extracted
Set oSourceElem = GetExistingNode(oConfXML, oChildPart.UsedPartNumber)

'The structure already exist. We clone the structure and append it
If Not oSourceElem Is Nothing Then

    bAlreadyExist = True
    
    'Clone
    Set oClonedElem = oSourceElem.cloneNode(True)
    oXMLParent.appendChild oClonedElem
    Set oInstanceElem = oXMLParent.lastChild

    'Edit RelationID
    iRelationID = iRelationID + 1
    Call Add_Attribute(oConfXML, oInstanceElem, "RelationID", CStr(iRelationID))

    'Edit position
    Set oPosElem = oInstanceElem.SelectSingleNode("./Position")

    oPosElem.Attributes.getNamedItem("Position0").nodeValue = CDbl(dPosMatrix(0))
    oPosElem.Attributes.getNamedItem("Position1").nodeValue = CDbl(dPosMatrix(1))
    oPosElem.Attributes.getNamedItem("Position2").nodeValue = CDbl(dPosMatrix(2))
    oPosElem.Attributes.getNamedItem("Position3").nodeValue = CDbl(dPosMatrix(3))
    oPosElem.Attributes.getNamedItem("Position4").nodeValue = CDbl(dPosMatrix(4))
    oPosElem.Attributes.getNamedItem("Position5").nodeValue = CDbl(dPosMatrix(5))
    oPosElem.Attributes.getNamedItem("Position6").nodeValue = CDbl(dPosMatrix(6))
    oPosElem.Attributes.getNamedItem("Position7").nodeValue = CDbl(dPosMatrix(7))
    oPosElem.Attributes.getNamedItem("Position8").nodeValue = CDbl(dPosMatrix(8))
    oPosElem.Attributes.getNamedItem("Position9").nodeValue = CDbl(dPosMatrix(9)) * 1000 'When using the web service position are in meter and not mm
    oPosElem.Attributes.getNamedItem("Position10").nodeValue = CDbl(dPosMatrix(10)) * 1000
    oPosElem.Attributes.getNamedItem("Position11").nodeValue = CDbl(dPosMatrix(11)) * 1000

'We add a new child to the structure
Else

    bAlreadyExist = False
    
    'Add instance to structure
    Set oInstanceElem = Add_Element(oConfXML, "Instance", oXMLParent)
    Call Add_Attribute(oConfXML, oInstanceElem, "PartNumber", oChildPart.UsedPartNumber)
    Call Add_Attribute(oConfXML, oInstanceElem, "InstanceName", "")
    Call Add_Attribute(oConfXML, oInstanceElem, "DocRev", oChildPart.UsedPartRev)
    Call Add_Attribute(oConfXML, oInstanceElem, "DocStatus", oChildPart.UsedPartStatus)
    Call Add_Attribute(oConfXML, oInstanceElem, "DocType", oChildPart.PartType)
    Call Add_Attribute(oConfXML, oInstanceElem, "SyncStatus", "")
    If oChildPart.PartType = "Component" Then
        Call Add_Attribute(oConfXML, oInstanceElem, "ComponentRename", "")
    End If

    iRelationID = iRelationID + 1
    Call Add_Attribute(oConfXML, oInstanceElem, "RelationID", CStr(iRelationID))

    'Add position
    Set oPosElem = Add_Element(oConfXML, "Position", oInstanceElem)
    Call Add_Attribute(oConfXML, oPosElem, "Position0", CDbl(dPosMatrix(0)))
    Call Add_Attribute(oConfXML, oPosElem, "Position1", CDbl(dPosMatrix(1)))
    Call Add_Attribute(oConfXML, oPosElem, "Position2", CDbl(dPosMatrix(2)))
    Call Add_Attribute(oConfXML, oPosElem, "Position3", CDbl(dPosMatrix(3)))
    Call Add_Attribute(oConfXML, oPosElem, "Position4", CDbl(dPosMatrix(4)))
    Call Add_Attribute(oConfXML, oPosElem, "Position5", CDbl(dPosMatrix(5)))
    Call Add_Attribute(oConfXML, oPosElem, "Position6", CDbl(dPosMatrix(6)))
    Call Add_Attribute(oConfXML, oPosElem, "Position7", CDbl(dPosMatrix(7)))
    Call Add_Attribute(oConfXML, oPosElem, "Position8", CDbl(dPosMatrix(8)))
    Call Add_Attribute(oConfXML, oPosElem, "Position9", CDbl(dPosMatrix(9) * 1000)) 'When using the web service position are in meter and not mm
    Call Add_Attribute(oConfXML, oPosElem, "Position10", CDbl(dPosMatrix(10) * 1000))
    Call Add_Attribute(oConfXML, oPosElem, "Position11", CDbl(dPosMatrix(11) * 1000))
End If
    

End Sub

Private Function GetExistingNode(ByVal oDoc As DOMDocument60, ByVal sPN As String) As IXMLDOMElement

Dim oNodeList As IXMLDOMNodeList

Set oNodeList = oDoc.SelectNodes("//Instance[@PartNumber='" & sPN & "' and @SyncStatus='']")
If oNodeList.Length >= 1 Then
    Set GetExistingNode = oNodeList.Item(0)
Else
    Set GetExistingNode = Nothing
End If

End Function

Private Function GetType(ByVal oProduct As Product) As String

'CATPart
If oProduct.HasAMasterShapeRepresentation Then
    GetType = "CATPart"
'Component
ElseIf oProduct.ReferenceProduct.Parent.Name = oProduct.Parent.Parent.ReferenceProduct.Parent.Name Then
    GetType = "Component"
'CATProduct
Else
    GetType = "CATProduct"
End If
 
End Function

Private Function Add_Element(ByVal oDoc As DOMDocument60, ByVal sTagName As String, ByVal oParentElem As IXMLDOMElement, Optional oBrotherElem = Nothing) As IXMLDOMElement

Dim oElement As IXMLDOMElement

'Add new element
Set oElement = oDoc.CreateElement(sTagName)

'Append node
If oBrotherElem Is Nothing Then
    oParentElem.appendChild oElement
Else
    Call oParentElem.InsertBefore(oElement, oBrotherElem)
End If

'Return
Set Add_Element = oElement

End Function

Private Sub Add_Attribute(ByVal oDoc As DOMDocument60, ByRef oElement As IXMLDOMElement, ByVal sAttributeName As String, ByVal sAttributeValue As String)

Dim oAttribute As IXMLDOMAttribute

'Add attribute
Set oAttribute = oDoc.createAttribute(sAttributeName)
oAttribute.nodeValue = sAttributeValue
oElement.setAttributeNode oAttribute

End Sub

Private Sub Add_Comment(ByVal oDoc As DOMDocument60, ByRef oElement As IXMLDOMElement, ByVal sCommentValue As String)

Dim oComment As IXMLDOMComment

'Add comment
Set oComment = oDoc.createComment(sCommentValue)
oElement.appendChild oComment

End Sub


Private Sub GenerateErrorLog()

Dim sText As String
Dim i As Integer

If oLoadErrDoc.Count > 0 Then
    
    sText = "Following error were found:"
    For i = 1 To oLoadErrDoc.Count
        sText = sText & vbCrLf & "-" & oLoadErrDoc.GetKey(i) & oLoadErrDoc.GetItemByIndex(i)
    Next
    
    Call WriteTextFile(sErroLogFilePath, sText, True)
    
    Call MsgBox("Refer to " & sErroLogFilePath, vbCritical)
End If


End Sub

Private Sub GetRelatedPVRandNone(ByRef oParentPart As clsPartInfo, ByVal sSourceDocNumber As String, ByVal sSourceDocRev As String, ByVal sUpdateType As String)

'From a source document find the related PLMACTION.
'In the PLMACTION we retrieve:
' - The PVRREF equal to oParentPart.RefPartNumber & "PVRREF"
' - The NONE equal to oParentPart.RefPartNumber
' - The drawing equal to string left of "-" from oParentPart.RefPartNumber
'If we can't find the PVR and the NONE we get the PLMACTION linked to the previous drawing revision

Dim oActionsElem As IXMLDOMElement
Dim oActionElem As IXMLDOMElement
Dim oAttachedDocElem As IXMLDOMElement
Dim oOrderedRevList As clsCollection
Dim oPLMAction As clsCollection
Dim oDocColl As clsCollection
Dim oTemp
Dim sPLMAction As String, sDrawingNb As String, sDrawingRev As String, sPVRNb As String, sPVRRev As String, sNoneRev As String
Dim i As Integer

Call oTracking.AddInformationLine("GetRelatedPVRandNone", "EXECUTING", "Getting the related PVR and None document")

'Initialize
Set oRelatedPVRReport = New DOMDocument60
oRelatedPVRReport.setProperty "SelectionLanguage", "XPath"
oRelatedPVRReport.async = True

'Create root node: PLMActionsEnoviaDocuments
Set oActionsElem = oRelatedPVRReport.CreateElement("PLMACTIONS")
oRelatedPVRReport.appendChild oActionsElem

Do

    'Create PLMACTION node
    Set oActionElem = Add_Element(oRelatedPVRReport, "PLMACTION", oActionsElem)
    Call Add_Attribute(oRelatedPVRReport, oActionElem, "SourceDocNumber", sSourceDocNumber)
    Call Add_Attribute(oRelatedPVRReport, oActionElem, "SourceDocRev", sSourceDocRev)
    
    'Get the PLMACTION
    sPLMAction = ""
    Set oTemp = WebServiceAccessTool.GetPLMActionDataFromDocument(sSourceDocNumber, sSourceDocRev)
    Call oTracking.WebServiceResultAnalysis("GetPLMActionDataFromDocument", oTemp, sSourceDocNumber, sSourceDocRev)
    
    Set oPLMAction = New clsCollection
    Call oPLMAction.InitializeWithDLLclsColObject(oTemp, oPLMAction)
    sPLMAction = oPLMAction.GetItemByKey("FIELD_ACTION_ID")
    
    'Add PLMAction to report
    Call Add_Attribute(oRelatedPVRReport, oActionElem, "PLMAction", sPLMAction)
    
    If sPLMAction <> "" Then
    
        'Get the document collection
        Call WebServiceAccessTool.GetPLMActionDocuments(sPLMAction, oTemp)
        Call oTracking.WebServiceResultAnalysis("GetPLMActionDocuments", oTemp, sPLMAction)
        
        Set oDocColl = New clsCollection
        Call oDocColl.InitializeWithDLLclsColObject(oTemp, oDocColl)
        
        'Add Attached Document to report
        If oDocColl.Count > 0 Then
            For i = 1 To oDocColl.Count
                Set oAttachedDocElem = Add_Element(oRelatedPVRReport, "AttachedDocument", oActionElem)
                Call Add_Attribute(oRelatedPVRReport, oAttachedDocElem, "DocNumber", oDocColl.GetKey(i))
                Call Add_Attribute(oRelatedPVRReport, oAttachedDocElem, "DocRev", oDocColl.GetItem(i))
            Next
        End If
    
        'Looking for the PVR
        If oDocColl.Exists(oParentPart.RefPartNumber & "PVRREF") And sPVRNb = "" Then
            sPVRNb = oParentPart.RefPartNumber & "PVRREF"
            sPVRRev = oDocColl.GetItem(oParentPart.RefPartNumber & "PVRREF")
        End If
        
        'Looking for the NONE
        If oDocColl.Exists(oParentPart.RefPartNumber) And sNoneRev = "" Then
            sNoneRev = oDocColl.GetItem(oParentPart.RefPartNumber)
        End If
        
        'Looking for the drawing
        If oDocColl.Exists(Split(oParentPart.RefPartNumber, "-")(0)) Then
            sDrawingNb = Split(oParentPart.RefPartNumber, "-")(0)
            sDrawingRev = oDocColl.GetItem(Split(oParentPart.RefPartNumber, "-")(0))
        End If
        
        'PVR and None were found we exit
        If sPVRRev <> "" And sNoneRev <> "" Then GoTo UpdateBeforeExit
        
        'No drawing was found
        If sDrawingNb = "" Then GoTo UpdateBeforeExit
        
    'No PLMACTION found
    Else
        GoTo UpdateBeforeExit
    End If
    
    'Get the list ordered document revision
    If oOrderedRevList Is Nothing Then
        Set oOrderedRevList = GetDocRevOrderedList(sDrawingNb)
    End If
    
    'PVR or NONE not found, get the previous revision of the drawing
    If oOrderedRevList.GetIndex(sDrawingRev) > 1 Then
        sSourceDocRev = oOrderedRevList.GetItemByIndex(oOrderedRevList.GetIndex(sSourceDocRev) - 1)
        sSourceDocNumber = sDrawingNb
    'We reached the document with lowest revision
    Else
        GoTo UpdateBeforeExit
    End If
  
Loop

UpdateBeforeExit:
If UCase(sUpdateType) Like "*NONE*" Then
    oParentPart.RefPartRev = sNoneRev
    oParentPart.RefPartStatus = oAttList.GetEnoviaAttributes(oParentPart.RefPartNumber, sNoneRev, False, "Revision Status")
    
    If oParentPart.OriginalPartNumber = oParentPart.RefPartNumber Then
        oParentPart.OriginalPartRev = oParentPart.RefPartRev
        oParentPart.OriginalPartStatus = oParentPart.RefPartStatus
    End If
End If

If UCase(sUpdateType) Like "*PROPOSEDPVR*" Then
    oParentPart.ProposedSourceNb = sPVRNb
    oParentPart.ProposedSourceRev = sPVRRev
End If
If UCase(sUpdateType) Like "*SELECTEDPVR*" Then
    oParentPart.SelectedSourceNb = sPVRNb
    oParentPart.SelectedSourceRev = sPVRRev
End If

Call oTracking.AddInformationLine("GetRelatedPVRandNone", "EXECUTING", "Related PVR and None document - DONE ")

End Sub

Private Function GetEffectiveBDIDoc(ByVal oBDIPartColl As clsCollection, ByVal sTail As String) As String

Dim oRootElem As IXMLDOMNode
Dim oDocumentsElem As IXMLDOMElement
Dim oDocumentElem As IXMLDOMElement
Dim oEffectivityElem As IXMLDOMElement
Dim oNode As IXMLDOMNode

Dim bIsDocEffectiveForTail As Boolean
Dim bEffectiveDocFound As Boolean

Dim sBDIDocNumber, sBDIDocRev As String
Dim i, j, k As Integer

Dim oBDIDoc As clsBDIDoc

Call oTracking.AddInformationLine("GetEffectiveBDIDoc", "EXECUTING", "Getting the effective BDI document")


'Initialize
Set oEffDocReport = New DOMDocument60
oEffDocReport.setProperty "SelectionLanguage", "XPath"
oEffDocReport.async = True

'Create Root element
Set oRootElem = oEffDocReport.CreateElement("EffectiveDocReport")
oEffDocReport.appendChild oRootElem

'BDI Primary Docs Elem
Set oDocumentsElem = Add_Element(oEffDocReport, "BDIDocuments", oRootElem)

'Initial set
bEffectiveDocFound = False

'Scan all BDI Part Rev under Part object
For i = 1 To oBDIPartColl.Count

    'Documents (Primary and Prev)
    For j = 1 To 2
        
        'Get document
        If j = 1 Then
            Set oBDIDoc = oBDIPartColl.GetItemByIndex(i).PrimaryDoc
        Else
            Set oBDIDoc = oBDIPartColl.GetItemByIndex(i).PrevDoc
            If oBDIDoc.DocNumber = "" Then Exit For
        End If
        
        'If the document is already in the list we skip it
        Set oNode = Nothing
        Set oNode = oDocumentsElem.SelectSingleNode("./BDIDocument[@DocRev='" & oBDIDoc.DocRev & "']")
        If Not oNode Is Nothing Then GoTo NextDocument
               
        'Doc element
        Set oDocumentElem = Add_Element(oEffDocReport, "BDIDocument", oDocumentsElem)
        Call Add_Attribute(oEffDocReport, oDocumentElem, "DocNumber", oBDIDoc.DocNumber)
        Call Add_Attribute(oEffDocReport, oDocumentElem, "DocRev", oBDIDoc.DocRev)
        Call Add_Attribute(oEffDocReport, oDocumentElem, "Status", oBDIDoc.Status)
        Call Add_Attribute(oEffDocReport, oDocumentElem, "DocEffectiveForTail", "")
        
        'Effectivity elements
        If oBDIDoc.Effectivity.Count >= 1 Then
        
            'Check if doc rev is effective for the given tail
            bIsDocEffectiveForTail = IsDocInTail(oBDIDoc.Effectivity, sTail)
    
            'Add all effectivities to report
            For k = 1 To oBDIDoc.Effectivity.Count
                
                Set oEffectivityElem = Add_Element(oEffDocReport, "Effectivity", oDocumentElem)
                Call Add_Attribute(oEffDocReport, oEffectivityElem, "From", oBDIDoc.Effectivity.GetItemByIndex(k).GetItem("IN"))
                Call Add_Attribute(oEffDocReport, oEffectivityElem, "To", oBDIDoc.Effectivity.GetItemByIndex(k).GetItem("OUT"))
    
            Next
        End If
        
        'Define if the doc rev is the one for the tail
        If bEffectiveDocFound = True Then
            oDocumentElem.Attributes.getNamedItem("DocEffectiveForTail").nodeValue = "N/A"
        ElseIf bIsDocEffectiveForTail = False Then
            oDocumentElem.Attributes.getNamedItem("DocEffectiveForTail").nodeValue = "False"
        ElseIf oBDIDoc.Status = "CLOSED" Then
            oDocumentElem.Attributes.getNamedItem("DocEffectiveForTail").nodeValue = "True"
            bEffectiveDocFound = True
        Else
            oDocumentElem.Attributes.getNamedItem("DocEffectiveForTail").nodeValue = "False"
        End If

NextDocument:
    Next
Next

'Return
Set oNode = oEffDocReport.SelectSingleNode("EffectiveDocReport/BDIDocuments/BDIDocument[@DocEffectiveForTail='True']")


If Not oNode Is Nothing Then
    sBDIDocNumber = oNode.Attributes.getNamedItem("DocNumber").nodeValue
    sBDIDocRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    
    GetEffectiveBDIDoc = sBDIDocNumber & sBDIDocRev
Else
    GetEffectiveBDIDoc = ""
End If

Call oTracking.AddInformationLine("GetEffectiveBDIDoc", "EXECUTING", "The effective BDI doc is " & sBDIDocNumber & sBDIDocRev)
End Function


Private Sub GenerateLogReport()

Dim oTopElem As IXMLDOMElement
Dim oElem As IXMLDOMElement
Dim oNode As IXMLDOMNode
Dim oNodeList As IXMLDOMNodeList
Dim i As Integer
Dim oPart As clsPartInfo

'Create new xml structure
Set oLogReport = New DOMDocument60
oLogReport.setProperty "SelectionLanguage", "XPath"
oLogReport.async = True

'Create top node
Set oTopElem = oLogReport.CreateElement("Parts")
oLogReport.appendChild oTopElem

'Create comment on top node
If oTopNode.PartIsCI Then
    Call Add_Comment(oLogReport, oTopElem, oTopNode.OriginalPartNumber & " is a CI")
Else
    Call Add_Comment(oLogReport, oTopElem, oTopNode.OriginalPartNumber & " is not a CI")
End If
If oTopNode.SyncFromBSF = True Then
    Call Add_Comment(oLogReport, oTopElem, "PVR updated using best so far structure from ENOVIA")
Else
    Call Add_Comment(oLogReport, oTopElem, "Structure configured for " & oTopNode.Project & "/" & oTopNode.Tail)
End If
If oTopNode.NonCIOption = "BSF" Then
    Call Add_Comment(oLogReport, oTopElem, "User choose to update the non CIs using the Best so Far from Enovia")
Else
    Call Add_Comment(oLogReport, oTopElem, "User choose to update the non CIs using the latest released revision")
End If

Call Add_Comment(oLogReport, oTopElem, "Report generated on " & Format(Now(), "yyyy mm dd") & " at " & Format(Time, "hh:mm:ss"))

'Scan all part in oPartList
For i = 1 To oPartList.Count
    
    'Get the part
    Set oPart = oPartList.GetItemByIndex(i)
    
    'Create new elem
    Set oElem = Add_Element(oLogReport, "Part", oTopElem)
    
    'Add attributes
    Call Add_Attribute(oLogReport, oElem, "OriginalPartNumber", oPart.OriginalPartNumber)
    Call Add_Attribute(oLogReport, oElem, "PartNumber", oPart.UsedPartNumber)
    Call Add_Attribute(oLogReport, oElem, "Revision", oPart.UsedPartRev)
    Call Add_Attribute(oLogReport, oElem, "Status", oPart.UsedPartStatus)
    If oPart.UsedPartRev <> "NA" Then
        Call Add_Attribute(oLogReport, oElem, "DocumentOrganization", oAttList.GetEnoviaAttributes(oPart.UsedPartNumber, oPart.UsedPartRev, False, "Document Organization"))
    Else
        Call Add_Attribute(oLogReport, oElem, "DocumentOrganization", "?")
    End If
    Call Add_Attribute(oLogReport, oElem, "Title", oPart.PartTitle)
    Call Add_Attribute(oLogReport, oElem, "IsCI", CStr(oPart.PartIsCI))
    Call Add_Attribute(oLogReport, oElem, "Type", IIf(oPart.PartType = "Component", "NONE", oPart.PartType))
    Call Add_Attribute(oLogReport, oElem, "EffectiveDwg", IIf(oPart.EffectiveDwgNb = "N/A" Or oPart.EffectiveDwgNb = "", "N/A", oPart.EffectiveDwgNb & oPart.EffectiveDwgRev))
    Call Add_Attribute(oLogReport, oElem, "ProposedSource", IIf(oPart.ProposedSourceNb = "", "N/A", oPart.ProposedSourceNb & oPart.ProposedSourceRev))
    Call Add_Attribute(oLogReport, oElem, "SelectedDwg", IIf(oPart.SelectedDwgNb = "N/A", "N/A", oPart.SelectedDwgNb & oPart.SelectedDwgRev))
    Call Add_Attribute(oLogReport, oElem, "SelectedSource", IIf(oPart.SelectedSourceNb = "", "N/A", oPart.SelectedSourceNb & oPart.SelectedSourceRev))
    
    If oPart.Comment = "No comment" And (oPart.ProposedSourceNb & oPart.ProposedSourceRev) <> (oPart.SelectedSourceNb & oPart.SelectedSourceRev) Then
        Call Add_Attribute(oLogReport, oElem, "Comment", "Selected source is different from proposed source")
    Else
       Call Add_Attribute(oLogReport, oElem, "Comment", oPart.Comment)
    End If
    
    Call Add_Attribute(oLogReport, oElem, "RefPartNumber", IIf(oPart.RefPartNumber = "", "N/A", oPart.RefPartNumber & oPart.RefPartRev))
Next

'Save report
oLogReport.Save sXMLPath & "PVRSync_ConfiguredStructure_Report_" & oTopNode.UsedPartNumber & ".xml"

End Sub

Private Function GetDocRevOrderedList(ByVal sDocNumber As String) As clsCollection

Dim oDocs As clsCollection
Dim oTemp
Dim oOrderedList As New clsCollection
Dim oOrderedColl As New Collection
Dim i, j As Integer

'Retrieve all documents with same base number with web service
Set oTemp = WebServiceAccessTool.GetDocumentByBaseNumber(sDocNumber)
Call oTracking.WebServiceResultAnalysis("GetDocumentByBaseNumber", oTemp, sDocNumber)
Set oDocs = New clsCollection
Call oDocs.InitializeWithDLLclsColObject(oTemp, oDocs)

'Error management
If oDocs.Count = 0 Then
    Set GetDocRevOrderedList = Nothing
    Exit Function
End If

'Remove from the list all document that don't have the same Part Number
For i = oDocs.Count To 1 Step -1
    If oDocs.GetItemByIndex(i).GetItemByKey("FIELD_PART_NUMBER") <> sDocNumber Then oDocs.RemoveByIndex i
Next

'Add all revision to oOrderedColl
For i = 1 To oDocs.Count
    
    'First item
    If oOrderedColl.Count = 0 Then
        oOrderedColl.Add (oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION"))
    
    'Other object
    Else
        For j = 1 To oOrderedColl.Count
            If oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION") < oOrderedColl.Item(j) Then
                oOrderedColl.Add oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION"), , j
                Exit For
            End If
        Next
        If oOrderedColl.Count < i Then oOrderedColl.Add oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION")
    End If
Next

'Transfer object from oOrderedColl to oOrderedList
For i = 1 To oOrderedColl.Count
    Call oOrderedList.Add(oOrderedColl.Item(i), oOrderedColl.Item(i))
Next

'Return
Set GetDocRevOrderedList = oOrderedList

End Function

Private Sub GenerateEffectiveDocReport(ByVal sBDIDocNumber As String, ByVal sBDIDocRev As String, ByVal sPN As String)

Dim sLine() As String
Dim oText As New clsTextGenerator
Dim oDocNode As IXMLDOMNode
Dim oEffNode As IXMLDOMNode
Dim oDocNodeList As IXMLDOMNodeList
Dim oEffNodeList As IXMLDOMNodeList
Dim oAttachedDocList As IXMLDOMNodeList
Dim oAttachedDoc As IXMLDOMNode
Dim sEffectivity As String
Dim i As Integer

Call oTracking.AddInformationLine("GenerateEffectiveDocReport", "EXECUTING", "Generating Effective Document Report")

'Date
Erase sLine
ReDim sLine(0, 1)
sLine(0, 0) = "Date: " & Now()
sLine(0, 1) = 0
Call oText.AddNewLine(sLine)
Call oText.AddEmptyLine

'Step 1
Erase sLine
ReDim sLine(0, 1)
sLine(0, 0) = "Finding effective BDI document revision for " & oTopNode.Tail
sLine(0, 1) = 0
Call oText.AddNewLine(sLine)

'Divider
Call oText.AddLineDivider("-", 127)

'Header
Erase sLine
ReDim sLine(4, 1)
sLine(0, 0) = "BDI Doc Number"
sLine(0, 1) = 0
sLine(1, 0) = "BDI Doc Rev"
sLine(1, 1) = 20
sLine(2, 0) = "BDI Doc Status"
sLine(2, 1) = 40
sLine(3, 0) = "BDI Doc Effectivity"
sLine(3, 1) = 60
sLine(4, 0) = "BDI Doc Effective for " & oTopNode.Tail
sLine(4, 1) = 90
Call oText.AddNewLine(sLine)

'Divider
Call oText.AddLineDivider("-", 127)

'Scan each document
Set oDocNodeList = oEffDocReport.SelectNodes("EffectiveDocReport/BDIDocuments/BDIDocument")
For Each oDocNode In oDocNodeList
    
    'Get all the effectivities
    Set oEffNodeList = oDocNode.SelectNodes("./Effectivity")
    
    Erase sLine
    ReDim sLine(5, 1)
    sLine(0, 0) = oDocNode.Attributes.getNamedItem("DocNumber").nodeValue
    sLine(0, 1) = 0
    sLine(1, 0) = oDocNode.Attributes.getNamedItem("DocRev").nodeValue
    sLine(1, 1) = 20
    sLine(2, 0) = oDocNode.Attributes.getNamedItem("Status").nodeValue
    sLine(2, 1) = 40
    sLine(3, 0) = ""
    sLine(3, 1) = 60
    sLine(4, 0) = ""
    sLine(4, 1) = 75


    'First effectivity
    If oEffNodeList.Length > 0 Then
        Set oEffNode = oEffNodeList.Item(0)
        sLine(3, 0) = "From:" & oEffNode.Attributes.getNamedItem("From").nodeValue
        sLine(3, 1) = 60
        sLine(4, 0) = "To:" & oEffNode.Attributes.getNamedItem("To").nodeValue
        sLine(4, 1) = 75

    End If
    
    sLine(5, 0) = oDocNode.Attributes.getNamedItem("DocEffectiveForTail").nodeValue
    sLine(5, 1) = 90
    Call oText.AddNewLine(sLine)

    'Other effectivities
    If oEffNodeList.Length > 1 Then
        For i = 2 To oEffNodeList.Length
            Set oEffNode = oEffNodeList.Item(i - 1)
            Erase sLine
            ReDim sLine(1, 1)
            sLine(0, 0) = "From:" & Trim(oEffNode.Attributes.getNamedItem("From").nodeValue)
            sLine(0, 1) = 60
            sLine(1, 0) = "To:" & Trim(oEffNode.Attributes.getNamedItem("To").nodeValue)
            sLine(1, 1) = 75
            Call oText.AddNewLine(sLine)
        Next
    End If
    
    Call oText.AddLineDivider("*", 127)
Next


Call WriteTextFile(sEffectiveDocReportPath & "PVRSync_GetEffectiveDocument_" & sPN & ".txt", oText.GetText, True)
Call oTracking.AddInformationLine("GenerateEffectiveDocReport", "EXECUTING", "Effective Document Report was saved")
End Sub


Private Sub Add3DMarker()

Dim oWindow As Window
Dim oProduct As Product
Dim oMarkers3Ds
Dim sArray1(2)
Dim sArray2(2)
Dim oMarker3D As Marker3D

Set oWindow = CATIA.ActiveWindow
Set oProduct = CATIA.ActiveDocument.Product
Set oMarkers3Ds = oProduct.GetTechnologicalObject("Marker3Ds")

'Delete existing 3D Marker
For Each oMarker3D In oMarkers3Ds
    If oMarker3D.Name = "PVR_Sync_Message" Then
        oMarkers3Ds.Remove ("PVR_Sync_Message")
        Exit For
    End If
Next

'Position
sArray1(0) = 0
sArray1(1) = 0
sArray1(2) = 0

sArray2(0) = 0
sArray2(1) = 0
sArray2(2) = 0

'Create 3D Marker
Set oMarker3D = oMarkers3Ds.Add3DText(sArray1, o3DMarkerText.GetText, sArray2, oProduct)

'Format
oMarker3D.TextFont = "Swiss.pfb"
oMarker3D.TextSize = 4#
oMarker3D.Name = "PVR_Sync_Message"
oMarker3D.Update

'Reframe in window
oWindow.Viewers.Item(1).Reframe
End Sub



Private Function CompareBOM(ByVal sPartNumber As String, ByVal sPVRNb As String, sPVRRev As String) As String

Dim i As Integer
Dim oPVRDoc As ProductDocument
Dim oProduct As Product
Dim oChild As Product
Dim sPN As String
Dim sChildPN As String
Dim dTimer As Double
Dim EnoviaDoc As EnoviaDocument
Dim oBOM1 As clsCollection
Dim oBOM2 As clsCollection
Dim bPVRAlreadyOpen As Boolean
Dim EV5product As Product
Dim sType As String

'Initialize
Set oBOM1 = New clsCollection
Set oBOM2 = New clsCollection

'Connect to API
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call ConnectToAPI
End If
On Error Resume Next
Set EnoviaDoc = CATIA.Application
If Err.Number <> 0 Then
    Call oTracking.AddInformationLine("ExtractProductStructureCIPVR", "CODEERROR", "Can't connect to CATIA API")
End If
On Error GoTo 0

'Find the component with Part Number = sPartNumber. It can only be the first level under the top node
Set oPVRDoc = CATIA.ActiveDocument
For Each oProduct In oPVRDoc.Product.Products

    sPN = ""
    On Error Resume Next
    sPN = oProduct.PartNumber
    On Error GoTo 0
    
    If sPartNumber = sPN Then Exit For
    
    Set oProduct = Nothing
Next

'No product found, we can't compare BOM.
If oProduct Is Nothing Then
    CompareBOM = "Different BOM"
    Exit Function
End If

'Extract the BOM
Call ExtractBOM(oBOM1, oProduct, False)

'Check if PVR already open
bPVRAlreadyOpen = True
Set oPVRDoc = Nothing
On Error Resume Next
Set oPVRDoc = CATIA.Documents.Item(sPVRNb & sPVRRev & ".CATProduct")
On Error GoTo 0

'Load PVR from Enovia
If oPVRDoc Is Nothing Then

    Call oTracking.AddInformationLine("CompareBOM", "EXECUTING", sPVRNb & sPVRRev & ".CATProduct will be loaded from ENOVIA.")
    
    dTimer = Timer
    On Error Resume Next
    bPVRAlreadyOpen = False
    Set EV5product = EnoviaDoc.OpenPartDocument(sPVRNb, sPVRRev)
    If Err.Description <> "" Then Call oTracking.AddInformationLine("CompareBOM", "WARNING", "After load the error description is: " & Err.Description)
    
    Do
        Set oPVRDoc = EV5product.ReferenceProduct.Parent
        
        If Not oPVRDoc Is Nothing Then Exit Do
        
        DoEvents
        Sleep 150
        Call oTracking.AddInformationLine("CompareBOM", "WARNING", "Sleeping while waiting for document to load")
        If Abs(Timer - dTimer) > 5 Then
            Exit Do
        End If
    Loop
End If
    
'No PVR found, we can't compare BOM. We use what we have in the loaded PVR
If oPVRDoc Is Nothing Then
    CompareBOM = "Same BOM"
    Exit Function
End If
    
'Extract the BOM
Call ExtractBOM(oBOM2, oPVRDoc.Product, False)

'Close de PVR
If bPVRAlreadyOpen = False Then
    oPVRDoc.Close
End If

'Compare the BOM
For i = oBOM1.Count To 1 Step -1

    'Item exist in BOM2
    If oBOM2.Exists(oBOM1.GetKey(i)) Then
    
        'If qty are the same we remove them from both collection
        If CInt(oBOM1.GetItemByIndex(i).GetItemByKey("Qty")) = CInt(oBOM2.GetItemByKey(oBOM1.GetKey(i)).GetItemByKey("Qty")) Then
            Call oBOM2.RemoveByKey(oBOM1.GetKey(i))
            Call oBOM1.RemoveByIndex(i)
        End If
    
    End If
Next

'Return
If oBOM1.Count = 0 And oBOM2.Count = 0 Then
    CompareBOM = "Same BOM"
Else
    CompareBOM = "Different BOM"
End If
End Function



Private Sub ExtractBOM(ByRef oBOM As clsCollection, ByVal oParent As Product, ByVal bInsideSP As Boolean)

Dim oChild As Product
Dim sChildPN As String
Dim sChildRev As String
Dim sString As String
Dim oItem As clsCollection
Dim sDefiningPart As String
Dim sDatasetType As String
Dim sType As String

For Each oChild In oParent.Products

    'Get PN
    On Error Resume Next
    sChildPN = ""
    sChildPN = oChild.PartNumber
    If sChildPN = "" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    sChildPN = oChild.PartNumber
    On Error GoTo 0
    
    'We have a PN
    If sChildPN <> "" Then
    
        'Get the type
        sType = GetType(oChild)

        'Override part number by using the file name for CATPart and CATProduct
        If sType = "CATPart" Or sType = "CATProduct" Then
            sChildPN = oChild.ReferenceProduct.Parent.Name
            sChildPN = Left(Split(sChildPN, ".")(0), Trim(Len(Split(sChildPN, ".")(0)) - 2))
        End If

        'Get dataset type and defining part attribute
        sDefiningPart = ""
        sDatasetType = ""
        If Not (sType = "Component" And bInsideSP = True) Then
            sDefiningPart = oAttList.GetEnoviaAttributes(sChildPN, "LatestRev", False, "Defining Part")
            sDatasetType = oAttList.GetEnoviaAttributes(sChildPN, "LatestRev", False, "Dataset Type")
        End If
        
        'Remove (Don't use this part) and (CANCELLED) from defining part
        sDefiningPart = UCase(sDefiningPart)
        sDefiningPart = Trim(Replace(sDefiningPart, "(DON'T USE THIS PART)", ""))
        sDefiningPart = Trim(Replace(sDefiningPart, "(CANCELLED)", ""))
        
        'Override part number with defining part
        If (sDatasetType = "CATALOG LIGHT VERSION" Or sDatasetType = "FLEXIBLE REPRESENTATION") And sDefiningPart <> "" Then
            sChildPN = sDefiningPart
        End If
        
        'Recursive
        If bInsideSP = False And sType = "CATProduct" And sChildPN Like "*SP*" Then
            Call ExtractBOM(oBOM, oChild, True)
        ElseIf bInsideSP = True And sType = "Component" Then
            Call ExtractBOM(oBOM, oChild, bInsideSP)
        Else
        
            'Add part to BOM collection or update quantity
            If Not oBOM.Exists(sChildPN) Then
                Set oItem = New clsCollection
                Call oItem.Add("Qty", 1)
                Call oBOM.Add(sChildPN, oItem)
            Else
                Call oBOM.GetItemByKey(sChildPN).SetItemByKey("Qty", CInt(oBOM.GetItemByKey(sChildPN).GetItemByKey("Qty")) + 1)
            End If
        End If
        
    'Can't find the Part Number
    Else
        'Do nothing
    End If
Next


End Sub




















