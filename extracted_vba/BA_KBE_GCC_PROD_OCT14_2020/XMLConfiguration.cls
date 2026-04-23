Attribute VB_Name = "XMLConfiguration"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private configNode As IXMLDOMNode
Private seedNode As IXMLDOMElement
Private paramNode As IXMLDOMElement
Private attNode As IXMLDOMAttribute
Private sMissingParamsInSeedAssy As String

Public Sub ExportData(ByVal sXMLFile As String, ByVal sConfigName As String, ByVal sConfigDesc As String, ByVal sToolVersion As String, ByVal Functionality As String, Optional ByVal oSelectedSource As AnyObject = Nothing, Optional ByVal oSelectedSourcePart As Part = Nothing)
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    timeStart = DateTime.Timer
    
    Call frmProgress.progressBarInitialize("Manage Configuration / Parameters")
    
    iProgressCount1Max = 4
    iProgressCount1 = 1
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Searching for DT Parameters to export..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    If Functionality = "New" Then
        Set xmlDoc = New MSXML2.DOMDocument50
        Set xmlRoot = xmlDoc.createElement("configurationData")
        Set xmlDoc.documentElement = xmlRoot
        
        Set attNode = xmlDoc.createAttribute("version")
        xmlRoot.attributes.setNamedItem attNode
        attNode.Value = sToolVersion
    ElseIf Functionality = "Add" Then
        Set xmlDoc = New MSXML2.DOMDocument50
        Set xmlDoc = New DOMDocument50
        xmlDoc.Load (sXMLFile)
        Set xmlRoot = xmlDoc.documentElement
    End If
    
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & sConfigName & Chr(34) & "]")
    If configNode Is Nothing Then
        Set configNode = xmlDoc.createElement("ConfigurationID")
        Call xmlRoot.appendChild(configNode)
        Set attNode = xmlDoc.createAttribute("id")
        configNode.attributes.setNamedItem attNode
        attNode.Value = sConfigName
    Else
        If MsgBox("Configuration already exits." & vbCrLf & "Would you like to overwrite the exiting Configuration?", vbYesNo + vbInformation) = vbNo Then
            Exit Sub
        End If
    End If
    
    iProgressCount1 = iProgressCount1 + 1 '2
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    If Not oSelectedSource Is Nothing Then
        'Manage Parameters > Copy / Save
        
        Set oParameters = oSelectedSourcePart.Parameters
        Set oParameters = oParameters.SubList(oSelectedSource, True)
        Set aCopyParams = New Collection
        
        For Each oParameter In oParameters
            If oParameter.Name Like "*_DT" Then
                If oParameter.OptionalRelation Is Nothing Then aCopyParams.Add oParameter, oParameter.Name
            End If
        Next
        
        'All the parameters under the Assemble objects
        If TypeName(oSelectedSource) = "Body" Then
            Call ScanSubBodies(oSelectedSource, aCopyParams)
        End If

        If aCopyParams.Count = 0 Then
            MsgBox "No DT Parameter found in Source!     ", vbExclamation, "Warning"
            Exit Sub
        End If
    Else
        'Export Data
        On Error Resume Next
        Set aCopyParams = New Collection
        
        For Each iObject In oSeedPartCollection 'Scan collection of Seed Part
            
            sSeedPartName = Split(iObject.Name, "-")(0)
            Set seedNode = Nothing
            Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
            
            If seedNode Is Nothing Then
                Set seedNode = xmlDoc.createElement("seedFile")
                Call configNode.appendChild(seedNode)
                Set attNode = xmlDoc.createAttribute("name")
                seedNode.attributes.setNamedItem attNode
                attNode.Value = sSeedPartName
            End If
            
            'Retrieve all the parameters of the Seed Part
            Set oParameters = iObject.ReferenceProduct.Parent.Part.Parameters
            For Each oParameter In oParameters
                If oParameter.Name Like "*_DT" Then
                    aCopyParams.Add oParameter
                End If
            Next
            
            If aCopyParams.Count = 0 Then
                MsgBox "No DT Parameter found in Source!     ", vbExclamation, "Warning"
                Exit Sub
            End If
        Next
        On Error GoTo 0
    End If
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    iProgressCount1Max = aCopyParams.Count
    iProgressCount1 = 0
    SkipParams = 0
    sLog = ""
    On Error Resume Next
    For i = 1 To aCopyParams.Count
        
        Set oParameter = aCopyParams.Item(i)
        iProgressCount1 = iProgressCount1 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Exporting Parameters (" & CStr(iProgressCount1) & "/" & CStr(iProgressCount1Max) & ")        [Skip : " & SkipParams & "]"
        Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
        
        Set oRelation = Nothing
        Set oRelation = oParameter.OptionalRelation
        If oRelation Is Nothing Then
            sSeedPartName = ""
            sSeedPartName = Split(oParameter.Context.Name, "-")(0)
            Set seedNode = Nothing
            Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
            
            If seedNode Is Nothing Then
                Set seedNode = xmlDoc.createElement("seedFile")
                Call configNode.appendChild(seedNode)
                Set attNode = xmlDoc.createAttribute("name")
                seedNode.attributes.setNamedItem attNode
                attNode.Value = sSeedPartName
            End If
            
            If Not seedNode Is Nothing Then
                Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & oParameter.Name & Chr(34) & "]")
                If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
                
                    If paramNode Is Nothing Then
                    
                        If oSelectedSource Is Nothing Then
                            If oParameter.Name Like "*\ConfigurationName_DT" Then oParameter.Value = sConfigName
                            If oParameter.Name Like "*\ConfigurationDescription_DT" Then oParameter.Value = sConfigDesc
                        End If
                        
                        Set paramNode = xmlDoc.createElement("parameter")
                        Call seedNode.appendChild(paramNode)
                            
                        Set attNode = xmlDoc.createAttribute("name")
                        paramNode.attributes.setNamedItem attNode
                        attNode.Value = oParameter.Name
                    
                        Set attNode = xmlDoc.createAttribute("value")
                        paramNode.attributes.setNamedItem attNode
                        attNode.Value = oParameter.ValueAsString
                    Else
                    
                        ' Value changed Params
                        paramNode.attributes.Item(0).Text = oParameter.Name
                        On Error Resume Next
                        paramNode.attributes.Item(1).Text = oParameter.ValueAsString
                        If Err.Number <> 0 Then
                            paramNode.attributes.Item(1).Text = oParameter.Value
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        Else
            sLog = sLog & oParameter.Name & vbCrLf
            SkipParams = SkipParams + 1
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next i
    On Error GoTo 0
    
    iProgressCount1Max = 4
    iProgressCount1 = 3
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Sorting the Seed File nodes..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    ' Sorting XML Seed Nodes
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & sConfigName & Chr(34) & "]")
    iNoOfParams = 0
    For i = configNode.childNodes.Length - 1 To 1 Step -1
        iNoOfParams = iNoOfParams + 1
        If configNode.childNodes.Item(i - 1).attributes.Item(0).Text > configNode.childNodes.Item(i).attributes.Item(0).Text Then
            Set xmlNode = configNode.childNodes.Item(i)
            configNode.removeChild xmlNode
            configNode.insertBefore xmlNode, configNode.childNodes.Item(i - 1)
        End If
    Next i
    
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Finalizing the Export..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    Set attNode = xmlDoc.createAttribute("numberOfParameters")
    configNode.attributes.setNamedItem attNode
    
    iNoOfParams = 0
    For i = 0 To configNode.childNodes.Length - 1
        iNoOfParams = iNoOfParams + configNode.childNodes.Item(i).childNodes.Length
    Next i
    
    attNode.Value = iNoOfParams
    
    iProgressCount1 = iProgressCount1 + 1 '4
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = frmProgress.lblMessageMain.Caption
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    Set attNode = xmlDoc.createAttribute("date")
    configNode.attributes.setNamedItem attNode
    attNode.Value = Date
    
    Set attNode = xmlDoc.createAttribute("start")
    configNode.attributes.setNamedItem attNode
    attNode.Value = Format(timeStart / 86400, "h:mm:ss AMPM")
    
    timeEnd = DateTime.Timer
    Set attNode = xmlDoc.createAttribute("finish")
    configNode.attributes.setNamedItem attNode
    attNode.Value = Format(timeEnd / 86400, "h:mm:ss AMPM")
    
    If oSelectedSource Is Nothing Then
        configNode.childNodes.Item(0).childNodes.Item(1).Text = sConfigName
        configNode.childNodes.Item(0).childNodes.Item(2).Text = sConfigDesc
    End If
    
    xmlDoc.loadXML (xmlDoc.xml)
    xmlDoc.Save (sXMLFile)
    
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_ExportData.txt"
        sLog = "The following Read Only parameters were not exported:" & vbCrLf & sLog
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If
    
    MsgBox "Current Configuration has been Exported Successfully!     ", vbInformation, "Manage Configuration / Parameters"
    frmProgress.Hide
    
End Sub

Public Sub ImportData(ByVal xmlFile As String, ByVal ConfigName As String)
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    timeStart = DateTime.Timer
    Call frmProgress.progressBarInitialize("Manage Configuration / Parameters")
    
    iProgressCount1Max = 4
    iProgressCount1 = 1
    iProgressCount2Max = 1
    iProgressCount2 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Preparing to Import Parameters..."
    Call frmProgress.progressBarRepaint("Searching for DT Parameters to import...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    'load xml file and get the Doc Element
    Set xmlDoc = New MSXML2.DOMDocument50
    xmlDoc.Load (xmlFile)
    Set xmlRoot = xmlDoc.documentElement
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & ConfigName & Chr(34) & "]")
    
    'Select all of the DT Params
    Set aPasteParams = New Collection
    For Each iObject In oSeedPartCollection
        If InStr(1, sXMLSeedPartsList, "SEED_PART_" & Split(Split(iObject.Name, "SEED_PART_")(1), "-")(0) & ";") > 0 Then
            'Retrieve all the parameters of the Seed Part
            Set oParameters = iObject.ReferenceProduct.Parent.Part.Parameters
            For Each oParameter In oParameters
                If oParameter.Name Like "*_DT" Then
                    aPasteParams.Add oParameter
                End If
            Next
        End If
    Next
    
    If Not aPasteParams.Count > 0 Then timeEnd = DateTime.Timer: Exit Sub
    sSeedPartName = ""
    UserDefinedParams = ""
    NodeCnt = 0
    
    iProgressCount1 = iProgressCount1 + 1 '2
    iProgressCount2Max = aPasteParams.Count
    iProgressCount2 = 0
    sLog = ""
    For i = 1 To aPasteParams.Count
        
        Set oParameter = aPasteParams.Item(i)
        Set oRelation = oParameter.OptionalRelation
        
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Importing Parameters (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
            If oRelation Is Nothing And InStr(1, sXMLSeedPartsList, "SEED_PART_" & Split(Split(oParameter.Context.Name, "-")(0), "SEED_PART_")(1) & ";") > 0 Then
                
                If sSeedPartName <> Split(oParameter.Context.Name, "-")(0) Then
                    sSeedPartName = Split(oParameter.Context.Name, "-")(0)
                    Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
                End If
                
                If oSeedAssyMigrateConfigFile = False Then
                    sParameterName = oParameter.Name
                Else
                    sParameterName = Replace(oParameter.Name, Split(sSeedAssyToolVersion, ".")(0), Split(sSeedAssyMigratedVersion, ".")(0), vbTextCompare)
                End If
                
                Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                
                If paramNode Is Nothing Then
                    On Error Resume Next
                        sParameterName = Replace(oParameter.Name, Split(Split(oParameter.Name, "\")(0), "-")(1), Split(Split(seedNode.childNodes(1).xml, "\")(0), "-")(1), vbTextCompare)
                        Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                    On Error GoTo 0
                End If
                
                On Error Resume Next
                If UCase(paramNode.attributes.Item(1).Text) Like "*USER*DEFINED*" Then
                    If Err.Number = 0 Then
                        UserDefinedParams = UserDefinedParams & "<||>" & oParameter.Name
                    End If
                End If
                Err.Clear
                
                If paramNode.attributes.Item(1).Text = "" Then
                    oParameter.ValuateFromString (paramNode.nodeTypedValue)
                Else
                    oParameter.ValuateFromString (paramNode.attributes.Item(1).Text)
                End If
                
                If Err.Number <> 0 Then
                    oParameter.Value = paramNode.attributes.Item(1).Text
                    Err.Clear
                    If UCase(paramNode.attributes.Item(1).Text) Like "*USER*DEFINED*" Then
                        If Err.Number = 0 Then
                            UserDefinedParams = UserDefinedParams & "<||>" & oParameter.Name
                        End If
                    End If
                End If
                On Error GoTo 0
            Else
                sLog = sLog & oParameter.Name & vbCrLf
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    vUserDefinedParams = Split(UserDefinedParams, "<||>")
    MsgStr = ""
    
    i = 0
    For i = 1 To UBound(vUserDefinedParams)
        MsgStr = MsgStr & vbCrLf & i & ". " & vUserDefinedParams(i)
    Next i
    
    If UBound(vUserDefinedParams) > 0 Then MsgBox "Make sure the following User Defined Parameter(s) have the Correct UDGs" & vbCrLf & MsgStr, vbInformation, "USERDEFINED Parameters..."
    
    'Report the missing params
    iProgressCount1 = iProgressCount1 + 1 '3
    iProgressCount2Max = 1
    iProgressCount2 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Checking the Missing Parameters..."
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    MissingParamFileName = Replace(xmlFile, Split(xmlFile, "\")(UBound(Split(xmlFile, "\"))), "") & "MissingParams.xml"
    Set ParamFSO = CreateObject("Scripting.FileSystemObject")
    ParamFSO.CopyFile xmlFile, MissingParamFileName
    Set xmlDoc = New MSXML2.DOMDocument50
    xmlDoc.Load MissingParamFileName
    Set xmlRoot = xmlDoc.documentElement
    
    'Keep only the selected Configuration
    For i = xmlRoot.childNodes.Length - 1 To 0 Step -1
        If xmlRoot.childNodes.Item(i).attributes.Item(0).nodeTypedValue <> ConfigName Then
            Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & xmlRoot.childNodes.Item(i).attributes.Item(0).nodeTypedValue & Chr(34) & "]")
            xmlRoot.removeChild configNode
        End If
    Next i
    
    'Keep only the selected SeedParts(s)
    For i = xmlRoot.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1
        sSeedPartName = xmlRoot.childNodes.Item(0).childNodes.Item(i).attributes.Item(0).nodeValue
        If Not InStr(1, sXMLSeedPartsList, "_" & Split(sSeedPartName, "_")(UBound(Split(sSeedPartName, "_"))) & ";", vbTextCompare) > 0 Then
            xmlRoot.childNodes.Item(0).removeChild xmlRoot.childNodes.Item(0).childNodes.Item(i)
        End If
    Next i
    
    sSeedPartName = ""
    MissParams = ""
    
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & ConfigName & Chr(34) & "]")
    
    iProgressCount2 = 0
    iProgressCount2Max = aPasteParams.Count
    For i = 1 To aPasteParams.Count
        
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Checking the Missing Parameters (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        Set oParameter = aPasteParams.Item(i)
        Set oRelation = oParameter.OptionalRelation
        
        If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
            If oRelation Is Nothing And InStr(1, sXMLSeedPartsList, "SEED_PART_" & Split(Split(oParameter.Context.Name, "-")(0), "SEED_PART_")(1) & ";") > 0 Then
                
                If sSeedPartName <> Split(oParameter.Context.Name, "-")(0) Then
                    sSeedPartName = Split(oParameter.Context.Name, "-")(0)
                    Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
                End If
                
                If oSeedAssyMigrateConfigFile = False Then
                    sParameterName = oParameter.Name
                Else
                    sParameterName = Replace(oParameter.Name, Split(sSeedAssyToolVersion, ".")(0), Split(sSeedAssyMigratedVersion, ".")(0), vbTextCompare)
                End If
                MissParamparamNode = ""
                Set MissParamparamNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                If MissParamparamNode Is Nothing Then   ' missing in xml
                    On Error Resume Next
                        sParameterName = Replace(oParameter.Name, Split(Split(oParameter.Name, "\")(0), "-")(1), Split(Split(seedNode.childNodes(1).xml, "\")(0), "-")(1), vbTextCompare)
                        Set MissParamparamNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                    On Error GoTo 0
                End If
                If MissParamparamNode Is Nothing Then   ' missing in xml
                    MissParams = MissParams & "<||>" & sParameterName & " value=" & oParameter.ValueAsString
                Else
                    Err.Number = 0
                    On Error Resume Next
                    seedNode.removeChild MissParamparamNode   ' missing in Seed
                    On Error GoTo 0
                End If
                
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    On Error GoTo 0
    
    iProgressCount1 = iProgressCount1 + 1 '4
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Finalizing the Import..."
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    For i = xmlRoot.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1
        
        MissSeedPartName = xmlRoot.childNodes.Item(0).childNodes.Item(i).attributes.Item(0).nodeValue
        sMissingParamsInSeedAssy = ""
        For j = xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Length - 1 To 0 Step -1 'For each parameter under SeedPart node
            
            If Not Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1) Like "*" & MissSeedPartName & "*" Then
                sMissingParamsInSeedAssy = sMissingParamsInSeedAssy & "<||>" & MissSeedPartName & "\" & Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1)
            Else
                sMissingParamsInSeedAssy = sMissingParamsInSeedAssy & "<||>" & Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1)
            End If
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        Next j
        Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & MissSeedPartName & Chr(34) & "]")
        configNode.removeChild seedNode
    Next i
    sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, """", "")
    sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, "/>", "")
    sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, "</seedFile>", "")
    
    TotalParamNodeCnt = 0
    For ParamNodeCnt = 0 To configNode.childNodes.Length - 1
        TotalParamNodeCnt = TotalParamNodeCnt + configNode.childNodes.Item(ParamNodeCnt).childNodes.Length
    Next ParamNodeCnt
    
    Set MissingParamsinXMLNode = xmlDoc.createElement("ParametersInSEED_NOT_FoundInXML")
    Call configNode.appendChild(MissingParamsinXMLNode)
    Set attNode = xmlDoc.createAttribute("ParametersInSEED_NOT_FoundInXML")
    
    For i = 0 To UBound(Split(MissParams, "<||>")) - 1
        Set paramNode = xmlDoc.createElement("parameter")
        Call MissingParamsinXMLNode.appendChild(paramNode)
        
        Set attNode = xmlDoc.createAttribute("name")
        paramNode.attributes.setNamedItem attNode
        attNode.Value = Split(Split(MissParams, "<||>")(i + 1), " value=")(0)
        
        Set attNode = xmlDoc.createAttribute("value")
        paramNode.attributes.setNamedItem attNode
        attNode.Value = Split(Split(MissParams, "<||>")(i + 1), " value=")(1)
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    Set MissingParamsinSeedNode = xmlDoc.createElement("ParametersInXML_NOT_FoundInSEED")
    Call configNode.appendChild(MissingParamsinSeedNode)
    Set attNode = xmlDoc.createAttribute("ParametersInXML_NOT_FoundInSEED")
    
    For i = 0 To UBound(Split(sMissingParamsInSeedAssy, "<||>")) - 1
        Set paramNode = xmlDoc.createElement("parameter")
        Call MissingParamsinSeedNode.appendChild(paramNode)
        
        Set attNode = xmlDoc.createAttribute("name")
        paramNode.attributes.setNamedItem attNode
        attNode.Value = Split(Split(sMissingParamsInSeedAssy, "<||>")(i + 1), " value=")(0)
        
        Set attNode = xmlDoc.createAttribute("value")
        paramNode.attributes.setNamedItem attNode
        attNode.Value = Split(Split(sMissingParamsInSeedAssy, "<||>")(i + 1), " value=")(1)
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    configNode.attributes.Item(1).nodeValue = UBound(Split(sMissingParamsInSeedAssy, "<||>")) + UBound(Split(MissParams, "<||>"))
    
    'Write Ending Time
    timeEnd = DateTime.Timer
    configNode.attributes.Item(2).nodeValue = Format(timeStart / 86400, "h:mm:ss AMPM")
    configNode.attributes.Item(3).nodeValue = Format(timeEnd / 86400, "h:mm:ss AMPM")
    
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_ImportData.txt"
        sLog = "The following Read Only parameters were not modified:" & vbCrLf & sLog
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If
    
    xmlDoc.Save MissingParamFileName
    TotalMissParams = IIf(UBound(Split(sMissingParamsInSeedAssy, "<||>")) < 0, 0, UBound(Split(sMissingParamsInSeedAssy, "<||>"))) + IIf(UBound(Split(MissParams, "<||>")) < 0, 0, UBound(Split(MissParams, "<||>")))
    If TotalMissParams > 0 Then
        If MsgBox("Selected Configuration has been Imported Successfully!" & vbCrLf & "There are few Parameters missing in XML/SEED_ASSY.     " & vbCrLf & "Do you want to view the Report?     ", vbYesNo + vbInformation, "Manage Configuration / Parameters") = vbYes Then
            ShellExecute 0, "", MissingParamFileName, "", "", 0
        End If
    Else
        ParamFSO.DeleteFile MissingParamFileName
        MsgBox "Selected Configuration has been Imported Successfully!     ", vbInformation, "Manage Configuration / Parameters"
    End If
    
    frmProgress.Hide
    
End Sub

Public Sub CompareData(ByVal xmlFile As String, ByVal ConfigName As String, Optional ByVal xmlFile2 As String, Optional ByVal ConfigName2 As String)
    
    Dim sXMLFileName2 As String
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    timeStart = DateTime.Timer
    Call frmProgress.progressBarInitialize("Manage Configuration / Parameters")
    
    iProgressCount1Max = 4
    iProgressCount1 = 0
    iProgressCount2Max = 2
    iProgressCount2 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Preparing to Compare Parameters..."
    Call frmProgress.progressBarRepaint("Comparing Data...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
    If xmlFile2 <> "" Then
        
        Set xmlDoc = New MSXML2.DOMDocument50
        xmlDoc.Load xmlFile
        
        sXMLFileName2 = Replace(xmlFile2, Split(xmlFile2, "\")(UBound(Split(xmlFile2, "\"))), "") & "RefinedParamsList.xml"
        Set ParamFSO = CreateObject("Scripting.FileSystemObject")
        ParamFSO.CopyFile xmlFile2, sXMLFileName2
        Set xmlDoc2 = New MSXML2.DOMDocument50
        xmlDoc2.Load sXMLFileName2
        Set xmlRoot2 = xmlDoc2.documentElement
        
        'Keep only the selected configuration
        If ConfigName2 <> "" Then
            For i = xmlRoot2.childNodes.Length - 1 To 0 Step -1
                If xmlRoot2.childNodes.Item(i).attributes.Item(0).nodeTypedValue <> ConfigName2 Then
                    xmlRoot2.removeChild xmlRoot2.childNodes.Item(i)
                End If
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next i
        End If
        
        'Keep only the selected SeedParts(s)
        For i = xmlRoot2.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1
            sSeedPartName = xmlRoot2.childNodes.Item(0).childNodes.Item(i).attributes.Item(0).nodeValue
            If Not InStr(1, sXMLSeedPartsList, "_" & Split(sSeedPartName, "_")(UBound(Split(sSeedPartName, "_"))) & ";", vbTextCompare) > 0 Then
                xmlRoot2.childNodes.Item(0).removeChild xmlRoot2.childNodes.Item(0).childNodes.Item(i)
            End If
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        Next i
        
        For i = xmlRoot2.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1 'For each SeedPart in selected config
            Set seedNode = xmlRoot2.childNodes.Item(0).childNodes.Item(i)
            For j = seedNode.childNodes.Length - 1 To 0 Step -1 'For each parameter under SeedPart node
                
                Set xmlNode2 = seedNode.childNodes.Item(j)
                
                xmlDoc.setProperty "SelectionLanguage", "XPath"
                
                On Error Resume Next
                Set xmlNode = Nothing
                Set xmlNode = xmlDoc.selectSingleNode("//*[@name=" & Chr(34) & xmlNode2.attributes.Item(0).nodeTypedValue & Chr(34) & "]")
                On Error GoTo 0
                
                If Not xmlNode Is Nothing Then
                    If xmlNode2.attributes.Item(0).nodeTypedValue = xmlNode.attributes.Item(0).nodeTypedValue Then
                        If xmlNode2.attributes.Item(1).nodeTypedValue = xmlNode.attributes.Item(1).nodeTypedValue Then
                            seedNode.removeChild xmlNode2
                        End If
                    End If
                End If
                
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next j
        Next i
        
        k = 0
        For i = 0 To xmlRoot2.childNodes.Length - 1
            For j = 0 To xmlRoot2.childNodes.Item(i).childNodes.Length - 1
                xmlRoot2.childNodes.Item(i).attributes.Item(1).nodeValue = xmlRoot2.childNodes.Item(i).childNodes.Item(j).childNodes.Length
                k = k + xmlRoot2.childNodes.Item(i).attributes.Item(1).nodeValue
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next j
        Next i
        
        xmlDoc2.Save sXMLFileName2
        
        If k > 0 Then
            If MsgBox("Selected Configuration has been Compared Successfully.     " & vbCrLf & "Do you want to view the Report?     ", vbYesNo + vbInformation, "Manage Configuration / Parameters") = vbYes Then
                ShellExecute 0, "", sXMLFileName2, "", "", 0
            End If
        Else
            ParamFSO.DeleteFile sXMLFileName2
            MsgBox "Selected Configuration has been Compared Successfully!     ", vbInformation, "Manage Configuration / Parameters"
        End If
        
    Else
        'load xml file and get the Doc Element
        Set xmlDoc = New MSXML2.DOMDocument50
        xmlDoc.Load (xmlFile)
        Set xmlRoot = xmlDoc.documentElement
        
        Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & ConfigName & Chr(34) & "]")
        
        iProgressCount2 = 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Preparing to Compare Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        'Select all of the DT Params
        oSelectionFromSeedPart.Clear
        Set aCopyParams = New Collection
        For Each iObject In oSeedPartCollection
            If InStr(1, sXMLSeedPartsList, "SEED_PART_" & Split(Split(iObject.Name, "SEED_PART_")(1), "-")(0) & ";") > 0 Then
                Set oParameters = iObject.Parameters
                For Each oParameter In oParameters
                    If oParameter.Name Like "*_DT" Then
                        aCopyParams.Add oParameter
                    End If
                Next
            End If
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        Next
        
        If Not aCopyParams.Count > 0 Then timeEnd = DateTime.Timer: Exit Sub
        
        iProgressCount2 = 2
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Preparing to Compare Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
        sSeedPartName = ""
        UserDefinedParams = ""
        MisMatchedParams = ""
        
        iProgressCount1 = iProgressCount1 + 1 '1
        iProgressCount2 = 0
        iProgressCount2Max = aCopyParams.Count
        For i = 1 To aCopyParams.Count
            
            Set oParameter = aCopyParams.Item(i)
            Set oRelation = oParameter.OptionalRelation
            
            iProgressCount2 = iProgressCount2 + 1
            timeEnd = DateTime.Timer
            sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
            sProgBarComment = "Comparing Parameters (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
    
            If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
                If oRelation Is Nothing And InStr(1, sXMLSeedPartsList, "SEED_PART_" & Split(Split(oParameter.Context.Name, "-")(0), "SEED_PART_")(1) & ";") > 0 Then
                    
                    If sSeedPartName <> Split(oParameter.Context.Name, "-")(0) Then
                        sSeedPartName = Split(oParameter.Context.Name, "-")(0)
                        Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
                    End If
                    
                    If oSeedAssyMigrateConfigFile = False Then
                        sParameterName = oParameter.Name
                    Else
                        sParameterName = Replace(oParameter.Name, Split(sSeedAssyToolVersion, ".")(0), Split(sSeedAssyMigratedVersion, ".")(0), vbTextCompare)
                    End If
                    
                    On Error Resume Next
                    Set paramNode = Nothing
                    Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                    If paramNode Is Nothing Then
                            sParameterName = Replace(oParameter.Name, Split(Split(oParameter.Name, "\")(0), "-")(1), Split(Split(seedNode.childNodes(1).xml, "\")(0), "-")(1), vbTextCompare)
                            Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                    End If
                    
                    Err.Clear
                    If UCase(paramNode.attributes.Item(1).Text) Like "*USER*DEFINED*" Then
                        If Err.Number = 0 Then
                            UserDefinedParams = UserDefinedParams & "<||>" & oParameter.Name
                        End If
                    End If
                    Err.Clear
                    
                    XMLParamValue = paramNode.attributes.Item(1).Text
                    
                    If Err.Number = 0 Then
                        UnitVal = ""
                        UnitVal = oParameter.Unit.Symbol
                        If UnitVal = "" Then
                            If UCase(oParameter.Value) <> UCase(paramNode.attributes.Item(1).Text) Then
                                MisMatchedParams = MisMatchedParams & "<||>" & oParameter.Name & "MisMatchedValues==" & oParameter.Value & "<Value>" & paramNode.attributes.Item(1).Text
                            End If
                        Else
                            If UnitVal = "mm" Then
                                If InStr(paramNode.attributes.Item(1).Text, ".") > 0 Then
                                    PrecLen = Len(Replace(Split(paramNode.attributes.Item(1).Text, ".")(1), "in", ""))
                                Else
                                    PrecLen = 5
                                End If
                                
                                If Round((Val(oParameter.Value) / 25.4), PrecLen) & "in" <> paramNode.attributes.Item(1).Text Then
                                    MisMatchedParams = MisMatchedParams & "<||>" & oParameter.Name & "MisMatchedValues==" & Round((Val(oParameter.Value) / 25.4), PrecLen) & "in" & "<Value>" & paramNode.attributes.Item(1).Text
                                End If
                            ElseIf UnitVal = "deg" Then
                                If oParameter.Value & "deg" <> paramNode.attributes.Item(1).Text Then
                                    MisMatchedParams = MisMatchedParams & "<||>" & oParameter.Name & "MisMatchedValues==" & oParameter.Value & "deg" & "<Value>" & paramNode.attributes.Item(1).Text
                                End If
                            Else
                                MsgBox "Parameters in """ & UnitVal & """ are not supported!   ", vbExclamation, "Warning"
                            End If
                        End If
                    End If
                    
                    If Err.Number <> 0 Then
                        Err.Clear
                        If UCase(paramNode.attributes.Item(1).Text) Like "*USER*DEFINED*" Then
                            If Err.Number = 0 Then
                                UserDefinedParams = UserDefinedParams & "<||>" & oParameter.Name
                            End If
                        End If
                    End If
                    On Error GoTo 0
                End If
            End If
        
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        Next
        
        vMisMatchedParams = Split(MisMatchedParams, "<||>")
        
        ' Reporting the MisMatched Parameters
        iProgressCount1 = iProgressCount1 + 1 '2
        iProgressCount2 = 0.001
        iProgressCount2Max = 2
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Reporting Compared Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        strHTML = ""
        strHTML = "<html>"
        
        strHTML = strHTML & "<head>"
        strHTML = strHTML & "<meta http-equiv=Content-Type content=""text/html; charset=windows-1252"">"
        strHTML = strHTML & "<meta name=Generator content=""Microsoft Word 11 (filtered)"">"
        strHTML = strHTML & "<title>Mismatched Parameters</title>"
        strHTML = strHTML & "</head>"
        
        strHTML = strHTML & "<body lang=EN-US>"
        strHTML = strHTML & "<H4>Generated On: " & Format$(Now(), "dd-mmm-yyyy hh:mm:ss am/pm") & "</H4>" & vbCrLf & "<HR>" & vbCrLf
        strHTML = strHTML & "<div class=Section1>"
        strHTML = strHTML & "<p class=MsoNormal><b><u><span style='font-size:16.0pt;font-family:Arial;color:red'>Parameters Comparison</span></u></b></p>"
        strHTML = strHTML & "<p class=MsoNormal><b><u><span style='font-size:14.0pt;font-family:Arial;color:red'><span style='text-decoration:none'>&nbsp;</span></span></u></b></p>"
        strHTML = strHTML & "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>"
        
        strHTML = strHTML & "<tr>"
            strHTML = strHTML & "<td width=57 rowspan=2 valign=top style='width:43.1pt;border:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal align=center style='text-align:center'><b><span  style='font-family:Arial;color:blue'>Sl.No</span></b></p>  </td>"
            strHTML = strHTML & "<td width=300 rowspan=2 valign=top style='width:179.4pt;border:solid windowtext 1.0pt;  border-left:none;padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal align=center style='text-align:center'><b><span  style='font-family:Arial;color:blue'>Parameters</span></b></p>  </td>"
            strHTML = strHTML & "<td width=294 colspan=2 valign=top style='width:220.3pt;border:solid windowtext 1.0pt;  border-left:none;padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal align=center style='text-align:center'><b><span  style='font-family:Arial;color:blue'>Value in</span></b></p>  </td>"
        strHTML = strHTML & "</tr>"
        
        strHTML = strHTML & "<tr>"
            strHTML = strHTML & "<td width=147 valign=top style='width:110.15pt;border-top:none;border-left:  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal align=center style='text-align:center'><b><span  style='font-family:Arial;color:blue'>Seed Assy</span></b></p>  </td>"
            strHTML = strHTML & "<td width=147 valign=top style='width:110.15pt;border-top:none;border-left:  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal align=center style='text-align:center'><b><span  style='font-family:Arial;color:blue'>XML</span></b></p>  </td>"
        strHTML = strHTML & "</tr>"
        
        For lngCol = 1 To UBound(vMisMatchedParams)
            strHTML = strHTML & "<tr>"
            strHTML = strHTML & vbTab & "<td width=57 valign=top style='width:43.1pt;border:solid windowtext 1.0pt;  border-top:none;padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>" & lngCol & ". </span></p>  </td>"
            strHTML = strHTML & vbTab & "<td width=300 valign=top style='width:300.4pt;border-top:none;border-left:  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial;  color:red'>" & Split(vMisMatchedParams(lngCol), "MisMatchedValues==")(0) & "</span></p>  </td>"
            strHTML = strHTML & vbTab & "<td width=147 valign=top style='width:110.15pt;border-top:none;border-left:  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial;  color:red'>" & Split(Split(vMisMatchedParams(lngCol), "MisMatchedValues==")(1), "<Value>")(0) & "</span></p>  </td>"
            strHTML = strHTML & vbTab & "<td width=147 valign=top style='width:110.15pt;border-top:none;border-left:  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;  padding:0in 5.4pt 0in 5.4pt'>  <p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial;  color:red'>" & Split(Split(vMisMatchedParams(lngCol), "MisMatchedValues==")(1), "<Value>")(1) & "</span></p>  </td>"
            strHTML = strHTML & "</tr>"
        Next lngCol
        
        iProgressCount2 = 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Reporting Compared Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        strHTML = strHTML & "</table>"
        strHTML = strHTML & "<p class=MsoNormal><b><u><span style='font-size:14.0pt;font-family:Arial;color:red'><span style='text-decoration:none'>&nbsp;</span></span></u></b></p>"
        strHTML = strHTML & "</div>"
        strHTML = strHTML & "</body>"
        strHTML = strHTML & "</html>"
        
        CreateReport = strHTML
        
        iProgressCount2 = iProgressCount2 + 1 '2
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Reporting Compared Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
        'Report the missing params
        If UBound(vMisMatchedParams) <> 0 Then
            Set ParamFSO = CreateObject("Scripting.FileSystemObject")
            'strDir = mstrWebDirectory
            strFile = Replace(xmlFile, Split(xmlFile, "\")(UBound(Split(xmlFile, "\"))), "") & "ParamsCompare.htm"
            
            If Len(CreateReport) > 0 Then
                If ParamFSO.FileExists(strFile) Then
                    ParamFSO.DeleteFile strFile, True
                End If
                Set oText = ParamFSO.CreateTextFile(strFile, False)
                oText.Write CreateReport
                oText.Close
            End If
            
            If UBound(vMisMatchedParams) > 0 Then
                If MsgBox("Found " & UBound(vMisMatchedParams) & " Parameter(s) mis-matched!     " & vbCrLf & "Do you want to view the Report?     ", vbYesNo + vbInformation, "Manage Configuration / Parameters") = vbYes Then
                    ShellExecute 0, "", strFile, "", "", 0
                End If
            Else
                MsgBox "No mis-matched parameters found!       ", vbInformation, "Manage Configuration / Parameters"
            End If
            
        End If
        
        AppActivate "CATIA V5"
        
        iProgressCount1 = iProgressCount1 + 1 '3
        iProgressCount2 = 0.001
        iProgressCount2Max = 2
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Checking the Missing Parameters..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
        MissingParamFileName = Replace(xmlFile, Split(xmlFile, "\")(UBound(Split(xmlFile, "\"))), "") & "MissingParams.xml"
        Set ParamFSO = CreateObject("Scripting.FileSystemObject")
        ParamFSO.CopyFile xmlFile, MissingParamFileName
        Set xmlDoc = New MSXML2.DOMDocument50
        xmlDoc.Load MissingParamFileName
        Set xmlRoot = xmlDoc.documentElement
        
        'Keep only the selected Configuration
        For i = xmlRoot.childNodes.Length - 1 To 0 Step -1
            If xmlRoot.childNodes.Item(i).attributes.Item(0).nodeTypedValue <> ConfigName Then
                Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & xmlRoot.childNodes.Item(i).attributes.Item(0).nodeTypedValue & Chr(34) & "]")
                xmlRoot.removeChild configNode
            End If
        Next i
        
        'Keep only the selected SeedParts(s)
        For i = xmlRoot.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1
            sSeedPartName = xmlRoot.childNodes.Item(0).childNodes.Item(i).attributes.Item(0).nodeValue
            If Not InStr(1, sXMLSeedPartsList, "_" & Split(sSeedPartName, "_")(UBound(Split(sSeedPartName, "_"))) & ";", vbTextCompare) > 0 Then
                xmlRoot.childNodes.Item(0).removeChild xmlRoot.childNodes.Item(0).childNodes.Item(i)
            End If
        Next i
        
        sSeedPartName = ""
        MissParams = ""
        Set configNode = xmlRoot.childNodes.Item(0)
        
        'Remove all existing nodes, keep only nodes that are not found in config file (only selected configuration)
        iProgressCount2Max = aCopyParams.Count
        iProgressCount2 = 0
        For i = 1 To aCopyParams.Count
            
            iProgressCount2 = iProgressCount2 + 1
            timeEnd = DateTime.Timer
            sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
            sProgBarComment = "Checking the Missing Parameters (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
            
            Set oParameter = aCopyParams.Item(i)
            Set oRelation = oParameter.OptionalRelation
            sSeedPartName = Split(oParameter.Context.Name, "-")(0)
            
            On Error Resume Next
            If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
                If oRelation Is Nothing And InStr(1, sXMLSeedPartsList, "_" & Split(sSeedPartName, "_")(UBound(Split(sSeedPartName, "_"))) & ";", vbTextCompare) > 0 Then
                    
                    Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & sSeedPartName & Chr(34) & "]")
                    
                    If oSeedAssyMigrateConfigFile = False Then
                        sParameterName = oParameter.Name
                    Else
                        sParameterName = Replace(oParameter.Name, Split(sSeedAssyToolVersion, ".")(0), Split(sSeedAssyMigratedVersion, ".")(0), vbTextCompare)
                    End If
                    MissParamparamNode = ""
                    Set MissParamparamNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                    If MissParamparamNode Is Nothing Then   ' missing in xml
                        sParameterName = Replace(oParameter.Name, Split(Split(oParameter.Name, "\")(0), "-")(1), Split(Split(seedNode.childNodes(1).xml, "\")(0), "-")(1), vbTextCompare)
                        Set MissParamparamNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
                        MissParams = MissParams & "<||>" & sParameterName & " value=" & oParameter.ValueAsString
                    Else
                        seedNode.removeChild MissParamparamNode   ' missing in Seed
                    End If
                    
                End If
            End If
            On Error GoTo 0
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        Next
        On Error GoTo 0
        
        iProgressCount1 = iProgressCount1 + 1 '4
        iProgressCount2 = 0.001
        iProgressCount2Max = 2
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Finalizing the Compare..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        For i = xmlRoot.childNodes.Item(0).childNodes.Length - 1 To 0 Step -1 'For each SeedPart node
            
            MissSeedPartName = xmlRoot.childNodes.Item(0).childNodes.Item(i).attributes.Item(0).nodeValue
            
            sMissingParamsInSeedAssy = ""
            For j = xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Length - 1 To 0 Step -1 'For each parameter under SeedPart node
                If Not Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1) Like "*" & MissSeedPartName & "*" Then
                    sMissingParamsInSeedAssy = sMissingParamsInSeedAssy & "<||>" & MissSeedPartName & "\" & Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1)
                Else
                    sMissingParamsInSeedAssy = sMissingParamsInSeedAssy & "<||>" & Split(xmlRoot.childNodes.Item(0).childNodes.Item(i).childNodes.Item(j).xml, "<parameter name=")(1)
                End If
            
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            Next j
            Set seedNode = configNode.selectSingleNode(".//*[@name=" & Chr(34) & MissSeedPartName & Chr(34) & "]")
            configNode.removeChild seedNode
        Next i
        
        sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, """", "")
        sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, "/>", "")
        sMissingParamsInSeedAssy = Replace(sMissingParamsInSeedAssy, "</seedFile>", "")
        
        iProgressCount2 = 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Finalizing the Compare..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        TotalParamNodeCnt = 0
        For ParamNodeCnt = 0 To configNode.childNodes.Length - 1
            TotalParamNodeCnt = TotalParamNodeCnt + configNode.childNodes.Item(ParamNodeCnt).childNodes.Length
        Next ParamNodeCnt
        
        Set MissingParamsinXMLNode = xmlDoc.createElement("ParametersInSEED_NOT_FoundInXML")
        Call configNode.appendChild(MissingParamsinXMLNode)
        Set attNode = xmlDoc.createAttribute("ParametersInSEED_NOT_FoundInXML")
        
        For i = 0 To UBound(Split(MissParams, "<||>")) - 1
            Set paramNode = xmlDoc.createElement("parameter")
            Call MissingParamsinXMLNode.appendChild(paramNode)
            
            Set attNode = xmlDoc.createAttribute("name")
            paramNode.attributes.setNamedItem attNode
            attNode.Value = Split(Split(MissParams, "<||>")(i + 1), " value=")(0)
            
            Set attNode = xmlDoc.createAttribute("value")
            paramNode.attributes.setNamedItem attNode
            attNode.Value = Split(Split(MissParams, "<||>")(i + 1), " value=")(1)
        Next
        
        Set MissingParamsinSeedNode = xmlDoc.createElement("ParametersInXML_NOT_FoundInSEED")
        Call configNode.appendChild(MissingParamsinSeedNode)
        Set attNode = xmlDoc.createAttribute("ParametersInXML_NOT_FoundInSEED")
        
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Finalizing the Compare..."
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
        For i = 0 To UBound(Split(sMissingParamsInSeedAssy, "<||>")) - 1
            Set paramNode = xmlDoc.createElement("parameter")
            Call MissingParamsinSeedNode.appendChild(paramNode)
            
            Set attNode = xmlDoc.createAttribute("name")
            paramNode.attributes.setNamedItem attNode
            attNode.Value = Split(Split(sMissingParamsInSeedAssy, "<||>")(i + 1), " value=")(0)
            
            Set attNode = xmlDoc.createAttribute("value")
            paramNode.attributes.setNamedItem attNode
            attNode.Value = Split(Split(sMissingParamsInSeedAssy, "<||>")(i + 1), " value=")(1)
        Next
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        configNode.attributes.Item(1).nodeValue = UBound(Split(sMissingParamsInSeedAssy, "<||>")) + UBound(Split(MissParams, "<||>"))
        
        'Write Ending Time
        timeEnd = DateTime.Timer
        configNode.attributes.Item(2).nodeValue = Format(timeStart / 86400, "h:mm:ss AMPM")
        configNode.attributes.Item(3).nodeValue = Format(timeEnd / 86400, "h:mm:ss AMPM")
        
        xmlDoc.Save MissingParamFileName
        TotalMissParams = IIf(UBound(Split(sMissingParamsInSeedAssy, "<||>")) < 0, 0, UBound(Split(sMissingParamsInSeedAssy, "<||>"))) + IIf(UBound(Split(MissParams, "<||>")) < 0, 0, UBound(Split(MissParams, "<||>")))
        If TotalMissParams > 0 Then
            If MsgBox("There are few Parameters missing in XML/SEED_ASSY.     " & vbCrLf & "Do you want to view the Report?     ", vbYesNo + vbInformation, "Manage Configuration / Parameters") = vbYes Then
                ShellExecute 0, "", MissingParamFileName, "", "", 0
            End If
        Else
            ParamFSO.DeleteFile MissingParamFileName
            MsgBox "Selected Configuration has been Compared Successfully!     ", vbInformation, "Manage Configuration / Parameters"
        End If
        
    End If
    
    frmProgress.Hide
    
End Sub

Public Sub ReleaseData(ByVal InputConfigFile As String, ByVal InputConfigName As String, ByVal ReleaseFile As String, ByVal ReleaseConfigName As String, ByVal ReleaseConfigDesc As String)
    
    Set xmlDoc = New DOMDocument50
    xmlDoc.Load (InputConfigFile)
    Set xmlRoot = xmlDoc.documentElement
    
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & InputConfigName & Chr(34) & "]")
    Set seedNode = configNode.cloneNode(True)
    seedNode.attributes(0).nodeValue = ReleaseConfigName
    
    Set xmlDoc2 = New DOMDocument50
    Set xmlRoot2 = xmlDoc2.createElement("configurationData")
    Set xmlDoc2.documentElement = xmlRoot2
    xmlRoot2.appendChild seedNode
    
    Set attNode = xmlDoc2.createAttribute("version")
    xmlRoot2.attributes.setNamedItem attNode
    attNode.Value = ReleaseConfigName
    
    Set ParamFSO = CreateObject("Scripting.FileSystemObject")
    If ParamFSO.FileExists(ReleaseFile) = True Then
        If MsgBox("File Already exist! Do you want to overwrite the existing file with the new version?", vbYesNo + vbInformation) = vbNo Then
            frmManageData.txtNewConfigFileName.Text = ""
            frmManageData.txtNewConfigFileName.SetFocus
            Exit Sub
        End If
    End If
    
    seedNode.childNodes.Item(0).childNodes.Item(1).Text = ReleaseConfigName
    seedNode.childNodes.Item(0).childNodes.Item(2).Text = ReleaseConfigDesc
    
    xmlDoc2.Save (ReleaseFile)
    
    MsgBox "Selected Configuration has been Released Successfully!     ", vbInformation, "Manage Configuration / Parameters"
    
End Sub

Public Sub LoadPaste(ByVal xmlFile As String, ByVal ConfigName As String, ByVal oSelectedTarget As AnyObject, ByVal oSelectedTargetPart As Part)
    
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    timeStart = DateTime.Timer
    Call frmProgress.progressBarInitialize("Manage Configuration / Parameters")
    
    iProgressCount1Max = 1
    iProgressCount1 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Searching DT Parameters to import..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    Set aPasteParams = New Collection
    
    'Target
    Set oParameters = oSelectedTargetPart.Parameters
    Set oParameters = oParameters.SubList(oSelectedTarget, True)
    
    For Each oParameter In oParameters
        If oParameter.Name Like "*_DT" Then
            If oParameter.OptionalRelation Is Nothing Then aPasteParams.Add oParameter, oParameter.Name
        End If
    Next
    
    'All the parameters under the Assemble objects
    If TypeName(oSelectedTarget) = "Body" Then
        Call ScanSubBodies(oSelectedTarget, aPasteParams)
    End If

    If aPasteParams.Count = 0 Then MsgBox "No DT Parameter found in Target!     ", vbExclamation, "Warning": Exit Sub
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    'load xml file and get the Doc Element
    Set xmlDoc = New MSXML2.DOMDocument50
    xmlDoc.Load (xmlFile)
    Set xmlRoot = xmlDoc.documentElement
    Set configNode = xmlRoot.selectSingleNode("//*[@id=" & Chr(34) & ConfigName & Chr(34) & "]")
    
    NodeCnt = 0
    sLog = ""
    iProgressCount1Max = aPasteParams.Count
    iProgressCount1 = 0
    For Each oParameter In aPasteParams
        
        iProgressCount1 = iProgressCount1 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Pasting Parameters of - " & ConfigName & " (" & CStr(iProgressCount1) & "/" & CStr(iProgressCount1Max) & ")"
        Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
        
        Set oRelation = oParameter.OptionalRelation
        
        Set paramNode = configNode.childNodes(0).childNodes.Item(NodeCnt)
        
        On Error Resume Next
        If paramNode Is Nothing Then
                sParameterName = Replace(oParameter.Name, Split(Split(oParameter.Name, "\")(0), "-")(1), Split(Split(seedNode.childNodes(1).xml, "\")(0), "-")(1), vbTextCompare)
                Set paramNode = seedNode.selectSingleNode(".//*[@name=" & Chr(34) & sParameterName & Chr(34) & "]")
            On Error GoTo 0
        End If
        
        If Not paramNode Is Nothing Then
            Set oRelation = oParameter.OptionalRelation
            If oRelation Is Nothing Then
                If Not (oParameter.Name Like "*KnowledgeTemplates*" Or oParameter.Name Like "*PowerCopy*" Or oParameter.Name Like "*External Parameters*") Then
                    If Split(oParameter.Name, "\")(UBound(Split(oParameter.Name, "\"))) = Split(paramNode.attributes.Item(0).nodeValue, "\")(UBound(Split(paramNode.attributes.Item(0).nodeValue, "\"))) Then
                        NodeCnt = NodeCnt + 1
                        Err.Clear
                        
                        If paramNode.attributes.Item(1).Text = "" Then
                            oParameter.ValuateFromString (paramNode.nodeTypedValue)
                        Else
                            oParameter.ValuateFromString (paramNode.attributes.Item(1).Text)
                        End If
                        
                        If Err.Number <> 0 Then
                            oParameter.Value = paramNode.attributes.Item(1).Text
                        End If
                        
                    End If
                End If
            Else
                sLog = sLog & oParameter.Name & vbCrLf
            End If
        End If
        
        On Error GoTo 0
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_LoadPasteData.txt"
        sLog = "The following Read Only parameters were not modified:" & vbCrLf & sLog
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If
    
    MsgBox "Selected Configuration(s) has/have been Imported Successfully!     ", vbInformation, "Manage Configuration / Parameters"
    frmProgress.Hide
    
End Sub

Public Sub CopyPasteParam(ByVal oSelectedSource As AnyObject, ByVal oSelectedTarget As AnyObject, ByVal oSelectedSourcePart As Part, ByVal oSelectedTargetPart As Part)
    
    Dim iCopiedParams As Long
    Dim iLockedParams As Long
    Dim iI As Integer
    Dim oShape As Shape
    Dim oBody As Body
    
    Dim oSourceParameter As Parameter
    Dim sSourcePartName As String
    Dim oTargetParameter As Parameter
    Dim sTargetPartName As String
    
    sLogComplete = ""
    timeStart = DateTime.Timer
    Call frmProgress.progressBarInitialize("Manage Configuration / Parameters")
    
    iProgressCount1Max = 1
    iProgressCount1 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Searching DT Parameters in Source..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    Set aCopyParams = New Collection
    
    'Source
    'All the parameters under the oSelectedSource except the ones under the Assemble objects
    Set oParameters = oSelectedSourcePart.Parameters
    Set oParameters = oParameters.SubList(oSelectedSource, True)
    
    For Each oParameter In oParameters
        If oParameter.Name Like "*_DT" Then
            If oParameter.OptionalRelation Is Nothing Then
                aCopyParams.Add oParameter, oParameter.Name
            End If
        End If
    Next
    
    'All the parameters under the Assemble objects
    If TypeName(oSelectedSource) = "Body" Then
        Call ScanSubBodies(oSelectedSource, aCopyParams)
    End If
    
    If aCopyParams.Count = 0 Then MsgBox "No DT Parameter found in Source!     ", vbExclamation, "Warning": Exit Sub
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    iProgressCount1Max = aCopyParams.Count + 1
    iProgressCount1 = 1
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Searching DT Parameters in Target..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    Set aPasteParams = New Collection
    
    'Target
    'All the parameters under the oSelectedTarget except the ones under the Assemble objects
    Set oParameters = oSelectedTargetPart.Parameters
    Set oParameters = oParameters.SubList(oSelectedTarget, True)
    
    For Each oParameter In oParameters
        If oParameter.Name Like "*_DT" Then
            If oParameter.OptionalRelation Is Nothing Then
                aPasteParams.Add oParameter, oParameter.Name
            End If
        End If
    Next
    
    'All the parameters under the Assemble objects
    If TypeName(oSelectedTarget) = "Body" Then
        Call ScanSubBodies(oSelectedTarget, aPasteParams)
    End If
    
    If aPasteParams.Count = 0 Then MsgBox "No DT Parameter found in Source!     ", vbExclamation, "Warning": Exit Sub
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    sSourcePartName = aCopyParams.Item(1).Context.Name
    sTargetPartName = aPasteParams.Item(1).Context.Name
    
    On Error Resume Next
    iCopiedParams = 0
    iLockedParams = 0
    sLog = ""
    For Each oSourceParameter In aCopyParams
        
        iProgressCount1 = iProgressCount1 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Pasting Parameters (" & CStr(iProgressCount1) & "/" & CStr(iProgressCount1Max) & ")"
        Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
        
        Set oTargetParameter = Nothing
        For Each oParameter In aPasteParams
            If oParameter.Name = FindParameterPath(oSourceParameter.Name, oParameter.Name) Then Set oTargetParameter = oParameter: Exit For
        Next
        
        Err.Clear
        If Not oTargetParameter Is Nothing Then
            If oTargetParameter.Value <> oSourceParameter.Value Then
                oTargetParameter.Value = oSourceParameter.Value
                If Err = 0 Then iCopiedParams = iCopiedParams + 1
                If Err <> 0 Then iLockedParams = iLockedParams + 1: sLog = sLog & oTargetParameter.Name & vbCrLf
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    Next
    On Error GoTo 0
    
endsub:
    If sLog <> "" Then
        sLogFileName = sTempDirectory & "LogFile_LoadPasteData.txt"
        sLog = "The following Read Only parameters were not modified:" & vbCrLf & sLog
        Call WriteTextFile(sLogFileName, sLog)
        If OpenWithNotepad <> 0 Then
            MsgBox "Problem opening Notepad. Refer to log file: " & sLogFileName, vbExclamation
        End If
    End If

    If Not iCopiedParams = 0 And Not iLockedParams = 0 Then
        MsgBox CStr(iCopiedParams) & " parameters copied successfully!     " & vbCrLf & _
              CStr(iLockedParams) & " parameters not modified (locked)!     ", vbInformation, "Manage Configuration / Parameters"
    ElseIf Not iCopiedParams = 0 And iLockedParams = 0 Then
        MsgBox CStr(iCopiedParams) & " parameters copied successfully!     ", vbInformation, "Manage Configuration / Parameters"
    ElseIf iCopiedParams = 0 And Not iLockedParams = 0 Then
        MsgBox "No parameter copied! Target parameters are Read Only (" & CStr(iLockedParams) & ")!     ", vbExclamation, "Manage Configuration / Parameters"
    End If
    
End Sub

Public Function IsValidToolVersion(xmlRootValid As IXMLDOMElement) As Boolean
    
    'Find the Tool Version Parameter in CATIA
    Dim sXMLToolVersion As String
    
    If (xmlRootValid Is Nothing) Then
        IsValidToolVersion = False
        Exit Function
    Else
        sXMLToolVersion = xmlRootValid.attributes.Item(0).Text
    End If
    
    'Compare it to the sXMLToolVersion in Import File
    If sXMLToolVersion <> sSeedAssyToolVersion Then
        IsValidToolVersion = False
        sSeedAssyMigratedVersion = sXMLToolVersion
    Else
        IsValidToolVersion = True
    End If
    
End Function

Public Function IsValidMonumentID(xmlDocValid As DOMDocument50) As Boolean
    
    IsValidMonumentID = True
    
    'Check if data is related to current monument
    Dim NodeList As IXMLDOMNodeList
    Set NodeList = xmlDocValid.getElementsByTagName("seedFile")
    SeedMonumentID = Mid(oSeedAssyConfigName.Context.Name, 1, InStr(1, oSeedAssyConfigName.Context.Name, "_PART_") + 4)
    For i = 0 To NodeList.Length - 1
        XMLFileMonumentID = Mid(NodeList.Item(i).attributes.Item(0).Text, 1, InStr(1, NodeList.Item(i).attributes.Item(0).Text, "_PART_") + 4)
        If XMLFileMonumentID <> SeedMonumentID Then
            unrelatedXMLFileMonumentID = XMLFileMonumentID
            IsValidMonumentID = False
            Exit For
        End If
    Next
    
End Function

Private Sub ScanSubBodies(oBody As Body, ByRef oColl As Collection)

Dim oChildBody As Body

For Each oShape In oBody.Shapes
    If TypeName(oShape) = "Assemble" Then
        Set oChildBody = oShape.Body
        Set oParameters = oParameters.SubList(oChildBody, True)
        
        For Each oParameter In oParameters
            If oParameter.Name Like "*_DT" Then
                If oParameter.OptionalRelation Is Nothing Then
                    oColl.Add oParameter, oParameter.Name
                End If
            End If
        Next
    
        Call ScanSubBodies(oChildBody, oColl)
        
    End If
Next

End Sub

Private Function FindParameterPath(ByVal sSourceParameterName As String, ByVal sTargetParameterName As String) As String
    
    FindParameterPath = ""
    
    On Error Resume Next
    If Split(sSourceParameterName, "\")(1) = "" Then FindParameterPath = sSourceParameterName: Exit Function
    If Split(sTargetParameterName, "\")(1) = "" Then FindParameterPath = sTargetParameterName: Exit Function
    
    For i = LBound(Split(sSourceParameterName, "\")) To UBound(Split(sSourceParameterName, "\"))
        If Not Split(sSourceParameterName, "\")(i) Like "*_M##_*" Then FindParameterPath = Split(sSourceParameterName, Split(sSourceParameterName, "\")(i - 1))(1): Exit For
    Next
    FindParameterPath = Mid(FindParameterPath, 2)
    
    For i = UBound(Split(sTargetParameterName, "\")) To LBound(Split(sTargetParameterName, "\")) Step -1
        If Split(sTargetParameterName, "\")(i) Like "*_M##_*" Then FindParameterPath = Split(sTargetParameterName, Split(sTargetParameterName, "\")(i + 1))(0) & FindParameterPath: Exit For
    Next
    On Error GoTo 0
    
End Function
