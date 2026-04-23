Attribute VB_Name = "CollectorSP"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private bPartFound As Boolean
Private oSPCollNumber  As Long
Private sSPCollName As String
Private dNHAMatrix(11) As Variant
Private d3DCollMatrix(11) As Variant
Private dSPCollMatrix(11) As Variant
Private sSPCollDescription As String
Private sActiveBoolParamList As Collection
Private sAllBoolParamList As Collection
Private bStdPartBoolean As BoolParam
Private bInstActivity As Boolean

'************************************************************************************************************
'* Please make sure any modification is in sync between Prod and CD toolbars                                *
'* Version Log:                                                                                             *
'*                                                                                                          *
'* Version 1.0 August 31st 2015                                                                             *
'* - Initial Release                                                                                        *
'************************************************************************************************************

Public Sub launchCollSP()
    
'-------------------------------------------------------
'     FIND & SELECT SUBSET
'-------------------------------------------------------
    
    If frmSPCollector.cbSubSet.Value = True Then
        'Scan Subsets in part
        Call ScanSubsets(oSeedPartCollection)
        
        'Ask user to select a Subset
        Call frmSubsetsSelection.FormInitialize(frmSPCollector.Left + 15, frm3DCollector.Top + frmSPCollector.cmdOK.Top - frmSubsetsSelection.Height)
        If sCurrentSSParameterName = "" Then
            frmSPCollector.Show
            Exit Sub
        End If
    End If
    
'-------------------------------------------------------
'     PROGRESS BAR SETTINGS
'-------------------------------------------------------
    timeStart = DateTime.Timer
    iProgressCount1 = 0
    iProgressCount1Max = 3
    Call frmProgress.progressBarInitialize("SP Collector")
    
'-------------------------------------------------------
'     CREATE NEW COLLECTOR
'-------------------------------------------------------
    If frmSPCollector.optCreateNew.Value = True Then
        Call CreateNewCollector
    End If
    
'-------------------------------------------------------
'     APPLY BA STANDARD
'-------------------------------------------------------
    Call oTemplateData.SetProprietaryNotice(oSPCollInstance)
    
'-------------------------------------------------------
'     COPY STANDARD PARTS
'-------------------------------------------------------
    iProgressCount1 = iProgressCount1 + 1
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    Call FindBoolParams
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
    iProgressCount1 = iProgressCount1 + 1
    If frmSPCollector.optModifyExisting.Value Then Call CleanSPCollector
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
    iProgressCount1 = iProgressCount1 + 1
    Call CopySTDParts
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then GoTo endsub
    
'-------------------------------------------------------
'     END
'-------------------------------------------------------
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, 1, 1, frmProgress.lblMessage.Caption, 1, 1, sTimer)
    
    oSelectionFromSPColl.Clear
    
    On Error Resume Next
    CATIA.StatusBar = "Successfully Done! "
    On Error GoTo 0
    
    MsgBox "SP Collector Successfully Done!    ", vbInformation, "SP Collector"
  
endsub:
'-------------------------------------------------------
'     POSITION SP COLLECTOR
'-------------------------------------------------------
    If frmSPCollector.cbSnap.Value And Not oNHAAssy Is Nothing Then
        Call SnapPartToPart(oSPCollInstance, oNHAAssy, oSeedLingInstance)
    End If
    
End Sub

Private Sub CreateNewCollector()
    
    Dim iColSPIndex As Integer, sColSPIndex As String
    
    On Error Resume Next
    
    'List all NHA products
    sNHAAssyProducts = ""
    For Each oProduct In oNHAAssy.Products
        If oProduct.Name Like "*-*SP*" And oProduct.Products.Count > 0 Then
            oProduct.ApplyWorkMode DEFAULT_MODE
            sNHAAssyProducts = sNHAAssyProducts & oProduct.PartNumber & ";"
        End If
    Next
    
    'Create new part
    oNHAAssy.Products.AddNewComponent "Product", ""
    Set oSPCollInstance = oNHAAssy.ReferenceProduct.Products.Item(oNHAAssy.Products.Count)
    
    'Rename new part and instance
    iColSPIndex = 0
    Do
        'Find Collector index (no duplication)
        Do
            iColSPIndex = iColSPIndex + 1
            sColSPIndex = Format(CStr(iColSPIndex), "0#")
        Loop Until Not sNHAAssyProducts Like "*" & oNHAAssy.PartNumber & "SP" & sColSPIndex & "*" And _
                   Not sNHAAssyProducts Like "*" & oNHAAssy.PartNumber & "-SP" & sColSPIndex & "*"
        
        If InStr(1, oNHAAssy.PartNumber, "-", vbTextCompare) = 0 Then
            oSPCollInstance.ReferenceProduct.PartNumber = oNHAAssy.PartNumber & "-SP" & sColSPIndex
            oSPCollInstance.Name = oNHAAssy.PartNumber & "-SP" & sColSPIndex & ".1"
        Else
            oSPCollInstance.ReferenceProduct.PartNumber = oNHAAssy.PartNumber & "SP" & sColSPIndex
            oSPCollInstance.Name = oNHAAssy.PartNumber & "SP" & sColSPIndex & ".1"
        End If
        
    Loop Until oSPCollInstance.ReferenceProduct.PartNumber Like "*-*SP*" And oSPCollInstance.Name Like "*-*SP*"
    
    '-------------------------------------------------------
    '     APPLY BA STANDARD
    '-------------------------------------------------------
    Call oTemplateData.CreateProperties(oSPCollInstance, , "KBE", True)
    oSPCollInstance.ReferenceProduct.UserRefProperties.GetItem("Title").Value = oSeedAssy.UserRefProperties.Item("Title").ValueAsString
    oSPCollInstance.Definition = "SP_" & Split(oSeedAssy.PartNumber, "_")(1)
    
    On Error GoTo 0
    
End Sub

Private Sub FindBoolParams()
        
'***********FIND ALL BOOLEAN PARAMS
    Dim bInvertSel As Boolean
    
    iProgressCount2 = 0.001
    iProgressCount2Max = 1
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Searching for boolean parameters..."
    Call frmProgress.progressBarRepaint("Initializing SP Collector...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
                    
    oSelectionFromSeedPart.Clear
    For Each iObject In oSeedPartCollection
        oSelectionFromSeedPart.Add iObject
    Next
    Set oSelectionFromSeedPart = KBESearch("(CATKnowledgeSearch.InternalParameter.Name_CAP=SP_*_Boolean + CATKnowledgeSearch.InternalParameter.Name_CAP=SA_*_Boolean),sel", , oSelectionFromSeedPart)
    
    On Error Resume Next
    Set sAllBoolParamList = New Collection
    Set sActiveBoolParamList = New Collection
    iProgressCount2 = 0
    iProgressCount2Max = oSelectionFromSeedPart.Count2
    For i = 1 To oSelectionFromSeedPart.Count2
    
        iProgressCount2 = iProgressCount2 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        sParameterName = Split(oSelectionFromSeedPart.Item(i).Value.Name, "\")(UBound(Split(oSelectionFromSeedPart.Item(i).Value.Name, "\")))
        sAllBoolParamList.Add oSelectionFromSeedPart.Item(i).Value, sParameterName
        
        If oSelectionFromSeedPart.Item(i).Value.Value Then
            
            'If Subset is used, verify if the current boolparam has to be collected
            '*****Update18-March-2015 If InvertSelectin is checked then collect everything except bodies in Subset
            If frmSPCollector.cbSubSet.Value = True Then
                bCopySelection = False
                For Each oSSParameterValue In SSParameters.Item(sCurrentSSParameterName).SubsetValues
                    If oSSParameterValue = sParameterName Then
                        bCopySelection = True
                        Exit For
                    End If
                Next
                If Not frmSubsetsSelection.cbInvertSelection.Value Then
                    If bCopySelection Then
                        sActiveBoolParamList.Add oSelectionFromSeedPart.Item(i).Value, sParameterName
                    End If
                Else
                    If Not bCopySelection Then
                     sActiveBoolParamList.Add oSelectionFromSeedPart.Item(i).Value, sParameterName
                    End If
                End If
            Else
                sActiveBoolParamList.Add oSelectionFromSeedPart.Item(i).Value, sParameterName
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
    Next i
    
    oSelectionFromSeedPart.Clear
    On Error GoTo 0
    
End Sub

Private Sub CleanSPCollector()
'***********ERASE ALL STD PARTS THAT ARE NOT FOUND IN SEED ASSY, POSITION ALL OTHER EXISTING STD PARTS
    
    On Error Resume Next
    If frmSPCollector.optCreateNew.Value = False Then
        
        iProgressCount2 = 0
        iProgressCount2Max = oSPCollInstance.Products.Count
        For Each iObject In oSPCollInstance.Products
            
            iProgressCount2 = iProgressCount2 + 1
            timeEnd = DateTime.Timer
            sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
            sProgBarComment = "Scanning SP Collector..."
            Call frmProgress.progressBarRepaint(frmProgress.lblMessageMain.Caption, iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)

            bPartFound = False
            For Each oProduct In oSeedAssy.Products
                If iObject.DescriptionInst = oProduct.DescriptionInst And (iObject.DescriptionInst <> "" Or iObject.Name Like "*MONUMENT*ORIGIN*") Then
                    
                    bInstActivity = GetInstanceActivity(oProduct)
                    
                    If Not bInstActivity Then
                        Set bStdPartBoolean = Nothing
                        Set bStdPartBoolean = sAllBoolParamList.Item(iObject.DescriptionInst)
                        
                        If bStdPartBoolean.Value Then 'Continues if bStdPartBoolean is nothing
                            bPartFound = True
                            Set oProdPosition = oProduct.Position
                            oProdPosition.GetComponents dSPCollMatrix
                            Set oProdPosition = iObject.Position
                            oProdPosition.SetComponents dSPCollMatrix
                            Exit For
                        End If
                    End If
                End If
            Next
            
            If Not bPartFound Then
                oSPCollInstance.Products.Remove iObject
            End If
            
            If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
            
        Next
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub CopySTDParts()
    
'***************COPY PARTS********************
    iProgressCount2 = 0
    iProgressCount2Max = oSeedAssy.Products.Count
    oSPCollNumber = oSPCollInstance.Products.Count
    
    On Error Resume Next
    For Each iObject In oSeedAssy.Products
        
        iProgressCount2 = iProgressCount2 + 1
        
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Scanning Seed Assy (" & CStr(iProgressCount2) & "/" & CStr(iProgressCount2Max) & ")"
        Call frmProgress.progressBarRepaint("Importing Standard Parts...", iProgressCount1Max, iProgressCount1, sProgBarComment, iProgressCount2Max, iProgressCount2, sTimer)
        
        'Find parts to be copied in Seed Assy
        bPartFound = False
        If Not iObject.Name Like "KBE_*" Then
            Set bStdPartBoolean = Nothing
            Set bStdPartBoolean = sActiveBoolParamList.Item(iObject.DescriptionInst)
            If (iObject.DescriptionInst <> "" And Not bStdPartBoolean Is Nothing) Or iObject.Name Like "*MONUMENT*ORIGIN*" Then
                
                bInstActivity = GetInstanceActivity(iObject)

                If bInstActivity Then
                    If iObject.PartNumber = "" Then iObject.ApplyWorkMode DEFAULT_MODE
                    bPartFound = True
                End If
            End If
        End If
        
        'Verify if part already exist in SP Collector, not to copy a STD part twice
        If bPartFound And frmSPCollector.optCreateNew.Value = False Then
            For i = 1 To oSPCollNumber
                If oSPCollInstance.Products.Item(i).DescriptionInst = iObject.DescriptionInst Then
                    bPartFound = False
                    Exit For
                End If
            Next
        End If
        
        If bPartFound Then
            Set oProdPosition = iObject.Position
            oProdPosition.GetComponents dSPCollMatrix
            
            If iObject.DescriptionInst Like "SA_*" Or iObject.DescriptionInst Like "SP_*" Or iObject.Name Like "*MONUMENT*ORIGIN*" Then
                
                oSPCollInstance.Products.AddExternalComponent iObject.ReferenceProduct.Parent
                
                Set oProdPosition = oSPCollInstance.Products.Item(oSPCollInstance.Products.Count).Position
                oProdPosition.SetComponents dSPCollMatrix
                
                oSPCollInstance.ReferenceProduct.Products.Item(oSPCollInstance.Products.Count).Name = iObject.Name
                oSPCollInstance.ReferenceProduct.Products.Item(oSPCollInstance.Products.Count).DescriptionInst = iObject.DescriptionInst
                
            End If
            
            'If copying MONUMENT ORIGIN PART, reorder it to be first in the SP Collector tree
            If iObject.Name Like "*MONUMENT*ORIGIN*" Then
                j = 1
                oSPCollNumber = oSPCollNumber + 1
                For i = 1 To oSPCollInstance.Products.Count
                    
                    sSPCollName = oSPCollInstance.Products.Item(j).Name
                    sSPCollDescription = oSPCollInstance.Products.Item(j).DescriptionInst
                    
                    If sSPCollName Like "*MONUMENT*ORIGIN*" Then
                        j = j + 1
                    Else
                        Set oProdPosition = oSPCollInstance.Products.Item(j).Position
                        oProdPosition.GetComponents dSPCollMatrix
                        
                        oSPCollInstance.Products.AddComponent oSPCollInstance.Products.Item(j).ReferenceProduct
                        oSPCollInstance.Products.Remove (j)
                        
                        Set oProdPosition = oSPCollInstance.Products.Item(oSPCollInstance.Products.Count).Position
                        oProdPosition.SetComponents dSPCollMatrix

                        oSPCollInstance.ReferenceProduct.Products.Item(oSPCollInstance.Products.Count).Name = sSPCollName
                        oSPCollInstance.ReferenceProduct.Products.Item(oSPCollInstance.Products.Count).DescriptionInst = sSPCollDescription
                    End If
                Next i
            End If
            
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
    Next
    On Error GoTo 0
    
'*******************************************
'*************** OBSOLETE??? ***************
'*******************************************
'    'Loop through SP Collector to Find Configuration Params
'    'If found check if a relation exists that points back to the SEED ASSY, if so delete it and create a new one using the existing value
'    Dim oSPCollPart as part
'    For Each iObject In oSPCollInstance.Products
'
'        Set oSPCollPart = Nothing
'        Set oSPCollPart = iObject.ReferenceProduct.Parent.Part
'        If Not oSPCollPart Is Nothing Then
'            Set oParameter = Nothing
'            Set oParameter = oSPCollPart.FindObjectByName("Configuration")
'            If Not (oParameter Is Nothing) Then
'                If Not (oParameter.ValueAsString Like ("*C1")) Then
'                    Set oRelation = oParameter.OptionalRelation
'                    If Not (oRelation Is Nothing) Then
'                        strConfig = "L" & Mid(oParameter.ValueAsString, 2, 2)
'                        Set oRelation = oSPCollInstance.Relations.CreateFormula(iObject.Name & "Configuration", "", oParameter, Chr(34) & strConfig & Chr(34))
'                    End If
'                End If
'            End If
'            oSPCollPart.Update
'        Else
'            Set oParameter = iObject.Parameters.Item("Configuration")
'            If Not (oParameter Is Nothing) Then
'                If oParameter.ValueAsString Like ("?C#") Then
'                    If Not (oParameter.ValueAsString Like ("*C1")) Then
'                        Set oRelation = oParameter.OptionalRelation
'                        If Not (oRelation Is Nothing) Then
'                            strConfig = "L" & Mid(oParameter.ValueAsString, 2, 2)
'                            Set oRelation = oSPCollInstance.Relations.CreateFormula(iObject.Name & "Configuration", "", oParameter, Chr(34) & strConfig & Chr(34))
'                        End If
'                    End If
'                Else
'                    Set oRelation = oParameter.OptionalRelation
'                    If Not (oRelation Is Nothing) Then
'                        Set oRelation = oSPCollInstance.Relations.CreateFormula(iObject.Name & "Configuration", "", oParameter, Chr(34) & oParameter.ValueAsString & Chr(34))
'                    End If
'                End If
'            End If
'            iObject.Update
'        End If
'    Next
'*******************************************
'*******************************************
'*******************************************

End Sub




