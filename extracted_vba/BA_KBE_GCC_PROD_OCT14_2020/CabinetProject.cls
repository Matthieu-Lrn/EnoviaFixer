Attribute VB_Name = "CabinetProject"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private j As Long
Private bInstActivity As Boolean
Public Sub launchCabinetProject()
    
    Dim sProjectNumber As String
    Dim sMonumentNumber As String
    Dim sMonumentTypeNumber As String
    Dim sSceneNumber As String
    Dim iProjectIndex As Long
    
    sProjectNumber = ""
    sMonumentNumber = ""
    sMonumentTypeNumber = ""
    sSceneNumber = ""
    
    timeStart = DateTime.Timer
    Call frmProgress.progressBarInitialize("Cabinet Project")
    
    iProgressCount1Max = 1
    iProgressCount1 = 0.001
    timeEnd = DateTime.Timer
    sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
    sProgBarComment = "Initializing Project..."
    Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
    
    On Error Resume Next
    sProjectNumber = Replace(Split(oNHAAssy.ReferenceProduct.Name, "-")(0), "GXRS", "")
    sMonumentNumber = Mid(Split(oNHAAssy.ReferenceProduct.Name, "-")(1), 1, 2)
    sMonumentTypeNumber = Mid(Split(oNHAAssy.ReferenceProduct.Name, "-")(1), 3, 2)
    sSceneNumber = Split(oNHAAssy.ReferenceProduct.Name, "-")(3)
    
    If sSceneNumber = "" Then
        sSceneNumber = sProjectNumber
        sProjectNumber = ""
        sMonumentNumber = ""
        sMonumentTypeNumber = ""
    End If
     
    'Find Project number
    iProjectIndex = 1
    For Each iObject In oNHAAssy.Products
        If iObject.PartNumber Like oNHAAssy.PartNumber & "CAB" & Format(iProjectIndex, "0#") Then
            iProjectIndex = iProjectIndex + 1
        End If
    Next

    oNHAAssy.ReferenceProduct.Products.AddNewComponent "Product", oNHAAssy.ReferenceProduct.Name & "CAB" & Format(iProjectIndex, "0#")
    Set oProduct = oNHAAssy.ReferenceProduct.Products.Item(oNHAAssy.Products.Count)
    
    'Apply BA Standard
    Call oTemplateData.SetProprietaryNotice(oProduct)
    Call oTemplateData.CreateProperties(oProduct, , "KBE", True)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    
    sSeedAssyName = Split(oSeedAssy.PartNumber, "-")(0) & "-" & Mid(Split(oSeedAssy.PartNumber, "-")(1), 1, 3)
    oProduct.ReferenceProduct.Products.AddNewComponent "Product", sSeedAssyName
    Set oSeedLingInstance = oProduct.ReferenceProduct.Products.Item(oProduct.Products.Count)
    oSeedLingInstance.ReferenceProduct.PartNumber = CStr(sSeedAssyName & sProjectNumber & sMonumentNumber & "-" & sMonumentTypeNumber & sSceneNumber & "CAB" & Format(iProjectIndex, "0#"))
    
    'Apply BA Standard
    Call oTemplateData.SetProprietaryNotice(oSeedLingInstance)
    Call oTemplateData.CreateProperties(oSeedLingInstance, , "KBE", True)
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
        
    'Verify if the Assembly Design workbench is Active
    If Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    If CATIA.GetWorkbenchId = "KnowledgeAdvisor" Then CATIA.StartWorkbench "KnowledgeAdvisor"  '****This is a "bug" that we use: when starting KWA when being in KWA, it switches to the previous workbench
    If CATIA.GetWorkbenchId <> "Assembly" And CATIA.GetWorkbenchId <> "PrsConfiguration" Then
        CATIA.StartWorkbench ("Assembly")
    End If
    If Not CATIA.ActiveWindow Is sTargetWindow Then sTargetWindow.Activate
    
    'Allows only  "ABCDEFGH JK MN P R TUVWXY "
    i = 0
    iProgressCount1 = 0
    iProgressCount1Max = oSeedAssy.ReferenceProduct.Products.Count
    Set oConstraints = oSeedLingInstance.ReferenceProduct.Connections("CATIAConstraints")
    
    For Each iObject In oSeedAssy.ReferenceProduct.Products
        
        iProgressCount1 = iProgressCount1 + 1
        timeEnd = DateTime.Timer
        sTimer = "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss")
        sProgBarComment = "Collecting Parts from Seed Assy  (" & CStr(iProgressCount1) & "/" & CStr(iProgressCount1Max) & ")"
        Call frmProgress.progressBarRepaint(sProgBarComment, iProgressCount1Max, iProgressCount1, , , , sTimer)
        
        bInstActivity = GetInstanceActivity(iObject, True)
        
        If Not bInstActivity Then
            
            If iObject.Name Like "KBE_*" Then
                If iObject.Name Like "KBE_*_SEED_PART_*" Then
                    
                    i = i + 1
                    sSeedLingLetter = ""
                    If i Mod 20 = 0 Then
                        sSeedLingLetter = (Choose(Fix(i / 20) - 1, "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "M", "N", "P", "R", "T", "U", "V", "W", "X", "Y"))
                        sSeedLingLetter = sSeedLingLetter & (Choose(20, "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "M", "N", "P", "R", "T", "U", "V", "W", "X", "Y"))
                    Else
                        If i > 20 Then sSeedLingLetter = (Choose(Fix(i / 20), "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "M", "N", "P", "R", "T", "U", "V", "W", "X", "Y"))
                        sSeedLingLetter = sSeedLingLetter & (Choose(i Mod 20, "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "M", "N", "P", "R", "T", "U", "V", "W", "X", "Y"))
                    End If
                    
                    oSelectionFromSeedPart.Clear
                    oSelectionFromSeedPart.Add iObject
                    oSelectionFromSeedPart.Copy
    
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oSeedLingInstance
                    oSelectionFromSeedLing.PasteSpecial "CATSpecBreakLink"

                    sPartNumber = ""
                    sPartNumber = Split(iObject.PartNumber, "-")(0) & "-" & Mid(Split(iObject.PartNumber, "-")(1), 1, 3)
                    sSeedPartLetter = Split(Mid(sPartNumber, InStrRev(sPartNumber, "_") + 1), "-")(0)
                    sPartNumber = Replace(sPartNumber, "SEED_PART_" & sSeedPartLetter, "SEED_PART_" & sSeedLingLetter)
                    oSeedLingInstance.Products.Item(oSeedLingInstance.Products.Count).PartNumber = "KBE_" & Split(sPartNumber, "KBE_")(1) & sProjectNumber & sMonumentNumber & "-" & sMonumentTypeNumber & sSceneNumber & "CAB" & Format(iProjectIndex, "0#")
                    
                    'Create Fix constraint and hide it
                    Set oReference = oSeedLingInstance.ReferenceProduct.CreateReferenceFromName(oSeedLingInstance.PartNumber & "/" & oSeedLingInstance.Products.Item(oSeedLingInstance.Products.Count).Name & "/!" & oSeedLingInstance.PartNumber & "/" & oSeedLingInstance.Products.Item(oSeedLingInstance.Products.Count).Name & "/")
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oConstraints.AddMonoEltCst(catCstTypeReference, oReference)
                    oSelectionFromSeedLing.VisProperties.SetShow (catVisPropertyNoShowAttr)
                
                ElseIf iObject.Name Like "KBE_*_MOLDING*" Then
                    'Create an instance
                    iObject.ApplyWorkMode DESIGN_MODE
                    oSeedLingInstance.Products.AddExternalComponent (iObject.ReferenceProduct.Parent)
                End If
            Else
                'We only copy the Std Part instance ending with ".0", and rename them with a .0 at the end
                If CDbl(Split(iObject.Name, ".")(UBound(Split(iObject.Name, ".")))) = 0 Then
                    oSelectionFromSeedLing.Clear
                    oSelectionFromSeedLing.Add oSeedLingInstance.Products.AddExternalComponent(iObject.ReferenceProduct.Parent)
                    oSelectionFromSeedLing.VisProperties.SetShow catVisPropertyNoShowAttr
                    
                    sSeedLingPartName = oSeedLingInstance.ReferenceProduct.Products.Item(oSeedLingInstance.Products.Count).Name
                    sSeedLingPartName = Split(sSeedLingPartName, ".")(UBound(Split(sSeedLingPartName, ".")) - 1) & ".0"
                    oSeedLingInstance.ReferenceProduct.Products.Item(oSeedLingInstance.Products.Count).Name = sSeedLingPartName
                End If
            End If
        End If
        sTargetWindow.ActiveViewer.Reframe
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
    Next
    
    oSelectionFromSeedPart.Clear
    oSelectionFromSeedLing.Clear
    CATIA.StatusBar = "Successfully Done! "
    On Error GoTo 0
    
    MsgBox "Cabinet Project Successfully Done!    ", vbInformation, "Cabinet Project"

End Sub
