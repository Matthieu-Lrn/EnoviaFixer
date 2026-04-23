Attribute VB_Name = "frmCheckFastenerCompatibility"
Attribute VB_Base = "0{FF1755B9-43D2-44D5-9856-EEBE082E6278}{BFCF2BB6-3113-4F50-8F59-23AFE1627492}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'********************************************************************************
'* Name: Fastener Compatibility
'* Purpose: Analyse compatibility of fasteners in a selected assembly
'*
'* Assumption:
'*
'* Author:Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Private sKBEFastenerCompatibilityParametersPath As String '*** Txt file with all the parameters value
Private InchConversion As Double
Private dAngularTolerance As Double                '*** For angular misalignment
Private dLinearTolerance As Double                 '*** For linear misalignmente
Private dClashTolerance As Double                  '*** Clash tolerance inside an assy
Private dCSKTolerance As Double                    '*** To establish if first insert should be CSK or not in an assy
Private dStackupForThruInsert As Double            '*** Max gap between end points of male & last femal fastener in case of thru insert
Private dStackupForBlindInsert As Double           '*** Max gap between end points of male & last femal fastener in case of blind insert
Private dStackupFornutplateMin As Double           '*** Min gap between end points of male & last femal fastener in case of nutplate
Private dStackupFornutplateMax As Double           '*** Max gap between end points of male & last femal fastener in case of nutplate
Private dRatioForMaleFastenerTooShort As Double    '*** Min penetration percentage between male & last female fastener. it is a ratio of female fastener length
Private dLocusToFindFemaleFastener As Double       '*** If female fastener is withing radius of locus then it is added to the fastener assy

Private cSections As Sections
Private oMasterSection 'As Section
Private oTopAssy As ProductDocument
Private cClashes As Clashes
Private oClash As Clash
Private cGroups As Groups
Private oFirstGroup As Group

Private oSelectedAssembly As Product
Private ListOfFastenersAssemblies As Collection
Private ListFastenerAssemblyReport As Collection     '*** Collection of Report of checks done on fasteners assemblies
Private ListOfAllFasteners As New Collection
Private ListOfAllMaleFasteners As Collection
Private ListOfFemaleFasteners As Collection          '*** List of remaining fasteners without male fasteners
Private oInstance As Product
Private oTest As AnyObject                           '*** PlaceHolder
Private iObject As Variant                           '*** PlaceHolder
Private oFast As clsFastener                         '*** PlaceHolder
Private oFastenerAssembly As New Collection          '*** To access a fastener assembly inside collection ListOfFastenersAssemblies
Private i As Integer
Private j As Integer
Private k As Integer
Private testcollection As New Collection
Private dBoundingBoxMin(9 To 11) As Double           '*** To store min value of bounding box
Private dBoundingBoxMax(9 To 11) As Double           '*** To Store Max value of bounding box

Private bCheck As Boolean
Private bIsMale As Boolean

'**** to analyse fastassy
Private sFastassyReport As String
Private sCheckResult As String

'*** for importing tables from excel
'Public oExcel As clsExcel
Private sImportedFastLengthFromExcel As Variant           '*** MultiDimensional array of Fastener Length Imported from Excel Table
Private sImportedFastDiaFromExcel As Variant              '*** MultiDimensional array of Fastener Dia Imported from Excel Table
Private sExportResultToExcel As Variant
Private oExcelWorkbook 'As workbook
Private sReportName As String

'*** column number of excel file having fasteners length & other properties
Private iGender As Integer
Private iPNum As Integer                                  '***PartNumber string to search in Fastener Length table
Private iFunctLength As Integer
Private iPaxis_i As Integer
Private iPaxis_j As Integer
Private iPaxis_k As Integer
Private iOriginIsEndPoint As Integer
Private iLMaxPenetration As Integer
Private iLMinPenetration As Integer
Private iGripLength As Integer

Private dProgressCounter As Double
Private dMax As Double

Private iFastenerType As Integer                          '***INSERT,NUT,NUTPLATE,SCREW,WASHER
Private iPartNumber As Integer                            '***PartNumber string to search in Fastener dia table
Private iDia As Integer
Private iType As Integer                                  '***Blind or Thru
Private iThreaded As Integer
Private iCSK As Integer
Private Const USEGLOBALVALUE As Double = -5000            '*** arbitrary value to ensure that global value is used instead
Private Const EXTENVIRON As String = "OUT"
Private Const ENVOFFSET As Double = 152.4
Private Const ANALYSISSHT As String = "AnalysisResult"

Public Sub LaunchAssemblyAnalysis()
    Dim dInstanceGlobalMatrix(11) As Variant
    Dim dVector(9 To 11) As Double          '*** To enter local coordinate in array compatible with position matrix
    Dim dVectorReturn(9 To 11) As Variant   '*** To return global coordinate in array compatible with position matrix
    Dim dVectorThreadStart(9 To 11) As Double '*** to enter local coordinate of start of thread for fastener grip length
    Dim dVectorreturnThreadStart(9 To 11) As Double '*** To return global coordinate of start of threadfor fastener with GripLength
    Dim oRootAssy As Document
    Dim bLoadRefDoc As Boolean              '*** To read & reset reference document property in Tool /Options
    
    Dim dThreadstart As Double
    Dim bRunExtended As Boolean
    Dim HardWareTableSilentLog As New clsCollection
    '*** Ask User to Choose analysis type
    If MsgBox("Do you want to run Extended Analysis", vbYesNo + vbDefaultButton1 + vbQuestion, "Choose Analysis Type") = vbYes Then bRunExtended = True
    '****
    bLoadRefDoc = CATIA.SettingControllers.Item("CATCafGeneralSessionSettingCtrl").RefDoc
    '***Set up Analysis parameters
    sKBEFastenerCompatibilityParametersPath = oPathDict("KBEFastCompatibilityParams")
    
    dAngularTolerance = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "AngularTolerance"))
    dLinearTolerance = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "LinearTolerance"))
    dClashTolerance = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "ClashTolerance"))
    dCSKTolerance = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "CSKTolerance"))
    dStackupForThruInsert = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "StackupForThruInsert"))
    dStackupForBlindInsert = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "StackupForBlindInsert"))
    dStackupFornutplateMin = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "StackupFornutplateMin"))
    dStackupFornutplateMax = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "StackupFornutplateMax"))
    dRatioForMaleFastenerTooShort = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "RatioForMaleFastenerTooShort"))
    dLocusToFindFemaleFastener = CDbl(GetTagValueFromFile(sKBEFastenerCompatibilityParametersPath, "LocusToFindFemaleFastener"))

    
    '*** ask user to select an assembly
    Call selectAssy(oSelectedAssembly, oRootAssy, , "Please select an Assembly to validate!   ")
    Set sSourceWindow = CATIA.ActiveWindow
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSelectedAssembly Is Nothing Then
            MsgBox "Please select a valid Assy!  ", vbCritical
            GoTo NoSelect
        ElseIf oSelectedAssembly.Products.Count = 0 Then
            MsgBox "No Parts Found in the assembly!  ", vbCritical
            GoTo NoSelect
        End If
    Else
        GoTo NoSelect
    End If
    
    '*** get clashes object for provisional fastener analysis
    Set cClashes = CATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")
    Set cGroups = CATIA.ActiveDocument.Product.GetTechnologicalObject("Groups")
    Set oFirstGroup = cGroups.Add()
    oFirstGroup.Name = "AutomationGroup_001"
    Set oClash = cClashes.Add()
    oClash.FirstGroup = oFirstGroup
    oClash.ComputationType = catClashComputationTypeAgainstAll
    oClash.InterferenceType = catClashInterferenceTypeClearance
    oClash.Clearance = 0.1 '*** 1 in
    oClash.Name = "AutomationClash_001"
    '***
    '*** Launch import from Excel table
    Call oImportTablefromExcel.LaunchImport(sImportedFastLengthFromExcel, sKbe_Hardwaretable, "FastenerLength")
    If UBound(sImportedFastLengthFromExcel) = 0 Then
        MsgBox "Cannot find HardwareTable Please contact KBE developer!  ", vbCritical
        Exit Sub
    End If
    
    Call oImportTablefromExcel.LaunchImport(sImportedFastDiaFromExcel, sKbe_Hardwaretable, "FastenerDiameter")
    If UBound(sImportedFastDiaFromExcel) = 0 Then
        MsgBox "Cannot find HardwareTable Please contact KBE developer!  ", vbCritical
        Exit Sub
    End If
    
    '***Accessing Column name from each table
    For i = 1 To UBound(sImportedFastLengthFromExcel, 2)
        If sImportedFastLengthFromExcel(1, i) = "Gender" Then iGender = i
        If sImportedFastLengthFromExcel(1, i) = "PartNumber" Then iPNum = i
        If sImportedFastLengthFromExcel(1, i) = "FunctionalLength_mm" Then iFunctLength = i
        If sImportedFastLengthFromExcel(1, i) = "Paxis_i" Then iPaxis_i = i
        If sImportedFastLengthFromExcel(1, i) = "Paxis_j" Then iPaxis_j = i
        If sImportedFastLengthFromExcel(1, i) = "Paxis_k" Then iPaxis_k = i
        If sImportedFastLengthFromExcel(1, i) = "OriginIsEndPoint" Then iOriginIsEndPoint = i
        If sImportedFastLengthFromExcel(1, i) = "L-MaxPenetration_mm" Then iLMaxPenetration = i
        If sImportedFastLengthFromExcel(1, i) = "L-MinPenetration_mm" Then iLMinPenetration = i
        If sImportedFastLengthFromExcel(1, i) = "GripLength" Then iGripLength = i
        
    Next
    
    For i = 1 To UBound(sImportedFastDiaFromExcel, 2)
        If sImportedFastDiaFromExcel(1, i) = "FastenerType" Then iFastenerType = i
        If sImportedFastDiaFromExcel(1, i) = "PartNumber" Then iPartNumber = i
        If sImportedFastDiaFromExcel(1, i) = "Dia" Then iDia = i
        If sImportedFastDiaFromExcel(1, i) = "Type" Then iType = i
        If sImportedFastDiaFromExcel(1, i) = "Threaded" Then iThreaded = i
        If sImportedFastDiaFromExcel(1, i) = "CSK" Then iCSK = i
    Next
    
    '***Initialize progress bar
    timeStart = DateTime.Timer  'Time set for Progress Bar
    Call frmProgress.progressBarInitialize("Fastener Assemblies Analysis")
    
    '*** Building Parent List of all instance objects inside SP collectors in the root assy
    '***Turn external reference "on" if not already on before building parents list
    
    If bLoadRefDoc = False Then CATIA.SettingControllers.Item("CATCafGeneralSessionSettingCtrl").RefDoc = True
    
    Dim sRefParentLst As String
    Dim sParentLst As String
    Dim sAvoidAsy As String
    sParentLst = BuildParentList(oSelectedAssembly)
    sRefParentLst = BuildParentListWithRefProducts(oSelectedAssembly)
    sAvoidAsy = sParentLst
    Call GenerateItemlist(oSelectedAssembly, , sParentLst, sRefParentLst) 'Scanning selected assy
    If bCancelAction = True Then '***if user presses cancel on progress toolbar
        GoTo NoSelect:
    End If
    If ListOfAllFasteners.Count = 0 Then
        MsgBox "No FastenersFound"
        GoTo NoSelect:
    End If
    '*** if extended analysis is chosen
    If bRunExtended Then
        '*** initialize bounding box
        dBoundingBoxMin(9) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(9)
        dBoundingBoxMin(10) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(10)
        dBoundingBoxMin(11) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(11)
        dBoundingBoxMax(9) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(9)
        dBoundingBoxMax(10) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(10)
        dBoundingBoxMax(11) = ListOfAllFasteners.Item(1).GlobalPositionMatrixIndex(11)
        
        If ListOfAllFasteners.Count >= 2 Then
            For i = 2 To ListOfAllFasteners.Count
                If dBoundingBoxMin(9) > ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(9) Then _
                    dBoundingBoxMin(9) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(9)
                If dBoundingBoxMin(10) > ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(10) Then _
                    dBoundingBoxMin(10) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(10)
                If dBoundingBoxMin(11) > ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(11) Then _
                    dBoundingBoxMin(11) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(11)
                If dBoundingBoxMax(9) < ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(9) Then _
                    dBoundingBoxMax(9) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(9)
                If dBoundingBoxMax(10) < ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(10) Then _
                    dBoundingBoxMax(10) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(10)
                If dBoundingBoxMax(11) < ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(11) Then _
                    dBoundingBoxMax(11) = ListOfAllFasteners.Item(i).GlobalPositionMatrixIndex(11)
            Next
        End If
        '***Scan environment
        sParentLst = BuildParentList(oRootAssy.Product)
        sRefParentLst = BuildParentListWithRefProducts(oRootAssy.Product)
        Call GenerateItemlist(oRootAssy.Product, sAvoidAsy, sParentLst, sRefParentLst, EXTENVIRON, True) 'Scanning Environment
        If bCancelAction = True Then '***if user presses cancel on progress toolbar
            GoTo NoSelect:
        End If
    End If
    '*** Following code will look for part number of fastener in the imported fastener table & if it finds then it will populate
    '*** ListOfAllMaleFasteners & ListOfFemaleFasteners and populate attributes of each fastener.In case it doesn't find it then
    '*** it removes the fastener from ListofAllfasteners list.
    Set ListOfAllMaleFasteners = New Collection
    Set ListOfFemaleFasteners = New Collection
    dProgressCounter = 0
    dMax = ListOfAllFasteners.Count
    
    For Each oFast In ListOfAllFasteners
        bCheck = False
        If oFast.InstanceObject.HasAMasterShapeRepresentation Then '*** ensure that oFast is a catpart
            '***progress bar
            dProgressCounter = dProgressCounter + 1
            Call frmProgress.progressBarRepaint("Assigning attributes to each fastener...", dMax, dProgressCounter)
            If bCancelAction = True Then '***if user presses cancel on progress toolbar
                   GoTo NoSelect:
            End If
            '***----
            Err.Clear
            On Error Resume Next
            Set oTest = oFast.InstanceObject.ReferenceProduct
            If Err.Number <> 0 Then
                ListOfAllFasteners.Item(oFast.UniqueNameInAssembly).InstanceObject.ApplyWorkMode DEFAULT_MODE
            End If
            Err.Clear
            oFast.ReferenceProductName = oFast.InstanceObject.ReferenceProduct.Name
            If Err.Number <> 0 Then '*** If instance cannot be loaded then remove it from the list
                ListOfAllFasteners.Remove oFast.UniqueNameInAssembly
                On Error GoTo 0
            Else
                On Error GoTo 0
                Dim bFastInLengthTable As Boolean: bFastInLengthTable = False
                Dim iFastInLengthTable As Integer: iFastInLengthTable = 0
                '***Populating Properties from Table FasterLength
                For i = 2 To UBound(sImportedFastLengthFromExcel, 1)
                    If oFast.ReferenceProductName Like sImportedFastLengthFromExcel(i, iPNum) Or oFast.ReferenceProductName Like sImportedFastLengthFromExcel(i, iPNum) + "D##" Then
                      bFastInLengthTable = True
                      iFastInLengthTable = iFastInLengthTable + 1
                        
                      oFast.Length = sImportedFastLengthFromExcel(i, iFunctLength)
                      dVector(9) = sImportedFastLengthFromExcel(i, iFunctLength) * sImportedFastLengthFromExcel(i, iPaxis_i)
                      dVector(10) = sImportedFastLengthFromExcel(i, iFunctLength) * sImportedFastLengthFromExcel(i, iPaxis_j)
                      dVector(11) = sImportedFastLengthFromExcel(i, iFunctLength) * sImportedFastLengthFromExcel(i, iPaxis_k)
                      Call MultiplyVector(dVector, oFast.GlobalPositionMatrix, dVectorReturn, True)                       '*** Find global position of Functional length point
                      '*** Populating start, end and thread point
                      If UCase(sImportedFastLengthFromExcel(i, iOriginIsEndPoint)) = "TRUE" Then
                        Call oFast.SetStartPoint(dVectorReturn)
                        Call oFast.SetEndPoint(oFast.GlobalPositionMatrix)
                        If sImportedFastLengthFromExcel(i, iGripLength) <> "-" And IsNumeric(sImportedFastLengthFromExcel(i, iGripLength)) Then
                            oFast.HasGripLength = True
                            dThreadstart = sImportedFastLengthFromExcel(i, iFunctLength) - sImportedFastLengthFromExcel(i, iGripLength)
                            dVectorThreadStart(9) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_i)
                            dVectorThreadStart(10) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_j)
                            dVectorThreadStart(11) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_k)
                            Call MultiplyVector(dVectorThreadStart, oFast.GlobalPositionMatrix, dVectorreturnThreadStart, True)
                            Call oFast.SetThreadPoint(dVectorreturnThreadStart)
                        End If
                      Else
                        Call oFast.SetStartPoint(oFast.GlobalPositionMatrix)
                        Call oFast.SetEndPoint(dVectorReturn)
                        If sImportedFastLengthFromExcel(i, iGripLength) <> "-" And IsNumeric(sImportedFastLengthFromExcel(i, iGripLength)) Then
                            oFast.HasGripLength = True
                            dThreadstart = sImportedFastLengthFromExcel(i, iGripLength)
                            dVectorThreadStart(9) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_i)
                            dVectorThreadStart(10) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_j)
                            dVectorThreadStart(11) = dThreadstart * sImportedFastLengthFromExcel(i, iPaxis_k)
                            Call MultiplyVector(dVectorThreadStart, oFast.GlobalPositionMatrix, dVectorreturnThreadStart, True)
                            Call oFast.SetThreadPoint(dVectorreturnThreadStart)
                        End If
                      End If
                      '*** Populate Gender
                      If UCase(sImportedFastLengthFromExcel(i, iGender)) = "MALE" Then
                        oFast.IsMale = True
                      End If
                      '***Population Min and Max penetration minus length
                      If sImportedFastLengthFromExcel(i, iLMaxPenetration) = "-" Then
                        oFast.StackupOffset = USEGLOBALVALUE
                      Else
                        oFast.StackupOffset = sImportedFastLengthFromExcel(i, iLMaxPenetration)
                      End If
                      If sImportedFastLengthFromExcel(i, iLMinPenetration) = "-" Then
                        oFast.StackupOffsetMin = USEGLOBALVALUE
                      Else
                        oFast.StackupOffsetMin = sImportedFastLengthFromExcel(i, iLMinPenetration)
                      End If
                    End If
                Next
                '***Populating Properties from Table FastenerDiameter
                Dim bFastInDIaTable As Boolean: bFastInDIaTable = False
                Dim iTestFastenerDiaTable As Integer: iTestFastenerDiaTable = 0
                For j = 2 To UBound(sImportedFastDiaFromExcel, 1)
                    On Error Resume Next
                    If oFast.ReferenceProductName Like sImportedFastDiaFromExcel(j, iPartNumber) Or oFast.ReferenceProductName Like sImportedFastDiaFromExcel(j, iPartNumber) + "D##" Then
                     If Err.Number <> 0 Then Stop
                     On Error GoTo 0
                        bFastInDIaTable = True
                        iTestFastenerDiaTable = iTestFastenerDiaTable + 1

                        oFast.Dia = UCase(sImportedFastDiaFromExcel(j, iDia))
                        oFast.FastenerType = UCase(sImportedFastDiaFromExcel(j, iFastenerType))
                        If sImportedFastDiaFromExcel(j, iType) = "BLIND" Then oFast.IsBlind = True Else oFast.IsBlind = False
                        If sImportedFastDiaFromExcel(j, iThreaded) = "THREADED" Then oFast.IsThreaded = True Else oFast.IsThreaded = False
                        If sImportedFastDiaFromExcel(j, iCSK) = "CSK" Then oFast.IsCounterSunk = True Else oFast.IsCounterSunk = False
                    End If
                Next
                '******* Silent Log of Table info ********
                If iTestFastenerDiaTable > 1 Then
                    If Not HardWareTableSilentLog.Exists(oFast.ReferenceProductName & "|" & "AmbiguousInDiaTable") Then HardWareTableSilentLog.Add oFast.ReferenceProductName & "|" & "AmbiguousInDiaTable", ""
                ElseIf iTestFastenerDiaTable = 0 And iFastInLengthTable <> 0 Then
                    If Not HardWareTableSilentLog.Exists(oFast.ReferenceProductName & "|" & "MissingInDiaTable") Then HardWareTableSilentLog.Add oFast.ReferenceProductName & "|" & "MissingInDiaTable", ""
                End If
                If iFastInLengthTable > 1 Then
                    If Not HardWareTableSilentLog.Exists(oFast.ReferenceProductName & "|" & "AmbiguousInLengthTable") Then HardWareTableSilentLog.Add oFast.ReferenceProductName & "|" & "AmbiguousInLengthTable", ""
                ElseIf iFastInLengthTable = 0 And iTestFastenerDiaTable <> 0 Then
                    If Not HardWareTableSilentLog.Exists(oFast.ReferenceProductName & "|" & "MissingInLengthTable") Then HardWareTableSilentLog.Add oFast.ReferenceProductName & "|" & "MissingInLengthTable", ""
                End If
                '******************************************
                bCheck = bFastInLengthTable And bFastInDIaTable
                If Not bCheck Then
                    ListOfAllFasteners.Remove oFast.UniqueNameInAssembly
                Else
                    If oFast.IsMale Then
                      On Error Resume Next    ' error may happen due to ambiguous entry in table
                      ListOfAllMaleFasteners.Add oFast, oFast.UniqueNameInAssembly
                      On Error GoTo 0
                    Else
                      On Error Resume Next    ' error may happen due to ambiguous entry in table
                      ListOfFemaleFasteners.Add oFast, oFast.UniqueNameInAssembly
                      On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next

    '*** create list of fastener assemblies
    Set ListOfFastenersAssemblies = New Collection
    Call CreateListsOfFastenersAssemblies(dLocusToFindFemaleFastener)
    
    '*** if extended analysis is chosen
    If bRunExtended Then
        '***Filter list of FemaleFastener
        For Each oFast In ListOfFemaleFasteners
            If oFast.Tag = EXTENVIRON Then ListOfFemaleFasteners.Remove oFast.UniqueNameInAssembly
        Next
        
        '*** Filter Fastener Assyneeded
        For Each oFastenerAssembly In ListOfFastenersAssemblies
            bCheck = False
            For i = 1 To oFastenerAssembly.Count
              If oFastenerAssembly.Item(i).Tag = "" Then bCheck = True
            Next i
            If Not (bCheck) Then ListOfFastenersAssemblies.Remove oFastenerAssembly.Item(1).UniqueNameInAssembly + "_Fastassy"
        Next
    End If
    '*** analysing fastener assemblies
    dProgressCounter = 0
    dMax = ListOfFastenersAssemblies.Count + ListOfFemaleFasteners.Count
    Set ListFastenerAssemblyReport = New Collection

    For Each oFastenerAssembly In ListOfFastenersAssemblies
        '*** progress bar
        dProgressCounter = dProgressCounter + 1
        If bCancelAction = True Then '***if user presses cancel on progress toolbar
            GoTo NoSelect:
        End If
        '***----
        Call frmProgress.progressBarRepaint("Analysing fastener assemblies...", dMax, dProgressCounter)
        sFastassyReport = ""
        sCheckResult = ""
        sCheckResult = sCheckMisAlignment(oFastenerAssembly)                            '**** Checking MisAlignment of Assembly
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + sCheckResult
        
        sCheckResult = sCheckClashInsideAssy(oFastenerAssembly)                         '**** Checking Clash of Assembly
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + ";" + sCheckResult
        
        sCheckResult = sCheckThreadofAssy(oFastenerAssembly)                            '**** Checking Threaded fasteners & dia of assemby
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + ";" + sCheckResult
        
        sCheckResult = sCheckCSK(oFastenerAssembly)                                     '**** Checking CSK fasteners
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + ";" + sCheckResult
        
        sCheckResult = sCheckLastFastenerOfAssy(oFastenerAssembly)                      '**** Checking Last fastener for stack length  etc
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + ";" + sCheckResult
        
        sCheckResult = sCheckShankInThreadedFemaleFastener(oFastenerAssembly)           '**** checking shank in thead for bolts with grip length
        If sCheckResult <> "" Then sFastassyReport = sFastassyReport + ";" + sCheckResult
        
        '***Remove useless ";" from start & end
        While Right(sFastassyReport, 1) = ";"
            sFastassyReport = Mid(sFastassyReport, 1, Len(sFastassyReport) - 1)
        Wend
        
        While Left(sFastassyReport, 1) = ";"
            sFastassyReport = Mid(sFastassyReport, 2, Len(sFastassyReport))
        Wend
        '*** add one ";" at the end to give line break between assemblies
        
        If sFastassyReport <> "" Then sFastassyReport = sFastassyReport + ";"
        
        ListFastenerAssemblyReport.Add sFastassyReport
    Next
    '*** Adding isolated Female Fasteners to Assembly report
    On Error GoTo 0
    Call SwapToAssemblyWorkbench
    
    For Each oFast In ListOfFemaleFasteners
        dProgressCounter = dProgressCounter + 1
        Call frmProgress.progressBarRepaint("Analysing fastener assemblies...", dMax, dProgressCounter)
        If bCancelAction = True Then '***if user presses cancel on progress toolbar
            GoTo NoSelect:
        End If
        '***----
        ListOfFastenersAssemblies.Add New Collection, oFast.UniqueNameInAssembly + "_Fastassy"
        ListOfFastenersAssemblies.Item(oFast.UniqueNameInAssembly + "_Fastassy").Add oFast, oFast.UniqueNameInAssembly
        If bCheckProvisionalFasteners(oFast) Then       '*** check for provisional inserts
            ListFastenerAssemblyReport.Add "---without male fastener;---May be a provisional fastener;"
        Else
            ListFastenerAssemblyReport.Add "---without male fastener;"
        End If

    Next
    '*** Log to HardwareTableSilent log
    If HardWareTableSilentLog.Count > 0 Then
        For i = 1 To HardWareTableSilentLog.Count
            Call AddToLogFile("HardwareTableSilentLog", Split(HardWareTableSilentLog.GetKey(i), "|")(0), , , , Split(HardWareTableSilentLog.GetKey(i), "|")(1))
        Next
    End If
    '***Generate Report In Excel
    Call ExtractReportToExcel
    
    With Me
        .Height = dTitleBarHeight + 68
        .Width = 110
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 29
        .imgDSGlobe.Left = -3
        .cmdSelect.Top = .Height - dTitleBarHeight - 60
        .cmdSelect.Left = .Width - .cmdSelect.Width - .cmdSelectExcel.Width - 15
        .cmdSelectExcel.Top = .Height - dTitleBarHeight - 60
        .cmdSelectExcel.Left = .Width - .cmdSelectExcel.Width - 10
        .cmdSelect.ControlTipText = "Select a fastassy from Excel report and press Catia button to visualize in 3D"
        .cmdSelectExcel.ControlTipText = "Press Excel button and select an instance to highlight in the report"
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        
        .cmdSelect.SetFocus
        .Show vbModless
    End With
NoSelect:
On Error Resume Next
    frmProgress.Hide
    cClashes.Remove ("AutomationClash_001")
    cGroups.Remove ("AutomationGroup_001")
On Error GoTo 0
    '*** Reset external Reference properties as per user's setting
    CATIA.SettingControllers.Item("CATCafGeneralSessionSettingCtrl").RefDoc = bLoadRefDoc
End Sub
Private Sub cmdSelectExcel_Click()
    Dim oRT As Document
    Dim oInstance As Product
    Dim sReturn As String
    Dim oRange As Object
    
    bStopMultiSelection = True  '*** to cancel previously running multiselection
    Call SelectInstance(oInstance, oRT, "Select a fastener from the assy")
    If Not (oInstance Is Nothing) Then sReturn = sSearchStringForExcel(oInstance)
    bCheck = False
    If Not sReturn = "" Then
        On Error Resume Next
        Set oRange = oExcelWorkbook.Sheets.Item(ANALYSISSHT).Range("a:a").Find(sReturn)
        If Err.Number <> 0 Then
            MsgBox "Link between Catia and Excel broken." + vbCrLf + _
                    "This could be due to the fact that user has closed " + vbCrLf + _
                    "the fastener assembly report.", vbOKOnly + vbExclamation
            Exit Sub
        End If
        On Error GoTo 0
        If Not (oRange Is Nothing) Then
            oExcelWorkbook.Sheets.Item(ANALYSISSHT).Activate
            oExcelWorkbook.Sheets.Item(ANALYSISSHT).Cells(oRange.Row, 2).Select
            oExcelGlobal.App.ActiveWindow.Zoom = True
            DoEvents
            oExcelGlobal.App.ActiveWindow.Zoom = 100
            DoEvents
            bCheck = True
        End If
    End If
    If Not (bCheck) Then
        On Error Resume Next
        Err.Clear
        Set oTest = oInstance.ReferenceProduct
        If Err.Number <> 0 Then
            oInstance.ApplyWorkMode DEFAULT_MODE
        End If
        Err.Clear
        Set oTest = oInstance.ReferenceProduct
        If Err.Number <> 0 Then
            MsgBox "Instance cannot be loaded"
            On Error GoTo 0
        Else
            On Error GoTo 0
            For i = 2 To UBound(sImportedFastDiaFromExcel, 1)
                If oInstance.ReferenceProduct.Name Like sImportedFastDiaFromExcel(i, iPartNumber) Or _
                oInstance.ReferenceProduct.Name Like sImportedFastDiaFromExcel(i, iPartNumber) + "D##" Then bCheck = True
            Next
            If bCheck Then
                If sReturn = "" Then
                    MsgBox "Instance out of scope of current analysis"
                Else
                    MsgBox "No error found in fastassy analysis report"
                End If
            Else
                MsgBox "Instance not a preferred fastener"
            End If
        End If
    End If
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        On Error Resume Next
        cSections.Remove "Automation_001Section" '*** remove any previously created c Section
        On Error GoTo 0
        Me.Hide
        Call oExcelGlobal.CloseAlreadyOpenWorkbook(sReportName)
        Call frmKBEMain.resetToolbar
    End If
End Sub
Private Sub cmdSelect_Click()
    Dim sAssyName As String
    
    sAssyName = ""
    '**Getting value of coloum1 and selected row
    On Error Resume Next
        sAssyName = oExcelGlobal.App.Cells(oExcelGlobal.App.ActiveCell.Row, 1).Value
    On Error GoTo 0
    If sAssyName = "" Or Not (sAssyName Like "*_Fastassy") Then
        MsgBox ("Please select the row with assembly name in Excel")
    Else
        Call SelectAssemblyInCatia(sAssyName)
    End If
End Sub
Private Function bInsideEnvironment(ByRef GlobalPosMtx() As Variant) As Boolean
    
  If GlobalPosMtx(9) > (dBoundingBoxMin(9) - ENVOFFSET) And _
    GlobalPosMtx(10) > (dBoundingBoxMin(10) - ENVOFFSET) And _
    GlobalPosMtx(11) > (dBoundingBoxMin(11) - ENVOFFSET) And _
    GlobalPosMtx(9) < (dBoundingBoxMax(9) + ENVOFFSET) And _
    GlobalPosMtx(10) < (dBoundingBoxMax(10) + ENVOFFSET) And _
    GlobalPosMtx(11) < (dBoundingBoxMax(11) + ENVOFFSET) Then bInsideEnvironment = True _
    Else bInsideEnvironment = False
End Function

Private Sub GenerateItemlist(ByVal oAssembly As Product, Optional sAvoidAssy As String = "!!!!!", _
                             Optional sListName As String = "", Optional sRefListName As String = "", _
                             Optional sFastenerTag As String = "", Optional bCheckEnv As Boolean = False)
'*** This functions scans everthing underneath oAssembly & populates
'*** List of all fasteners collection with catparts of fasteners
'*** A fastener at this point is everthing under SPXX assembly
    Dim oInstance As Product
    Dim sParentList As String
    Dim sRefParentList As String
    Dim dInstanceGlobalMatrix(11) As Variant
    
    For Each oInstance In oAssembly.Products
        '***Progress Bar
        If bCancelAction Then  '***if user presses cancel on progress toolbar
            Exit Sub
        End If
        timeEnd = DateTime.Timer
        Call frmProgress.progressBarRepaint("Initial Scan...", , , , , , "Elapsed time: " & Format((timeEnd - timeStart) / 86400, "hh:mm:ss"))
        If bCancelAction Then  '***if user presses cancel on progress toolbar
            Exit Sub
        End If
        '****
        If oInstance.Products.Count > 0 Then
            sParentList = oInstance.Name & ";" & sListName
            On Error GoTo ExtFcnt
            sRefParentList = oInstance.ReferenceProduct.Name & ";" & sListName
            On Error GoTo 0
            If Not (sParentList Like "*" & sAvoidAssy & "*") Then Call GenerateItemlist(oInstance, sAvoidAssy, sParentList, sRefParentList, sFastenerTag, bCheckEnv) ' recursion to scan all sub branches
        Else
            sRefParentList = oInstance.Name & ";" & sRefListName
            sParentList = oInstance.Name & ";" & sListName
            bCheck = False
            If sRefParentList Like "*SP##*" Or sRefParentList Like "*SP###*" Then bCheck = True
            If bCheck = True Then
                Err.Clear
                On Error Resume Next
                ListOfAllFasteners.Add New clsFastener, sParentList
                If Err.Number <> 0 Then
                    MsgBox "Duplicate Instance Name " & Right(sParentList, Len(sParentList) - InStrRev(sParentList, ";")) & "In " & Split(sParentList, ";")(UBound(Split(sParentList, ";")) - 1)
                    bCancelAction = True
                End If
                On Error GoTo 0
                ListOfAllFasteners.Item(sParentList).UniqueNameInAssembly = sParentList
                Set ListOfAllFasteners.Item(sParentList).InstanceObject = oInstance
                Call CalculateInstanceGlobalPositionMatrix(sParentList, dInstanceGlobalMatrix)       '*** Find global position of Origin point
                ListOfAllFasteners.Item(sParentList).GlobalPositionMatrix = dInstanceGlobalMatrix
                If Not (sFastenerTag) = "" Then ListOfAllFasteners.Item(sParentList).Tag = sFastenerTag '*** tag fasteners
                If bCheckEnv Then
                    If Not (bInsideEnvironment(dInstanceGlobalMatrix)) Then ListOfAllFasteners.Remove sParentList
                End If
            End If
        End If
    Next
ExtFcnt:
End Sub
Public Function BuildParentListWithRefProducts(ByVal iObject As Variant) As String

    'This is a function similar to Build Parent List but we build Parent list with Reference product name, instead of instance name
    BuildParentListWithRefProducts = iObject.Name & ";"
    
    Do
        Set iObject = iObject.Parent
        
        'We don't care about the "Products" or "ProductDocuments" objects
        If TypeName(iObject) = "Product" Then
            BuildParentListWithRefProducts = BuildParentListWithRefProducts & iObject.ReferenceProduct.Name & ";"
        End If
    Loop Until TypeName(iObject) = "Application"

End Function
Private Sub CreateListsOfFastenersAssemblies(Optional ByRef dComparisonDimension As Double = 5)

    Dim dStartPoint(2) As Double
    Dim oDummyFast As New clsFastener
    Dim oMaleFast As New clsFastener
    Dim oFemaleFast As New clsFastener
    Dim oNextFemaleFast As New clsFastener
    Dim iCheck As Integer
   
    '*** First element of Fastassy is always male fastners
    '*** Fastassies are named as MaleFast.UniqueNameInAssembly + "_Fastassy"
    '*** All elements inside a fastassy can be accessed with fastner name generated by buildparent list
    For Each oMaleFast In ListOfAllMaleFasteners
      ListOfFastenersAssemblies.Add New Collection, oMaleFast.UniqueNameInAssembly + "_Fastassy"
      ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Add oMaleFast, oMaleFast.UniqueNameInAssembly
      For Each oFemaleFast In ListOfFemaleFasteners
          If bCheckFemaleFastenerBelongToAssy(oMaleFast, oFemaleFast, dComparisonDimension) Then
              ListOfFemaleFasteners.Remove oFemaleFast.UniqueNameInAssembly
              ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Add oFemaleFast, oFemaleFast.UniqueNameInAssembly
          End If
      Next
    Next
    '*** Sort Fastener List---Using Bubble Sort
    For Each oMaleFast In ListOfAllMaleFasteners
        If ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Count > 2 Then  '*** Sort only if there are more than two fasterens in the list
            For i = 2 To ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Count - 1
               For j = i + 1 To ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Count
                   Set oFemaleFast = ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Item(i)
                   Set oNextFemaleFast = ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Item(j)
                   If dDistancePointToLineSeg(oMaleFast.StartPoint, oFemaleFast.StartPoint, oFemaleFast.EndPoint) > dDistancePointToLineSeg(oMaleFast.StartPoint, oNextFemaleFast.StartPoint, oNextFemaleFast.EndPoint) Then
                       Set oDummyFast = ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Item(j)
                       ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Remove j
                       ListOfFastenersAssemblies.Item(oMaleFast.UniqueNameInAssembly + "_Fastassy").Add oDummyFast, oDummyFast.UniqueNameInAssembly, i
                   End If
                Next j
             Next i
        End If
    Next
End Sub
Private Function bCheckFemaleFastenerBelongToAssy(oMaleFastener As clsFastener, oFemaleFastener As clsFastener, dComparisonDimension As Double) As Boolean
    Dim dDistFemaleStarttoMaleStartPlane As Double
    Dim dDistFemaleEndtoMaleStartPlane As Double
    Dim dDistFemaleStartToMaleEndPlane As Double
    Dim dDistFemaleEndToMaleEndPlane As Double
    Dim bFastenerAboveMaleFastener As Boolean
    Dim bFastenerbelowMaleFastener As Boolean
    
    bFastenerAboveMaleFastener = False
    bFastenerbelowMaleFastener = False

    
    If dDistanceLineSegToLineSeg(oMaleFastener.StartPoint, oMaleFastener.EndPoint, oFemaleFastener.StartPoint, oFemaleFastener.EndPoint) <= dComparisonDimension Then
        dDistFemaleStarttoMaleStartPlane = dDistancePointToPlane(oFemaleFastener.StartPoint, oMaleFastener.StartPoint, oMaleFastener.EndPoint)
        dDistFemaleEndtoMaleStartPlane = dDistancePointToPlane(oFemaleFastener.EndPoint, oMaleFastener.StartPoint, oMaleFastener.EndPoint)
        dDistFemaleStartToMaleEndPlane = dDistancePointToPlane(oFemaleFastener.StartPoint, oMaleFastener.EndPoint, oMaleFastener.StartPoint)
        dDistFemaleEndToMaleEndPlane = dDistancePointToPlane(oFemaleFastener.EndPoint, oMaleFastener.EndPoint, oMaleFastener.StartPoint)
        
        If dDistFemaleStarttoMaleStartPlane < 0 And dDistFemaleEndtoMaleStartPlane < 0 Then bFastenerAboveMaleFastener = True
        If dDistFemaleStartToMaleEndPlane < 0 And dDistFemaleEndToMaleEndPlane < 0 Then bFastenerbelowMaleFastener = True
        
        If bFastenerAboveMaleFastener Or bFastenerbelowMaleFastener Then
             bCheckFemaleFastenerBelongToAssy = False
        Else
             bCheckFemaleFastenerBelongToAssy = True
        End If
    End If
End Function

Private Sub ExtractReportToExcel()
    Dim sString As String
    Dim iNbRow As Integer
    Dim iNbColumn As Integer
    Dim iRowCounter As Integer
    Dim ReportCollection As New Collection
    Dim ReportKeysCollection As New Collection
    Dim sStatus As String
    
    iRowCounter = 2
    Set ReportCollection = New Collection
   '*** Building ReportCollection
    For i = 1 To ListFastenerAssemblyReport.Count
        If ListFastenerAssemblyReport.Item(i) <> "" Then
            sString = ""
            For j = 1 To ListOfFastenersAssemblies.Item(i).Count
                Set oFast = ListOfFastenersAssemblies.Item(i).Item(j)
                If sString = "" Then sString = sString + Left(oFast.UniqueNameInAssembly, InStr(oFast.UniqueNameInAssembly, ";") - 1) Else sString = sString + "; " + Left(oFast.UniqueNameInAssembly, InStr(oFast.UniqueNameInAssembly, ";") - 1)
            Next
            ReportCollection.Add sString
            ReportKeysCollection.Add ListOfFastenersAssemblies.Item(i).Item(1).UniqueNameInAssembly + "_Fastassy"

            iObject = Split(ListFastenerAssemblyReport.Item(i), ";")
            For k = 0 To UBound(iObject)
                ReportCollection.Add iObject(k)
                ReportKeysCollection.Add ListOfFastenersAssemblies.Item(i).Item(1).UniqueNameInAssembly + "_Fastassy"
            Next
        End If
    Next
        
    '***** Launch Excel
    Set oExcelGlobal = New clsExcel
    Call oExcelGlobal.GetExcel
    oExcelGlobal.App.ScreenUpdating = False
    Call oExcelGlobal.CloseAlreadyOpenWorkbook(sKbe_ReportTemplate)
    Call oExcelGlobal.CloseAlreadyOpenWorkbook("C:\Temp\FastenersAssembliesResult.xls")
    Call oExcelGlobal.OpenExcelFile(sKbe_ReportTemplate)
    Set oExcelWorkbook = oExcelGlobal.App.Workbooks.Item(Right(sKbe_ReportTemplate, Len(sKbe_ReportTemplate) - InStrRev(sKbe_ReportTemplate, "\")))
    oExcelWorkbook.Sheets.Item(ANALYSISSHT).Activate
    '*** write name of analysed assembly
    oExcelGlobal.App.Cells(1, 2).Value = "Analysed Assy : " + oSelectedAssembly.Name
    '***Generating Array for export
    If ReportCollection.Count > 0 Then
        sStatus = "ERRORS"
        ReDim sExportResultToExcel(1 To ReportCollection.Count, 1 To 2)
        
        For i = 1 To ReportKeysCollection.Count
                sExportResultToExcel(i, 1) = ReportKeysCollection.Item(i)
                sExportResultToExcel(i, 2) = ReportCollection.Item(i)
        Next
        '*** Writing To Excel
        iNbRow = ReportKeysCollection.Count + 2
        iNbColumn = 2
        oExcelGlobal.App.Range(oExcelGlobal.App.Cells(3, 1), oExcelGlobal.App.Cells(iNbRow, iNbColumn)).NumberFormat = "@"
        oExcelGlobal.App.Range(oExcelGlobal.App.Cells(3, 1), oExcelGlobal.App.Cells(iNbRow, iNbColumn)).Value = sExportResultToExcel
        '***Formatting
        For iNbRow = 2 To ReportKeysCollection.Count + 2
            If Not (oExcelGlobal.App.Cells(iNbRow, 2).Value2 Like "---*") And oExcelGlobal.App.Cells(iNbRow, 2).Value2 <> "" Then
                oExcelGlobal.App.Cells(iNbRow, 2).Interior.ColorIndex = 35
                With oExcelGlobal.App.Cells(iNbRow - 1, 2)
                .Borders(9).LineStyle = 1    'xlContinuous
                .Borders(9).Weight = -4138   'xlMedium
                End With
            End If
        Next
        '***Last Row format
        With oExcelGlobal.App.Cells(ReportKeysCollection.Count + 2, 2)
                .Borders(9).LineStyle = 1    'xlContinuous
                .Borders(9).Weight = -4138   'xlMedium
        End With
     Else '*** if no error found
        oExcelGlobal.App.Range(oExcelGlobal.App.Cells(3, 2), oExcelGlobal.App.Cells(3, 2)).NumberFormat = "@"
        oExcelGlobal.App.Range(oExcelGlobal.App.Cells(3, 2), oExcelGlobal.App.Cells(3, 2)).Value = "Congratulations! No error Found."
        sStatus = "OK"
     End If
        oExcelGlobal.App.Columns("B").Select
        oExcelGlobal.App.Selection.EntireColumn.AutoFit
        oExcelGlobal.App.Cells(1, 2).Select
    '*** Change Worksheet & adding tolerances
    oExcelGlobal.App.Workbooks.Item(Right(sKbe_ReportTemplate, Len(sKbe_ReportTemplate) - InStrRev(sKbe_ReportTemplate, "\"))).Sheets.Item("TolerancesHardwareCompatibility").Activate
    oExcelGlobal.App.Cells(2, 4).Value = CStr(Round(Abs(dAngularTolerance), 3)) + "Degs"
    oExcelGlobal.App.Cells(3, 4).Value = CStr(Round(Abs(dLinearTolerance / 25.4), 3)) + "in"
    oExcelGlobal.App.Cells(4, 4).Value = CStr(Round(Abs(dClashTolerance / 25.4), 3)) + "in"
    oExcelGlobal.App.Cells(5, 4).Value = CStr(Round(Abs(dRatioForMaleFastenerTooShort), 3))
    oExcelGlobal.App.Cells(6, 4).Value = CStr(Round(Abs(dStackupForThruInsert / 25.4), 3)) + "in"
    oExcelGlobal.App.Cells(7, 4).Value = CStr(Round(Abs(dStackupForBlindInsert / 25.4), 3)) + "in"
    oExcelGlobal.App.Cells(8, 4).Value = CStr(Round(Abs(dStackupFornutplateMin / 25.4), 3)) + "in"
    oExcelGlobal.App.Cells(9, 4).Value = CStr(Round(Abs(dStackupFornutplateMax / 25.4), 3)) + "in"
    
    '***Lock the sheet
    oExcelGlobal.App.DisplayAlerts = False
    oExcelWorkbook.Sheets.Item("TolerancesPanelAssyCheck").Delete
    oExcelWorkbook.Sheets.Item(ANALYSISSHT).Activate
    
    On Error Resume Next
    oExcelWorkbook.Sheets.Item(ANALYSISSHT).Protect "AutomationLock", True, True
    oExcelWorkbook.Sheets.Item("TolerancesHardwareCompatibility").Protect "AutomationLock", True, True
    On Error GoTo 0
    
    oExcelWorkbook.bNoRunBeforesave = True
    sReportName = "C:\Temp\FastenerAnalysisResult_" & Format((DateTime.Timer) / 86400, "hh_mm_ss") & ".xls"
    oExcelWorkbook.SaveAs sReportName
    oExcelWorkbook.bNoRunBeforesave = False
    oExcelGlobal.App.DisplayAlerts = True
    '***Save excel file in local drive
    oExcelGlobal.App.ScreenUpdating = True
    Call oExcelGlobal.ShowExcelWindow
    '******Log File
    On Error Resume Next
        Call AddToLogFile("FastenerAssemblyAnalysis", , oSelectedAssembly.Name, , , sStatus)
    On Error GoTo 0
End Sub

Private Sub SelectAssemblyInCatia(ByVal sAssyName As String)
    Dim dGlobalMatrix As Variant
    Dim yVector(3 To 5) As Double
    Dim zVector(6 To 8) As Double
    '*** make sure that active window is the one in which tested assembly was selected
    If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    Set cSections = CATIA.ActiveDocument.Product.GetTechnologicalObject("Sections")
    
    If cSections.Count > 0 Then
        For i = 1 To cSections.Count
            If cSections.Item(i).Name = "Automation_001Section" Then cSections.Remove i
        Next
    End If
    Set oMasterSection = cSections.Add
    oMasterSection.Type = catSectionTypePlane
    oMasterSection.CutMode = 1
    oMasterSection.Name = "Automation_001Section"
    oMasterSection.Width = 50
    oMasterSection.Height = ListOfFastenersAssemblies.Item(sAssyName).Item(1).Length * 2
    '*** creating global matrix for section view
    dGlobalMatrix = ListOfFastenersAssemblies.Item(sAssyName).Item(1).GlobalPositionMatrix
    '*** y becomes z & z becomes - y because axis is rotated about x in anticlockwisedirection
    yVector(3) = dGlobalMatrix(3): yVector(4) = dGlobalMatrix(4): yVector(5) = dGlobalMatrix(5)
    zVector(6) = dGlobalMatrix(6): zVector(7) = dGlobalMatrix(7): zVector(8) = dGlobalMatrix(8)
    
    dGlobalMatrix(3) = zVector(6): dGlobalMatrix(4) = zVector(7): dGlobalMatrix(5) = zVector(8)
    dGlobalMatrix(6) = -yVector(3): dGlobalMatrix(7) = -yVector(4): dGlobalMatrix(8) = -yVector(5)
    
    On Error Resume Next
    oMasterSection.SetPosition dGlobalMatrix
    If Err.Number <> 0 Then
         MsgBox ("Assembly cannot be found" & vbExclamation)
         cSections.Remove "Automation_001Section"
         On Error GoTo 0
         Exit Sub
    End If
    On Error GoTo 0
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutOrigin Array(dGlobalMatrix(9), dGlobalMatrix(10), dGlobalMatrix(11))
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutSightDirection Array(-yVector(3), -yVector(4), -yVector(5))
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutUpDirection Array(-zVector(6), -zVector(7), -zVector(8))
    CATIA.ActiveDocument.Selection.Clear
    CATIA.ActiveDocument.Selection.Add ListOfFastenersAssemblies.Item(sAssyName).Item(1).InstanceObject
    CATIA.StartCommand ("Reframe on")
    CATIA.ActiveDocument.Selection.Clear
    CATIA.ActiveWindow.ActiveViewer.Update
End Sub

Private Function sCheckMisAlignment(oFastAssy As Collection) As String
    '*** check misalignment between male and female fasteners
    '*** do not use it for checking alignment between female fasteners as this logic will not work
    '*** this logic assumes that fasteners are already clashing --- which is true for male & female fasteners
    
    Dim PEnd As Variant
    Dim PStart As Variant
    Dim VectorM(2) As Double
    Dim VectorF(2) As Double
    Dim oMale As clsFastener
    Dim oFeMale As clsFastener
    Dim dAngle As Double
    Dim sReturnString As String
    Dim dGap As Double
    
    sReturnString = ""
    Set oMale = oFastAssy.Item(1)
    PEnd = oFastAssy.Item(1).EndPoint
    PStart = oFastAssy.Item(1).StartPoint
    
    VectorM(0) = PEnd(0) - PStart(0): VectorM(1) = PEnd(1) - PStart(1): VectorM(2) = PEnd(2) - PStart(2)

    For i = 2 To oFastAssy.Count
        bCheck = True
        Set oFeMale = oFastAssy.Item(i)

        '*** Calculation Miminum distance between fastassy
        dGap = dDistanceLineSegToLineSeg(oMale.StartPoint, oMale.EndPoint, oFeMale.StartPoint, oFeMale.EndPoint)
        If dGap > dLinearTolerance Then bCheck = False

        '***Calculating Angle between Fastassy
        PEnd = oFeMale.EndPoint
        PStart = oFeMale.StartPoint
        VectorF(0) = PEnd(0) - PStart(0): VectorF(1) = PEnd(1) - PStart(1): VectorF(2) = PEnd(2) - PStart(2)
        dAngle = ArcCos(DotProd(VectorM, VectorF) / (LengthVector(VectorM) * LengthVector(VectorF)))
        dAngle = Abs(Round(dAngle * 180 / PI, 9))
        If dAngle > dAngularTolerance Then
            If dAngle > 180 + dAngularTolerance Or dAngle < 180 - dAngularTolerance Then bCheck = False
        End If
        If bCheck = False Then
            If sReturnString = "" Then
                sReturnString = sReturnString + "---" + Left(oFeMale.UniqueNameInAssembly, InStr(oFeMale.UniqueNameInAssembly, ";") - 1) + " is Misaligned. [Angle " + CStr(Round(dAngle, 4)) + "deg / Gap " + CStr(Round(dGap / 25.4, 4)) + "in]"
            Else
                sReturnString = sReturnString + ";" + "---" + Left(oFeMale.UniqueNameInAssembly, InStr(oFeMale.UniqueNameInAssembly, ";") - 1) + " is Misaligned. [Angle " + CStr(Round(dAngle, 4)) + "deg / Gap " + CStr(Round(dGap / 25.4, 4)) + "in]"
            End If
        End If
    Next
        sCheckMisAlignment = sReturnString
End Function
Private Function sCheckClashInsideAssy(oFastAssy As Collection) As String
    Dim oFirstFast As clsFastener
    Dim oSecondFast As clsFastener
    Dim oMaleFast As clsFastener
    
    Dim dDistSecondFastStart_ToFirstFastStartPlane As Double
    Dim dDistSecondFastEnd_ToFirstFastStartPlane As Double
    
    Dim dDistSecondFastStart_ToFirstFastEndPlane As Double
    Dim dDistSecondFastEnd_ToFirstFastEndPlane As Double

    Dim dDistSecondFastStart_ToMaleFastStartPlane As Double
    Dim dDistSecondFastEnd_ToMaleFastStartPlane As Double
    
    Dim bFastAbove As Boolean
    Dim bFastBelow As Boolean
    
    Dim bSecondFastenerInClash As Boolean
    Dim sReturnString As String

    sReturnString = ""
    '*** check clash among female fasteners
    If oFastAssy.Count > 2 Then
        For i = 2 To oFastAssy.Count - 1
            Set oFirstFast = oFastAssy.Item(i)
            For j = i + 1 To oFastAssy.Count
                Set oSecondFast = oFastAssy.Item(j)
                '*** second fastener  to first fastener start plane ****
                dDistSecondFastStart_ToFirstFastStartPlane = dDistancePointToPlane(oSecondFast.StartPoint, oFirstFast.StartPoint, oFirstFast.EndPoint)
                dDistSecondFastEnd_ToFirstFastStartPlane = dDistancePointToPlane(oSecondFast.EndPoint, oFirstFast.StartPoint, oFirstFast.EndPoint)
                '*** second fastener  to first fastener end plane ****
                dDistSecondFastStart_ToFirstFastEndPlane = dDistancePointToPlane(oSecondFast.StartPoint, oFirstFast.EndPoint, oFirstFast.StartPoint)
                dDistSecondFastEnd_ToFirstFastEndPlane = dDistancePointToPlane(oSecondFast.EndPoint, oFirstFast.EndPoint, oFirstFast.StartPoint)
                
                If dDistSecondFastStart_ToFirstFastStartPlane <= dClashTolerance And dDistSecondFastEnd_ToFirstFastStartPlane <= dClashTolerance Then bFastAbove = True Else bFastAbove = False ' ***Second fastener is completely above first one
                If dDistSecondFastStart_ToFirstFastEndPlane <= dClashTolerance And dDistSecondFastEnd_ToFirstFastEndPlane <= dClashTolerance Then bFastBelow = True Else bFastBelow = False ' ***Second fastener is completely below first one
                If bFastAbove = True Or bFastBelow = True Then bSecondFastenerInClash = False Else bSecondFastenerInClash = True '****if either of them is true then fastener is not clashing, otherwise it is
                
                If bSecondFastenerInClash Then
                    If sReturnString = "" Then
                        sReturnString = sReturnString + "---" + Left(oSecondFast.UniqueNameInAssembly, InStr(oSecondFast.UniqueNameInAssembly, ";") - 1) + " is clashing with " + Left(oFirstFast.UniqueNameInAssembly, InStr(oFirstFast.UniqueNameInAssembly, ";") - 1)
                    Else
                        sReturnString = sReturnString + ";" + "---" + Left(oSecondFast.UniqueNameInAssembly, InStr(oSecondFast.UniqueNameInAssembly, ";") - 1) + " is clashing with " + Left(oFirstFast.UniqueNameInAssembly, InStr(oFirstFast.UniqueNameInAssembly, ";") - 1)
                    End If
                End If
            Next
        Next
    End If
    '*** check clash with male fastener
    If oFastAssy.Count > 1 Then
            Set oMaleFast = oFastAssy.Item(1) '***Male Fastener
            For i = 2 To oFastAssy.Count
                Set oSecondFast = oFastAssy.Item(i)
                dDistSecondFastStart_ToMaleFastStartPlane = dDistancePointToPlane(oSecondFast.StartPoint, oMaleFast.StartPoint, oMaleFast.EndPoint)
                dDistSecondFastEnd_ToMaleFastStartPlane = dDistancePointToPlane(oSecondFast.EndPoint, oMaleFast.StartPoint, oMaleFast.EndPoint)
                
                If dDistSecondFastStart_ToMaleFastStartPlane < -dClashTolerance Or dDistSecondFastEnd_ToMaleFastStartPlane < -dClashTolerance Then ' *** if either is true then female is in clash with male
                    If sReturnString = "" Then
                        sReturnString = sReturnString + "---" + Left(oSecondFast.UniqueNameInAssembly, InStr(oSecondFast.UniqueNameInAssembly, ";") - 1) + " is clashing with " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1)
                    Else
                        sReturnString = sReturnString + ";" + "---" + Left(oSecondFast.UniqueNameInAssembly, InStr(oSecondFast.UniqueNameInAssembly, ";") - 1) + " is clashing with " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1)
                    End If
                End If
            Next

    End If
        sCheckClashInsideAssy = sReturnString
End Function
Private Function sCheckThreadofAssy(oFastAssy As Collection) As String
    '***
    '*** This check first looks for threaded female part :- there should be one & only one of it
    '*** female threaded fastener should be the last fastener of the assembly
    '*** dia of all fasteners in assy should match
    
    Dim oMaleFast As clsFastener
    Dim oFemaleFast As clsFastener
    Dim iNumberOfThreadedFemale As Integer
    Dim sReturnString As String
    Dim sCheckString As String
    
    Set oMaleFast = oFastAssy.Item(1) '***Male Fastener
    iNumberOfThreadedFemale = 0
    sReturnString = ""
    If oMaleFast.IsThreaded Then
        '*** check female threaded fasteners that are not the last one
        If oFastAssy.Count > 1 Then
            For i = 2 To oFastAssy.Count
                Set oFemaleFast = oFastAssy.Item(i)
                If oFemaleFast.IsThreaded = True Then
                iNumberOfThreadedFemale = iNumberOfThreadedFemale + 1
                    If i <> oFastAssy.Count Then
                        If sReturnString = "" Then
                            sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is Threaded & not the last fastener in assy"
                        Else
                            sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is Threaded & not the last fastener in assy"
                        End If
                    End If
                End If
                '*** female fastener can have more than one diameter below method compares male to any one of female diameter
                On Error Resume Next
                Err.Clear
                sCheckString = ConvertStringToCollection(oFemaleFast.Dia, ";", True).Item(oMaleFast.Dia)
                If Err.Number <> 0 Then
                    If sReturnString = "" Then
                        sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " Dia/Thread is not compatible with " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1)
                    Else
                        sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " Dia/Thread is not compatible with " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1)
                    End If
                End If
                On Error GoTo 0
            Next
        End If
        '*** check multiple or no female threaded fasteners
        If iNumberOfThreadedFemale > 1 Then
            If sReturnString = "" Then
                sReturnString = sReturnString + "---" + CStr(iNumberOfThreadedFemale) + " female threaded fasteners"
            Else
                sReturnString = sReturnString + ";" + "---" + CStr(iNumberOfThreadedFemale) + " female threaded fasteners"
            End If
        ElseIf iNumberOfThreadedFemale = 0 Then
            If sReturnString = "" Then
                sReturnString = sReturnString + "---" + "Female threaded fastener not found"
            Else
                sReturnString = sReturnString + ";" + "---" + "Female threaded fastener not found"
            End If
        End If
    Else '**** if male fastener is not threaded
        If oFastAssy.Count > 1 Then
            For i = 2 To oFastAssy.Count
                Set oFemaleFast = oFastAssy.Item(i)
                If oFemaleFast.IsThreaded Then
                    If sReturnString = "" Then
                        sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is Threaded whereas " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " is not"
                    Else
                        sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is Threaded whereas " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " is not"
                    End If
                End If
            Next
        End If
    End If
    
    sCheckThreadofAssy = sReturnString
End Function

Private Function sCheckCSK(oFastAssy As Collection) As String
    '***
    '*** This check first looks for male CSK, if it exist then if starting distance
    '*** CSK should be toward start point of male fastener
    '*** between male & female is zero then first female fastener should be CSK
    '*** if distance is more than dCSKTolerance  then it should be a non CSK fastener
    '***
    Dim oMaleFast As clsFastener
    Dim oFemaleFast As clsFastener
    Dim iNumberOfThreadedFemale As Integer
    Dim sReturnString As String
    Dim dDistance As Double
    
    Set oMaleFast = oFastAssy.Item(1) '***Male Fastener
    iNumberOfCSKFemale = 0
    sReturnString = ""
    
    If oFastAssy.Count > 1 Then
        If oMaleFast.IsCounterSunk Then
            For i = 2 To oFastAssy.Count
                Set oFemaleFast = oFastAssy.Item(i)
                '*** check that only first female fastener is countersunk
                If oFemaleFast.IsCounterSunk Then
                    '*** check if fastener is inverted
                    If dDistancePointToPlane(oFemaleFast.StartPoint, oMaleFast.StartPoint, oMaleFast.EndPoint) > dDistancePointToPlane(oFemaleFast.EndPoint, oMaleFast.StartPoint, oMaleFast.EndPoint) Then
                        If sReturnString = "" Then
                            sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is inverted"
                        Else
                            sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " is inverted"
                        End If
                    End If
                    If i <> 2 Then
                        If sReturnString = "" Then
                            sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk"
                        Else
                            sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk "
                        End If
                    End If
                End If
            Next
            Set oFemaleFast = oFastAssy.Item(2) '*** first female fastener should be countersunk only if distance between it & male fastener is less than 3 mm
            dDistance = dDistancePointToPlane(oFemaleFast.StartPoint, oMaleFast.StartPoint, oMaleFast.EndPoint)
            If dDistance < dCSKTolerance And oFemaleFast.IsCounterSunk = False Then
                If sReturnString = "" Then
                    sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should be countersunk & " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " seems misaligned"
                Else
                    sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should be countersunk & " + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " seems misaligned"
                End If
            ElseIf dDistance > dCSKTolerance And oFemaleFast.IsCounterSunk Then
                If sReturnString = "" Then
                    sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk"
                Else
                    sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk"
                End If
            End If
        Else '*** check that none of the female fasteners are countersunk
        For i = 2 To oFastAssy.Count
            Set oFemaleFast = oFastAssy.Item(i)
                If oFemaleFast.IsCounterSunk Then
                    If sReturnString = "" Then
                        sReturnString = sReturnString + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk"
                    Else
                        sReturnString = sReturnString + ";" + "---" + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1) + " should not be countersunk "
                    End If
            End If
        Next
        End If
    End If
    sCheckCSK = sReturnString
End Function
Private Function sCheckLastFastenerOfAssy(oFastAssy As Collection) As String
    Dim oMaleFast As clsFastener
    Dim oLastFast As clsFastener
    
    Dim dDistFemaleFastStart_ToMaleFastEndPlane As Double
    Dim dDistFemaleFastEnd_ToMaleFastEndPlane As Double
    
    Dim dDistFemaleFastStart_ToMaleFastStartPlane As Double
    Dim dDistFemaleFastEnd_ToMaleFastStartPlane As Double
    
    Dim bFastAbove As Boolean
    Dim bFastBelow As Boolean
    Dim bFastenerInverted As Boolean
    Dim bTooLong As Boolean
    Dim bTooShort As Boolean
    Dim dMinPenetration As Double
    
    Dim dBlindInsertSTK As Double
    Dim dThruInsertSTK As Double
    Dim dNutplateSTKMin As Double
    Dim dNutplateSTKMax As Double
    
    If oFastAssy.Count > 1 Then
        Set oMaleFast = oFastAssy.Item(1)    '***Male Fastener
        Set oLastFast = oFastAssy.Item(oFastAssy.Count)
        dMinPenetration = oLastFast.Length * dRatioForMaleFastenerTooShort
        '*** save values to temp parameters
        dBlindInsertSTK = dStackupForBlindInsert
        dThruInsertSTK = dStackupForThruInsert
        dNutplateSTKMin = dStackupFornutplateMin
        dNutplateSTKMax = dStackupFornutplateMax
        '*** if fastener has specific value defined then use it, else use global values
        If oLastFast.StackupOffset <> USEGLOBALVALUE Then
            dStackupForBlindInsert = -oLastFast.StackupOffset
            dStackupForThruInsert = -oLastFast.StackupOffset
            dStackupFornutplateMin = oLastFast.StackupOffsetMin
            dStackupFornutplateMax = oLastFast.StackupOffset
        End If
        If oLastFast.StackupOffsetMin <> USEGLOBALVALUE Then dMinPenetration = oLastFast.Length - oLastFast.StackupOffsetMin
        
        '*** check if fastener is inverted
        dDistFemaleFastStart_ToMaleFastStartPlane = dDistancePointToPlane(oLastFast.StartPoint, oMaleFast.StartPoint, oMaleFast.EndPoint)
        dDistFemaleFastEnd_ToMaleFastStartPlane = dDistancePointToPlane(oLastFast.EndPoint, oMaleFast.StartPoint, oMaleFast.EndPoint)
        
        If dDistFemaleFastEnd_ToMaleFastStartPlane < dDistFemaleFastStart_ToMaleFastStartPlane Then bFastenerInverted = True Else bFastenerInverted = False
        '*** if nutplate or blind fastener is inverted then flag it
        If bFastenerInverted Then
            If UCase(oLastFast.FastenerType) = "NUTPLATE" Or oLastFast.IsBlind Then
                If sReturnString = "" Then
                    sReturnString = sReturnString + "---" + Left(oLastFast.UniqueNameInAssembly, InStr(oLastFast.UniqueNameInAssembly, ";") - 1) + " is Inverted"
                Else
                    sReturnString = sReturnString + ";" + "---" + Left(oLastFast.UniqueNameInAssembly, InStr(oLastFast.UniqueNameInAssembly, ";") - 1) + " is Inverted"
                End If
            End If
        End If
        
        '*** check intersection with male fastener
        dDistFemaleFastStart_ToMaleFastEndPlane = dDistancePointToPlane(oLastFast.StartPoint, oMaleFast.EndPoint, oMaleFast.StartPoint)
        dDistFemaleFastEnd_ToMaleFastEndPlane = dDistancePointToPlane(oLastFast.EndPoint, oMaleFast.EndPoint, oMaleFast.StartPoint)
    
        '*** check if fastener is too long
        '*** check thru insert & trisert
        If UCase(oLastFast.FastenerType) = "TRISERT" Then
            If bFastenerInverted = False Then
                If dDistFemaleFastEnd_ToMaleFastEndPlane < dStackupForThruInsert Then bTooLong = False Else bTooLong = True
            Else
                If dDistFemaleFastStart_ToMaleFastEndPlane < dStackupForThruInsert Then bTooLong = False Else bTooLong = True
            End If
        End If
        
        If UCase(oLastFast.FastenerType) = "INSERT" And oLastFast.IsBlind = False Then
            If bFastenerInverted = False Then
                If dDistFemaleFastEnd_ToMaleFastEndPlane < dStackupForThruInsert Then bTooLong = False Else bTooLong = True
            Else
                If dDistFemaleFastStart_ToMaleFastEndPlane < dStackupForThruInsert Then bTooLong = False Else bTooLong = True
            End If
        End If
        '*** check Blind insert
        If UCase(oLastFast.FastenerType) = "INSERT" And oLastFast.IsBlind Then
            If bFastenerInverted = False Then
                If dDistFemaleFastEnd_ToMaleFastEndPlane < dStackupForBlindInsert Then bTooLong = False Else bTooLong = True
            End If
        End If
        '***check for other types , nuts, nutplates
        '*** it checks only bTooLong for trisert, bTooShort is set by next if statement
        If oLastFast.IsThreaded And (UCase(oLastFast.FastenerType) = "NUT") And oLastFast.IsBlind = False Or (UCase(oLastFast.FastenerType) = "NUTPLATE") Then
            If bFastenerInverted = False Then
                If dDistFemaleFastEnd_ToMaleFastEndPlane > dStackupFornutplateMax Then bTooLong = True Else bTooLong = False
                If dDistFemaleFastEnd_ToMaleFastEndPlane < dStackupFornutplateMin Then bTooShort = True Else bTooShort = False
            Else
                If dDistFemaleFastStart_ToMaleFastEndPlane > dStackupFornutplateMax Then bTooLong = True Else bTooLong = False
                If dDistFemaleFastStart_ToMaleFastEndPlane < dStackupFornutplateMin Then bTooShort = True Else bTooShort = False
            End If
        End If
        '*** check if fastener is too short for inserts and helicoil
        '*** if overlap is less than dRatioForMaleFastenerTooShort * length of female fastener then it is male fastener is too short
        If UCase(oLastFast.FastenerType) = "INSERT" Or UCase(oLastFast.FastenerType) = "HELICOIL" Or UCase(oLastFast.FastenerType) = "TRISERT" Then
            If bFastenerInverted = False Then
                If dDistFemaleFastStart_ToMaleFastEndPlane < dMinPenetration Then bTooShort = True Else bTooShort = False
            Else
                If dDistFemaleFastEnd_ToMaleFastEndPlane < dMinPenetration Then bTooShort = True Else bTooShort = False
            End If
        End If
        
        If bTooShort Then
            If sReturnString = "" Then
                sReturnString = sReturnString + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " is too short"
            Else
                sReturnString = sReturnString + ";" + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " too short"
            End If
        End If
        
        If bTooLong Then
            If sReturnString = "" Then
                sReturnString = sReturnString + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " is too long"
            Else
                sReturnString = sReturnString + ";" + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " too long"
            End If
        End If
        sCheckLastFastenerOfAssy = sReturnString
        
        '*** Reset Global Values
        dStackupForBlindInsert = dBlindInsertSTK
        dStackupForThruInsert = dThruInsertSTK
        dStackupFornutplateMin = dNutplateSTKMin
        dStackupFornutplateMax = dNutplateSTKMax
    End If
End Function
Private Function sCheckShankInThreadedFemaleFastener(oFastAssy As Collection) As String
    Dim oMaleFast As clsFastener
    Dim oFemaleFast As clsFastener
        
    Dim dDistFemaleFastStart_ToMaleFastThreadPlane As Double
    Dim dDistFemaleFastEnd_ToMaleFastThreadPlane As Double
    
    If oFastAssy.Count > 1 Then
        Set oMaleFast = oFastAssy.Item(1)    '***Male Fastener
        If oMaleFast.HasGripLength Then
            For i = 2 To oFastAssy.Count
                Set oFemaleFast = oFastAssy.Item(i)
                If oFemaleFast.IsThreaded Then
                    If oFemaleFast.HasGripLength Then
                        dDistFemaleFastStart_ToMaleFastThreadPlane = dDistancePointToPlane(oFemaleFast.ThreadPoint, oMaleFast.ThreadPoint, oMaleFast.EndPoint)
                    Else
                        dDistFemaleFastStart_ToMaleFastThreadPlane = dDistancePointToPlane(oFemaleFast.StartPoint, oMaleFast.ThreadPoint, oMaleFast.EndPoint)
                    End If
                    dDistFemaleFastEnd_ToMaleFastThreadPlane = dDistancePointToPlane(oFemaleFast.EndPoint, oMaleFast.ThreadPoint, oMaleFast.EndPoint)
                    If dDistFemaleFastStart_ToMaleFastThreadPlane < SMALLNUMBER Or dDistFemaleFastEnd_ToMaleFastThreadPlane < SMALLNUMBER Then
                        If sReturnString = "" Then
                            sReturnString = sReturnString + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " shank in " + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1)
                        Else
                            sReturnString = sReturnString + ";" + "---" + Left(oMaleFast.UniqueNameInAssembly, InStr(oMaleFast.UniqueNameInAssembly, ";") - 1) + " shank in " + Left(oFemaleFast.UniqueNameInAssembly, InStr(oFemaleFast.UniqueNameInAssembly, ";") - 1)
                        End If
                    End If
                End If
            Next
        End If
    End If
    sCheckShankInThreadedFemaleFastener = sReturnString
End Function
Private Function bCheckProvisionalFasteners(ByRef oInsert As clsFastener) As Boolean

'*** this method strives to find provisional inserts... its algorithm is bit fishy
'*** if all instance of conflicts are design parts ("Material Specification Production"<>"") or skeleton and atleast one is a panel
'*** then insert is a provisional fastener

    Dim oConflict As Conflict
    Dim cConflicts As Conflicts
    
    
    
    bCheckProvisionalFasteners = False

    oFirstGroup.AddExplicit oInsert.InstanceObject
    oClash.Compute
    Dim bPanelFound As Boolean: bPanelFound = False
    Dim bDesignPart As Boolean: bNonDesignPart = False
    Set cConflicts = oClash.Conflicts
    If oInsert.FastenerType = "INSERT" And cConflicts.Count > 0 Then
        bCheckProvisionalFasteners = False
        For Each oConflict In oClash.Conflicts
        
            Dim sMatSpec As String: sMatSpec = ""
            Dim bLoaded As Boolean: bLoaded = False
            Dim oInst As Product
            
            Set oInst = oConflict.secondProduct
            
            On Error Resume Next
            Set oTest = oInst.ReferenceProduct
            If Err.Number <> 0 Then oInst.ApplyWorkMode DEFAULT_MODE
            Err.Clear
            
            Set oTest = oInst.ReferenceProduct  '*** if instance cannot be loaded
            If Err.Number = 0 Then
                bLoaded = True
                Dim sValue As String: sValue = ""
                sValue = oInst.ReferenceProduct.UserRefProperties.Item("Material Specification Production").Value
                If sValue = "-" Or sValue = "N/A" Then sValue = ""
                sMatSpec = Trim(sValue)
                If sMatSpec = "" Then ' Check for Skeleton part and skip it
                    sValue = oInst.ReferenceProduct.UserRefProperties.Item("Dataset Type").Value
                    If sValue = "SKELETON" Then sMatSpec = "NOTBLANK"
                End If
            End If
            On Error GoTo 0
            
            If sMatSpec <> "" And bLoaded And Not bPanelFound Then
                bPanelFound = sGetPanelAttribute(oInst, "Material Form") = "SANDWICH PANEL"
            End If
            bNonDesignPart = sMatSpec = ""
            If bNonDesignPart Then Exit For
        Next
    End If
    bCheckProvisionalFasteners = bPanelFound And (Not bNonDesignPart)
    
    oFirstGroup.RemoveExplicit 1
End Function
Private Sub SwapToAssemblyWorkbench()
    Dim sWkb As String
    
    sWkb = CATIA.GetWorkbenchId
    If sWkb <> "Assembly" Then
        CATIA.StartWorkbench "Assembly"
    End If
End Sub
Private Function sSearchStringForExcel(oInstance As Product) As String
    Dim sParentList As String
    
    sParentList = BuildParentList(oInstance)
    bCheck = False
    
    For i = 1 To ListOfFastenersAssemblies.Count
        For Each oFast In ListOfFastenersAssemblies.Item(i)
            If sParentList = oFast.UniqueNameInAssembly Then
                bCheck = True
                sSearchStringForExcel = ListOfFastenersAssemblies.Item(i).Item(1).UniqueNameInAssembly + "_Fastassy"
                Exit Function
            End If
        Next
    Next
    If Not (bCheck) Then sSearchStringForExcel = ""
End Function
