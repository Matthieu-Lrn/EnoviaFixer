Attribute VB_Name = "mdlPanelAssyCheck"
Option Explicit
'*******************************************
'       For WinAPIs
'*******************************************
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private ParentForm_hwnd As Long
Private Child_hwnd As Long
Private sChildCLASS As String
Private sChildSTRING As String

Private Const WM_GETTEXT = &HD
Private Const GW_HWNDNEXT = 2
'*******************************************
Public Enum EnumPanelThicknessStatus
    OK = 0
    ThicknessUndefined = 1
    NonFlatPanel = 2
    NonStandardPanelThickness = 4
End Enum

Public Enum EnumPanelAssyStatus
    OK
    SPCollectorNotFound
    BaseNumberError
End Enum

Public sImportedInsertHoleTable As Variant

Public dMiminumGapAllowed As Double
Public dMinGapTongue As Double                 '*** insert center to Tongue edge
Public dFlushnessTol As Double                 '*** 1 thou tolerance
Public dMinMisalignmentTol As Double           '*** 1 thou tolerance
Public dLTonguewidthmin As Double              '*** min width of std L tongue in T & L  tongue & groove joint
Public dLTonguewidthmax As Double              '*** max width of std L tongue in T & L  tongue & groove joint
Public dTTonguewidthmin As Double              '*** min width of std T tongue in T & L  tongue & groove joint
Public dTTonguewidthmax As Double              '*** max width of std T tongue in T & L  tongue & groove joint
Public sInsertdistanceCol As New Collection
'***
Public Const sMsgStressIncompatible As String = "Insert is not compatible with the panel as per stress requirements"    ': update 28Jan2018
'***
Private oSelectedAssembly As Product
Private ListOfAllPanels As Collection
Private oExcel As clsExcel
Sub ScanPanelAssies(ByVal sdummy As String)

    Dim i As Integer
    Dim oPanelAna As clsPanelCheck
    Dim oWind As Window
    
    Call selectAssy(oSelectedAssembly, CATIA.ActiveDocument, , "Please select an Assembly to validate!   ")
    If oSelectedAssembly Is Nothing Then Exit Sub
        
    '***** setting global parameters
    dMiminumGapAllowed = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "MinimumDIstance_mm"))
    dMinGapTongue = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "MinimumDistanceTongue_mm"))
    dFlushnessTol = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "FlushnessTolerance_mm"))
    dMinMisalignmentTol = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "MinMisalignmentTolerance_mm"))
    dLTonguewidthmin = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "LTongueWidthMin_mm"))
    dLTonguewidthmax = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "LTongueWidthMax_mm"))
    dTTonguewidthmin = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "TTonguewidthMin_mm"))
    dTTonguewidthmax = CDbl(GetTagValueFromFile(sPanelAssyParametersFile, "TTonguewidthMax_mm"))
    Set sInsertdistanceCol = ConvertStringToCollection(GetTagValueFromFile(sPanelAssyParametersFile, "InsertDistance_CenterToCenter_mm"), "|")
    '*****
    '***** Reading Insert Hole Dia Table
    Call oImportTablefromExcel.LaunchImport(sImportedInsertHoleTable, sInsertHoleDefinition, "G7000")

    Set oWind = CATIA.ActiveWindow
    '*****
    timeStart = DateTime.Timer
    Set ListOfAllPanels = New Collection
    If Not oSelectedAssembly Is Nothing Then
        Call frmProgress.progressBarInitialize("Panel Assy Check")
        Call GenerateItemlist(oSelectedAssembly)
    End If
    i = 1
    If ListOfAllPanels.Count > 0 Then
    For Each oPanelAna In ListOfAllPanels
        If oPanelAna.PanelAssyStatus = EnumPanelAssyStatus.OK And oPanelAna.PanelInstanceStatus = "" And oPanelAna.PanelThicknessStatus = EnumPanelThicknessStatus.OK Then
            Call frmProgress.progressBarRepaint("Analysing Panel assies", ListOfAllPanels.Count, i, , , , "Elapsed time: " & Format((DateTime.Timer - timeStart) / 86400, "hh:mm:ss"))
            Call oCheckPanelAssy.LaunchAnalysis(oPanelAna, CATIA.ActiveWindow, ListOfAllPanels.Count, i)
        End If
        oPanelAna.DeleteThicknessGeo    ' Bug correction GSUP-2176
        i = i + 1
    Next
    Else
        MsgBox "No Panel found."
        Exit Sub
    End If
    If bCancelAction Then
        Exit Sub
    End If
    
    Call ExtractReportToExcel(ListOfAllPanels)
    Call frmDisplayInCatia.LaunchDisplayInCatia(oWind, oExcel)
    frmProgress.Hide
End Sub
Private Sub GenerateItemlist(ByVal oAssembly As Product, _
                             Optional sListName As String = "", _
                             Optional sRefListName As String = "")
'*** This functions scans everthing underneath oAssembly & populates
'*** ListofallPanels collection

    Dim oInstance As Product
    Dim oPanelassyInstance As Product
    Dim sParentList As String
    Dim sRefParentList As String
    Dim sName As String
    Dim iErr As Integer
    Dim bCheck As Boolean
    Dim oPanelThicknessCol As Collection
    Dim oPanelCheckObject As clsPanelCheck
    
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
            If Not (sParentList Like "*SP##*") Then Call GenerateItemlist(oInstance, sParentList, sRefParentList) ' recursion to scan all sub branches
        Else
            sRefParentList = oInstance.Name & ";" & sRefListName
            sParentList = oInstance.Name & ";" & sListName
            bCheck = False
            
            On Error Resume Next
            sName = oInstance.ReferenceProduct.Name
            If Err.Number <> 0 Then
                oInstance.ApplyWorkMode DEFAULT_MODE
            End If
            Err.Clear
            sName = oInstance.ReferenceProduct.Name
            iErr = Err.Number
            On Error Resume Next
            If iErr <> 0 Then GoTo ExtFcnt
            bCheck = sCheckPanelAttribute(oInstance)

            If bCheck = True Then
                Err.Clear
                Set oPanelCheckObject = New clsPanelCheck
                On Error Resume Next
                ListOfAllPanels.Add oPanelCheckObject, sParentList
                If Err.Number <> 0 Then
                    oPanelCheckObject.PanelInstanceStatus = "Duplicate Instance Name " & Right(sParentList, Len(sParentList) - InStrRev(sParentList, ";")) & "In " & Split(sParentList, ";")(UBound(Split(sParentList, ";")) - 1)
                    GoTo ExtFcnt
                End If
                On Error GoTo 0
                Set oPanelCheckObject.PanelInstance = oInstance
                '**** Get material of the panel: update 28Jan2018
                oPanelCheckObject.GetMaterial
                '**** Update 20March2018
                Dim i As Integer
                Dim dOffsetValue As Double
                For i = 0 To 2
                    dOffsetValue = Array(6.25, 14, 3.25)(i)
                    Call GetPanelThickness(oInstance.ReferenceProduct, oInstance.ReferenceProduct.Parent.Part, oPanelCheckObject, dOffsetValue)
                    If oPanelCheckObject.PanelThicknessStatus <> EnumPanelThicknessStatus.OK Then
                        oPanelCheckObject.DeleteThicknessGeo
                    Else
                        Exit For
                    End If
                Next
                If oPanelCheckObject.PanelThicknessStatus = EnumPanelThicknessStatus.OK Then
                    oPanelCheckObject.PanelThicknessStatus = CheckPanelThickness(oPanelCheckObject.PanelThickness)
                End If
                Set oPanelCheckObject.PanelAssyInstance = oInstance.Parent.Parent
                oPanelCheckObject.PanelAssyStatus = CheckPanelAssyStatus(oPanelCheckObject.PanelAssyInstance, oInstance)
            End If
        End If
    Next
ExtFcnt:
End Sub
Private Function sCheckPanelAttribute(oPanelInstance As Product) As Boolean
    '*** check for attribute 15 which is material form
    Dim bCheck As String
    Dim sString As String
    If sGetPanelAttribute(oPanelInstance, "Material Form") = "SANDWICH PANEL" Then
        sCheckPanelAttribute = True
    End If
'    On Error Resume Next
'    sString = oPanelInstance.ReferenceProduct.UserRefProperties.Item("Material Form").Value
'    On Error GoTo 0
'    If sString = "SANDWICH PANEL" Then
'        sCheckPanelAttribute = True
'    Else
'        Dim sMakeFromPartnumber As String: sMakeFromPartnumber = ""
'        On Error Resume Next
'        sMakeFromPartnumber = oPanelInstance.ReferenceProduct.UserRefProperties.Item("Eng. Make From").Value
'        On Error GoTo 0
'        If sMakeFromPartnumber <> "" Then
'            Dim sAttributes As Variant
'            Call GetAttributesOfLatestDocRevision(sAttributes, sMakeFromPartnumber)
'            If Not IsEmpty(sAttributes) Then
'             If sAttributes.Count > 0 Then
'                If sAttributes("Material Form") = "SANDWICH PANEL" Then sCheckPanelAttribute = True
'             End If
'            End If
'        End If
'    End If
End Function
Public Function sGetPanelAttribute(oPanelInstance As Product, ByVal sAttribute As String) As String
    '**** Get Attribute from refrenceProduct however
    '**** if Eng. Make From is not blank then attribute is extracted from engg make from part
    Static oPanelAttCollection As New clsCollection
    Dim sPartNumber As String: sPartNumber = ""
    Dim sAttributes As Variant
    Dim sReturn As String: sReturn = ""
        
    On Error Resume Next
    If oPanelAttCollection.Exists(oPanelInstance.ReferenceProduct.Name) Then
        sAttributes = oPanelAttCollection.GetItemByKey(oPanelInstance.ReferenceProduct.Name)
        sReturn = sAttributes(sAttribute)
    End If
    On Error GoTo 0
    If sReturn = "" Then
        Dim sMakeFromPartnumber As String: sMakeFromPartnumber = ""
        On Error Resume Next
        sMakeFromPartnumber = oPanelInstance.ReferenceProduct.UserRefProperties.Item("Eng. Make From").Value
        On Error GoTo 0
        If sMakeFromPartnumber <> "" Then
            
            Call GetAttributesOfDocRevision(sAttributes, sMakeFromPartnumber)
            If Not IsEmpty(sAttributes) Then
                If sAttributes.Count > 0 Then
                    '*** add to static clscollection
                    If Not oPanelAttCollection.Exists(oPanelInstance.ReferenceProduct.Name) Then oPanelAttCollection.Add oPanelInstance.ReferenceProduct.Name, sAttributes
                    '***
                    sReturn = sAttributes(sAttribute)
                End If
            End If
        Else
            On Error Resume Next
            sReturn = oPanelInstance.ReferenceProduct.UserRefProperties.Item(sAttribute).Value
            On Error GoTo 0
        End If
    End If
    sGetPanelAttribute = sReturn
End Function
Public Function CheckPanelAssyStatus(ByRef oAssy As Product, ByRef oPanelInstance As Product, Optional ByRef oSPColInstance As Product) As EnumPanelAssyStatus
    Dim eStatus As EnumPanelAssyStatus
    Dim oPartBody As Body
    Dim iCounter As Integer
    Dim oInstance As Product
    Dim oTest As AnyObject
    eStatus = EnumPanelAssyStatus.SPCollectorNotFound
    '***Panel assy should have atleast one SP collector
    For Each oInstance In oAssy.Products
        On Error Resume Next
        Err.Clear
        Set oTest = oInstance.ReferenceProduct
        If Err.Number <> 0 Then oInstance.ApplyWorkMode DEFAULT_MODE
        Err.Clear
        Set oTest = oInstance.ReferenceProduct  '*** if instance cannot be loaded
        If Err.Number = 0 Then
            On Error GoTo 0
            If oInstance.ReferenceProduct.Name Like "*SP##" Or oInstance.ReferenceProduct.Name Like "*SP###" Then
                Set oSPColInstance = oInstance
                eStatus = EnumPanelAssyStatus.OK
            End If
        End If
        On Error GoTo 0
    Next
    If Split(oAssy.PartNumber, "-")(0) <> Split(oPanelInstance.PartNumber, "-")(0) Then
        eStatus = EnumPanelAssyStatus.BaseNumberError
    End If
    CheckPanelAssyStatus = eStatus
End Function
Public Function CheckPanelThickness(ByVal dPanelThickness As Double) As EnumPanelThicknessStatus
    '*** This sub compares panelthickness to acceptable panel thickness
    '*** if thickness is not standard then it clears selection of first & second face
    
    Dim PanelThicknesCol As Collection
    Dim sThk As String
    Dim bCheckThk As Boolean
    Dim i As Integer
    
    Set PanelThicknesCol = New Collection
    Set PanelThicknesCol = ConvertStringToCollection(GetTagValueFromFile(sPanelAssyParametersFile, "Allowed_Panel_Thickness_mm"), "|")
    '***SPAWorkbench for Measurable object
    
    '*** compares thickness found with allowed thicknesses & if found then sets bCheck to true
    bCheckThk = False
    For i = 1 To PanelThicknesCol.Count
        If PanelThicknesCol.Item(i) = CStr(dPanelThickness) Then bCheckThk = True
    Next
    If bCheckThk Then
        CheckPanelThickness = EnumPanelThicknessStatus.OK
    Else
        CheckPanelThickness = EnumPanelThicknessStatus.NonStandardPanelThickness
    End If
End Function
'Modified 20March2018
Public Sub GetPanelThickness(oRefProduct As ProductStructureTypeLib.Product, _
                             oPanelPart As MECMOD.Part, _
                             ByRef oPanelCheck As clsPanelCheck, _
                             Optional dOffsetVal As Double = 12.25)
    '**** Output in form of Collection
    '****   Item 1= Error Number
    '****   Item 2= Flat Part Thickness in mm ----If calculated

    '****
    '**** Error Number description is as follows:
    '**** iErrNumber = 0 No error
    '**** iErrNumber = 1 Cannot measure thickness.
    '**** iErrNumber = 2 Non Flat Part.
                            
    Dim oPartBodyRef As INFITF.Reference
    Dim oPointOut As HybridShapeTypeLib.Point
    Dim oPlane As HybridShapeTypeLib.Plane
    Dim oMeasureable 'As Measurable
    Dim oMeasurePlane 'As Measurable
    Dim oExtr1 As HybridShapeExtremum
    Dim oExtr2 As HybridShapeExtremum
    Dim oExtr3 As HybridShapeExtremum
    Dim oExtr4 As HybridShapeExtremum

    Dim oSPAWorkbench As INFITF.Workbench
    Dim oHSF As HybridShapeTypeLib.HybridShapeFactory
    Dim oAutomationGeo As MECMOD.HybridBody
    Dim oCloseSurface As HybridShapeTypeLib.HybridShapeExtract 'Closed surface representation of solid
    
    Dim oInertia     'As Inertia
    Dim Comp(8)      'Value in mm
    Dim COG(2)       'Value in Meters
    Dim oInertiaAxis 'As AxisSystem
    Dim dInertiaPosMatrix As Variant
    Dim dVector(9 To 11) As Double
    Dim dVectorReturn(9 To 11) As Double
    Dim oSel As Selection
    Dim dDistancePts(8) ' As Variant
    Dim dGap 'As Double
    Dim dAngle 'As Double
    Dim dPMaxLength As Double   'Distance of Measured extrenum from COG
    Dim dOffset As Double
    Dim dPartPosMatrix(11) As Variant
    Dim dPMaxInIMS As Variant '*** To Store value of Extrenum in Inertia Position Matrix System
    Dim cMeasurementCol As New Collection
    Dim iErrNumber As Integer
    Dim iCount As Integer
    Dim dPanelThickness As Double
    Dim cExtrCol As New Collection
    Dim dBoxMin(1) As Double
    Dim dBoxMax(1) As Double
    Dim oOutput As New Collection
    Dim dL1 As Double
    Dim dL2 As Double
    Dim TOLERANCE As Double
    Dim Plane1(8), Plane2(8)
    TOLERANCE = 0.001
    
    iErrNumber = 0
    
    dPanelThickness = 0
    dPartPosMatrix(0) = 1:  dPartPosMatrix(1) = 0:  dPartPosMatrix(2) = 0
    dPartPosMatrix(3) = 0:  dPartPosMatrix(4) = 1:  dPartPosMatrix(5) = 0
    dPartPosMatrix(6) = 0:  dPartPosMatrix(7) = 0:  dPartPosMatrix(8) = 1
    dPartPosMatrix(9) = 0: dPartPosMatrix(10) = 0: dPartPosMatrix(11) = 0
    '*** To avoid bug in catia
    '*** if a catpart is open in another window and we try to .GetTechnologicalObject("Inertia") then it switches windows
    '*** Hence if refprod is open in another window then it is closed.
    Dim oCurWindow As Window
    Set oCurWindow = CATIA.ActiveWindow
    For iCount = 1 To CATIA.Documents.Count
        If CATIA.Documents.Item(iCount).FullName = oRefProduct.Parent.FullName Then
            CATIA.DisplayFileAlerts = False
            CATIA.Documents.Item(iCount).Close
            DoEvents
            CATIA.DisplayFileAlerts = True
        End If
    Next
    If Not CATIA.ActiveWindow Is oCurWindow Then oCurWindow.Activate
    DoEvents
    
    Set oInertia = oRefProduct.GetTechnologicalObject("Inertia")
    Set oPartBodyRef = oPanelPart.CreateReferenceFromObject(oPanelPart.MainBody)
    
    Set oHSF = oPanelPart.HybridShapeFactory
    Set oSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    
    Set oAutomationGeo = oPanelPart.HybridBodies.Add
    oAutomationGeo.Name = "AutomationTemp_To_Delete"
    Call SwapToAssemblyWorkbench '*** to ensure that geoset is hidden & not instance
    Set oSel = oRefProduct.Parent.Selection
    oSel.Clear
    oSel.Add oAutomationGeo
    oSel.VisProperties.SetShow catVisPropertyNoShowAttr
    
    Set oCloseSurface = oHSF.AddNewExtract(oPartBodyRef)
    oAutomationGeo.AppendHybridShape oCloseSurface
    oCloseSurface.Compute
    Set oMeasureable = oSPAWorkbench.GetMeasurable(oPanelPart.CreateReferenceFromObject(oCloseSurface))
    '***Create a temporary Point
    Set oPointOut = oHSF.AddNewPointCoord(0, 0, 0)
    oAutomationGeo.AppendHybridShape oPointOut
    oPointOut.Name = "PTemp"
    oPointOut.Compute
    
    oInertia.GetCOGPosition COG
    oInertia.GetPrincipalAxes Comp
    '*** Position Matrix of Inertia w.r.t Part Matrix
    dInertiaPosMatrix = Array(Comp(0), Comp(3), Comp(6), _
                              Comp(1), Comp(4), Comp(7), _
                              Comp(2), Comp(5), Comp(8), _
                              COG(0) * 1000, COG(1) * 1000, COG(2) * 1000)
    
    '*** Far away point in ZX direction in Inertia coordinate system
    dVector(9) = 25.4 * 100: dVector(10) = 0: dVector(11) = 25.4 * 5
    '*** to convert Coordinate System of COG to Coordinate system of Part Origin
    Call MultiplyVector(dVector, dInertiaPosMatrix, dVectorReturn, True) ' Find Position of Coordinate w.r.t Partorigin
    oPointOut.X.Value = dVectorReturn(9): oPointOut.Y.Value = dVectorReturn(10): oPointOut.Z.Value = dVectorReturn(11)
    oPointOut.Compute
    dGap = oMeasureable.GetMinimumDistancePoints(oPanelPart.CreateReferenceFromObject(oPointOut), dDistancePts)
    '*** To Conver Global Coordinate value to value in COG coordinate system
    dPMaxInIMS = dValueOfPointInLocalCoordSystem(dPartPosMatrix, dInertiaPosMatrix, dDistancePts)
    '*** verify that Extrenum is in positive direction of X axis of Inertia matrix if not then recreate the extrenum in -ve X direction
    If dPMaxInIMS(0) < 0 Then
        '*** dVector is in Inertia coordinate system
        dVector(9) = -25.4 * 100: dVector(10) = 0: dVector(11) = 25.4 * 5
        '*** to convert Coordinate System of COG to Coordinate system of Part Origin
        Call MultiplyVector(dVector, dInertiaPosMatrix, dVectorReturn, True) ' Find Position of Coordinate w.r.t Partorigin
        oPointOut.X.Value = dVectorReturn(9): oPointOut.Y.Value = dVectorReturn(10): oPointOut.Z.Value = dVectorReturn(11)
        oPointOut.Compute
        dGap = oMeasureable.GetMinimumDistancePoints(oPanelPart.CreateReferenceFromObject(oPointOut), dDistancePts)
        dPMaxInIMS = dValueOfPointInLocalCoordSystem(dPartPosMatrix, dInertiaPosMatrix, dDistancePts)
    End If
    
    dPMaxLength = LengthVector(dPMaxInIMS)
    dOffset = dOffsetVal
    '*** Create a Point at 0.125in and 1 in away from PMAX on PMaxInIMS Positionvector and add 5in to their Z axis value
    If dPMaxLength > 25.4 Then
        Call PopulatePointCollection(cMeasurementCol, dPMaxInIMS, dPMaxLength, dOffset, _
                                    dPartPosMatrix, dInertiaPosMatrix, oMeasureable, oPointOut, oPanelPart, oHSF, oAutomationGeo)
        dOffset = 2 * dOffsetVal
        Call PopulatePointCollection(cMeasurementCol, dPMaxInIMS, dPMaxLength, dOffset, _
                                    dPartPosMatrix, dInertiaPosMatrix, oMeasureable, oPointOut, oPanelPart, oHSF, oAutomationGeo)
    End If
    '*** Far away point in ZY direction in Inertia coordinate system
    dVector(9) = 0: dVector(10) = 25.4 * 100: dVector(11) = 25.4 * 5
    '*** to convert Coordinate System of COG to Coordinate system of Part Origin
    Call MultiplyVector(dVector, dInertiaPosMatrix, dVectorReturn, True) ' Find Position of Coordinate w.r.t Partorigin
    oPointOut.X.Value = dVectorReturn(9): oPointOut.Y.Value = dVectorReturn(10): oPointOut.Z.Value = dVectorReturn(11)
    oPointOut.Compute
    dGap = oMeasureable.GetMinimumDistancePoints(oPanelPart.CreateReferenceFromObject(oPointOut), dDistancePts)
    dPMaxInIMS = dValueOfPointInLocalCoordSystem(dPartPosMatrix, dInertiaPosMatrix, dDistancePts)
    If dPMaxInIMS(1) < 0 Then
        '*** dVector Inertia coordinate system
        dVector(9) = -25.4 * 100: dVector(10) = 0: dVector(11) = 25.4 * 5
        '*** to convert Coordinate System of COG to Coordinate system of Part Origin
        Call MultiplyVector(dVector, dInertiaPosMatrix, dVectorReturn, True) ' Find Position of Coordinate w.r.t Partorigin
        oPointOut.X.Value = dVectorReturn(9): oPointOut.Y.Value = dVectorReturn(10): oPointOut.Z.Value = dVectorReturn(11)
        oPointOut.Compute
        dGap = oMeasureable.GetMinimumDistancePoints(oPanelPart.CreateReferenceFromObject(oPointOut), dDistancePts)
        dPMaxInIMS = dValueOfPointInLocalCoordSystem(dPartPosMatrix, dInertiaPosMatrix, dDistancePts)
    End If
    dPMaxLength = LengthVector(dPMaxInIMS)
    dOffset = 2 * dOffsetVal
    '*** Create a Point at 0.5in and 1 in away from PMAX on PMaxInIMS Positionvector and add 1.55in to their Z axis value
    If dPMaxLength > 25.4 Then
        Call PopulatePointCollection(cMeasurementCol, dPMaxInIMS, dPMaxLength, dOffset, _
                                    dPartPosMatrix, dInertiaPosMatrix, oMeasureable, oPointOut, oPanelPart, oHSF, oAutomationGeo)
        dOffset = 4 * dOffsetVal
        Call PopulatePointCollection(cMeasurementCol, dPMaxInIMS, dPMaxLength, dOffset, _
                                    dPartPosMatrix, dInertiaPosMatrix, oMeasureable, oPointOut, oPanelPart, oHSF, oAutomationGeo)
    End If
    '*** Verify That Part is a flat body
    If cMeasurementCol.Count = 4 Then
        On Error Resume Next
        Set oPlane = oHSF.AddNewPlane3Points(oPanelPart.CreateReferenceFromObject(cMeasurementCol.Item(1).Item("PointObject")), _
                                             oPanelPart.CreateReferenceFromObject(cMeasurementCol.Item(2).Item("PointObject")), _
                                             oPanelPart.CreateReferenceFromObject(cMeasurementCol.Item(3).Item("PointObject")))
        oAutomationGeo.AppendHybridShape oPlane
        oPlane.Compute
        If Err.Number <> 0 Then
            On Error GoTo 0
            iErrNumber = 1 '*** Cannot measure panel thickness
            GoTo endsub
        End If
        Err.Clear
        Set oMeasurePlane = oSPAWorkbench.GetMeasurable(oPanelPart.CreateReferenceFromObject(oPlane))
        dGap = oMeasurePlane.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(cMeasurementCol.Item(4).Item("PointObject")))
        If Err.Number <> 0 Then
            On Error GoTo 0
            iErrNumber = 1 '*** Cannot measure panel thickness
            GoTo endsub
        End If
        On Error GoTo 0
        If Abs(dGap) > TOLERANCE Then
            iErrNumber = 2 '***Non Flat Panel.
            GoTo endsub
        End If
    Else
        iErrNumber = 1 '*** Cannot measure panel thickness
        GoTo endsub
    End If
    '*** Measure Panel Thickness
    oPointOut.X.Value = cMeasurementCol.Item(2).Item("PtOppInPartAxis")(0)
    oPointOut.Y.Value = cMeasurementCol.Item(2).Item("PtOppInPartAxis")(1)
    oPointOut.Z.Value = cMeasurementCol.Item(2).Item("PtOppInPartAxis")(2)
    oPointOut.Compute
    dGap = oMeasureable.GetMinimumDistancePoints(oPanelPart.CreateReferenceFromObject(oPointOut), dDistancePts)
    oPointOut.X.Value = dDistancePts(0): oPointOut.Y.Value = dDistancePts(1): oPointOut.Z.Value = dDistancePts(2)
    oPointOut.Compute
    dPanelThickness = oMeasurePlane.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oPointOut))
    dPanelThickness = Round(dPanelThickness, 3)
    
    Set oPanelCheck.PanelPlane1 = oPlane
    oMeasurePlane.GetPlane Plane1
    oPanelCheck.SetArPlane1 Plane1
    Set oPanelCheck.PanelPlane2 = oHSF.AddNewPlaneOffsetPt(oPanelPart.CreateReferenceFromObject(oPlane), _
                                                           oPanelPart.CreateReferenceFromObject(oPointOut))
    oAutomationGeo.AppendHybridShape oPanelCheck.PanelPlane2
    oPanelCheck.PanelPlane2.Compute
    Set oMeasurePlane = oSPAWorkbench.GetMeasurable(oPanelPart.CreateReferenceFromObject(oPanelCheck.PanelPlane2))

    oMeasurePlane.GetPlane Plane2
    oPanelCheck.SetArPlane2 Plane2
    oPanelCheck.PanelThickness = dPanelThickness
endsub:
    On Error GoTo 0
    oPanelCheck.PanelThicknessStatus = iErrNumber
   
    'If oOutput.Count > 0 Then oOutput.Add iErrNumber, , 1 Else oOutput.Add iErrNumber
    'Set GetPanelThickness = oOutput
    Set oPanelCheck.oPanelThicknessGeometery = oAutomationGeo
    Set oPanelCheck.PanelPart = oPanelPart
    'If Not oAutomationGeo Is Nothing Then oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oAutomationGeo)
    DoEvents
End Sub

Private Sub PopulatePointCollection(ByRef PtCollection As Collection, _
                                    ByRef dExtrenumVector As Variant, _
                                    ByVal dVectorLength As Double, _
                                    ByVal dOffsetfromExtrnmInmm As Double, _
                                    ByRef dPartPosMatrix As Variant, _
                                    ByRef dInertiaPosMatrix As Variant, _
                                    ByRef oMeasSurf, _
                                    ByRef oPoint As HybridShapeTypeLib.Point, _
                                    ByRef oPanelPrt As MECMOD.Part, _
                                    ByRef oHYSF As HybridShapeTypeLib.HybridShapeFactory, _
                                    ByRef oGeoset As MECMOD.HybridBody)
                                    
    Dim dPtInIMS(9 To 11) As Double     ' Point 1.5in away in +ve Z Direction to find thickness  of Panel in InertiaAxis
'    Dim dPtOppInIMS(9 To 11) As Double  ' Point 1.5in away in -ve Z Direction to find thickness  of Panel in InertiaAxis
    Dim dPtOnFaceinIMS As Variant       ' Point of shortest Distance from dPtInIMS in Inertia Axis
    Dim dPtOppInPartAxis(2) As Double
    Dim PColName As String
    Dim dPtInPartAxis(9 To 11) As Double
    Dim dVectorReturn(9 To 11) As Double
    Dim dDistancePoints(8) ' As Variant
    Dim dGap As Double
    Dim oPointP As HybridShapeTypeLib.Point
    
    PColName = "P" & CStr(PtCollection.Count + 1)
    
    PtCollection.Add New Collection, PColName
    PtCollection.Item(PColName).Add PColName, PColName  'Name of Point Collection in Format P#

    '--------------For Pt in +ve Z axis of Inertia Axis--------------
    dPtInIMS(9) = ((dVectorLength - dOffsetfromExtrnmInmm) / dVectorLength) * dExtrenumVector(0)
    dPtInIMS(10) = ((dVectorLength - dOffsetfromExtrnmInmm) / dVectorLength) * dExtrenumVector(1)
    dPtInIMS(11) = ((dVectorLength - dOffsetfromExtrnmInmm) / dVectorLength) * dExtrenumVector(2) + 38 '*** 1.5 Inch away from Panel COG in Z axis of IMS
    
    '*** Find Position of Pt w.r.t Partorigin
    Call MultiplyVector(dPtInIMS, dInertiaPosMatrix, dVectorReturn, True)
    '*** Populate P1 Collection
    PtCollection.Item(PColName).Add Array(dVectorReturn(9), dVectorReturn(10), dVectorReturn(11)), "PtInPartAxis"
    '*** Calculate Pt on Face of Panel
    oPoint.X.Value = dVectorReturn(9): oPoint.Y.Value = dVectorReturn(10): oPoint.Z.Value = dVectorReturn(11)
    oPoint.Compute
    dGap = oMeasSurf.GetMinimumDistancePoints(oPanelPrt.CreateReferenceFromObject(oPoint), dDistancePoints)
    '*** Populate P1 Collection
    PtCollection.Item(PColName).Add Array(dDistancePoints(0), dDistancePoints(1), dDistancePoints(2)), "PtOnFaceInPartAxis"
    Set oPointP = oHYSF.AddNewPointCoord(dDistancePoints(0), dDistancePoints(1), dDistancePoints(2))
    oGeoset.AppendHybridShape oPointP
    oPointP.Compute
    PtCollection.Item(PColName).Add oPointP, "PointObject"
    '--------------For Pt opposite side of the panel --------------
    '*** using Equation P(t)=(1-t)P0 +tP1
    '*** where P0 is start coordinate and P1 is end coordinate of line
    '*** t is the ratio of length where distance from P0 to P1 is equat to t
    dPtOppInPartAxis(0) = -1.5 * dVectorReturn(9) + 2.5 * dDistancePoints(0)
    dPtOppInPartAxis(1) = -1.5 * dVectorReturn(10) + 2.5 * dDistancePoints(1)
    dPtOppInPartAxis(2) = -1.5 * dVectorReturn(11) + 2.5 * dDistancePoints(2)

    PtCollection.Item(PColName).Add dPtOppInPartAxis, "PtOppInPartAxis"
End Sub
Public Function dValueOfPointInLocalCoordSystem(ByVal dGlobalPositionMatrix As Variant, ByVal dLocalPositionMatrix As Variant, ByVal dPointCoordInGlobalMatrix) As Variant
    
    Dim dLocalMatrixPoint(11) As Variant
    Dim dGlobalMatrix As Variant
    Dim dLocalPointCoord(2) As Variant
    dGlobalMatrix = dGlobalPositionMatrix
    dGlobalMatrix(9) = dPointCoordInGlobalMatrix(0): dGlobalMatrix(10) = dPointCoordInGlobalMatrix(1): dGlobalMatrix(11) = dPointCoordInGlobalMatrix(2)
    
    '***Calculate position matrix in Local Coordinate system
    Call SubstractVector(dGlobalMatrix, dLocalPositionMatrix)
    Call InvertMatrix(dLocalPositionMatrix)
    Call MultiplyVector(dGlobalMatrix, dLocalPositionMatrix, dLocalMatrixPoint)
    Call MultiplyMatrix(dGlobalMatrix, dLocalPositionMatrix, dLocalMatrixPoint)
    dLocalPointCoord(0) = dLocalMatrixPoint(9): dLocalPointCoord(1) = dLocalMatrixPoint(10): dLocalPointCoord(2) = dLocalMatrixPoint(11)
    dValueOfPointInLocalCoordSystem = dLocalPointCoord
End Function

Private Sub SwapToAssemblyWorkbench()
    Dim sWkb As String
    
    sWkb = CATIA.GetWorkbenchId
    If sWkb <> "Assembly" Then
        CATIA.StartWorkbench "Assembly"
    End If
End Sub
Private Sub ExtractReportToExcel(ByRef oPanelCol As Collection)
    Dim iNbRow As Integer
    Dim iNbColumn As Integer
    'Dim iRowCounter As Integer
    Dim ReportCol As Collection
    Dim ReportKeysCol As Collection
    Dim sExportResultToExcel As Variant
    Dim bError As Boolean
    Dim sPropertiesThickness As String
    Dim oWorksheet ' As Worksheet
    Dim oTemplateWorksheet ' As Worksheet
    Dim oWorkbook 'As Workbook
    Dim sdummy As String
    Dim oPanelResult As clsPanelCheck
    '****for log
    Dim sDumpString As String
    Dim sStatus As String
    '****
    Dim i As Integer, j As Integer
    Dim k As Integer
    '***** Launch Excel
    Set oExcel = New clsExcel
    Call oExcel.GetExcel
    oExcel.App.ScreenUpdating = False
    
    Call oExcel.CloseAlreadyOpenWorkbook(sKbe_ReportTemplate)
    Call oExcel.CloseAlreadyOpenWorkbook("C:\Temp\EdgeDistanceAndInsertGapReport.xls")
    Call oExcel.OpenExcelFile(sKbe_ReportTemplate)
    Set oWorkbook = oExcel.App.Workbooks.Item(Right(sKbe_ReportTemplate, Len(sKbe_ReportTemplate) - InStrRev(sKbe_ReportTemplate, "\")))
    oWorkbook.Worksheets.Item("AnalysisResult").Activate
    Set oTemplateWorksheet = oWorkbook.Worksheets.Item("AnalysisResult")
    Set oWorksheet = oWorkbook.Worksheets.Item("AnalysisResult")
    oWorksheet.Range("C:C").EntireColumn.Insert
    oWorksheet.Range("C:C").EntireColumn.Hidden = True
    oWorksheet.LoadComments.Visible = False
    '***Create Validation List to for insert gaps
    For i = 1 To sInsertdistanceCol.Count
        oWorksheet.Cells(8000 + i, 3).Value = CStr(sInsertdistanceCol.Item(i) / 25.4) & "in"
    Next
    '****One work sheet per panel
    For Each oPanelResult In oPanelCol
        '**** get all SPCollectors for Logs update 28Jan2018
        oPanelResult.GetSPs
        '****
        sStatus = ""
        k = k + 1
        oTemplateWorksheet.Copy Before:=oWorkbook.Worksheets.Item("TolerancesHardwareCompatibility")
        oWorkbook.ActiveSheet.Name = "(" & CStr(k) & ")" & oPanelResult.PanelAssyInstance.Name
        Set oWorksheet = oWorkbook.ActiveSheet
        On Error Resume Next
            sPropertiesThickness = sGetPanelAttribute(oPanelResult.PanelInstance, "Thickness")
        On Error GoTo 0
        'On Error Resume Next
            oWorksheet.Cells(1, 2).Value = "Analysed Assy : " + oPanelResult.PanelAssyInstance.Name
            oWorksheet.Cells(3, 1).Value = "xxx"
            oWorksheet.Cells(3, 2).Value = "---Properties Thickness : " + sPropertiesThickness
            oWorksheet.Cells(4, 1).Value = "xxx"
            If oPanelResult.PanelThicknessStatus = EnumPanelThicknessStatus.OK Then
                oWorksheet.Cells(4, 2).Value = "---Actual Thickness : " + CStr(Round(oPanelResult.PanelThickness / 25.4, 8))
            Else
                oWorksheet.Cells(4, 2).Value = "---Actual Thickness : Unknown"
            End If
            oWorksheet.Cells(5, 1).Value = "xxx"
            oWorksheet.Cells(5, 2).Value = "Choose minimum distance between Inserts"
            oWorksheet.Cells(6, 1).Value = "xxx"
            oWorksheet.Range("B5").HorizontalAlignment = -4108 'xlCenter
            oWorksheet.Range("B6").Validation.Add Type:=3, AlertStyle:=1, _
                                    Formula1:="='" & oWorksheet.Name & "'!" & _
                                    oWorksheet.Range(oWorksheet.Cells(8001, 3), oWorksheet.Cells(8000 + sInsertdistanceCol.Count, 3)).Address
            oWorksheet.Range("B6").HorizontalAlignment = -4108 'xlCenter
            oWorksheet.Range("B6").Locked = False
            If sPropertiesThickness <> CStr(Round(oPanelResult.PanelThickness / 25.4, 8)) Then
                oWorksheet.Cells(4, 2).Interior.ColorIndex = 3
                sStatus = sStatus + "|ThicknessPropertyError"
                bError = True
            End If
        'On Error GoTo 0
        Set ReportCol = New Collection
        Set ReportKeysCol = New Collection
        '*** Building ReportCol
        If oPanelResult.InsertReportCollection.Count > 0 Then
            For i = 1 To oPanelResult.InsertReportCollection.Count
                If oPanelResult.InsertReportCollection.Item(i) <> "" Then
                    'sDummy = Split(oPanelResult.InsertkeyCollection.Item(i), "||")(0)
                    ReportCol.Add Left(oPanelResult.InsertkeyCollection.Item(i), InStr(oPanelResult.InsertkeyCollection.Item(i), ";") - 1)
                    ReportKeysCol.Add oPanelResult.InsertkeyCollection.Item(i)
                    iObject = Split(oPanelResult.InsertReportCollection.Item(i), ";")
                    For j = 0 To UBound(iObject)
                        ReportCol.Add iObject(j)
                        ReportKeysCol.Add oPanelResult.InsertkeyCollection.Item(i)
                    Next
                End If
                '*** Get stress non compatible result in log: update 28Jan2018
                If oPanelResult.InsertReportCollection.Item(i) Like "*" & sMsgStressIncompatible & "*" Then
                    sStatus = sStatus + "|[" + Split(oPanelResult.InsertkeyCollection.Item(i), ";")(0) + "] IncompatibleInsertPerStressReq"
                End If
                '***
            Next
        Else
            If oPanelResult.PanelAssyStatus = EnumPanelAssyStatus.BaseNumberError Then
                ReportCol.Add "---" & oPanelResult.PanelInstance.PartNumber & " doesn't have any panel assy!"
                ReportKeysCol.Add "xxx"
                sStatus = sStatus + "|NoPanelAssyFound"
            End If
            If oPanelResult.PanelAssyStatus = EnumPanelAssyStatus.SPCollectorNotFound Then
                ReportCol.Add "---SP Collector Not Found!"
                ReportKeysCol.Add "xxx"
                sStatus = sStatus + "|SPCollectorNotFound"
            End If
            If oPanelResult.PanelThicknessStatus = EnumPanelThicknessStatus.NonStandardPanelThickness Then
                ReportCol.Add "---Panel has a non standard thickness!"
                ReportKeysCol.Add "xxx"
                sStatus = sStatus + "|NonStandardPanelThickness"
            End If
            If oPanelResult.PanelThicknessStatus = EnumPanelThicknessStatus.NonFlatPanel Then
                ReportCol.Add "---Not a Flat panel, Please verify manually!"
                ReportKeysCol.Add "xxx"
                sStatus = sStatus + "|NonFlatPanel"
            End If
            If oPanelResult.PanelThicknessStatus = EnumPanelThicknessStatus.ThicknessUndefined Then
                ReportCol.Add "---Cannot ascertain panel's thickness, Please verify manually!"
                ReportKeysCol.Add "xxx"
                sStatus = sStatus + "|ThicknessUndefined"
            End If
        End If
        
        
        '***Generating Array for export
        'iRowCounter = 2

        If ReportCol.Count > 0 Then
            ReDim sExportResultToExcel(1 To ReportCol.Count, 1 To 3)
            For i = 1 To ReportKeysCol.Count
                sExportResultToExcel(i, 1) = ReportKeysCol.Item(i)
                If InStr(ReportCol.Item(i), "|") > 0 Then
                    sExportResultToExcel(i, 2) = Split(ReportCol.Item(i), "|")(0)
                    sExportResultToExcel(i, 3) = CStr(CDbl(Split(ReportCol.Item(i), "|")(1)) / 25.4) & "in"
                    If Not (ReportCol(i - 1) Like "---*") And Not (ReportCol(i + 1) Like "---*") Then
                        sExportResultToExcel(i - 1, 3) = CStr(CDbl(Split(ReportCol.Item(i), "|")(1)) / 25.4) & "in"
                        sExportResultToExcel(i + 1, 3) = CStr(CDbl(Split(ReportCol.Item(i), "|")(1)) / 25.4) & "in"
                    End If
                Else
                    sExportResultToExcel(i, 2) = ReportCol.Item(i)
                End If
            Next
            '*** Writing To Excel
            iNbRow = ReportKeysCol.Count + 6
            iNbColumn = 3
            oWorksheet.Range(oWorksheet.Cells(7, 1), oWorksheet.Cells(iNbRow, iNbColumn)).NumberFormat = "@"
            oWorksheet.Range(oWorksheet.Cells(7, 1), oWorksheet.Cells(iNbRow, iNbColumn)).Value = sExportResultToExcel
            '*** ---------Formatting---------------
            '***First Row Format
            oWorksheet.Cells(7, 2).Interior.ColorIndex = 35
            With oWorksheet.Cells(6, 2)
                .Borders(9).LineStyle = 1    'xlContinuous
                .Borders(9).Weight = -4138   'xlMedium
            End With
            '*** Rest of the rows format
            For iNbRow = 8 To ReportKeysCol.Count + 6
                If Not (oWorksheet.Cells(iNbRow, 2).Value2 Like "---*") And oWorksheet.Cells(iNbRow, 2).Value2 <> "" Then
                    oWorksheet.Cells(iNbRow, 2).Interior.ColorIndex = 35
                    With oWorksheet.Cells(iNbRow - 1, 2)
                        .Borders(9).LineStyle = 1    'xlContinuous
                        .Borders(9).Weight = -4138   'xlMedium
                    End With
                End If
            Next
            '***Last Row format
            With oWorksheet.Cells(ReportKeysCol.Count + 6, 2)
                .Borders(9).LineStyle = 1    'xlContinuous
                .Borders(9).Weight = -4138   'xlMedium
            End With
            '---------End Formatting-----------
            sStatus = sStatus + "|InsertRelatedErrorsFound"
        ElseIf Not bError Then '*** if no error found
            oWorksheet.Range(oWorksheet.Cells(7, 2), oWorksheet.Cells(7, 2)).NumberFormat = "@"
            oWorksheet.Range(oWorksheet.Cells(7, 2), oWorksheet.Cells(7, 2)).Value = "Congratulations! No error Found."
        End If
        '***Lock the sheet
        oWorksheet.Cells(6, 2).Value = oWorksheet.Cells(8000 + sInsertdistanceCol.Count, 3).Value
        oWorksheet.Protect "AutomationLock", True, True
        '****Log File   : update 28Jan2018
        If sStatus = "" Then
            Call AddToLogFile("PanelAssyCheck", oPanelResult.PanelPart.Parent.Name, , , , "OK", , , , , , , , , , , , , True)
            If oPanelResult.oListOfSPCollectors.Count > 0 Then
               For i = 1 To oPanelResult.oListOfSPCollectors.Count
                    Call AddToLogFile("PanelAssyCheck", oPanelResult.oListOfSPCollectors.Item(i).ReferenceProduct.Parent.Name, , , , "OK", , , , , , , , , , , , , True)
               Next
            End If
        Else
            Call AddToLogFile("PanelAssyCheck", oPanelResult.PanelPart.Parent.Name, , , , "KO", , sStatus, , , , , , , , , , , True)
            If oPanelResult.oListOfSPCollectors.Count > 0 Then
               For i = 1 To oPanelResult.oListOfSPCollectors.Count
                    Call AddToLogFile("PanelAssyCheck", oPanelResult.oListOfSPCollectors.Item(i).ReferenceProduct.Parent.Name, , , , "KO", , sStatus, , , , , , , , , , , True)
               Next
            End If
        End If
    Next
    oExcel.App.DisplayAlerts = False
    oWorkbook.Sheets.Item("TolerancesHardwareCompatibility").Delete
    oWorkbook.Sheets.Item("AnalysisResult").Delete
    
    Set oWorksheet = oWorkbook.Sheets.Item("TolerancesPanelAssyCheck")
    oWorksheet.Cells(2, 4).Value = CStr(Round(Abs(dFlushnessTol), 3)) + "in"
    oWorksheet.Cells(3, 4).Value = CStr(Round(Abs(dMinMisalignmentTol / 25.4), 3)) + "in"
    oWorksheet.Protect "AutomationLock", True, True
    
    oWorkbook.Sheets.Item("(1)" & oPanelCol.Item(1).PanelAssyInstance.Name).Activate
    
    Set oWorksheet = oWorkbook.Sheets.Item("(1)" & oPanelCol.Item(1).PanelAssyInstance.Name)
    oWorksheet.Range(oWorksheet.Cells(6, 2), oWorksheet.Cells(6, 2)).Select
    On Error GoTo 0
    '***Save excel file in local drive
    oExcel.App.ScreenUpdating = True
    Call oExcel.ShowExcelWindow
    oWorkbook.bNoRunBeforesave = True
    oWorkbook.SaveAs "C:\Temp\EdgeDistanceAndInsertGapReport_" & Format((DateTime.Timer) / 86400, "hh_mm_ss") & ".xls"
    oWorkbook.bNoRunBeforesave = False
    DoEvents
    oExcel.App.DisplayAlerts = True
    
    Set oExcel = Nothing

End Sub
'********************************************
'****           WINAPI Functions
'********************************************
Public Function lChildWindow(ByVal lParentHwnd As Long, Optional ByVal sChildCls As String = "", Optional ByVal sChildText As String = "") As Long
    sChildCLASS = ""
    sChildSTRING = ""
    If sChildCls <> "" Then sChildCLASS = sChildCls
    If sChildText <> "" Then sChildSTRING = sChildText
    
    EnumChildWindows lParentHwnd, AddressOf GetChildHwndByNameAndClass, 0
    lChildWindow = Child_hwnd
End Function
Public Function FindWindowLike(strPartOfCaption As String) As Long
    '**** Adapted from http://www.vbforums.com/showthread.php?277398-VB-FindWindowLike

    Dim strCurrentWindowText As String
    Dim r As Integer
    Dim hWnd As Long
    CATIA.RefreshDisplay = True 'To make catia the foreground window
    hWnd = GetForegroundWindow
    Do Until hWnd = 0
        strCurrentWindowText = Space$(255)
        r = GetWindowText(hWnd, strCurrentWindowText, 255)
        strCurrentWindowText = Left$(strCurrentWindowText, r)
        If InStr(1, LCase(strCurrentWindowText), LCase(strPartOfCaption)) <> 0 Then GoTo Found
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    Exit Function
Found:
    FindWindowLike = hWnd
End Function
Private Function GetChildHwndByNameAndClass(ByVal hChild As Long, ByVal lParam As Long) As Long
    
    Dim sClass As String
    Dim sText As String
    Dim j As Integer
    Dim bClassFound As Boolean
    Dim bTextFound As Boolean
    
    If sChildCLASS = "" Then bClassFound = True
    If sChildSTRING = "" Then bTextFound = True
    
    sClass = Space(256)
    j = GetClassName(hChild, sClass, 63)
    sClass = Left(sClass, j)
    sText = Space(256)
    j = SendMessageStr(hChild, WM_GETTEXT, 255, sText)
    sText = Left(sText, j)
    
    If sClass = sChildCLASS Then bClassFound = True
    If sText = sChildSTRING Then bTextFound = True
    
    If bClassFound And bTextFound Then
        Child_hwnd = hChild
        Exit Function
    End If
    If hChild <> 0 Then
        GetChildHwndByNameAndClass = 1 ' Continue enumeration
    Else
        GetChildHwndByNameAndClass = 0 ' exit function
    End If

End Function

