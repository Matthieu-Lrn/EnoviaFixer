Attribute VB_Name = "mdlCheckDocAttFromForm"
'--------------------------------------------------
Private Const sAttributeValidationToolSheetName As String = "Attributes_Template"
Private Const sAttributeValidationToolSkeletonConsumableSheetName As String = "Consumable-Compatibility"
Private Const sAttributeValidationToolCATIAMatlSheetName As String = "CatiaMaterial"
Private Const sAttributeValidationToolMaterialThicknessSheetName As String = "MaterialThickness"

'--------------------------------------------------
Public Enum EnumDocumentType ' >>> Update Fctn GetDocumentType as required!
    UNKNOWN = 0
    NONE = 1
    detail = 2
    COLLECTOR = 3
    Drawing = 4
    PVR = 5
    SKEL = 6
    RV = 7
    edrn = 8
    PLMAction = 9
    FLEXASSY = 10
    FLEXPART = 11
    SPLITPART = 12
    DesignReview = 13
    ERGONOMIC = 14
    FITTING = 15
    ELECBDL = 16
    INTERMEDIATE = 17
    MPP = 18
    MBP = 19
    KIN = 20
    KBE = 21
    STRESS = 22
    WEIGHTPART = 23
    TOLERANCE = 24
    CATALOG = 25
End Enum

Private Const bUseBDIDevtDatabase As Boolean = False
Private sAttributeTemplateArray As Variant
Private oListViewHwnd As New Collection
Public Sub RunAttributeAnalysis(ByVal oAnalysisCol As Collection)

    Dim oPartInfo As Collection
    Dim enumDocType As EnumDocumentType
    Dim sDocumentFamilyType As String
    
    'Clear cache
    Call StartWebServiceTool
    WebServiceAccessTool.ClearCache
    
    For Each oPartInfo In oAnalysisCol
        Dim oDocAttributes As Variant
        enumDocType = GetDocTypeFromDashNumber(oPartInfo.Item("PartNumber"))
        sDocumentFamilyType = GetDocumentFamilytype(oPartInfo.Item("PartNumber"), _
                                                        oPartInfo.Item("Revision"), _
                                                        enumDocType, _
                                                        oDocAttributes, _
                                                        oPartInfo.Item("Title"))
        oPartInfo.Add sDocumentFamilyType, "DocumentFamily"
        oPartInfo.Add oDocAttributes, "Attributes"
    Next

    Call ExportToExcel(oAnalysisCol)
End Sub
Private Function GetDocumentFamilytype(ByVal sPartNumber As String, _
                                       ByVal sRevision As String, _
                                       ByVal DocType As EnumDocumentType, _
                                       ByRef oDocumentAttributes As Variant, _
                                       Optional ByVal sTitle As String = "") As String

'*** It returns false in case of a mismatch of one or more attributes
'*** in case of a mismatch, collection oReturnMisMatch will have list of erroneous attributes in following format
'*** Key, Catia Value|Enovia Value
    Dim params()
    Dim Attributes As Variant
    Dim AttributeErrors As clsCollection
    Dim DocumentAttributFamilyType As String

    Dim ofrmSelectAttributeFamily As frmSelectAttributeFamily
    Dim sDocumentAttributeFamilyUnknown  As String
    
    sDocumentAttributeFamilyUnknown = "UNKNOWN"


    Set Attributes = WebServiceAccessTool.GetENOVIADocumentAttributs(sPartNumber, sRevision)
    'WebServices.GetENOVIADocumentAttributs(sPartNumber, sRevision, bUseBDIDevtDatabase)
    Set AttributeErrors = New clsCollection
    
    ' Reset Doc Family Type:
    DocumentAttributFamilyType = ""
    
    '*** Convert Security Check Key v1.5.1
    If Attributes.Exists("Security Check") Then
        Select Case Attributes.Item("Security Check")
            Case "SecurityCheck_YX": Attributes.Item("Security Check") = "RA Check / EC To Review"
            Case "SecurityCheck_YY": Attributes.Item("Security Check") = "RA Check / EC Check"
            Case "SecurityCheck_YN": Attributes.Item("Security Check") = "RA Check"
            Case "SecurityCheck_NY": Attributes.Item("Security Check") = "EC Check"
            Case "SecurityCheck_NN": Attributes.Item("Security Check") = "No Check"
        End Select
    End If

    ' Retrieve Excel Attribute Validation Tool Array:
    If IsEmpty(sAttributeTemplateArray) Then
        sAttributeTemplateArray = XLApp_ImportTable(sAttributesTemplateFile, sAttributeValidationToolSheetName, 3, 1)
    End If
    '------------------------------------------
    Dim oResetArray As Variant
    
    Call FilterAttributeTemplateBasedOnDocType(sAttributeTemplateArray, DocType, oResetArray)
    '----------------------------
    '*** Result Collection in following format
    '*** FirstItem:         Type of Document
    '*** Subsequent Items:  Erroneous Attribut|CurrentValue| ExpectedValue or values
    '***--------
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bCheck As Boolean
    Dim oCombinedAttributsDict As New clsCollection
    oCombinedAttributsDict.RemoveAll
    Dim oPartTypeCollectionOfError As New Collection
    If oPartTypeCollectionOfError.Count <> 0 Then Stop
    Dim sValue
    
    
    '***Generate Combined attributes dictionary: ----- superfluous operation
    Dim AttributesKeys
    Set AttributesKeys = Attributes.GetKeys
    For Each sKey In AttributesKeys
        If Not oCombinedAttributsDict.Exists(sKey) Then oCombinedAttributsDict.Add sKey, Attributes.Item(sKey)
    Next
    '*** bug in Enovia Attribute for Interchangeability Parts
    If oCombinedAttributsDict.Exists("Interchangeability Parts") And oCombinedAttributsDict.GetItem("Interchangeability Parts") = 0 Then
        oCombinedAttributsDict.SetItem "Interchangeability Parts", ""
    End If
    '***-----
    Set oDocumentAttributes = oCombinedAttributsDict
    '*** PerformCheck
    '*** j = column i = rows of sAttributeTemplateArray
    For j = 3 To UBound(sAttributeTemplateArray, 2)
        '*** Perform Check only for same document type:
        'If Split(sAttributeTemplateArray(2, j), "|")(1) = GetDocumentType(GetDocTypeFromDashNumber(sPartNumber)) Then
            k = 0
            oPartTypeCollectionOfError.Add New clsCollection, sAttributeTemplateArray(1, j)   '*** Type of Document
            oPartTypeCollectionOfError.Item(sAttributeTemplateArray(1, j)).Add sAttributeTemplateArray(1, j), sAttributeTemplateArray(1, j)
            
            For i = 3 To UBound(sAttributeTemplateArray, 1)
                If oCombinedAttributsDict.Exists(sAttributeTemplateArray(i, 1)) Then     '*** Perform Check only if attributes exist in dictionary
                    If Not sAttributeTemplateArray(i, j) = "NO_CHECK" Then
                        k = k + 1
                        bCheck = False
                        ' Trim the last ";" character:
                        If Right(sAttributeTemplateArray(i, j), 1) = ";" Then sAttributeTemplateArray(i, j) = Left(sAttributeTemplateArray(i, j), Len(sAttributeTemplateArray(i, j)) - 1)
                        ' Loop through each possible value:
                        For Each sValue In Split(Trim(sAttributeTemplateArray(i, j)), ";")
                            If sValue <> "" Then
                                If Left(sValue, 4) = "NOT:" Then
                                    sValue = Split(sValue, "NOT:")(1)
                                    bCheck = Not (DocAttributs_CheckValue(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), sValue))
                                    If bCheck Then
                                        Exit For
                                    End If
                                Else
                                    bCheck = DocAttributs_CheckValue(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), sValue)
                                    If bCheck Then
                                        Exit For
                                    End If
                                    ' Specific loop for I/R/C/S:
                                    Dim sValue2 As Variant
                                    If bCheck = False And sAttributeTemplateArray(i, 1) = "I/R/C/S" And InStr(1, oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), "/") <> 0 Then
                                        For Each sValue2 In Split(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), "/")
                                            bCheck = DocAttributs_CheckValue(sValue2, sValue)
                                            If Not bCheck Then Exit For
                                        Next
                                    End If
                                    ' Specific loop for ATA Chapter to accept multiple values, i.e. "25-4;25-2;25-3" (v1.7.3):
                                    If bCheck = False And sAttributeTemplateArray(i, 1) = "ATA Chapter Section / SNS" And InStr(1, oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), ";") <> 0 Then
                                        For Each sValue2 In Split(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), ";")
                                            bCheck = CBool(InStr(1, sAttributeTemplateArray(i, j), sValue2) > 0)
                                            If Not bCheck Then Exit For
                                        Next
                                    End If
                                    ' Specific loop for FT_LOCATION_ZONE:
                                    If bCheck = False And sAttributeTemplateArray(i, 1) = "FT Location Zone" And InStr(1, oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), ";") <> 0 Then
                                        For Each sValue2 In Split(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1)), ";")
                                            bCheck = DocAttributs_CheckValue(sValue2, sValue)
                                            If Not bCheck Then Exit For
                                        Next
                                    End If
                                    ' Specific loop for Title of SKELETON-CONSUMABLE (v1.5.5):
                                    If bCheck = False And sAttributeTemplateArray(i, 1) = "Title" And InStr(1, sAttributeTemplateArray(i, j), sAttributeValidationToolSkeletonConsumableSheetName) <> 0 Then
                                        ' Retrieve Excel Skeleton-Consumable Title Table:
                                        If IsEmpty(sTitleSkeletonConsumableTitleArray) Then
                                            sTitleSkeletonConsumableTitleArray = XLApp_ImportTable(sAttributesTemplateFile, sAttributeValidationToolSkeletonConsumableSheetName, 3, 1)
                                        End If
                                        
                                        Dim sRefTitle As String, sMatlSpec As String, sPartTitle As String 'v1.5.11
                                        sMatlSpec = oCombinedAttributsDict.GetItem("Material Specifications")
                                        sPartTitle = oCombinedAttributsDict.GetItem("Title")
                                        
                                        For k = 2 To UBound(sTitleSkeletonConsumableTitleArray, 1)

                                            If sTitleSkeletonConsumableTitleArray(k, 1) = sMatlSpec And sRefTitle = "" Then sRefTitle = sTitleSkeletonConsumableTitleArray(k, 2)
                                            If sTitleSkeletonConsumableTitleArray(k, 1) = sMatlSpec And sRefTitle = sPartTitle Then
                                                bCheck = True
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        Next
                        If Not (bCheck) Then
                            oPartTypeCollectionOfError.Item(sAttributeTemplateArray(1, j)).Add sAttributeTemplateArray(i, 1), sAttributeTemplateArray(i, 1) + "|" + CStr(oCombinedAttributsDict.GetItem(sAttributeTemplateArray(i, 1))) + "|" + sAttributeTemplateArray(i, j)
                        End If
                    End If
                End If
            Next
            oPartTypeCollectionOfError.Item(sAttributeTemplateArray(1, j)).Add "NoOfAttbs", k
        'End If
    Next
        
    '*** Sort oPartTypeCollectionOfError
    Dim oDummyList As clsCollection
    If oPartTypeCollectionOfError.Count > 0 Then
        For i = 1 To oPartTypeCollectionOfError.Count - 1
            For j = i + 1 To oPartTypeCollectionOfError.Count
                If (oPartTypeCollectionOfError.Item(i).Count / oPartTypeCollectionOfError.Item(i).GetItem("NoOfAttbs")) > (oPartTypeCollectionOfError.Item(j).Count / oPartTypeCollectionOfError.Item(j).GetItem("NoOfAttbs")) Then
                    Set oDummyList = oPartTypeCollectionOfError.Item(j)
                    oPartTypeCollectionOfError.Remove j
                    oPartTypeCollectionOfError.Add oDummyList, oDummyList.GetItem(1), i
                End If
            Next j
        Next i
        Set ofrmSelectAttributeFamily = New frmSelectAttributeFamily
        ' Find Multiple Possible Document Type:
        sMultiKey = ""
        sMultiValue = ""
        sValue = ""
        j = 1
        If oPartTypeCollectionOfError.Item(1).Count > 2 Then
           Do
               ' Record Multiple Attribute Errors for DES (v1.5.9):
               If sMultiValue <> "" Then sMultiValue = sMultiValue & Chr(11) & Chr(11)
               sMultiValue = sMultiValue & oPartTypeCollectionOfError.Item(j).GetItem(1)
    
               j = j + 1
               
               If j > oPartTypeCollectionOfError.Count Then Exit Do
           Loop Until oPartTypeCollectionOfError.Item(j).Count > oPartTypeCollectionOfError.Item(1).Count + 20
        Else
            Do
               'Record Multiple documnet family type possible that is all the family type without error
               If sMultiValue <> "" Then sMultiValue = sMultiValue & Chr(11) & Chr(11)
               sMultiValue = sMultiValue & oPartTypeCollectionOfError.Item(j).GetItem(1)
    
               j = j + 1
               
               If j > oPartTypeCollectionOfError.Count Then Exit Do
           Loop Until oPartTypeCollectionOfError.Item(j).Count > oPartTypeCollectionOfError.Item(1).Count
           DocumentAttributFamilyType = oPartTypeCollectionOfError.Item(1).GetItem(1)
        End If
        '***more than one possible document family type possible then ask user to confirm the document type
        If InStr(1, sMultiValue, Chr(11)) <> 0 Then
            ofrmSelectAttributeFamily.DocName = sPartNumber & sRevision
            ofrmSelectAttributeFamily.DocTitle = sTitle
            ofrmSelectAttributeFamily.ComboBoxSelect.Clear
            
            If ofrmSelectAttributeFamily.ComboBoxSelect.ListCount > 0 Then
               Do Until ofrmSelectAttributeFamily.ComboBoxSelect.ListCount < 1: ofrmSelectAttributeFamily.ComboBoxSelect.RemoveItem 0: Loop
            End If
            For i = 0 To UBound(Split(sMultiValue, Chr(11)))
               If Trim(Split(sMultiValue, Chr(11))(i)) <> "" Then ofrmSelectAttributeFamily.ComboBoxSelect.AddItem Split(sMultiValue, Chr(11))(i)
            Next
            ofrmSelectAttributeFamily.Show vbModal
            DoEvents
            Sleep 100
            DocumentAttributFamilyType = ofrmSelectAttributeFamily.DocumentAttributFamilyType
        Else
            'if multivalue is actually a single value  :) then no need to ask user
            DocumentAttributFamilyType = sMultiValue
        End If
    Else
        'if oPartTypeCollectionOfError.Count = 0 then cannot ascertain family type
        DocumentAttributFamilyType = sDocumentAttributeFamilyUnknown
    End If
    
ExitFnct:
    If Not IsEmpty(oResetArray) Then sAttributeTemplateArray = oResetArray
     GetDocumentFamilytype = DocumentAttributFamilyType
     
End Function

Private Function DocAttributs_CheckValue(ByRef sDictionaryValue, ByRef sValueforComparison) As Boolean '  (v1.2.7a)
    Dim bReturnValue As Boolean
    If UCase(sValueforComparison) = "BLANK" Then
        bReturnValue = CStr(sDictionaryValue) = ""
    ElseIf IsNumeric(sDictionaryValue) And IsNumeric(sValueforComparison) Then
        bReturnValue = CDbl(sDictionaryValue) = CDbl(sValueforComparison)
    ElseIf CStr(sDictionaryValue) = CStr(sValueforComparison) Then
        bReturnValue = CStr(sDictionaryValue) = CStr(sValueforComparison)
    End If
    DocAttributs_CheckValue = bReturnValue
End Function

Public Function XLApp_ImportTable(sExcelFilePath As String, Optional ByVal sWorksheetName As String = "", Optional ByRef StartRow As Integer = 2, Optional ByRef StartCol As Integer = 2) As Variant
'*** Imports values from excel file & dumps it in an array
'*** if aWorksheet name is give then it will import from a particular sheet
'*** otherwise it will import from active sheet
    
    Dim sAnswer As String
    Dim iNbRow As Integer
    Dim iNbColumn As Integer
    Dim sArray As Variant
    
    Dim XLApp As Object ' As New Excel.Application '
    Set XLApp = CreateObject("EXCEL.Application")
    XLApp.DisplayAlerts = False
    XLApp.ScreenUpdating = False
    Err.Clear
    On Error GoTo 0
    Dim oXLWorkbook 'As Workbook
    Set oXLWorkbook = XLApp.Workbooks.Add(sExcelFilePath)
    Dim oXLSheet 'As Worksheet
    If sWorksheetName = "" Then sWorksheetName = oXLWorkbook.Sheets.Item(1).Name
    Set oXLSheet = oXLWorkbook.Sheets(sWorksheetName)

    
    If XLApp.Workbooks.Count = 0 Then
        sAnswer = CATIA.MsgBox("Cannot Open Excel workbook" & vbCrLf & "Macro aborted.", vbExclamation, "AttributeCheck", "", 0)
        GoTo endsub
    End If
    
    iNbRow = StartRow
    Do While oXLSheet.Cells(iNbRow, StartCol).Value <> ""
        iNbRow = iNbRow + 1
    Loop
    iNbRow = iNbRow - 1
'    iNbRow = oXLSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    iNbColumn = StartCol
    Do While oXLSheet.Cells(StartRow, iNbColumn).Value <> ""
        iNbColumn = iNbColumn + 1
    Loop
    iNbColumn = iNbColumn - 1
'    iNbColumn = oXLSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    '***** Transfer values from Excel file to sArray
    ReDim sArray(1 To iNbRow, 1 To iNbColumn)
    sArray = oXLSheet.Range(oXLSheet.Cells(1, 1), oXLSheet.Cells(iNbRow, iNbColumn)).Value
 
    XLApp_ImportTable = sArray

    ' v1.5.16 -----------------
    XLApp.DisplayAlerts = False
    oXLWorkbook.Close
    DoEvents
    XLApp.ScreenUpdating = True
    '---------------------------

endsub:
    XLApp.Quit
    DoEvents
End Function
Private Function GetDocTypeFromDashNumber(ByVal sPartNumber As String) As EnumDocumentType
    Dim DashNumber As String
    Dim DocType As EnumDocumentType
    sPartNumber = Replace(sPartNumber, "-", "|", , 1)
    If InStr(sPartNumber, "|") > 0 Then DashNumber = "-" & Split(sPartNumber, "|")(1)
    
    On Error Resume Next
    Set regex = CreateObject("vbscript.regexp")
    regex.IgnoreCase = True
    regex.Global = False
    
    DocType = EnumDocumentType.UNKNOWN
    
    ' DRAWING
    If DashNumber = "" Then DocType = EnumDocumentType.Drawing: GoTo ExitSub
    
    ' NONE
    regex.Pattern = "-0\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.NONE: GoTo ExitSub

    ' DETAIL PART
    regex.Pattern = "-\d{3}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.detail: GoTo ExitSub
    
    ' SKEL or ARM (v1.5.21)
    regex.Pattern = "-\d{3}(SKEL|ARM)\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.SKEL: GoTo ExitSub

    ' SP:
    regex.Pattern = "(-\d{3}SP\d{2}|-\d{3}(I|D)\d{2}SP\d{2})($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.COLLECTOR: GoTo ExitSub
    
    ' PVR
    regex.Pattern = "((-\d{3}PVR\w{3}\d*)($|-XXXX$|-YYYY$)|(-\d{3}-XXXXPVR\w{3}\d*$))" 'v1.5.12
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.PVR: GoTo ExitSub
    
    ' FLEX ASSY
    regex.Pattern = "-\d*I\d{2,9}($|-XXXX$|-YYYY$)" 'v1.6.0
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.FLEXASSY: GoTo ExitSub
    
    ' FLEX PART
    regex.Pattern = "-\d*D\d{2,9}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.FLEXPART: GoTo ExitSub
    
    ' SPLIT PART
    regex.Pattern = "-\d*DL\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.SPLITPART: GoTo ExitSub
    
    ' DESIGN REVIEW
    regex.Pattern = "-\d*DR\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.DesignReview: GoTo ExitSub
    
    ' ERGONOMICS
    regex.Pattern = "-\d*ERGO\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.ERGONOMIC: GoTo ExitSub

    ' FITTING
    regex.Pattern = "-\d*FTG\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.FITTING: GoTo ExitSub

    ' ELEC. BUNDLE
    regex.Pattern = "-\d*GBN\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.ELECBDL: GoTo ExitSub

    ' INTERMEDIATE
    regex.Pattern = "-\d*INT\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.INTERMEDIATE: GoTo ExitSub

    ' KBE
    regex.Pattern = "-\d*KBE\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.KBE: GoTo ExitSub

    ' KIN
    regex.Pattern = "-\d*KIN\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.KIN: GoTo ExitSub

    ' MultiPARTS PART
    regex.Pattern = "-\d*M\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.MPP: GoTo ExitSub

    ' MultiBODY PART
    regex.Pattern = "-\d*MBP\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.MBP: GoTo ExitSub

    ' REPLACEABLE COMP. CONNECTOR
    regex.Pattern = "-\d*RCC\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.COLLECTOR: GoTo ExitSub

    ' KBE SUB SET
    regex.Pattern = "-\d*SS\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.KBE: GoTo ExitSub

    ' STRESS
    regex.Pattern = "-\d*STRE\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.STRESS: GoTo ExitSub

    ' TOL ANALYSIS
    regex.Pattern = "-\d*TOL\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.TOLERANCE: GoTo ExitSub

    ' TUBING COLLECTOR
    regex.Pattern = "-\d*TUB\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.COLLECTOR: GoTo ExitSub

    ' WEIGHTS
    regex.Pattern = "-\d*WEIG\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.WEIGHTPART: GoTo ExitSub

    ' ELEC HARNESS
    regex.Pattern = "-\d*WIRE\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.COLLECTOR: GoTo ExitSub

    ' CATALOG LIGGHT VERSION
    regex.Pattern = "-\d*CAT\d{2}($|-XXXX$|-YYYY$)"
    Set regExp_Matches = regex.Execute(DashNumber)
    If Err.Number = 0 Then If regExp_Matches.Count > 0 Then DocType = EnumDocumentType.CATALOG: GoTo ExitSub
  
ExitSub:
    GetDocTypeFromDashNumber = DocType
    Set regex = Nothing
End Function
Private Function GetDocumentType(i As EnumDocumentType) As String
    Select Case i
        Case 0: GetDocumentType = "UNKNOWN"
        Case 1: GetDocumentType = "NONE"
        Case 2: GetDocumentType = "DETAIL"
        Case 3: GetDocumentType = "COLLECTOR"
        Case 4: GetDocumentType = "DRAWING"
        Case 5: GetDocumentType = "PVR"
        Case 6: GetDocumentType = "SKELETON"
        Case 7: GetDocumentType = "RELEASE VEHICLE"
        Case 8: GetDocumentType = "EDRN"
        Case 9: GetDocumentType = "PLMACTION"
        Case 10: GetDocumentType = "FLEXASSY" ' v1.5.22
        Case 11: GetDocumentType = "FLEXPART" ' v1.5.22
        Case 12: GetDocumentType = "SPLIT PART"
        Case 13: GetDocumentType = "DESIGN REVIEW"
        Case 14: GetDocumentType = "ERGONOMIC"
        Case 15: GetDocumentType = "FITTING"
        Case 16: GetDocumentType = "ELEC. BUNDLE"
        Case 17: GetDocumentType = "INTERMEDIATE"
        Case 18: GetDocumentType = "MULTIPARTS PART"
        Case 19: GetDocumentType = "MULTIBODY PART"
        Case 20: GetDocumentType = "KINEMATIC"
        Case 21: GetDocumentType = "KBE"
        Case 22: GetDocumentType = "STRESS"
        Case 23: GetDocumentType = "WEIGHT"
        Case 24: GetDocumentType = "TOLERANCE"
        Case 25: GetDocumentType = "CATALOG"
    End Select
    ' mmaincheck.GetDocumentType(EnumDocumentType)
End Function
Private Sub FilterAttributeTemplateBasedOnDocType(ByRef sAttributeTemplateArray As Variant, ByVal DocType As EnumDocumentType, ByRef oResetArray As Variant)

    Dim iColumn As Integer, iTypeRow As Integer, iClassRow As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer

    Dim sClassNewVal As String, sTypeNewVal As String
    Dim sClassoldVal As String, sTypeOldValue As String
    Dim sdummy As String
    Dim sDocumentType As String
    Dim NewAttributeTemplate As Variant
    Dim NewAttributeCollection As New Collection
    Dim bDocTypeFoundInTemplate As Boolean
    oResetArray = sAttributeTemplateArray

    sDocumentType = GetDocumentType(DocType)
    Set NewAttributeCollection = New Collection
    
    For j = 1 To 2
        NewAttributeCollection.Add New Collection
        For i = 1 To UBound(sAttributeTemplateArray, 1)
            NewAttributeCollection.Item(NewAttributeCollection.Count).Add sAttributeTemplateArray(i, j)
        Next
    Next
    '*** Make sure that sDocument type is found in Template, Not all sDocumentTypes are available in Template
    For j = 3 To UBound(sAttributeTemplateArray, 2)
        If Split(sAttributeTemplateArray(2, j), "|")(1) = sDocumentType Then
            bDocTypeFoundInTemplate = True
            Exit For
        End If
    Next
    '*** filter only of Document Type is found in the Template
    If bDocTypeFoundInTemplate Then
        For j = 3 To UBound(sAttributeTemplateArray, 2)
            If Split(sAttributeTemplateArray(2, j), "|")(1) = sDocumentType Then
                NewAttributeCollection.Add New Collection
                For i = 1 To UBound(sAttributeTemplateArray, 1)
                    NewAttributeCollection.Item(NewAttributeCollection.Count).Add sAttributeTemplateArray(i, j)
                Next
            End If
        Next
        ReDim sAttributeTemplateArray(1 To NewAttributeCollection.Item(1).Count, 1 To NewAttributeCollection.Count)
        For i = 1 To NewAttributeCollection.Item(1).Count
            For j = 1 To NewAttributeCollection.Count
                sAttributeTemplateArray(i, j) = NewAttributeCollection.Item(j).Item(i)
            Next
        Next
    End If
End Sub
Private Sub ExportToExcel(ByRef oAnalysisCol As Collection)
     Dim iNbRow As Integer, i As Integer, iTypeRow As Integer
     Dim oExcel As clsExcel
     Dim oWorkbook
     Dim oWorksheet
     Dim sAttributes As Variant 'as Object
     Dim iBlankCell As Integer
     Dim oPartInforCol As Collection
     
     Set oExcel = New clsExcel
     oExcel.GetExcel
     Call oExcel.OpenExcelFile(sAttributesTemplateFile)
     Set oWorkbook = oExcel.App.Workbooks.Item(Right(sAttributesTemplateFile, Len(sAttributesTemplateFile) - InStrRev(sAttributesTemplateFile, "\")))
     Set oWorksheet = oWorkbook.Worksheets.Item("Attributes_Check")
     oWorksheet.Activate
     
    iNbRow = 1
    iTypeRow = 2 '*** row with document type
    Do While oWorksheet.Cells(iNbRow, 1).Value <> ""
        iNbRow = iNbRow + 1
    Loop
    iNbRow = iNbRow - 1
    
    For Each oPartInforCol In oAnalysisCol
        iBlankCell = 2
        Do While CStr(oWorksheet.Cells(3, iBlankCell).Value) <> ""
            iBlankCell = iBlankCell + 1
            If iBlankCell = 500 Then Exit Sub
        Loop
        Set sAttributes = oPartInforCol.Item("Attributes")
        For i = 3 To iNbRow
            If sAttributes.Exists(oWorksheet.Cells(i, 1).Value) Then
                oWorksheet.Cells(i, iBlankCell).Value = sAttributes.GetItem(oWorksheet.Cells(i, 1).Value)
            End If
        Next
        If oPartInforCol.Item("DocumentFamily") <> "UNKNOWN" Then oWorksheet.Cells(iTypeRow, iBlankCell).Value = oPartInforCol.Item("DocumentFamily")
        oWorksheet.Cells(i, iBlankCell).EntireColumn.AutoFit
    Next
    oExcel.App.ScreenUpdating = True
    Call oExcel.ShowExcelWindow
    oExcel.App.Run oWorkbook.Name & "!ValidateValues.CheckValues", True
    
    Set oExcel = Nothing
End Sub


