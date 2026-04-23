Attribute VB_Name = "clsAttributesList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oListEnovia As clsCollection
Private oListLatestRev As clsCollection
Private oListLatestRelRev As clsCollection

Private Sub Class_Initialize()
    Set oListEnovia = New clsCollection
    Set oListLatestRev = New clsCollection
    Set oListLatestRelRev = New clsCollection
End Sub

Private Sub Class_Terminate()
    Set oListEnovia = Nothing
    Set oListLatestRev = Nothing
    Set oListLatestRelRev = Nothing
End Sub

Public Function GetEnoviaAttributes(ByVal sDMUPartNumber As String, ByVal sDMUDocRev As String, Optional bDefiningPart = False, Optional sAttributeName = "") As Variant

Dim oTemp
Dim sAttDMUDoc As clsCollection
Dim sAttDefiningDoc As clsCollection
Dim oItem As clsCollection
Dim sDefiningPartNumber As String
Dim sDefiningDocRev As String
Dim sDatasetType As String
Dim sString As String

'Initial set
sDatasetType = ""

'Get "LatestRev" or "LatestRelRev" or "LatestRelRevIfExist"
If sDMUDocRev = "LatestRev" Or sDMUDocRev = "LatestRelRev" Or sDMUDocRev = "LatestRelRevIfExist" Then
    sDMUDocRev = GetLatestDocRevision(sDMUPartNumber, sDMUDocRev)
    bIsLatestRev = True
End If

'If rev is "", we return empty data
If sDMUDocRev = "" Then GoTo ReturnEmptyData

'Check if DMU document can be found in oListEnovia
If Not oListEnovia.Exists(sDMUPartNumber & sDMUDocRev) Then

    'Get attributes
    Set oTemp = WebServiceAccessTool.GetENOVIADocumentAttributs(sDMUPartNumber, sDMUDocRev)
    If Not oTracking Is Nothing Then Call oTracking.WebServiceResultAnalysis("GetENOVIADocumentAttributs", oTemp, sDMUPartNumber & " " & sDMUDocRev)
    Set sAttDMUDoc = New clsCollection
    Call sAttDMUDoc.InitializeWithDLLclsColObject(oTemp, sAttDMUDoc)
    
    'Add to oListEnovia
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaAttributes", sAttDMUDoc)
    Call oItem.Add("DefiningPart", "")
    Call oListEnovia.Add(sDMUPartNumber & sDMUDocRev, oItem)

    'If there is not attribute, we return empty data
    If sAttDMUDoc.Count = 0 Then GoTo ReturnEmptyData
    
    'Get Dataset Type
    sDatasetType = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").GetItemByKey("Dataset Type")
    
Else

    'If there is not attribute, we return empty data
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").Count = 0 Then GoTo ReturnEmptyData
    
    'Get Dataset Type
    sDatasetType = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").GetItemByKey("Dataset Type")

End If


'Defining Part required
If bDefiningPart = True Then

    'Initial set
    sDefiningPartNumber = ""
    sDefiningDocRev = ""
    
    'Defining Part not required
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") = "N/A" Then GoTo DefiningPartNotRequired
           
    'Defining Part doesn't exits
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") = "Does not exist" Then GoTo ReturnEmptyData
           
    'Defining Part already found
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") <> "" Then
        sString = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart")
        sDefiningPartNumber = Left(sString, Len(sString) - 2)
        sDefiningDocRev = Right(sString, 2)
    End If
    
    'Get Defining Part
    If sDefiningPartNumber = "" Then

        '******************************
        ' I need to update this section
        '******************************

        'Defining Part not required
        If sDatasetType <> "FLEXIBLE REPRESENTATION" And sDatasetType <> "CATALOG LIGHT VERSION" Then
            Call oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).SetItemByKey("DefiningPart", "N/A")
            GoTo DefiningPartNotRequired
        End If

        'Get value of Defining Part attribute
        sDefiningPartNumber = Trim(oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").GetItemByKey("Defining Part"))

        'If sDefiningPartNumber is "", we return empty data
        If sDefiningPartNumber = "" Then GoTo ReturnEmptyData
                
    End If
    
    'Get the latest rev
    If sDefiningDocRev = "" Then
        sDefiningDocRev = GetLatestDocRevision(Replace(sDefiningPartNumber, "/", ""), "LatestRev")
        bIsLatestRev = True
    End If
    
    'If rev is "", we return empty data
    If sDefiningDocRev = "" Then
        Call oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).SetItemByKey("DefiningPart", "Does not exist")
        GoTo ReturnEmptyData
    End If

    'Check if Defining Part document can be found in oListEnovia
    If Not oListEnovia.Exists(sDefiningPartNumber & sDefiningDocRev) Then

        'Get attributes
        
        Set oTemp = WebServiceAccessTool.GetENOVIADocumentAttributs(Replace(sDefiningPartNumber, "/", ""), sDefiningDocRev)
        Set sAttDefiningDoc = New clsCollection
        Call sAttDefiningDoc.InitializeWithDLLclsColObject(oTemp, sAttDefiningDoc)
        
        'Add to oListEnovia
        Set oItem = New clsCollection
        Call oItem.Add("EnoviaAttributes", sAttDefiningDoc)
        Call oItem.Add("DefiningPart", "N/A")
        Call oItem.Add("IsLatestRev", bIsLatestRev)
        Call oListEnovia.Add(sDefiningPartNumber & sDefiningDocRev, oItem)
        
        'Update Defining Part info of the DMU Doc entry
        Call oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).SetItemByKey("DefiningPart", sDefiningPartNumber & sDefiningDocRev)

    End If

    'Return attributes of Defining Part
    If sAttributeName = "" Then
        Set GetEnoviaAttributes = oListEnovia.GetItemByKey(sDefiningPartNumber & sDefiningDocRev).GetItemByKey("EnoviaAttributes")
    Else

        If oListEnovia.GetItemByKey(sDefiningPartNumber & sDefiningDocRev).GetItemByKey("EnoviaAttributes").Exists(sAttributeName) Then
            GetEnoviaAttributes = oListEnovia.GetItemByKey(sDefiningPartNumber & sDefiningDocRev).GetItemByKey("EnoviaAttributes").GetItemByKey(sAttributeName)
        Else
            GetEnoviaAttributes = "Part or Attribute doesn't exist"
        End If

    End If

    Exit Function
End If
    
DefiningPartNotRequired:
    
'Return attributes of DMU Part
If sAttributeName = "" Then
    Set GetEnoviaAttributes = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes")
Else

    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").Exists(sAttributeName) Then
        GetEnoviaAttributes = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("EnoviaAttributes").GetItemByKey(sAttributeName)
    Else
        GetEnoviaAttributes = "Part or Attribute doesn't exist"
    End If
End If

Exit Function

'************************************
ReturnEmptyData:
Dim oEmptyData
Set oEmptyData = CreateObject("Scripting.Dictionary")

If sAttributeName = "" Then
    Set GetEnoviaAttributes = oEmptyData
Else
    GetEnoviaAttributes = "Part or Attribute doesn't exist"
End If

End Function

Private Function GetOriginalPart(ByVal sDMUPartNumber As String, ByVal sDMURev As String) As String

Dim sPN, sDF, sNO, sBN As String
Dim sSearchResult As clsCollection

sPN = sDMUPartNumber
sDF = GetEnoviaAttributes(sDMUPartNumber, sDMURev, False, "Defining Part")
sNO = GetEnoviaAttributes(sDMUPartNumber, sDMURev, False, "Nomenclature")
sBN = Split(sDMUPartNumber, "-")(0)

Set sSearchResult = New clsCollection

End Function
Private Function GetLatestDocRevision(ByVal sPartNumber As String, ByVal sReqdRev As String) As String

    Dim oResults As clsCollection
    Dim oTemp
    Dim i As Integer
    Dim sPN As String
    Dim sRev As String
    Dim sStatus As String
    Dim iLatestRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sLatestRev As String
    Dim iLatestRelRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sLatestRelRev As String
    

    '1.1) Check if latest rev is already defined
    If sReqdRev = "LatestRev" And oListLatestRev.Exists(sPartNumber) Then
        GetLatestDocRevision = oListLatestRev.GetItemByKey(sPartNumber)
        Exit Function
    End If
    
    '1.2) Check if latest rel rev is already defined
    If sReqdRev = "LatestRelRev" And oListLatestRelRev.Exists(sPartNumber) Then
        GetLatestDocRevision = oListLatestRelRev.GetItemByKey(sPartNumber)
        Exit Function
    End If
    
    '1.3) Latest rel rev (if it exist, otherwise get latest rev)
    If sReqdRev = "LatestRelRevIfExist" And oListLatestRelRev.Exists(sPartNumber) Then
        If oListLatestRelRev.GetItemByKey(sPartNumber) <> "" Then
            GetLatestDocRevision = oListLatestRelRev.GetItemByKey(sPartNumber)
        Else
            GetLatestDocRevision = oListLatestRev.GetItemByKey(sPartNumber)
        End If
        Exit Function
    End If
    
    '2) Find the latest rev and latest rel rev using the Web Service
    'Search all document where the part number like sPartNumber*
    Set oTemp = WebServiceAccessTool.GetDocumentByBaseNumber(sPartNumber)
    If Not oTracking Is Nothing Then Call oTracking.WebServiceResultAnalysis("GetDocumentByBaseNumber", oTemp, sPartNumber, , bStopOnError)
    Set oResults = New clsCollection
    Call oResults.InitializeWithDLLclsColObject(oTemp, oResults)

    'No result found
    If oResults.Count = 0 Then
        GetLatestDocRevision = ""
        Exit Function
    End If

    'Scan all results and find the document with the latest revision and latest released revision
    sLatestRev = ""
    iLatestRev = 0
    sLatestRelRev = ""
    iLatestRelRev = 0
    
    For i = 1 To oResults.Count
        
        If oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber Then
            
            sRev = oResults.GetItem(i).GetItem("FIELD_DOCUMENT_REVISION")
            sStatus = oResults.GetItem(i).GetItem("FIELD_STATUS")
            
            'Latest Rev
            If 26 * Asc(Left(sRev, 1)) + Asc(Right(sRev, 1)) > iLatestRev Then
                sLatestRev = sRev
                iLatestRev = 26 * Asc(Left(sLatestRev, 1)) + Asc(Right(sLatestRev, 1))
            End If
            
            'Latest Rel rev
            If sStatus = "REL" And 26 * Asc(Left(sRev, 1)) + Asc(Right(sRev, 1)) > iLatestRelRev Then
                sLatestRelRev = sRev
                iLatestRelRev = 26 * Asc(Left(sLatestRelRev, 1)) + Asc(Right(sLatestRelRev, 1))
            End If
            
        End If
    Next

    '3) Add
    Call oListLatestRev.Add(sPartNumber, sLatestRev)
    Call oListLatestRelRev.Add(sPartNumber, sLatestRelRev)
    
    'Return value
    If sReqdRev = "LatestRev" Then
        GetLatestDocRevision = sLatestRev
    ElseIf sReqdRev = "LatestRelRev" Then
        GetLatestDocRevision = sLatestRelRev
    Else
        If sLatestRelRev <> "" Then
            GetLatestDocRevision = sLatestRelRev
        Else
            GetLatestDocRevision = sLatestRev
        End If
    End If

End Function




