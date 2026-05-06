Attribute VB_Name = "SFB_clsAttributesList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oListEnovia As SFB_clsCollection
Private oListLatestRev As SFB_clsCollection
Private oListLatestRelRev As SFB_clsCollection

Private Sub Class_Initialize()
    Set oListEnovia = New SFB_clsCollection
    Set oListLatestRev = New SFB_clsCollection
    Set oListLatestRelRev = New SFB_clsCollection
End Sub

Private Sub Class_Terminate()
    Set oListEnovia = Nothing
    Set oListLatestRev = Nothing
    Set oListLatestRelRev = Nothing
End Sub

Public Function GetEnoviaAttributes(ByVal sDMUPartNumber As String, ByVal sDMUDocRev As String, Optional bDefiningPart = False, Optional sAttributeName = "") As Variant

Dim SFB_oTemp
Dim sAttDMUDoc As SFB_clsCollection
Dim sAttDefiningDoc As SFB_clsCollection
Dim SFB_oItem As SFB_clsCollection
Dim sDefiningPartNumber As String
Dim sDefiningDocRev As String
Dim sDatasetType As String
Dim SFB_sString As String

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
    Set SFB_oTemp = SFB_WebServiceAccessTool.GetENOVIADocumentAttributs(sDMUPartNumber, sDMUDocRev)
    Set sAttDMUDoc = New SFB_clsCollection
    Call sAttDMUDoc.InitializeWithDLLclsColObject(SFB_oTemp, sAttDMUDoc)
    
    'Add to oListEnovia
    Set SFB_oItem = New SFB_clsCollection
    Call SFB_oItem.Add("EnoviaAttributes", sAttDMUDoc)
    Call SFB_oItem.Add("DefiningPart", "")
    Call oListEnovia.Add(sDMUPartNumber & sDMUDocRev, SFB_oItem)

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
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") = "SFB_N/A" Then GoTo DefiningPartNotRequired
           
    'Defining Part doesn't exits
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") = "Does not exist" Then GoTo ReturnEmptyData
           
    'Defining Part already found
    If oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart") <> "" Then
        SFB_sString = oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).GetItemByKey("DefiningPart")
        sDefiningPartNumber = Left(SFB_sString, Len(SFB_sString) - 2)
        sDefiningDocRev = Right(SFB_sString, 2)
    End If
    
    'Get Defining Part
    If sDefiningPartNumber = "" Then

        '******************************
        ' I need to update this section
        '******************************

        'Defining Part not required
        If sDatasetType <> "FLEXIBLE REPRESENTATION" And sDatasetType <> "CATALOG LIGHT VERSION" Then
            Call oListEnovia.GetItemByKey(sDMUPartNumber & sDMUDocRev).SetItemByKey("DefiningPart", "SFB_N/A")
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
        
        Set SFB_oTemp = SFB_WebServiceAccessTool.GetENOVIADocumentAttributs(Replace(sDefiningPartNumber, "/", ""), sDefiningDocRev)
        Set sAttDefiningDoc = New SFB_clsCollection
        Call sAttDefiningDoc.InitializeWithDLLclsColObject(SFB_oTemp, sAttDefiningDoc)
        
        'Add to oListEnovia
        Set SFB_oItem = New SFB_clsCollection
        Call SFB_oItem.Add("EnoviaAttributes", sAttDefiningDoc)
        Call SFB_oItem.Add("DefiningPart", "SFB_N/A")
        Call SFB_oItem.Add("IsLatestRev", bIsLatestRev)
        Call oListEnovia.Add(sDefiningPartNumber & sDefiningDocRev, SFB_oItem)
        
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

Dim SFB_sPN, sDF, sNO, sBN As String
Dim sSearchResult As SFB_clsCollection

SFB_sPN = sDMUPartNumber
sDF = GetEnoviaAttributes(sDMUPartNumber, sDMURev, False, "Defining Part")
sNO = GetEnoviaAttributes(sDMUPartNumber, sDMURev, False, "Nomenclature")
sBN = Split(sDMUPartNumber, "-")(0)

Set sSearchResult = New SFB_clsCollection

End Function
Private Function GetLatestDocRevision(ByVal SFB_sPartNumber As String, ByVal sReqdRev As String) As String

    Dim oResults As SFB_clsCollection
    Dim SFB_oTemp
    Dim SFB_i As Integer
    Dim SFB_sPN As String
    Dim SFB_sRev As String
    Dim sStatus As String
    Dim iLatestRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sLatestRev As String
    Dim iLatestRelRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sLatestRelRev As String
    

    '1.1) Check if latest rev is already defined
    If sReqdRev = "LatestRev" And oListLatestRev.Exists(SFB_sPartNumber) Then
        GetLatestDocRevision = oListLatestRev.GetItemByKey(SFB_sPartNumber)
        Exit Function
    End If
    
    '1.2) Check if latest rel rev is already defined
    If sReqdRev = "LatestRelRev" And oListLatestRelRev.Exists(SFB_sPartNumber) Then
        GetLatestDocRevision = oListLatestRelRev.GetItemByKey(SFB_sPartNumber)
        Exit Function
    End If
    
    '1.3) Latest rel rev (if it exist, otherwise get latest rev)
    If sReqdRev = "LatestRelRevIfExist" And oListLatestRelRev.Exists(SFB_sPartNumber) Then
        If oListLatestRelRev.GetItemByKey(SFB_sPartNumber) <> "" Then
            GetLatestDocRevision = oListLatestRelRev.GetItemByKey(SFB_sPartNumber)
        Else
            GetLatestDocRevision = oListLatestRev.GetItemByKey(SFB_sPartNumber)
        End If
        Exit Function
    End If
    
    '2) Find the latest rev and latest rel rev using the Web Service
    'Search all document where the part number like SFB_sPartNumber*
    Set SFB_oTemp = SFB_WebServiceAccessTool.GetDocumentByBaseNumber(SFB_sPartNumber)
    Set oResults = New SFB_clsCollection
    Call oResults.InitializeWithDLLclsColObject(SFB_oTemp, oResults)

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
    
    For SFB_i = 1 To oResults.Count
        
        If oResults.GetItem(SFB_i).GetItem("FIELD_PART_NUMBER") = SFB_sPartNumber Then
            
            SFB_sRev = oResults.GetItem(SFB_i).GetItem("FIELD_DOCUMENT_REVISION")
            sStatus = oResults.GetItem(SFB_i).GetItem("FIELD_STATUS")
            
            'Latest Rev
            If 5 * Asc(Left(SFB_sRev, 1)) + Asc(Right(SFB_sRev, 1)) > iLatestRev Then
                sLatestRev = SFB_sRev
                iLatestRev = 5 * Asc(Left(sLatestRev, 1)) + Asc(Right(sLatestRev, 1))
            End If
            
            'Latest Rel rev
            If sStatus = "REL" And 5 * Asc(Left(SFB_sRev, 1)) + Asc(Right(SFB_sRev, 1)) > iLatestRelRev Then
                sLatestRelRev = SFB_sRev
                iLatestRelRev = 5 * Asc(Left(sLatestRelRev, 1)) + Asc(Right(sLatestRelRev, 1))
            End If
            
        End If
    Next

    '3) Add
    Call oListLatestRev.Add(SFB_sPartNumber, sLatestRev)
    Call oListLatestRelRev.Add(SFB_sPartNumber, sLatestRelRev)
    
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




