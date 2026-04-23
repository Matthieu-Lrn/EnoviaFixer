Attribute VB_Name = "BDIfunctions"


'****Mandatory for JSON webservices
Private Enum enumBDIObject
        BDIpart = 0
        BDIdoc = 1
        BDIrv = 2
End Enum

Public Function GetPrimaryDocFromPart(ByVal sPartNumber As String, ByVal sPartRev As String) As String

Dim sJsonInfo As String, sState As String
Dim vPartInfo As Variant
Dim bExitByUser As Boolean
Dim oPrimaryDoc

'Get the part info using the web service
sJsonInfo = ""
sJsonInfo = WebServiceAccessTool.JSON_BDIPartInfoFromPnRev(sPartNumber, sPartRev, bExitByUser, , oPasswordMgr.GetEncryptedPassword)

'Transform the Json in a variant
Call mdlParseJSON.ParseJSON(sJsonInfo, vPartInfo, sState)

'Web service doesn't return anything or return format is not good, return nothing
If TypeName(vPartInfo) <> "Variant()" Then GoTo ErrorFound
If UBound(vPartInfo) < 0 Then GoTo ErrorFound
If TypeName(vPartInfo(0)) <> "Dictionary" Then GoTo ErrorFound
If Not vPartInfo(0).Exists("primary_document") Then GoTo ErrorFound
If TypeName(vPartInfo(0)("primary_document")) <> "Dictionary" Then GoTo ErrorFound
Set oPrimaryDoc = vPartInfo(0)("primary_document")

If Not oPrimaryDoc.Exists("num") Then GoTo ErrorFound
If Not oPrimaryDoc.Exists("rev") Then GoTo ErrorFound


'Return OK
GetPrimaryDocFromPart = oPrimaryDoc("num") & oPrimaryDoc("rev")
Exit Function

'Return error
ErrorFound:
GetPrimaryDocFromPart = ""

End Function






