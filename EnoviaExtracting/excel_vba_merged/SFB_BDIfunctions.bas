Attribute VB_Name = "SFB_BDIfunctions"


'****Mandatory for JSON webservices
Private Enum SFB_enumBDIObject
        BDIpart = 0
        BDIdoc = 1
        BDIrv = 2
End Enum

Public Function SFB_GetPrimaryDocFromPart(ByVal SFB_sPartNumber As String, ByVal sPartRev As String) As String

Dim SFB_sJsonInfo As String, sState As String
Dim SFB_vPartInfo As Variant
Dim SFB_bExitByUser As Boolean
Dim SFB_oPrimaryDoc

'Get the part info using the web service
SFB_sJsonInfo = ""
SFB_sJsonInfo = SFB_WebServiceAccessTool.JSON_BDIPartInfoFromPnRev(SFB_sPartNumber, sPartRev, SFB_bExitByUser, , SFB_oPasswordMgr.GetEncryptedPassword)

'Transform the Json in SFB_a variant
Call SFB_mdlParseJSON.SFB_ParseJSON(SFB_sJsonInfo, SFB_vPartInfo, sState)

'Web service doesn't return anything or return format is not good, return nothing
If TypeName(SFB_vPartInfo) <> "Variant()" Then GoTo ErrorFound
If UBound(SFB_vPartInfo) < 0 Then GoTo ErrorFound
If TypeName(SFB_vPartInfo(0)) <> "Dictionary" Then GoTo ErrorFound
If Not SFB_vPartInfo(0).Exists("primary_document") Then GoTo ErrorFound
If TypeName(SFB_vPartInfo(0)("primary_document")) <> "Dictionary" Then GoTo ErrorFound
Set SFB_oPrimaryDoc = SFB_vPartInfo(0)("primary_document")

If Not SFB_oPrimaryDoc.Exists("num") Then GoTo ErrorFound
If Not SFB_oPrimaryDoc.Exists("rev") Then GoTo ErrorFound


'Return OK
SFB_GetPrimaryDocFromPart = SFB_oPrimaryDoc("num") & SFB_oPrimaryDoc("rev")
Exit Function

'Return error
ErrorFound:
SFB_GetPrimaryDocFromPart = ""

End Function






