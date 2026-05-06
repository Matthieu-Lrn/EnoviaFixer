Attribute VB_Name = "SFB_clsPasswordMgr"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private SFB_bExitByUser As Boolean
Private sUsername As String
Private sEncryptedString As String ' Username:Password
            
Private Sub GetUserNetworkName()

'Get username from network
Dim WSHnet

Set WSHnet = CreateObject("WScript.Network")
sUsername = WSHnet.UserName
Set WSHnet = Nothing

End Sub

Public Sub ResetPassword()

sEncryptedString = ""
sUsername = ""
SFB_bExitByUser = False

End Sub

Public Function GetEncryptedPassword() As String

Dim ofrmPasswordMgr 'As UserForm
Dim sData As String

If sEncryptedString <> "" Then

    'Return value
    GetEncryptedPassword = sEncryptedString


Else

    'Get user name from network
    Call GetUserNetworkName
        
    'Userform
    Set ofrmPasswordMgr = New SFB_frmPasswordMgr
    ofrmPasswordMgr.txtUserName.Text = sUsername
    ofrmPasswordMgr.lblMsg.Caption = ""
    
    Do
        'Show userform
        ofrmPasswordMgr.txtPassword.Text = ""
        ofrmPasswordMgr.Show
        
        'Quit by user
        If ofrmPasswordMgr.SFB_QuitByUser = True Then
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            Exit Function
        End If
        
        'Check if username & password are valid
        sData = ""
        sData = SFB_WebServiceAccessTool.JSON_BDIGetUserInfo(SFB_bExitByUser, , ofrmPasswordMgr.SFB_EncryptedString)
        
        'Exit by user
        If SFB_bExitByUser = True Then
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            Exit Function
        End If
        
        'Username & password are valid
        If sData Like "*badge*" Then
            sEncryptedString = ofrmPasswordMgr.SFB_EncryptedString
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            GetEncryptedPassword = sEncryptedString
            Exit Function
            
        'Wrong password
        ElseIf SFB_WebServiceAccessTool.WebResponseExceptionStatus = "Unauthorized" Then
            ofrmPasswordMgr.lblMsg.Caption = "Wrong password"
        
        'Problem connecting to BDI
        Else
            Call MsgBox("Can't connect to BDI. Process aborted", 16, "PVR Sync Tool")
            sEncryptedString = ""
            GetEncryptedPassword = sEncryptedString
        End If
    Loop
End If

End Function



