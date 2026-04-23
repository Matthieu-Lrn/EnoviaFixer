Attribute VB_Name = "clsPasswordMgr"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private bExitByUser As Boolean
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
bExitByUser = False

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
    Set ofrmPasswordMgr = New frmPasswordMgr
    ofrmPasswordMgr.txtUserName.Text = sUsername
    ofrmPasswordMgr.lblMsg.Caption = ""
    
    Do
        'Show userform
        ofrmPasswordMgr.txtPassword.Text = ""
        ofrmPasswordMgr.Show
        
        'Quit by user
        If ofrmPasswordMgr.QuitByUser = True Then
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            Exit Function
        End If
        
        'Check if username & password are valid
        sData = ""
        sData = WebServiceAccessTool.JSON_BDIGetUserInfo(bExitByUser, , ofrmPasswordMgr.EncryptedString)
        
        'Exit by user
        If bExitByUser = True Then
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            Exit Function
        End If
        
        'Username & password are valid
        If sData Like "*badge*" Then
            sEncryptedString = ofrmPasswordMgr.EncryptedString
            Unload ofrmPasswordMgr
            Set ofrmPasswordMgr = Nothing
            GetEncryptedPassword = sEncryptedString
            Exit Function
            
        'Wrong password
        ElseIf WebServiceAccessTool.WebResponseExceptionStatus = "Unauthorized" Then
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



