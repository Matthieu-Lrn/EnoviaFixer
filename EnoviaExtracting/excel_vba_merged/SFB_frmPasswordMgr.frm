Attribute VB_Name = "SFB_frmPasswordMgr"
Attribute VB_Base = "0{AB35784E-3252-4863-8DB2-8B496B5A96B5}{FF54FCEE-925D-48E6-9D60-8D07DE4615FD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Public SFB_QuitByUser As Boolean
Public SFB_EncryptedString As String

'If user presses enter then simulate SFB_cmdOK_Click
Private Sub SFB_txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Me.txtUserName.Text <> "" Then Call SFB_cmdOK_Click
End Sub

Private Sub SFB_UserForm_Activate()
    With Me
        .txtPassword.SetFocus
    End With
End Sub
'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub SFB_UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    If CloseMode = 0 Then Call SFB_cmdCancel_Click
End Sub
Private Sub SFB_cmdCancel_Click()
    SFB_QuitByUser = True
    Me.Hide
End Sub
Private Sub SFB_cmdOK_Click()
    SFB_EncryptedString = SFB_Base64Encode(Me.txtUserName.Text + ":" + Me.txtPassword.Text)
    Me.Hide
End Sub

Function SFB_Base64Encode(SFB_sText)

    Dim SFB_oXML, SFB_oNode
    Set SFB_oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set SFB_oNode = SFB_oXML.CreateElement("base64")
    SFB_oNode.DataType = "bin.base64"
    SFB_oNode.nodeTypedValue = SFB_Stream_StringToBinary(SFB_sText)
    SFB_Base64Encode = SFB_oNode.Text
    Set SFB_oNode = Nothing
    Set SFB_oXML = Nothing
    
End Function
'SFB_Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data

Function SFB_Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim SFB_BinaryStream 'As New Stream
  Set SFB_BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  SFB_BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  SFB_BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  SFB_BinaryStream.Open
  SFB_BinaryStream.WriteText Text

  'Change stream type To binary
  SFB_BinaryStream.Position = 0
  SFB_BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  SFB_BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  SFB_Stream_StringToBinary = SFB_BinaryStream.Read

  Set SFB_BinaryStream = Nothing
End Function



'SFB_Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To SFB_a string

Function SFB_Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim SFB_BinaryStream 'As New Stream
  Set SFB_BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  SFB_BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  SFB_BinaryStream.Open
  SFB_BinaryStream.Write Binary

  'Change stream type To binary
  SFB_BinaryStream.Position = 0
  SFB_BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  SFB_BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  SFB_Stream_BinaryToString = SFB_BinaryStream.ReadText
  Set SFB_BinaryStream = Nothing
End Function



