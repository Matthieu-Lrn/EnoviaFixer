Attribute VB_Name = "frmPVRSync"
Attribute VB_Base = "0{34C42739-2C63-4475-9DF2-C93BEDEA21DB}{53FF035C-C25B-43DE-87A8-228F7D74B08A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Option Explicit

Public SelectedDoc As String
Private Const Manual As Integer = 0
Private sReportPath As String

Public Sub InitializeForm(ByVal sPartNumber As String, ByVal sTitle As String, ByVal sDefaultValue As String, ByVal sPath As String, ByVal oEffDocReport As DOMDocument60, ByVal sTail As String)

Dim iIndex As Integer
Dim sDocNumber As String
Dim sDocRev As String
Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode


'Initial set
iIndex = 0
sReportPath = sPath

'Add BSF to combo box
ComboBox_Doc.AddItem ("BSF")

'Populate the combo box with all the document revison with status closed found in oEffDocReport
Set oNodeList = oEffDocReport.SelectNodes("//BDIDocument[@Status='CLOSED']")
For Each oNode In oNodeList
    
    sDocNumber = oNode.Attributes.getNamedItem("DocNumber").nodeValue
    sDocRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    ComboBox_Doc.AddItem (sDocNumber & " " & sDocRev)
    
    If sDefaultValue = sDocNumber & sDocRev Then
        iIndex = ComboBox_Doc.ListCount - 1
    End If

Next

With Me
    .lblCI.Caption = sPartNumber
    .lblTitle.Caption = sTitle
    .lblHeader.Caption = lblHeader.Caption & " " & sTail
    .StartUpPosition = Manual
    .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
    .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
          
    .ComboBox_Doc.ListIndex = iIndex
    .cmdOK.SetFocus
    .Show
End With

End Sub

Private Sub cmdCancel_Click()

SelectedDoc = ""

Me.Hide

End Sub

Private Sub cmdHelp_Click()

Dim lSuccess

lSuccess = ShellExecute(0, "Open", sReportPath, , , 1)

End Sub

Private Sub cmdOK_Click()

SelectedDoc = Me.ComboBox_Doc.Value

Me.Hide

End Sub


