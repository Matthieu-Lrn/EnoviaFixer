Attribute VB_Name = "frmAttributeAnalysis"
Attribute VB_Base = "0{E2F7B0D4-7DE6-4EE4-B216-F54E02349DE2}{D04CDC9E-F429-4D32-915C-76EF7E08855B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long
Private Declare Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) _
  As Long
Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

'******************Caption Messages*********************

Private Const sInitialText As String = "Double-click to execute search"
Private Const sDefaultCaption As String = "Double-click the search string to execute search"
Private Const sSearchMsg As String = "Searching for parts in the database...Please wait!"
Private Const sMgsTooLitteChars As String = "Enter atleast first 8 digits of Part number!"
Private Const sMsgAfterPopulatingList As String = "Select part numbers from the list." & vbCrLf & _
                                                  "'Ctrl +  Select' for multiselection."
Private Const sMsgNoSearchResult As String = "No results with this search string!"
Private Const sOkMsg As String = "Press OK to execute attribute analysis"
Private Const sAfterOKClick As String = "Analysing attributes...Please wait!"

'********************************************************

Private bBusy As Boolean
'********************************************************************************
'* Name: Form Search and Select Parts for Attribute analysis
'* Purpose:
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Language: VBA
'********************************************************************************
Public Sub InitializeForm()
    With Me
        .Height = dTitleBarHeight + 248.25
        .Width = 415
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 30
        .imgDSGlobe.Left = -3
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - 18
        .txtPartNumberSearch.ForeColor = &H80000006
        .txtPartNumberSearch.Text = sInitialText
        .StartUpPosition = Manual
        .LstBoxSearchResult.Height = .Height - 98.25
        .LstBoxSearchResult.Width = .Width - 15
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .RunTimeMsg.Top = .Height - 46.25
    End With
    With Me.LstBoxSearchResult
        .ColumnHeads = False
        .AddItem
        .List(0, 0) = "Part Number"
        .List(0, 1) = "Rev"
        .List(0, 2) = "Title"
    End With
    
    Me.Show vbModless
    Call MakeFormResizable
End Sub
Public Sub MakeFormResizable()
  '*** make a static form resizable
  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal
  
    hWnd = GetForegroundWindow
    'Get the basic window style
    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
    'Set the basic window styles
    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub
Private Sub UserForm_Resize()
    '*** Set Behaviour when resizing the form
    With Me
        If .Width < 380 Then .Width = 380
        If .Height < 215 Then .Height = 215
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 37
        .imgDSGlobe.Left = -3
        
        .cmdOK.Top = .Height - dTitleBarHeight - 31
        .cmdOK.Left = .Width - .cmdOK.Width - 18
        .LstBoxSearchResult.Height = .Height - 105.25
        .LstBoxSearchResult.Width = .Width - 22
        .RunTimeMsg.Top = .Height - 53.25
    End With
End Sub
Private Sub resetUserform()
        
    With Me
        If .LstBoxSearchResult.ListCount > 1 Then
           For i = .LstBoxSearchResult.ListCount To 2 Step -1
                .LstBoxSearchResult.RemoveItem i - 1
           Next
        End If
        .txtPartNumberSearch.ForeColor = &H80000006
        .txtPartNumberSearch.Text = sInitialText
        .cmdOK.ForeColor = -2147483631
        .cmdOK.Locked = True
        .RunTimeMsg.Caption = ""
    End With
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub
Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub
Private Sub LstBoxSearchResult_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Integer
    If Me.LstBoxSearchResult.Selected(0) = True Then Me.LstBoxSearchResult.Selected(0) = False
        
    Call cmdOKActivation
End Sub

Private Sub txtPartNumberSearch_Change()
    Dim sChr  As String
    Dim bCheck As Boolean
    
    If Not (txtPartNumberSearch.Text = "" Or txtPartNumberSearch.Text = sInitialText) Then
        On Error Resume Next
        txtPartNumberSearch.Text = UCase(txtPartNumberSearch.Text)
        On Error GoTo 0
        
        For i = 1 To Len(txtPartNumberSearch.Text)
            sChr = Mid(txtPartNumberSearch.Text, i, 1)
            If sChr Like "[-,A-Z,_]" Then
                bCheck = True
            ElseIf sChr Like "[0-9]" Then
                bCheck = True
            Else
                bCheck = False
            End If
            If Not bCheck Then
                txtPartNumberSearch.Text = Left(txtPartNumberSearch.Text, i - 1) & Replace(txtPartNumberSearch.Text, sChr, "", i, 1)
            End If
        Next
        
'        If Len(txtPartNumberSearch.Text) > 9 Then txtPartNumberSearch.Text = Left(txtPartNumberSearch.Text, 9)
    End If
End Sub

Private Sub txtPartNumberSearch_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '***Search for the Base nuber in Database & Populates the List in sorted order of Revision
    Dim oResult 'As clsCollection
    
    Dim oRevCollection As Collection

    Dim oRevDict As clsCollection
    Dim sTitle As String
    Dim sPartNumber As String, sRevision As String
    Dim i As Integer, j As Integer, k As Integer
    
    If bBusy Then Exit Sub
    
    'Clear cache
    Call StartWebServiceTool
    WebServiceAccessTool.ClearCache

    bBusy = True
    Set SysService = CATIA.SystemService
    Set oResult = Nothing
    Me.RunTimeMsg.Caption = sSearchMsg
    If Len(txtPartNumberSearch.Text) >= 8 Then
        
        '***GECK-191
        Select Case UCase(txtPartNumberSearch.Text) Like "PLMACTION*"
          Case True
            Dim oDumCol
            Set oDumCol = Nothing
            Call WebServiceAccessTool.GetPLMActionDocuments(txtPartNumberSearch.Text, oDumCol)
            Set oResult = New clsCollection
            For i = 1 To oDumCol.Count
               oResult.Add CStr(i), New clsCollection
               oResult.GetItem(CStr(i)).Add "FIELD_PART_NUMBER", oDumCol.GetKey(i)
               oResult.GetItem(CStr(i)).Add "FIELD_DOCUMENT_REVISION", oDumCol.GetItem(i)
            Next
          Case False
            Set oResult = WebServiceAccessTool.GetDocumentByBaseNumber(txtPartNumberSearch.Text)
        End Select
        '******
    Else
         Me.RunTimeMsg.Caption = sMgsTooLitteChars
         bBusy = False
         Exit Sub
    End If
    Set oRevDict = New clsCollection
    If Not oResult Is Nothing Then
        If oResult.Count > 0 Then
            For i = 1 To oResult.Count
                sTitle = ""
                sPartNumber = oResult.GetItem(i).GetItem("FIELD_PART_NUMBER")
                sRevision = oResult.GetItem(i).GetItem("FIELD_DOCUMENT_REVISION")
                On Error Resume Next
                sTitle = WebServiceAccessTool.GetENOVIADocumentAttributs(sPartNumber, sRevision).GetItem("Title")
                On Error GoTo 0
                If oRevDict.Exists(sPartNumber) Then
                    oRevDict.GetItem(sPartNumber).Add New Collection, sRevision
                    oRevDict.GetItem(sPartNumber).Item(sRevision).Add sRevision
                    oRevDict.GetItem(sPartNumber).Item(sRevision).Add sTitle
                Else
                    oRevDict.Add sPartNumber, New Collection
                    oRevDict.GetItem(sPartNumber).Add New Collection, sRevision
                    oRevDict.GetItem(sPartNumber).Item(sRevision).Add sRevision
                    oRevDict.GetItem(sPartNumber).Item(sRevision).Add sTitle
                End If
            Next
            For i = 1 To oRevDict.Count
                '*** Sort in increasing order of revsion
                Call SortRevCollection(oRevDict.GetItem(i))
            Next
        End If
    End If
    
    '***clear already filled list box
    If Me.LstBoxSearchResult.ListCount > 1 Then
       For i = Me.LstBoxSearchResult.ListCount To 2 Step -1
            Me.LstBoxSearchResult.RemoveItem i - 1
       Next
    End If
    '***Populating Result
    If oRevDict.Count > 0 Then
      With Me.LstBoxSearchResult
          k = 1
          For i = 1 To oRevDict.Count
            For j = 1 To oRevDict.GetItem(i).Count
                If Me.ChkLatestRev.Value = True Then j = oRevDict.GetItem(i).Count '*** if latest rev is chosen then take only last value of collection
                .AddItem
                .List(k, 0) = oRevDict.GetKey(i)
                .List(k, 1) = oRevDict.GetItem(i).Item(j).Item(1)
                .List(k, 2) = oRevDict.GetItem(i).Item(j).Item(2)
                k = k + 1
            Next
          Next
       End With
       Me.RunTimeMsg.Caption = sMsgAfterPopulatingList
    Else
'        Me.txtPartNumberSearch.ForeColor = &H80000006
'        Me.txtPartNumberSearch = sInitialText
        Me.RunTimeMsg.Caption = sMsgNoSearchResult
    End If

    bBusy = False
    If Me.RunTimeMsg.Caption = sSearchMsg Then Me.RunTimeMsg.Caption = sInitialText

End Sub

Private Sub txtPartNumberSearch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not bBusy Then
        If txtPartNumberSearch.Text = sInitialText Then txtPartNumberSearch.Text = ""
        txtPartNumberSearch.ForeColor = &H80000008
        Me.RunTimeMsg.Caption = sDefaultCaption
    End If
End Sub
Private Sub cmdOKActivation()
    '***If atleast on row is selected in list box then activate OK command
    Dim bCheck As Boolean

    
    If Me.LstBoxSearchResult.ListCount > 1 Then
        For i = 0 To Me.LstBoxSearchResult.ListCount
            If Me.LstBoxSearchResult.Selected(i) = True Then
                bCheck = True
                Exit For
            End If
        Next
    End If
    If bCheck Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
        Me.cmdOK.SetFocus
        Me.RunTimeMsg.Caption = sOkMsg
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        If Me.RunTimeMsg.Caption = sOkMsg Then Me.RunTimeMsg.Caption = ""
    End If

End Sub
Private Sub cmdOK_Click()
    Dim oAnalysisCol As New Collection
    If Not bBusy Then
        bBusy = True
        Me.RunTimeMsg.Caption = sAfterOKClick
        For i = 0 To Me.LstBoxSearchResult.ListCount - 1
            If Me.LstBoxSearchResult.Selected(i) = True Then
                oAnalysisCol.Add New Collection
                
                oAnalysisCol.Item(oAnalysisCol.Count).Add Me.LstBoxSearchResult.List(i, 0), "PartNumber"
                oAnalysisCol.Item(oAnalysisCol.Count).Add Me.LstBoxSearchResult.List(i, 1), "Revision"
                oAnalysisCol.Item(oAnalysisCol.Count).Add Me.LstBoxSearchResult.List(i, 2), "Title"
            End If
        Next
        Call RunAttributeAnalysis(oAnalysisCol)
    
        '***Log File
            If oAnalysisCol.Count > 0 Then
            For i = 1 To oAnalysisCol.Count
                Call AddToLogFile("Attribute Analysis", , oAnalysisCol.Item(i).Item("PartNumber"), oAnalysisCol.Item(i).Item("Revision"))
            Next
        End If
        
    End If
    bBusy = False
    For i = 0 To Me.LstBoxSearchResult.ListCount
        Me.LstBoxSearchResult.Selected(i) = False
    Next
    cmdOKActivation
    Me.RunTimeMsg.Caption = sDefaultCaption
End Sub

Private Sub SortRevCollection(ByRef oRevisionCol As Collection)
    '*** Sort
    Dim oDummyList As Collection
    Dim iComp As Integer, jComp As Integer
    Dim sRv As String
    Dim i As Integer, j As Integer
    For i = 1 To oRevisionCol.Count - 1
            sRv = Replace(oRevisionCol.Item(i).Item(1), "-", "0")
            iComp = 0
            Do While sRv <> ""
                iComp = iComp + Asc(sRv)
                sRv = Right(sRv, Len(sRv) - 1)
            Loop
        For j = i + 1 To oRevisionCol.Count
            sRv = Replace(oRevisionCol.Item(j).Item(1), "-", "0")
            jComp = 0
            Do While sRv <> ""
               jComp = jComp + Asc(sRv)
               sRv = Right(sRv, Len(sRv) - 1)
            Loop
            If iComp > jComp Then
                Set oDummyList = oRevisionCol.Item(j)
                oRevisionCol.Remove j
                oRevisionCol.Add oDummyList, oDummyList.Item(1), i
            End If
        Next j
    Next i
End Sub
