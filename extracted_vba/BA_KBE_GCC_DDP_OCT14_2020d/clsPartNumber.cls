Attribute VB_Name = "clsPartNumber"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'________________________________________________________________________________________
'________________________________________________________________________________________
'________________________________________________________________________________________
'________________________________________________________________________________________
'________________________________________________________________________________________
' Usage:
'
'    Dim oPartNumber As New clsPartNumber
'    Set oPartNumber = oPartNumber.Initialize(sInputPartNumber)
'
'________________________________________________________________________________________
'________________________________________________________________________________________
'________________________________________________________________________________________
'________________________________________________________________________________________
Option Explicit

Private Type TPartNumber
    PartNumber As String
    BaseNumber As String
    DashNumber As String
    IsColorCoded As Boolean
    Tag As String
    NumericDash As String
    NumericDashParity As String 'Result can be "ODD" or "EVEN"
    Suffix As String
End Type

Private this As TPartNumber


Public Property Get Initialize(Optional sInput As String) As clsPartNumber
    
    With New clsPartNumber
        
        ' Initialize values:
        .PartNumber = sInput
        .BaseNumber = ""
        .DashNumber = ""
        .IsColorCoded = False
        .Tag = "Unknown"
        .NumericDash = ""
        .NumericDashParity = ""
        

        ' Define Part Number attributes:
        If UBound(Split(sInput, "-")) > -1 Then
            
            .PartNumber = sInput
            
            ' Manage G7000 internal Part Number:
            If sInput Like "G#######" _
                Or sInput Like "G########" _
                Or sInput Like "G#######-*" _
                Or sInput Like "G########-*" _
                Or sInput Like "GE#######" _
                Or sInput Like "GE#######-*" _
                Or sInput Like "GE########" _
                Or sInput Like "GE########-*" _
                Or sInput Like "GV#######" _
                Or sInput Like "GV#######-*" _
                Or sInput Like "GV########" _
                Or sInput Like "GV########-*" _
                Or sInput Like "GF#######" _
                Or sInput Like "GF#######-*" _
                Or sInput Like "GF########" _
                Or sInput Like "GF########-*" _
                Or sInput Like "G-#######" _
                Or sInput Like "G-#######-*" Then
                
                If Left(sInput, 2) = "G-" Then
                    
                    If UBound(Split(sInput, "-")) > 0 Then
                        .BaseNumber = Split(sInput, "-")(0) & "-" & Split(sInput, "-")(1)
                    Else
                        .BaseNumber = sInput
                    End If
                
                Else
                    
                    If UBound(Split(sInput, "-")) > -1 Then
                        .BaseNumber = Split(sInput, "-")(0)
                    Else
                        .BaseNumber = sInput
                    End If
                    
                End If
            
                .Tag = "G7000"
                .DashNumber = Replace(.PartNumber, .BaseNumber, "")
            
            
            ' Manage G7000 String Part Number:
            ElseIf sInput Like "#######-E###*" _
                Or sInput Like "?????-#######-E###*" Then
                
                If sInput Like "?????-#######-E###*" Then
                    .BaseNumber = Split(sInput, "-")(0) & "-" & Split(sInput, "-")(1)
                Else
                    .BaseNumber = Split(sInput, "-")(0)
                End If
                
                .Tag = "String"
                .DashNumber = Replace(.PartNumber, .BaseNumber, "")
                
            ' Manage G5000 internal Part Number:
            ElseIf sInput Like "700#######-*" _
                Or sInput Like "700D#######-*" _
                Or sInput Like "700########-*" _
                Or sInput Like "700D########-*" Then
                
                .BaseNumber = Split(sInput, "-")(0)
                .Tag = "G5000"
                .DashNumber = Replace(.PartNumber, .BaseNumber, "")
            
            ' Manage SVCDs
            ElseIf IsNumeric(Replace(sInput, "-", "")) = True Then
            
                .BaseNumber = sInput
                .Tag = "NoDash (SVCD)"
                
            ' Manage S9000 Numbers
            ElseIf Split(sInput, "-")(0) Like "S####" _
                Or Split(sInput, "-")(0) Like "S#####" Then
            
                .BaseNumber = sInput
                .Tag = "NoDash (S Number)"
                
            ' Manage SC Numbers
            ElseIf Split(sInput, "-")(0) Like "SC#######" _
                Or Split(sInput, "-")(0) Like "SC#####" Then
            
                .BaseNumber = sInput
                .Tag = "NoDash (SC)"
                
            Else
                .BaseNumber = sInput
                .Tag = "NoDash (Unknown)"
            End If
        Else
            .BaseNumber = sInput
            .Tag = "NoDash(0x-)"
        End If
        
        
        ' Detect Numerical Dash:
        If .DashNumber <> "" Then
            .NumericDash = GetNumericDash(.DashNumber)
        End If
        
        ' Detect Numerical Dash Parity
        If .NumericDash <> "" Then
            If CInt(.NumericDash) Mod 2 > 0 Then
                .NumericDashParity = "ODD"
            Else
                .NumericDashParity = "EVEN"
            End If
        
        End If
        
        ' Detect Color coded Parts:
        If .PartNumber <> "" Then
            .IsColorCoded = IsColorCodedPart(.PartNumber)
        End If

        ' Detect the Suffix:
        If .NumericDash <> "" Then
            .Suffix = GetSuffixFromDashNumber(.NumericDash, .DashNumber)
        End If

        
        
        ' Return the newly created instance:
        Set Initialize = .Self
        
    End With
End Property

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function GetNumericDash(ByVal sInputPartNumber As String) As String
    
    ' Default Value
    GetNumericDash = ""
    
    If sInputPartNumber <> "" Then
        
        ' Keep only first part on PartNumber when split by "-"
        Dim sLocalInput As String
        sLocalInput = sInputPartNumber
        If Left(sLocalInput, 1) = "-" Then sLocalInput = Right(sLocalInput, Len(sLocalInput) - 1)
        If UBound(Split(sLocalInput, "-")) > -1 Then sLocalInput = "-" & Split(sLocalInput, "-")(0)
        
        Dim i As Integer
        Dim sTempDash As String
        sTempDash = ""
        For i = Len(sLocalInput) To 1 Step -1
        
            If IsNumeric(Mid(sLocalInput, i, 1)) Then
                If Len(sTempDash) < 3 Then
                    sTempDash = Mid(sLocalInput, i, 1) & sTempDash
                Else
                    sTempDash = ""
                End If
                
            ' Exit if finding an Hyphen after finding 3 digits:
            ElseIf Mid(sLocalInput, i, 1) = "-" And Len(sTempDash) = 3 Then
                Exit For
                
            ' Reset if Non-digit is found before reaching 3 consecutive digits
            ElseIf Len(sTempDash) <= 3 Then
                sTempDash = ""
                
            End If
        Next i
        GetNumericDash = sTempDash
    End If
End Function
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function IsColorCodedPart(ByVal sInput As String) As Boolean
    
    If sInput <> "" Then
        ' Build RegEx:
        Dim regex, regExp_Matches
        Set regex = CreateObject("vbscript.regexp")
        regex.IgnoreCase = True  'ignoring cases while regex engine performs the search.
        regex.Pattern = "(|-\.*)(-YYYY|-XXXX|-YYYY-Y|-XXXX-X)(|-\.*)"
        regex.Global = False     'restricting regex to find only first match.
        Set regExp_Matches = regex.Execute(sInput)

        ' Validate Revision Format using RegEx:
        If regExp_Matches.Count > 0 Then
            IsColorCodedPart = True
        End If

        Set regex = Nothing
    End If
End Function
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function GetSuffixFromDashNumber(ByVal sNumericalDashNumber As String, ByVal sDashNumber As String) As String
    
    ' Default Value
    GetSuffixFromDashNumber = ""
    
    ' Remove any Color Coding Patterns:
    sDashNumber = Replace(sDashNumber, "X-XXXX", "")
    sDashNumber = Replace(sDashNumber, "Y-YYYY", "")
    sDashNumber = Replace(sDashNumber, "-XXXX", "")
    sDashNumber = Replace(sDashNumber, "-YYYY", "")
    sDashNumber = Replace(sDashNumber, "XXXX", "")
    sDashNumber = Replace(sDashNumber, "YYYY", "")
    
    
    If sDashNumber <> "" Then
        
        On Error Resume Next
        Dim regex, regExp_Matches ' As RegularExpression
        Set regex = CreateObject("vbscript.regexp")
        regex.IgnoreCase = True
        regex.Global = False
                
        ' SP:
        regex.Pattern = "-" & sNumericalDashNumber & ".*(?:[A-Z]{2,4}|D|I)(?=\d{2,3})"
        Set regExp_Matches = regex.Execute(sDashNumber)
        If Err.Number = 0 Then
            If regExp_Matches.Count > 0 Then
                
                Dim sNewSuffix As String, sSuffix As String
                sSuffix = regExp_Matches.Item(0).Value
                sNewSuffix = ""
                
                ' Get only the last non numerical characters except "_" and "-"
                Dim i As Integer
                For i = Len(sSuffix) To 1 Step -1
                    If Not IsNumeric(Mid(sSuffix, i, 1)) Then
                        If Mid(sSuffix, i, 1) <> "_" And Mid(sSuffix, i, 1) <> "-" Then
                            sNewSuffix = Mid(sSuffix, i, 1) & sNewSuffix
                        End If
                    Else
                        Exit For
                    End If
                Next
                
                GetSuffixFromDashNumber = sNewSuffix
            
            End If
        End If
        Set regex = Nothing
    End If
End Function
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Public Property Get Self() As clsPartNumber
    Set Self = Me
End Property

Public Property Get PartNumber() As String
    PartNumber = this.PartNumber
End Property

Public Property Let PartNumber(ByVal Value As String)
    this.PartNumber = Value
End Property

Public Property Get BaseNumber() As String
    BaseNumber = this.BaseNumber
End Property

Public Property Let BaseNumber(ByVal Value As String)
    this.BaseNumber = Value
End Property

Public Property Get DashNumber() As String
    DashNumber = this.DashNumber
End Property

Public Property Let DashNumber(ByVal Value As String)
    this.DashNumber = Value
End Property

Public Property Get Suffix() As String
    Suffix = this.Suffix
End Property

Public Property Let Suffix(ByVal Value As String)
    this.Suffix = Value
End Property

Public Property Get Tag() As String
    Tag = this.Tag
End Property

Public Property Let Tag(ByVal Value As String)
    this.Tag = Value
End Property

Public Property Get IsColorCoded() As Boolean
    IsColorCoded = this.IsColorCoded
End Property

Public Property Let IsColorCoded(ByVal Value As Boolean)
    this.IsColorCoded = Value
End Property

Public Property Get NumericDash() As String
    NumericDash = this.NumericDash
End Property

Public Property Let NumericDash(ByVal Value As String)
    this.NumericDash = Value
End Property

Public Property Get NumericDashParity() As String
    NumericDashParity = this.NumericDashParity
End Property

Public Property Let NumericDashParity(ByVal Value As String)
    this.NumericDashParity = Value
End Property

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################


