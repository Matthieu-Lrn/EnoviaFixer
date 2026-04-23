Attribute VB_Name = "BATemplate"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public Sub PopulateNomenclature(ByVal ReferencePart As ProductStructureTypeLib.Product)
    ReferencePart.Nomenclature = ReferencePart.PartNumber
End Sub

Public Sub AddGeometricalSets(ByVal oPart As MECMOD.Part, ByVal oSelection As INFITF.Selection)
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Pilot_Holes"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 49
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Master_Geometry"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 3
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "References"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 4
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Geometrical Set for 3D Annotations"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 70
    
    oSelection.Clear
    Set oSubGeometricalSet = oGeometricalSet.HybridBodies.Add
    oSubGeometricalSet.Name = "No Show Geometry"
    oSelection.Add oSubGeometricalSet
    oSelection.VisProperties.SetShow catVisPropertyNoShowAttr
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Flat Pattern"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 47
    
    oSelection.Clear
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Construction_Geometry"
    oSelection.Add oGeometricalSet
    oSelection.VisProperties.SetLayer catVisLayerBasic, 46
    
    oSelection.Clear
    Set oSubGeometricalSet = oGeometricalSet.HybridBodies.Add
    oSubGeometricalSet.Name = "Pockets"
    Set oSubGeometricalSet = oGeometricalSet.HybridBodies.Add
    oSubGeometricalSet.Name = "Ruled_Surfaces"
    Set oSubGeometricalSet = oGeometricalSet.HybridBodies.Add
    oSubGeometricalSet.Name = "Sketches_Not_Used_By_PartBody"
    
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Skeleton"
    oPart.HybridShapeFactory.GSMVisibility oPart.CreateReferenceFromObject(oGeometricalSet), 0
    
    Set oGeometricalSet = oPart.HybridBodies.Add
    oGeometricalSet.Name = "Operations"
    oPart.HybridShapeFactory.GSMVisibility oPart.CreateReferenceFromObject(oGeometricalSet), 0
    
    oPart.InWorkObject = oPart.MainBody
    
End Sub


'*******************************************************************************
'* Procedure that hides Origin Planes XY, XZ and YZ according to R14 standards *
'*******************************************************************************
Public Sub HidePlanesAxis(ByVal oPart As MECMOD.Part, ByVal oSelection As INFITF.Selection)
    
    oSelection.Clear
    oSelection.Add oPart.OriginElements.PlaneXY
    oSelection.Add oPart.OriginElements.PlaneYZ
    oSelection.Add oPart.OriginElements.PlaneZX
    oSelection.VisProperties.SetLayer catVisLayerBasic, 4
    oSelection.VisProperties.SetShow catVisPropertyNoShowAttr
    
    oSelection.Clear
    For Each iObject In oPart.AxisSystems       'All ecountered axes set to No Show mode after execution
        oSelection.Add iObject
    Next
    oSelection.VisProperties.SetLayer catVisLayerBasic, 1
    oSelection.VisProperties.SetShow catVisPropertyNoShowAttr
    oSelection.Clear
    
End Sub

Public Sub AddPartBodyPublication(ByVal oPart As MECMOD.Part, ByVal oSelection As INFITF.Selection, ByVal oProduct As ProductStructureTypeLib.Product)
    
    oSelection.Clear
    oSelection.Add oPart.MainBody
    oSelection.VisProperties.SetLayer catVisLayerBasic, 50
    oSelection.Clear
    
    On Error Resume Next
    Set oPublication = Nothing
    Set oPublication = oProduct.ReferenceProduct.Publications.Item(oPart.MainBody.Name)
    If oPublication Is Nothing Then
        oProduct.ReferenceProduct.Publications.Add (oPart.MainBody.Name)
        oProduct.ReferenceProduct.Publications.SetDirect oPart.MainBody.Name, _
        oProduct.ReferenceProduct.CreateReferenceFromName(oProduct.PartNumber & "/!" & oPart.MainBody.Name)
    End If
    
End Sub

Public Sub SetProprietaryNotice(ByVal oProduct As ProductStructureTypeLib.Product)

    sDescriptionRef = ""
    sDescriptionRef = sDescriptionRef & "PROPRIETARY NOTICE (do not delete / modify)" & vbCrLf
    sDescriptionRef = sDescriptionRef & "***************************************************************************************************" & vbCrLf
    sDescriptionRef = sDescriptionRef & "BOMBARDIER INC.,  PAR L'INTERMEDIAIRE DE SES GROUPES, DIVISIONS ET/OU FILIALES  CORPORATIVES, FOURNIT LES" & vbCrLf
    sDescriptionRef = sDescriptionRef & "PRESENTS  RENSEIGNEMENTS AU GOUVERNEMENT  A TITRE CONFIDENTIEL. CES RENSEIGNEMENTS  ONT UNE TENEUR CONFI-" & vbCrLf
    sDescriptionRef = sDescriptionRef & "DENTIELLE,  AVEC  DROITS DE  PROPRIETE, ET NOUS  LES CONSIDERONS  EXCLUS DE L'APPLICATION  DE LA  LOI SUR" & vbCrLf
    sDescriptionRef = sDescriptionRef & "L 'ACCES A L'INFORMATION. ILS PEUVENT AUSSI CONSTITUER UN SECRET INDUSTRIEL DE NOTRE SOCIETE OU DE TIERCES" & vbCrLf
    sDescriptionRef = sDescriptionRef & "PARTIES. LA DIVULGATION  PUBLIQUE DE  CES RENSEIGNEMENTS NUIRAIT A BOMBARDIER INC. OU  A DES TIERCES PAR-" & vbCrLf
    sDescriptionRef = sDescriptionRef & "TIES. TOUTE INTENTION  DE LES DIFFUSER EN TOUT OU EN PARTIE  DEVRA ETRE SIGNIFIEE A L'AVANCE ET PAR ECRIT" & vbCrLf
    sDescriptionRef = sDescriptionRef & "AU VICE-PRESIDENT SERVICES JURIDIQUES, BOMBARDIER AERONAUTIQUE." & vbCrLf
    sDescriptionRef = sDescriptionRef & vbCrLf
    sDescriptionRef = sDescriptionRef & "BOMBARDIER INC., BY ITS GROUPS, DIVISIONS AND/OR CORPORATE SUBSIDIARIES, PROVIDES THIS INFORMATION TO THE" & vbCrLf
    sDescriptionRef = sDescriptionRef & "GOVERNMENT IN  CONFIDENCE. THIS  IS PROPRIETARY CONFIDENTIAL INFORMATION  WHICH WE CONSIDER EXCLUDED FROM" & vbCrLf
    sDescriptionRef = sDescriptionRef & "THE APPLICATION  OF THE ACCESS TO INFORMATION ACT AND MAY ALSO BE A TRADE  SECRET OF OUR COMPANY OR THIRD" & vbCrLf
    sDescriptionRef = sDescriptionRef & "PARTIES.  PUBLIC RELEASE  OF THIS INFORMATION WOULD BE  HARMFUL TO BOMBARDIER INC.  OR THIRD PARTIES. ANY" & vbCrLf
    sDescriptionRef = sDescriptionRef & "INTENTION TO  DISCLOSE SUCH INFORMATION, OR PART THEREOF, MUST BE INDICATED  IN ADVANCE AND IN WRITING TO" & vbCrLf
    sDescriptionRef = sDescriptionRef & "THE VICE-PRESIDENT LEGAL SERVICES, BOMBARDIER AEROSPACE." & vbCrLf
    sDescriptionRef = sDescriptionRef & vbCrLf & vbCrLf & vbCrLf
    sDescriptionRef = sDescriptionRef & "LES RENSEIGNEMENTS,  LES DONNEES ET DESSINS TECHNIQUES DIVULGUES DANS LES PRESENTES SONT LA PROPRIETE EX-" & vbCrLf
    sDescriptionRef = sDescriptionRef & "CLUSIVE DE BOMBARDIER INC. OU COMPORTENT DES DROITS DE PROPRIETE  APPARTENANT A AUTRUI ET NE PEUVENT ETRE" & vbCrLf
    sDescriptionRef = sDescriptionRef & "DIVULGUES  A QUICONQUE NI UTILISES PAR QUI QUE CE SOIT SANS  LE CONSENTEMENT ECRIT DE BOMBARDIER INC.  LE" & vbCrLf
    sDescriptionRef = sDescriptionRef & "DETENTEUR  DES PRESENTES S'ENGAGE, DANS TOUTES  CIRCONSTANCES, A GARDER CONFIDENTIELS LES DONNEES ET DES-" & vbCrLf
    sDescriptionRef = sDescriptionRef & "SINS TECHNIQUES QU'ELLES CONTIENNENT.  CE QUI PRECEDE NE  S'APPLIQUE PAS AUX PERSONNES QUI ONT DES DROITS" & vbCrLf
    sDescriptionRef = sDescriptionRef & "DE PROPRIETE SUR CES RENSEIGNEMENTS, DONNEES TECHNIQUES ET DESSINS DANS LA MESURE OU CES DROITS EXISTENT." & vbCrLf
    sDescriptionRef = sDescriptionRef & vbCrLf
    sDescriptionRef = sDescriptionRef & "THE INFORMATION,TECHNICAL DATA AND DESIGNS DISCLOSED HEREIN ARE THE EXCLUSIVE PROPERTY OF BOMBARDIER INC." & vbCrLf
    sDescriptionRef = sDescriptionRef & "OR CONTAIN PROPRIETARY RIGHTS OF OTHERS AND ARE NOT TO BE USED OR DISCLOSED TO OTHERS WITHOUT THE WRITTEN" & vbCrLf
    sDescriptionRef = sDescriptionRef & "CONSENT OF BOMBARDIER INC.THE RECIPIENT OF THIS DOCUMENT, BY ITS RETENTION AND USE AGREES TO HOLD IN CON-" & vbCrLf
    sDescriptionRef = sDescriptionRef & "FIDENCE THE TECHNICAL DATA AND DESIGNS CONTAINED HEREIN. THE FOREGOING SHALL  NOT APPLY TO PERSONS HAVING" & vbCrLf
    sDescriptionRef = sDescriptionRef & "PROPRIETARY RIGHTS TO SUCH INFORMATION, TECHNICAL DATA OR DESIGNS TO THE EXTENT THAT SUCH RIGHTS EXIST." & vbCrLf
    sDescriptionRef = sDescriptionRef & vbCrLf
    sDescriptionRef = sDescriptionRef & "Copyright 2005 Bombardier inc." & vbCrLf
    sDescriptionRef = sDescriptionRef & "***********************************"

    oProduct.DescriptionRef = sDescriptionRef

End Sub

Public Sub CreateProperties(ByVal oProduct As ProductStructureTypeLib.Product, Optional ByVal Title As String = "", Optional ByVal Dataset As String = "DETAIL", Optional ByVal bDocRevision As Boolean = False)
    
    Dim sPropertyFilePath As String
    Dim oProperty As StrParam
    Dim iTxtFileLine As Long
    Dim sParameterName As String
    Dim sParameterValues As String
    Dim aMultipleValues() As Variant
    Dim oParameterMultipleValues As Variant
    Dim iValueIndex As Long
    
    'Get property file
    sPropertyFilePath = TrimLine(ReadTextFile(sKBEPathFile, 16))        '***** Path for Property file (line 16)
    
    On Error Resume Next
    
    'Document Revision (Only in assembly nodes / CATProducts)
    If bDocRevision Then
        Set oProperty = Nothing
        Set oProperty = oProduct.ReferenceProduct.UserRefProperties.Item("Document Revision")
        If oProperty Is Nothing Then Set oProperty = oProduct.ReferenceProduct.UserRefProperties.CreateString("Document Revision", "")
    End If
    
    'All Parts (Products, CATParts)
    iTxtFileLine = 0
    Do
        iTxtFileLine = iTxtFileLine + 1
        sParameterName = ReadTextFile(sPropertyFilePath, iTxtFileLine)
        If sParameterName = "" Then Exit Do
        
        Set oProperty = Nothing
        Set oProperty = oProduct.ReferenceProduct.UserRefProperties.Item(sParameterName)
        If oProperty Is Nothing Then Set oProperty = oProduct.ReferenceProduct.UserRefProperties.CreateString(sParameterName, "")
        
        iTxtFileLine = iTxtFileLine + 1
        sParameterValues = ReadTextFile(sPropertyFilePath, iTxtFileLine)
        
        If Not sParameterValues = "" Then
            
            'Redim variable
            If Split(sParameterValues, ":")(UBound(Split(sParameterValues, ":"))) = "" Then
                ReDim aMultipleValues(UBound(Split(sParameterValues, ":")) - 1)
            Else
                ReDim aMultipleValues(UBound(Split(sParameterValues, ":")))
            End If
            
            'Populate aMultipleValues
            If UBound(aMultipleValues) > 0 Then
                For iValueIndex = 0 To UBound(aMultipleValues)
                    If Not Split(sParameterValues, ":")(iValueIndex) = "" Then aMultipleValues(iValueIndex) = Split(sParameterValues, ":")(iValueIndex)
                Next iValueIndex
                
                Set oParameterMultipleValues = oProperty
                oParameterMultipleValues.SetEnumerateValues aMultipleValues
            End If
            
            'Set default property value
            If Not UBound(Split(sParameterValues, ":")) > 2 Then
                oProperty.Value = Split(sParameterValues, ":")(0)
            Else
                oProperty.Value = "-"
            End If
            
        End If
        
        'Set Title and Dataset type
        If sParameterName = "Title" Then oProperty.Value = Title
        If sParameterName = "Dataset Type" Then oProperty.Value = Dataset
        
    Loop Until sParameterName = "" Or sParameterName = "END"
    
End Sub
