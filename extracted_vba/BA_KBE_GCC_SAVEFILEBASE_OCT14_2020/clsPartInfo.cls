Attribute VB_Name = "clsPartInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public PartNumber As String
Public PartRev As String
Public PartTitle As String
Public PartStatus As String
Public PartLatestRevStatus As String
Public PartIsCI As Boolean
Public PartType As String
Public PartExtension As String
Public DefiningPart As String
Public Nomenclature As String

Public SyncFromBSF As Boolean
Public Tail As String
Public Project As String
Public CIOption As String
Public NonCIOption As String
Public Mode As String

Public EffectiveDwgNb As String
Public EffectiveDwgRev As String
Public ProposedSourceNb As String
Public ProposedSourceRev As String
Public SelectedDwgNb As String
Public SelectedDwgRev As String
Public SelectedSourceNb As String
Public SelectedSourceRev As String


Public PVRDocNb As String
Public PVRDocRev As String
Public PVRDocStatus As String
Public PVRDocIteration As String

Public Comment As String



