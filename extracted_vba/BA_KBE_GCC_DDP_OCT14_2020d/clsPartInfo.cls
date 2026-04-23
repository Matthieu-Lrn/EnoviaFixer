Attribute VB_Name = "clsPartInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public OriginalPartNumber As String
Public OriginalPartRev As String
Public OriginalPartStatus As String

Public RefPartNumber As String
Public RefPartRev As String
Public RefPartStatus As String

Public UsedPartNumber As String
Public UsedPartRev As String
Public UsedPartStatus As String

Public PartTitle As String
Public PartIsCI As Boolean
Public PartType As String
Public PartExtension As String
Public DefiningPartAtt As String
Public DatasetTypeAtt As String

Public PartLatestRevStatus As String

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


'Public PartNumber As String
'Public PartRev As String
'Public PartStatus As String

