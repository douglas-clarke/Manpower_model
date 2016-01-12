VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'create ID, PALS/O&S array, Service array and make them read/write
Private pID As Long
Private pPALSOS() As Variant
Private pService() As Variant
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(value As Long)
    pID = value
End Property
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Property Get PALSOS() As Variant
    PALSOS = pPALSOS()
End Property
Public Property Let PALSOS(value As Variant)
    pPALSOS = value
End Property
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Property Get Service() As Variant
    Service = pService
End Property
Public Property Let Service(value As Variant)
    pService = value
End Property

