VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pHeader As String
Private pFieldValues() As Variant
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Property Get Header() As String
    Header = pHeader
End Property
Public Property Let Header(value As String)
    pHeader = value
End Property
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Property Get FieldValues() As Variant
    FieldValues = pFieldValues()
End Property
Public Property Let FieldValues(value As Variant)
    pFieldValues = value
End Property


