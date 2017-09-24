VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDataStart  As Collection
Private pDataSize  As Collection
Private pDataEnd  As Collection

Public Property Get DataStart() As Collection
    Set DataStart = pDataStart
End Property

Public Property Get DataEnd() As Collection
    Set DataEnd = pDataEnd
End Property


Public Property Get DataSize() As Collection
    Set DataSize = pDataSize
End Property

Private Sub Class_Terminate()
    Set pDataStart = Nothing
    Set pDataEnd = Nothing
    Set pDataSize = Nothing
End Sub

Public Sub Initialize(ByVal dataStartPositions As Collection, ByVal dataEndPositions As Collection, ByVal sizes As Collection)
    'side effect: mutates object properties
    Set pDataStart = dataStartPositions
    Set pDataEnd = dataEndPositions
    Set pDataSize = sizes
End Sub
