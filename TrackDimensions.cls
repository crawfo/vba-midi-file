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

Private pTrackStart  As Collection
Private pTrackSize  As Collection
Private pTrackEnd  As Collection

Public Property Get TrackStart() As Collection
    Set TrackStart = pTrackStart
End Property

Public Property Get TrackEnd() As Collection
    Set TrackEnd = pTrackEnd
End Property

Public Property Get TrackSize() As Collection
    Set TrackSize = pTrackSize
End Property

Private Sub Class_Terminate()
    Set pTrackStart = Nothing
    Set pTrackEnd = Nothing
    Set pTrackSize = Nothing
End Sub

Public Sub Initialize(ByVal dataStartPositions As Collection, _
                      ByVal dataEndPositions As Collection, _
                      ByVal sizes As Collection)
    Set pTrackStart = dataStartPositions
    Set pTrackEnd = dataEndPositions
    Set pTrackSize = sizes
End Sub
