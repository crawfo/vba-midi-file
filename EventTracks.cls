VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventTracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'a collection of EventTrack objects
Private pTracks As Collection

Private Sub Class_Initialize()
    Set pTracks = New Collection
End Sub

Private Sub Class_Terminate()
    Set pTracks = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = pTracks.[_NewEnum]
End Property

Public Property Get Tracks() As Collection
    Set Tracks = pTracks
End Property

Sub Add(Track As EventTrack)
    pTracks.Add Track
End Sub

Public Function Count() As Long
    Count = pTracks.Count
End Function

Public Sub Remove(trackIndex As Long)
    pTracks.Remove trackIndex
End Sub

Public Property Get Item(ByVal itemIndex As Long) As EventTrack
    Set Item = pTracks.Item(itemIndex)
End Property

Public Sub Initialize(ByVal eventTrks As Collection)
    Set pTracks = eventTrks
End Sub

Public Function ToBytesCollection() As Collection
    Dim trackBytesCollection As Collection
    
    Set trackBytesCollection = New Collection
    Set ToBytesCollection = GetAllTracksBytes(pTracks)
End Function
 
Private Function GetAllTracksBytes(ByVal eventTrks As Collection) As Collection
    'put all bytes from each track into a collection
    Dim eventTrk As Variant
    Dim eventTrackBytes As Collection
    Dim trackBytesCollection As Collection
    
    Set trackBytesCollection = New Collection
    For Each eventTrk In eventTrks
        Set eventTrackBytes = eventTrk.ToBytesCollection()
        AddEventTrackBytesToTrackBytesCollection eventTrackBytes, trackBytesCollection
    Next eventTrk
    
    Set GetAllTracksBytes = trackBytesCollection
End Function

Private Sub AddEventTrackBytesToTrackBytesCollection(ByVal eventTrackBytes As Collection, _
                                                     trackBytesCollection As Collection)
    'mutates parameter trackBytesCollection
    Dim trkByte As Variant
    For Each trkByte In eventTrackBytes
        trackBytesCollection.Add trkByte
    Next trkByte
End Sub





