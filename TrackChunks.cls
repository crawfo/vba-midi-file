VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackChunks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'a collection of TrackChunk objects
Private pTracks As Collection

Private Sub Class_Initialize()
    Set pTracks = New Collection
End Sub

Private Sub Class_Terminate()
    Set pTracks = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator Property"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pTracks.[_NewEnum]
End Property

Sub Add(ByVal Track As TrackChunk)
    pTracks.Add Track
End Sub

Function Count() As Long
    Count = pTracks.Count
End Function

Sub Remove(ByVal trackIndex As Long)
    pTracks.Remove trackIndex
End Sub

Public Property Get Item(ByVal itemIndex As Long) As TrackChunk
Attribute Item.VB_Description = "Default Property"
Attribute Item.VB_UserMemId = 0
    Set Item = pTracks.Item(itemIndex)
End Property

Public Sub Initialize()
    
End Sub

