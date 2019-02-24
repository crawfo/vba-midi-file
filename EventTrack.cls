VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'a collection of midi event objects: ChannelEvent, MetaEvent, SystemExclusiveEvent
Private pMidiEvents As Collection

Private Sub Class_Initialize()
    Set pMidiEvents = New Collection
End Sub

Private Sub Class_Terminate()
    Set pMidiEvents = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator Property"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pMidiEvents.[_NewEnum]
End Property

Public Property Get MidiEvents() As Collection
    Set MidiEvents = pMidiEvents
End Property

Sub Add(obj As Object)
    If TypeName(obj) = "ChannelEvent" Or _
       TypeName(obj) = "MetaEvent" Or _
       TypeName(obj) = "SystemExclusiveEvent" Then
          pMidiEvents.Add obj
    End If
End Sub

Function Count() As Long
    Count = pMidiEvents.Count
End Function

Sub Remove(eventIndex As Long)
    pMidiEvents.Remove eventIndex
End Sub

Public Property Get Item(ByVal itemIndex As Long) As Object
Attribute Item.VB_Description = "Default Property"
Attribute Item.VB_UserMemId = 0
    Set Item = pMidiEvents.Item(itemIndex)
End Property

Public Sub Initialize(ByVal trkEvents As Collection)
    Set pMidiEvents = trkEvents
End Sub

Public Function ToMidiTrackChunk() As TrackChunk
    Dim trackChunkBytes() As Byte
    trackChunkBytes = ToBytes()
    Set ToMidiTrackChunk = Factory.CreateNewMidiTrackChunk(trackChunkBytes)
End Function

Public Function ToBytes() As Byte()
    Dim trkByte As Variant
    Dim trkBytes As Collection
    Dim i As Long
    Dim midiTrkChunk As TrackChunk
    Dim trackChunkBytes() As Byte
    
    Set trkBytes = ToBytesCollection()
    ReDim trackChunkBytes(trkBytes.Count - 1)
    'copy event bytes to array
    i = 0
    For Each trkByte In trkBytes
        trackChunkBytes(i) = trkByte
        i = i + 1
    Next trkByte
    
    ToBytes = trackChunkBytes
End Function

Public Function ToBytesCollection() As Collection
    Dim trackLengths() As Byte
    Dim trackBytes As Collection
    Dim i As Long
    Dim length As Long
    
    Set trackBytes = New Collection
    
    trackBytes.Add TrackChunkTypeEnum.BYTE_1
    trackBytes.Add TrackChunkTypeEnum.BYTE_2
    trackBytes.Add TrackChunkTypeEnum.BYTE_3
    trackBytes.Add TrackChunkTypeEnum.BYTE_4
        
    AddBytesFromAllMidiEventsToTrackBytes trackBytes
    CheckForMissingEndOfTrackMetaEvent trackBytes
    AddTrackLengthBytes trackBytes

    Set ToBytesCollection = trackBytes
End Function

Private Sub AddBytesFromAllMidiEventsToTrackBytes(trackBytesCollection As Collection)
    'mutates parameter trackBytesCollection
    Dim midiEvt As Variant
    Dim evtBytes() As Byte
    
    For Each midiEvt In pMidiEvents
        evtBytes = midiEvt.ToBytes()
        AddBytesFromEventToTrackBytes evtBytes, trackBytesCollection
    Next midiEvt
End Sub

Private Sub AddBytesFromEventToTrackBytes(evtBytes() As Byte, trackBytesCollection As Collection)
    'mutates parameter trackBytesCollection
    Dim evtByte As Variant
    
    For Each evtByte In evtBytes
        trackBytesCollection.Add evtByte
    Next evtByte
End Sub

Private Sub CheckForMissingEndOfTrackMetaEvent(trackBytes As Collection)
    'checks for and adds end of track meta event bytes (if it doesn't exist)
    'mutates parameter trackBytes
    Const DELTA_TIME = 0
    Const END_OF_TRACK_DATA_BYTE = 0
    Dim isMetaEvt As Boolean
    Dim isNotMetaEvent As Boolean
    Dim isNotEndOfTrackMetaEvent As Boolean
    
    isNotMetaEvent = pMidiEvents(pMidiEvents.Count).Status <> StatusEnum.META_EVENT
    isMetaEvt = Not isNotMetaEvent
    isNotEndOfTrackMetaEvent = pMidiEvents(pMidiEvents.Count).Status = StatusEnum.META_EVENT
    If isMetaEvt Then
       isNotEndOfTrackMetaEvent = pMidiEvents(pMidiEvents.Count).MetaType <> MetaEventTypeEnum.END_OF_TRACK
    End If
    
    If isNotMetaEvent Or isNotEndOfTrackMetaEvent Then
        'add delta
        trackBytes.Add DELTA_TIME
        'add event
        trackBytes.Add StatusEnum.META_EVENT
        trackBytes.Add MetaEventTypeEnum.END_OF_TRACK
        trackBytes.Add END_OF_TRACK_DATA_BYTE
    End If
End Sub

Private Sub AddTrackLengthBytes(trackBytes As Collection)
    Const TRACK_ID_LENGTH = 4
    Dim length As Long
    Dim trackLengths() As Byte
    
    length = trackBytes.Count - TRACK_ID_LENGTH 'track id bytes were added earlier
    trackLengths = Convert.ToFourBytesFromLong(length)
    trackBytes.Add trackLengths(0), After:=4 'collections are 1-based, arrays 0-based
    trackBytes.Add trackLengths(1), After:=5
    trackBytes.Add trackLengths(2), After:=6
    trackBytes.Add trackLengths(3), After:=7
End Sub

