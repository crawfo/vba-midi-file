Attribute VB_Name = "TrackParser"
Option Explicit

Public Function ParseTrack(trkChunk As TrackChunk) As EventTrack
     Dim prevStatusChan As Byte 'used for running status
     Dim absoluteTime As Long
     Dim deltaTime As Long
     Dim trackPosition As Long
     Dim midiStatus As Byte
     Dim trkEvents As Collection
     Dim vlvBytes As Collection
     Dim upperBound As Long
     Dim trkEvent As Object
     Dim trkChunkBytes() As Byte
     'TODO: use a param object(s) for ReadEvent params
     
     Set trkEvents = New Collection
     trkChunkBytes = trkChunk.ChunkBytes
     trackPosition = Midi.TRACK_HEADER_LENGTH
     upperBound = UBound(trkChunkBytes)
     prevStatusChan = 0
     'loop thru track parsing msgs
     Do While trackPosition < upperBound
        'read delta vlv
        Set vlvBytes = Convert.GetVLVBytes(trkChunkBytes, trackPosition)
        deltaTime = Convert.DecodeVLV(vlvBytes)
        absoluteTime = absoluteTime + deltaTime
        'incr index to event start
        trackPosition = trackPosition + vlvBytes.Count
        'read status byte
        midiStatus = trkChunkBytes(trackPosition)
        'get event
        Set trkEvent = EventParser.ReadEvent(deltaTime, _
                                             absoluteTime, _
                                             midiStatus, _
                                             trackPosition, _
                                             prevStatusChan, _
                                             trkChunkBytes)
        trkEvents.Add trkEvent
        If MidiEventUtils.IsChannelEvent(midiStatus) Then
            prevStatusChan = midiStatus
        End If
        trackPosition = trackPosition + trkEvent.EventLength
    Loop
        
    Set ParseTrack = Factory.CreateNewEventTrack(trkEvents)
End Function

Public Function ParseTracks(ByVal trkChunks As TrackChunks) As EventTracks
    Dim trkChunk As TrackChunk
    Dim eventTrks As Collection
    Dim eventTrk As EventTrack
    
    Set eventTrks = New Collection
    For Each trkChunk In trkChunks
        Set eventTrk = ParseTrack(trkChunk)
        eventTrks.Add eventTrk
    Next trkChunk
    
    Set ParseTracks = Factory.CreateNewEventTracks(eventTrks)
End Function

