Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateNewMidiFileInfo(ByVal midiFileFormat As Integer, ByVal trackCount As Long, ByVal TimeDivision As Long) As MidiFileInfo
    Set CreateNewMidiFileInfo = New MidiFileInfo
    CreateNewMidiFileInfo.Initialize midiFileFormat, trackCount, TimeDivision
End Function

Public Function CreateNewMetaEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiMetaType As Byte, ByVal eventData As Collection) As MetaEvent
    Set CreateNewMetaEvent = New MetaEvent
    CreateNewMetaEvent.Initialize deltaTime, absoluteTime, midiMetaType, eventData
End Function

Public Function CreateNewChannelEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiStatus As Byte, ByVal midiChannel As Byte, ByVal eventData1 As Byte, Optional ByVal eventData2 As Variant) As ChannelEvent
    Set CreateNewChannelEvent = New ChannelEvent
    If IsMissing(eventData2) Then
        CreateNewChannelEvent.Initialize deltaTime, absoluteTime, midiStatus, midiChannel, eventData1
    Else
        CreateNewChannelEvent.Initialize deltaTime, absoluteTime, midiStatus, midiChannel, eventData1, eventData2
    End If
End Function

Public Function CreateNewSystemExclusiveEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiStatus As Byte, ByVal eventData As Collection, ByVal enmEvtType As SystemExclusiveType) As SystemExclusiveEvent
    Set CreateNewSystemExclusiveEvent = New SystemExclusiveEvent
    CreateNewSystemExclusiveEvent.Initialize deltaTime, absoluteTime, midiStatus, eventData, enmEvtType
End Function

Public Function CreateNewHeaderChunk(ByVal midiFileFormat As Integer, ByVal trackCount As Long, ByVal timeDiv As Long) As HeaderChunk
    Set CreateNewHeaderChunk = New HeaderChunk
    CreateNewHeaderChunk.Initialize midiFileFormat, trackCount, timeDiv
End Function

Public Function CreateNewTrackChunk(ByVal trackEventList As Collection) As TrackChunk
    Set CreateNewTrackChunk = New TrackChunk
    CreateNewTrackChunk.Initialize trackEventList
End Function

Public Function CreateNewTrackCollection() As TrackCollection
    Set CreateNewTrackCollection = New TrackCollection
    CreateNewTrackCollection.Initialize
End Function

Public Function CreateNewStandardMidiFile(ByVal midiFileFormat As Integer, ByVal timeDiv As Long, ByVal trackChunks As TrackCollection) As StandardMidiFile
    Set CreateNewStandardMidiFile = New StandardMidiFile
    CreateNewStandardMidiFile.InitA midiFileFormat, timeDiv, trackChunks
End Function

Public Function CreateNewTrackDimensions(ByVal dataStartPositions As Collection, ByVal dataEndPositions As Collection, ByVal sizes As Collection) As TrackDimensions
    Set CreateNewTrackDimensions = New TrackDimensions
    CreateNewTrackDimensions.Initialize dataStartPositions, dataEndPositions, sizes
End Function

Public Function CreateNewNoteOnEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiChannel As Byte, ByVal midiNoteNumber As Byte, ByVal velocity As Byte) As ChannelEvent
    Set CreateNewNoteOnEvent = CreateNewChannelEvent(deltaTime, absoluteTime, ceNoteOn, midiChannel, midiNoteNumber, velocity)
End Function

Public Function CreateNewNoteOffEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiChannel As Byte, ByVal midiNoteNumber As Byte, ByVal velocity As Byte) As ChannelEvent
    Set CreateNewNoteOffEvent = CreateNewChannelEvent(deltaTime, absoluteTime, ceNoteOff, midiChannel, midiNoteNumber, velocity)
End Function

Public Function CreateNewTempoMetaEvent(ByVal beatsPerMinute As Long, ByVal deltaTime As Long, ByVal absoluteTime As Long) As MetaEvent
    Dim tempoBytes As Collection
    Set tempoBytes = ToThreeBytes(ToMicrosecondsPerQuarterNote(beatsPerMinute))
    Set CreateNewTempoMetaEvent = CreateNewMetaEvent(deltaTime, absoluteTime, meSetTempo, tempoBytes)
End Function

Public Function CreateNewTimeSignatureMetaEvent(ByVal timeSignatureNumerator As Long, ByVal timeSignatureDenominatorPowerOfTwoExponent As Long, ByVal midiClocksPerMetronomeTick As Long, ByVal number32ndNotesPer24MidiClocks As Long, ByVal deltaTime As Long, ByVal absoluteTime As Long) As MetaEvent
    Dim dataBytes As Collection
    Set dataBytes = New Collection
    
    dataBytes.Add timeSignatureNumerator
    dataBytes.Add timeSignatureDenominatorPowerOfTwoExponent 'eg desired denominator of 8 would be 3 as 2^3 = 8
    dataBytes.Add midiClocksPerMetronomeTick 'normally 24, ie tick once every quarter note
    dataBytes.Add number32ndNotesPer24MidiClocks 'normally 8, ie 8 32nd notes per quarter note
    
    Set CreateNewTimeSignatureMetaEvent = CreateNewMetaEvent(deltaTime, absoluteTime, meTimeSignature, dataBytes)
End Function

Public Function CreateNewKeySignatureMetaEvent(ByVal KeyValue As KeySignatureKeyValue, ByVal modeValue As KeySignatureModeValue, ByVal deltaTime As Long, ByVal absoluteTime As Long) As MetaEvent
    Dim dataBytes As Collection
    Set dataBytes = New Collection
    
    dataBytes.Add KeyValue
    dataBytes.Add modeValue
    
    Set CreateNewKeySignatureMetaEvent = CreateNewMetaEvent(deltaTime, absoluteTime, meKeySignature, dataBytes)
End Function

Public Function CreateNewBarBeatTick(ByVal positionInTicks As Long, ticksPerQuarterNote As Long) _
                                     As BarBeatTick
    Set CreateNewBarBeatTick = New BarBeatTick
    CreateNewBarBeatTick.Initialize positionInTicks, ticksPerQuarterNote
End Function

Public Function ToThreeBytes(ByVal microsecondsPerQuarterNote As Long) As Collection
    'Splits a long into 3 bytes, range &H0-&H7F7F7F  (0-8355711)
    Set ToThreeBytes = New Collection
    ToThreeBytes.Add microsecondsPerQuarterNote \ 2 ^ 16 And &HFF
    ToThreeBytes.Add microsecondsPerQuarterNote \ 2 ^ 8 And &HFF
    ToThreeBytes.Add microsecondsPerQuarterNote And &HFF
End Function

Public Function ToMicrosecondsPerQuarterNote(ByVal beatsPerMinute As Long) As Long
    Const MICROSECONDS_PER_MINUTE = 60000000
    ToMicrosecondsPerQuarterNote = MICROSECONDS_PER_MINUTE \ beatsPerMinute
End Function

Public Function CreateNewNoteInfo(ByVal noteNum As Long, ByVal absoluteTime As Long, ByVal noteLen As Long, ByVal trackNum As Long) As noteInfo
    Set CreateNewNoteInfo = New noteInfo
    CreateNewNoteInfo.Initialize noteNum, absoluteTime, noteLen, trackNum
End Function
