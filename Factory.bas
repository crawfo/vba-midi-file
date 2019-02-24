Attribute VB_Name = "Factory"
'uses Utilities
Option Explicit

Public Function CreateNewTrackDimensions(ByVal dataStartPositions As Collection, _
                                         ByVal dataEndPositions As Collection, _
                                         ByVal sizes As Collection) As TrackDimensions
    Set CreateNewTrackDimensions = New TrackDimensions
    CreateNewTrackDimensions.Initialize dataStartPositions, dataEndPositions, sizes
End Function

Public Function CreateNewMidiTrackChunk(trackBytes() As Byte) As TrackChunk
    Set CreateNewMidiTrackChunk = New TrackChunk
    CreateNewMidiTrackChunk.Initialize trackBytes
End Function

Public Function CreateNewMidiTrackChunks() As TrackChunks
    Set CreateNewMidiTrackChunks = New TrackChunks
    CreateNewMidiTrackChunks.Initialize
End Function

Public Function CreateNewEventTracks(ByVal eventTrks As Collection) As EventTracks
    Set CreateNewEventTracks = New EventTracks
    CreateNewEventTracks.Initialize eventTrks
End Function

Public Function CreateNewEventTrack(ByVal trkEvents As Collection) As EventTrack
    Set CreateNewEventTrack = New EventTrack
    CreateNewEventTrack.Initialize trkEvents
End Function

Public Function CreateNewMidiFile(ByVal hdrChunk As HeaderChunk, _
                                  ByVal eventTrks As EventTracks) As MidiFile
    Set CreateNewMidiFile = New MidiFile
    CreateNewMidiFile.InitializeFromEventTracks hdrChunk, eventTrks
End Function

Public Function CreateNewMidiFileFromFile(ByVal fileNameFullyQualified As String) As MidiFile
    Dim bytes() As Byte
    
    bytes = FileUtils.ReadFile(fileNameFullyQualified)
    Set CreateNewMidiFileFromFile = CreateNewMidiFileFromArray(bytes)
End Function

Public Function CreateNewMidiFileFromArray(midiFileBytes() As Byte) As MidiFile
    Set CreateNewMidiFileFromArray = New MidiFile
    CreateNewMidiFileFromArray.Initialize midiFileBytes
End Function

Public Function CreateNewHeaderChunk(ByVal midiFileFormat As Integer, _
                                     ByVal trackCount As Long, _
                                     ByVal timeDivType As TimeDivisionType, _
                                     ByVal TimeDivision As Long) As HeaderChunk
    Set CreateNewHeaderChunk = New HeaderChunk
    CreateNewHeaderChunk.Initialize midiFileFormat, trackCount, timeDivType, TimeDivision
End Function

Public Function CreateNewChannelEvent(ByVal isRunStatus As Boolean, _
                                      ByVal deltaTime As Long, _
                                      ByVal absoluteTime As Long, _
                                      ByVal midiStatus As Byte, ByVal midiChannel As Byte, _
                                      ByVal eventData1 As Byte, _
                                      Optional ByVal eventData2 As Variant) As ChannelEvent
    Set CreateNewChannelEvent = New ChannelEvent
    CreateNewChannelEvent.Initialize isRunStatus, _
                                     deltaTime, _
                                     absoluteTime, _
                                     midiStatus, _
                                     midiChannel, _
                                     eventData1, _
                                     eventData2
End Function

Public Function CreateNewMetaEvent(ByVal deltaTime As Long, _
                                   ByVal absoluteTime As Long, _
                                   ByVal midiMetaType As Byte, _
                                   ByVal eventData As Collection) As MetaEvent
    Set CreateNewMetaEvent = New MetaEvent
    CreateNewMetaEvent.Initialize deltaTime, absoluteTime, midiMetaType, eventData
End Function

Public Function CreateNewSystemExclusiveEvent(ByVal deltaTime As Long, _
                                              ByVal absoluteTime As Long, _
                                              ByVal midiStatus As Byte, _
                                              ByVal eventData As Collection, _
                                              ByVal systemExType As SystemExclusiveType) As SystemExclusiveEvent
    Set CreateNewSystemExclusiveEvent = New SystemExclusiveEvent
    CreateNewSystemExclusiveEvent.Initialize deltaTime, absoluteTime, midiStatus, eventData, systemExType
End Function

Private Function CreateRawBytesCollectionMetaOrSystemEx(vlvBytes() As Byte, _
                                                        ByVal eventData As Collection) As Collection
    Dim vlvByte As Variant
    Dim evtByte As Variant
    Dim rawBytes As Collection
    
    Set rawBytes = New Collection
    
    'add length
    For Each vlvByte In vlvBytes
        rawBytes.Add vlvByte
    Next vlvByte
    
    'add data
    For Each evtByte In eventData
        rawBytes.Add evtByte
    Next evtByte
    
    Set CreateRawBytesCollectionMetaOrSystemEx = rawBytes
End Function

Private Function CreateRawBytesCollection(ByVal midiStatus As Byte, _
                                          ByVal midiChannel As Byte, _
                                          ByVal eventData1 As Byte, _
                                          Optional ByVal eventData2 As Variant) As Collection
    Dim rawBytes As Collection
    Dim ce As ChannelEvent
    
    Set ce = New ChannelEvent
    Set rawBytes = New Collection
    rawBytes.Add ce.JoinTwoNibbles(midiStatus, midiChannel)
    rawBytes.Add eventData1
    If Not IsMissing(eventData2) Then
        rawBytes.Add eventData2
    End If
    Set CreateRawBytesCollection = rawBytes
End Function

Public Function CreateNewCoreEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal statusByte As Byte, ByVal evtData As Collection, ByVal eventCoreLength As Long) As CoreEvent
    Set CreateNewCoreEvent = New CoreEvent
    CreateNewCoreEvent.Initialize deltaTime, absoluteTime, statusByte, evtData, eventCoreLength
End Function

'---------------------------------------------------------------------------------------------
'Convenience constructors
'---------------------------------------------------------------------------------------------
Public Function CreateNewNoteOnEvent(ByVal deltaTime As Long, _
                                     ByVal absoluteTime As Long, _
                                     ByVal midiChannel As Byte, _
                                     ByVal midiNoteNumber As Byte, _
                                     ByVal velocity As Byte) As ChannelEvent
    Set CreateNewNoteOnEvent = CreateNewChannelEvent(deltaTime, _
                                                     absoluteTime, _
                                                     StatusEnum.NOTE_ON, _
                                                     midiChannel, _
                                                     midiNoteNumber, _
                                                     velocity)
End Function

Public Function CreateNewNoteOffEvent(ByVal deltaTime As Long, _
                                      ByVal absoluteTime As Long, _
                                      ByVal midiChannel As Byte, _
                                      ByVal midiNoteNumber As Byte, _
                                      ByVal velocity As Byte) As ChannelEvent
    Set CreateNewNoteOffEvent = CreateNewChannelEvent(deltaTime, _
                                                      absoluteTime, _
                                                      StatusEnum.NOTE_OFF, _
                                                      midiChannel, _
                                                      midiNoteNumber, _
                                                      velocity)
End Function

Public Function CreateNewNoteAftertouchEvent(ByVal deltaTime As Long, _
                                             ByVal absoluteTime As Long, _
                                             ByVal midiChannel As Byte, _
                                             ByVal midiNoteNumber As Byte, _
                                             ByVal pressure As Byte) As ChannelEvent
    Set CreateNewNoteAftertouchEvent = CreateNewChannelEvent(deltaTime, _
                                                             absoluteTime, _
                                                             StatusEnum.NOTE_AFTERTOUCH, _
                                                             midiChannel, _
                                                             midiNoteNumber, _
                                                             pressure)
End Function

Public Function CreateNewControlChangeEvent(ByVal deltaTime As Long, _
                                            ByVal absoluteTime As Long, _
                                            ByVal midiChannel As Byte, _
                                            ByVal controllerNumber As ContinuousControllerType, _
                                            ByVal controllerValue As Byte) As ChannelEvent
    Set CreateNewControlChangeEvent = CreateNewChannelEvent(deltaTime, _
                                                            absoluteTime, _
                                                            StatusEnum.CONTROLLER, _
                                                            midiChannel, _
                                                            controllerNumber, _
                                                            controllerValue)
End Function

Public Function CreateNewChannelAftertouchEvent(ByVal deltaTime As Long, _
                                                ByVal absoluteTime As Long, _
                                                ByVal midiChannel As Byte, _
                                                ByVal pressure As Byte) As ChannelEvent
    Set CreateNewChannelAftertouchEvent = CreateNewChannelEvent(deltaTime, _
                                                                absoluteTime, _
                                                                StatusEnum.CHANNEL_AFTERTOUCH, _
                                                                midiChannel, pressure)
End Function

Public Function CreateNewPitchBendEvent(ByVal deltaTime As Long, _
                                        ByVal absoluteTime As Long, _
                                        ByVal midiChannel As Byte, _
                                        ByVal lsb As Byte, _
                                        ByVal msb As Byte) As ChannelEvent
    Set CreateNewPitchBendEvent = CreateNewChannelEvent(deltaTime, _
                                                        absoluteTime, _
                                                        StatusEnum.PITCH_BEND, _
                                                        midiChannel, _
                                                        lsb, _
                                                        msb)
End Function

Public Function CreateNewProgramChangeEvent(ByVal deltaTime As Long, _
                                            ByVal absoluteTime As Long, _
                                            ByVal midiChannel As Byte, _
                                            ByVal programNumber As Byte) As ChannelEvent
    Set CreateNewProgramChangeEvent = CreateNewChannelEvent(deltaTime, _
                                                            absoluteTime, _
                                                            StatusEnum.PROGRAM_CHANGE, _
                                                            midiChannel, _
                                                            programNumber)
End Function

Public Function CreateNewTimeSignatureMetaEvent(ByVal timeSignatureNumerator As Long, _
                                                ByVal timeSignatureDenominatorPowerOfTwoExponent As Long, _
                                                ByVal midiClocksPerMetronomeTick As Long, _
                                                ByVal number32ndNotesPer24MidiClocks As Long, _
                                                ByVal deltaTime As Long, _
                                                ByVal absoluteTime As Long) As MetaEvent
    Dim dataBytes As Collection
    Set dataBytes = New Collection
    
    dataBytes.Add timeSignatureNumerator
    dataBytes.Add timeSignatureDenominatorPowerOfTwoExponent 'eg desired denominator of 8 would be 3 as 2^3 = 8
    dataBytes.Add midiClocksPerMetronomeTick 'normally 24, ie tick once every quarter note
    dataBytes.Add number32ndNotesPer24MidiClocks 'normally 8, ie 8 32nd notes per quarter note
    
    Set CreateNewTimeSignatureMetaEvent = CreateNewMetaEvent(deltaTime, _
                                                            absoluteTime, _
                                                            MetaEventTypeEnum.TIME_SIGNATURE, _
                                                            dataBytes)
End Function

Public Function CreateNewKeySignatureMetaEvent(ByVal keyValue As KeySignatureKeyValue, _
                                               ByVal modeValue As KeySignatureModeValue, _
                                               ByVal deltaTime As Long, _
                                               ByVal absoluteTime As Long) As MetaEvent
    Dim dataBytes As Collection
    Set dataBytes = New Collection
    
    dataBytes.Add keyValue
    dataBytes.Add modeValue
    
    Set CreateNewKeySignatureMetaEvent = CreateNewMetaEvent(deltaTime, _
                                                            absoluteTime, _
                                                            meKeySignature, _
                                                            dataBytes)
End Function



