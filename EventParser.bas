Attribute VB_Name = "EventParser"
Public Function ReadEvent(ByVal deltaTime As Long, _
                          ByVal absoluteTime As Long, _
                          ByVal midiStatus As Byte, _
                          ByVal trackPosition As Long, _
                          ByVal prevStatusChan As Byte, _
                          trkChunkBytes() As Byte) As Object
    Dim trkEvent As Object

    'eval status byte
    If MidiEventUtils.IsChannelEvent(midiStatus) _
    Or MidiEventUtils.IsRunningStatus(midiStatus) Then
        Set trkEvent = ReadChannelEvent(deltaTime, _
                                        absoluteTime, _
                                        trkChunkBytes, _
                                        trackPosition, _
                                        prevStatusChan)
    ElseIf MidiEventUtils.IsMetaEvent(midiStatus) _
    Or MidiEventUtils.IsSysExEvent(midiStatus) Then
        Set trkEvent = ReadSystemExclusiveOrMetaEvent(deltaTime, _
                                                      absoluteTime, _
                                                      trkChunkBytes, _
                                                      trackPosition)
    Else
        Err.Raise Number:=vbObjectError + 1051, Source:="EventParser.ReadEvent", _
                  Description:="Source: EventParser.ReadEvent. Invalid midi status: " _
                               & midiStatus
     End If
    
    Set ReadEvent = trkEvent
End Function

Private Function ReadChannelEvent(ByVal deltaTime As Long, _
                                  ByVal absoluteTime As Long, _
                                  trackBytes() As Byte, _
                                  ByVal eventStartPosition As Long, _
                                  ByVal previousStatusByte As Byte) As ChannelEvent
    Const RUNNING_STATUS_OFFSET = 1
    Const NORMAL_OFFSET = 0
    Dim dataByte1  As Byte
    Dim dataByte2 As Byte
    Dim statusByte As Byte
    Dim statusNibble As Byte
    Dim channelNibble As Byte
    Dim isThreeByteChanEvt As Boolean
    Dim offset As Long
    Dim isRunStatus As Boolean
    Dim vlvByte As Variant
    Dim dataByte As Variant

    statusByte = trackBytes(eventStartPosition)
    isRunStatus = IsRunningStatus(statusByte)
    If isRunStatus Then 'TODO: can fail to stop here if running status byte is valid
                        'but event length is wrong but unclear if this case would ever happen
        statusByte = previousStatusByte
        offset = NORMAL_OFFSET
    Else
        offset = RUNNING_STATUS_OFFSET
    End If
    statusNibble = GetNibbleHigh(statusByte)
    channelNibble = GetNibbleLow(statusByte)
    isThreeByteChanEvt = IsThreeByteChannelEvent(statusByte)
    dataByte1 = trackBytes(eventStartPosition + offset)
    offset = offset + 1
    If isThreeByteChanEvt Then
        dataByte2 = trackBytes(eventStartPosition + offset)
    End If

    'return
    If isThreeByteChanEvt Then
        Set ReadChannelEvent = Factory.CreateNewChannelEvent(isRunStatus, _
                                                             deltaTime, _
                                                             absoluteTime, _
                                                             statusNibble, _
                                                             channelNibble, _
                                                             dataByte1, _
                                                             dataByte2)
    Else 'is 2 byte channel event TODO: should there be a 3rd branch to catch errors?
        Set ReadChannelEvent = Factory.CreateNewChannelEvent(isRunStatus, deltaTime, absoluteTime, statusNibble, channelNibble, dataByte1)
    End If
End Function

Private Function ReadSystemExclusiveOrMetaEvent(ByVal deltaTime As Long, _
                                                ByVal absoluteTime As Long, _
                                                trackBytes() As Byte, _
                                                ByVal eventStartPosition As Long) As Object
    'reads system exclusive or meta event that starts at specified position in
    'an array of track bytes.
    Dim midiStatus As Byte
    Dim evtDataLength As Long
    Dim vlvStartPosition As Long
    Dim currentPosition As Long
    Dim eventEndPosition As Long
    Dim evtData As Collection
    Dim vlvBytes As Collection
    Dim midiMetaType As Byte 'TODO: standardize var names
    Dim systemExType As SystemExclusiveType
    Dim isMetaEvt As Boolean

    currentPosition = eventStartPosition
    'TODO: refactor into small methods, if not too slow
    'status
    midiStatus = trackBytes(currentPosition)
    currentPosition = currentPosition + 1

    isMetaEvt = IsMetaEvent(midiStatus)
    If isMetaEvt Then
        'meta type
        midiMetaType = trackBytes(currentPosition)
        currentPosition = currentPosition + 1
    Else
        'exclusive type
        If midiStatus = StatusEnum.SYSTEM_EXCLUSIVE_START Then
            systemExType = NORMAL
        Else
            systemExType = DIVIDED
        End If
    End If

    'length of vlv
    vlvStartPosition = currentPosition
    Set vlvBytes = Convert.GetVLVBytes(trackBytes, vlvStartPosition)
    currentPosition = currentPosition + vlvBytes.Count

    'length of data
    evtDataLength = Convert.DecodeVLV(vlvBytes)

    'data
    eventEndPosition = currentPosition + evtDataLength - 1
    Set evtData = ListUtils.CollectionSliceFromArray(trackBytes, sliceStart:=currentPosition, _
                                                     sliceEnd:=eventEndPosition)

    'return
    If isMetaEvt Then
        Set ReadSystemExclusiveOrMetaEvent = Factory.CreateNewMetaEvent(deltaTime, _
                                                                        absoluteTime, _
                                                                        midiMetaType, evtData)
    Else
        Set ReadSystemExclusiveOrMetaEvent = _
        Factory.CreateNewSystemExclusiveEvent(deltaTime, absoluteTime, midiStatus, _
                                              evtData, systemExType)

    End If
End Function


