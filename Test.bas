Attribute VB_Name = "Test"
Option Explicit

Sub RunTests()
    Test_Utilities_CollectionSlice
    Test_Convert_EncodeVLV
    Test_Convert_DecodeVLV
    Test_Convert_ToFourBytesFromLong
    Test_ChannelEvent_Initialize_2ByteEvt
    Test_ChannelEvent_Initialize_3ByteEvt
    Test_ChannelEvent_InitializeFromEventWithRunnnigStatus_2ByteEvt
    Test_ChannelEvent_InitializeFromEventWithRunnnigStatus_3ByteEvt
    Test_MetaEvent_Initialize
    Test_SystemExclusiveEvent_Initialize
    Test_EventParser_ReadEvent_ChannelEvent
    Test_EventParser_ReadEvent_ChannelEventWithRunningStatus
    Test_EventParser_ReadEvent_MetaEvent
    Test_EventParser_ReadEvent_SyxEvent
    Test_TrackParser_ParseTrack
    Test_TrackParser_ParseTracks
    Test_MidiFile
    Test_MidiFileWithWrongTrackLength
End Sub

Sub Test_Utilities_CollectionSlice()
    Dim sourceCollection As Collection
    Dim sliceStart As Long
    Dim sliceEnd As Long
    Dim slice As Collection
    Dim i As Long
    Dim j As Long
    
    sliceStart = 2
    sliceEnd = 5
    Set sourceCollection = New Collection
    sourceCollection.Add 1
    sourceCollection.Add 2
    sourceCollection.Add 3
    sourceCollection.Add 4
    sourceCollection.Add 5
    sourceCollection.Add 6
    
    Set slice = ListUtils.CollectionSlice(sourceCollection, sliceStart, sliceEnd)
    i = sliceStart
    j = 1
    For i = sliceStart To sliceEnd
        Debug.Assert slice(j) = sourceCollection(i)
        j = j + 1
    Next i
End Sub

Sub Test_ChannelEvent_Initialize_3ByteEvt()
    Dim ce As ChannelEvent
    Dim evtBytes() As Byte
    
    Set ce = Factory.CreateNewChannelEvent(False, 0, 1, &H9, 2, 60, 68)
    Debug.Assert ce.Channel = 2
    Debug.Assert ce.Data1 = 60
    Debug.Assert ce.Data2 = 68
    Debug.Assert ce.Delta = 0
    Debug.Assert ce.ChannelEventType = ceThreeByte
    Debug.Assert ce.Status = &H9
    Debug.Assert ce.StatusName = "NoteOn"
    Debug.Assert ce.TimeStamp = 1
    Debug.Assert ce.EventLength = 3
    
    evtBytes = ce.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = CByte("&H" & (Hex(&H9) & Hex(&H2)))
    Debug.Assert evtBytes(2) = 60
    Debug.Assert evtBytes(3) = 68
End Sub

Sub Test_ChannelEvent_Initialize_2ByteEvt()
    Dim ce As ChannelEvent
    Dim evtBytes() As Byte
    
    Set ce = Factory.CreateNewChannelEvent(False, 0, 1, &HC, 2, 60)
    Debug.Assert ce.Channel = 2
    Debug.Assert ce.Data1 = 60
    Debug.Assert ce.Delta = 0
    Debug.Assert ce.ChannelEventType = ceTwoByte
    Debug.Assert ce.Status = &HC
    Debug.Assert ce.StatusName = "ProgramChange"
    Debug.Assert ce.TimeStamp = 1
    Debug.Assert ce.EventLength = 2
    
    evtBytes = ce.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = CByte("&H" & (Hex(&HC) & Hex(&H2)))
    Debug.Assert evtBytes(2) = 60
End Sub

Sub Test_ChannelEvent_InitializeFromEventWithRunnnigStatus_3ByteEvt()
    Dim ce As ChannelEvent
    Dim evtBytes() As Byte
    
    Set ce = Factory.CreateNewChannelEvent(True, 0, 1, &H9, 2, 60, 68)
    Debug.Assert ce.Channel = 2
    Debug.Assert ce.Data1 = 60
    Debug.Assert ce.Data2 = 68
    Debug.Assert ce.Delta = 0
    Debug.Assert ce.ChannelEventType = ceThreeByte
    Debug.Assert ce.Status = &H9
    Debug.Assert ce.StatusName = "NoteOn"
    Debug.Assert ce.TimeStamp = 1
    Debug.Assert ce.EventLength = 2
    
    evtBytes = ce.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = 60
    Debug.Assert evtBytes(2) = 68
End Sub

Sub Test_ChannelEvent_InitializeFromEventWithRunnnigStatus_2ByteEvt()
    Dim ce As ChannelEvent
    Dim evtBytes() As Byte
    
    Set ce = Factory.CreateNewChannelEvent(True, 0, 1, &HC, 2, 60)
    Debug.Assert ce.Channel = 2
    Debug.Assert ce.Data1 = 60
    Debug.Assert ce.Delta = 0
    Debug.Assert ce.ChannelEventType = ceTwoByte
    Debug.Assert ce.Status = &HC
    Debug.Assert ce.StatusName = "ProgramChange"
    Debug.Assert ce.TimeStamp = 1
    Debug.Assert ce.EventLength = 1
    
    evtBytes = ce.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = 60
End Sub

Sub Test_MetaEvent_Initialize()
    Dim mev As MetaEvent
    Dim evtBytes() As Byte
    Dim evtData As Collection

    Set evtData = New Collection
    evtData.Add 15
    
    Set mev = Factory.CreateNewMetaEvent(0, 1, 32, evtData)
    Debug.Assert mev.Data(1) = evtData(1)
    Debug.Assert mev.Delta = 0
    Debug.Assert mev.DataLength = 1
    Debug.Assert mev.MetaType = 32
    Debug.Assert mev.MetaTypeName = "MidiChannelPrefix"
    Debug.Assert mev.Status = 255
    Debug.Assert mev.TimeStamp = 1
    
    evtBytes = mev.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = 255
    Debug.Assert evtBytes(2) = 32
    Debug.Assert evtBytes(3) = 1
    Debug.Assert evtBytes(4) = evtData(1)
End Sub

Sub Test_SystemExclusiveEvent_Initialize()
    Dim se As SystemExclusiveEvent
    Dim evtBytes() As Byte
    Dim evtData As Collection
    
    Set evtData = New Collection
    evtData.Add 10
    evtData.Add &HF7
    
    Set se = Factory.CreateNewSystemExclusiveEvent(0, 1, &HF0, evtData, NORMAL)
    Debug.Assert se.Data(1) = evtData(1)
    Debug.Assert se.Data(2) = evtData(2)
    Debug.Assert se.Delta = 0
    Debug.Assert se.EvtType = NORMAL
    Debug.Assert se.DataLength = 2
    Debug.Assert se.Status = &HF0
    Debug.Assert se.EventLength = 4
    Debug.Assert se.TimeStamp = 1
        
    evtBytes = se.ToBytes()
    Debug.Assert evtBytes(0) = 0
    Debug.Assert evtBytes(1) = &HF0
    Debug.Assert evtBytes(2) = 2
    Debug.Assert evtBytes(3) = evtData(1)
    Debug.Assert evtBytes(4) = evtData(2)
End Sub

Sub Test_MidiFile()
    Dim mf As MidiFile
    Set mf = Factory.CreateNewMidiFileFromArray(Mock.GetTestFile())
    Debug.Assert IsEqual(Mock.GetTestFile(), mf.FileBytes)
End Sub

Sub Test_MidiFileWithWrongTrackLength()
    Dim mf As MidiFile
    Set mf = Factory.CreateNewMidiFileFromArray(Mock.GetTestFileWithWrongTrackLength())
    Debug.Assert Not IsEqualExitsOnFirstDiff(Mock.GetTestFileWithWrongTrackLength(), _
                                             mf.FileBytes)
End Sub

Sub Test_Convert_ToFourBytesFromLong()
    Dim fourByteArray() As Byte
    fourByteArray = Convert.ToFourBytesFromLong(&HABCDEF1)
    Debug.Assert fourByteArray(0) = &HA
    Debug.Assert fourByteArray(1) = &HBC
    Debug.Assert fourByteArray(2) = &HDE
    Debug.Assert fourByteArray(3) = &HF1
End Sub

Sub Test_Convert_EncodeVLV()
    Dim vlvBytes() As Byte
    
    vlvBytes = EncodeVLV(0)
    Debug.Assert vlvBytes(0) = 0
    Debug.Assert UBound(vlvBytes) = 0
    
    vlvBytes = EncodeVLV(127)
    Debug.Assert vlvBytes(0) = 127
    Debug.Assert UBound(vlvBytes) = 0
    
    vlvBytes = EncodeVLV(128)
    Debug.Assert vlvBytes(0) = &H81
    Debug.Assert vlvBytes(1) = &H0
    Debug.Assert UBound(vlvBytes) = 1
    
    vlvBytes = EncodeVLV(268)
    Debug.Assert vlvBytes(0) = &H82
    Debug.Assert vlvBytes(1) = &HC
    Debug.Assert UBound(vlvBytes) = 1
    
    vlvBytes = EncodeVLV(1000)
    Debug.Assert vlvBytes(0) = &H87
    Debug.Assert vlvBytes(1) = &H68
    Debug.Assert UBound(vlvBytes) = 1
    
    vlvBytes = EncodeVLV(&HF4240)
    Debug.Assert vlvBytes(0) = &HBD
    Debug.Assert vlvBytes(1) = &H84
    Debug.Assert vlvBytes(2) = &H40
    Debug.Assert UBound(vlvBytes) = 2
    
    vlvBytes = EncodeVLV(16383)
    Debug.Assert vlvBytes(0) = &HFF
    Debug.Assert vlvBytes(1) = &H7F
    Debug.Assert UBound(vlvBytes) = 1
    
    vlvBytes = EncodeVLV(32768)
    Debug.Assert vlvBytes(0) = &H82
    Debug.Assert vlvBytes(1) = &H80
    Debug.Assert vlvBytes(2) = &H0
    Debug.Assert UBound(vlvBytes) = 2
    
    vlvBytes = EncodeVLV(&HE7A14F5)
    Debug.Assert vlvBytes(0) = &HF3
    Debug.Assert vlvBytes(1) = &HE8
    Debug.Assert vlvBytes(2) = &HA9
    Debug.Assert vlvBytes(3) = &H75
    Debug.Assert UBound(vlvBytes) = 3
    
    vlvBytes = EncodeVLV(MAX_MIDI_VALUE)
    Debug.Assert vlvBytes(0) = &HFF
    Debug.Assert vlvBytes(1) = &HFF
    Debug.Assert vlvBytes(2) = &HFF
    Debug.Assert vlvBytes(3) = &H7F
    Debug.Assert UBound(vlvBytes) = 3
    
    vlvBytes = EncodeVLV(MAX_MIDI_VALUE + 1)
    Debug.Assert vlvBytes(0) = 0
    
    vlvBytes = EncodeVLV(-1)
    Debug.Assert vlvBytes(0) = 0
End Sub

Sub Test_Convert_DecodeVLV()
    Dim vlvBytes As Collection
    Dim vlv As Long
        
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(0))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 0
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(127))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 127
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(128))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 128
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(268))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 268
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(1000))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 1000
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(&HF4240))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = &HF4240
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(16383))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 16383
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(32768))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 32768
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(&HE7A14F5))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = &HE7A14F5
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(MAX_MIDI_VALUE))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = MAX_MIDI_VALUE
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(MAX_MIDI_VALUE + 1))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 0
    
    Set vlvBytes = ListUtils.ToCollectionFromByteArray(EncodeVLV(-1))
    vlv = DecodeVLV(vlvBytes)
    Debug.Assert vlv = 0
End Sub

Sub Test_TrackParser_ParseTrack()
    Dim trkChunk As TrackChunk
    Set trkChunk = Factory.CreateNewMidiTrackChunk(Mock.GetTestTrack())
    Dim evtTrack As EventTrack
    Set evtTrack = TrackParser.ParseTrack(trkChunk)
    Dim byteArray() As Byte
    byteArray = evtTrack.ToBytes()

    Debug.Assert IsEqual(trkChunk.ChunkBytes, byteArray)
End Sub

Sub Test_TrackParser_ParseTracks()
    Dim trkChunks As TrackChunks
    Set trkChunks = Factory.CreateNewMidiTrackChunks()
    Dim i As Long
    For i = 0 To 3
        trkChunks.Add Factory.CreateNewMidiTrackChunk(Mock.GetTestTrack())
    Next i
    
    Dim evtTracks As EventTracks
    Set evtTracks = TrackParser.ParseTracks(trkChunks)

    For i = 1 To evtTracks.Count()
        Debug.Assert IsEqual(evtTracks.Item(i).ToBytes, Mock.GetTestTrack())
    Next i
End Sub

Sub Test_EventParser_ReadEvent_ChannelEvent()
    Dim evtBytes() As Byte
    Dim evtBytesV() As Variant
    Dim evt2 As ChannelEvent
    evtBytesV = Array(&H83, 64, 127)
    evtBytes = ListUtils.ToByteArray(evtBytesV)
    Set evt2 = EventParser.ReadEvent(deltaTime:=50, _
                                     absoluteTime:=100, _
                                     midiStatus:=&H83, _
                                     trackPosition:=0, _
                                     prevStatusChan:=0, _
                                     trkChunkBytes:=evtBytes)
    Debug.Assert IsEqual(Mock.GetTestChannelEvent.ToBytes(), evt2.ToBytes())
End Sub

Sub Test_EventParser_ReadEvent_ChannelEventWithRunningStatus()
    Dim evtBytes() As Byte
    Dim evtBytesV() As Variant
    Dim evt2 As ChannelEvent
    evtBytesV = Array(64, 127)
    evtBytes = ToByteArray(evtBytesV)
    Set evt2 = EventParser.ReadEvent(deltaTime:=50, _
                                     absoluteTime:=100, _
                                     midiStatus:=64, _
                                     trackPosition:=0, _
                                     prevStatusChan:=&H83, _
                                     trkChunkBytes:=evtBytes)
    Debug.Assert IsEqual(Mock.GetTestRunningStatus.ToBytes(), evt2.ToBytes())
End Sub

Sub Test_EventParser_ReadEvent_MetaEvent()
    Dim evtBytes() As Byte
    Dim evtBytesV() As Variant
    Dim evt2 As MetaEvent
    evtBytesV = Array(&HFF, 32, 1, 15)
    evtBytes = ToByteArray(evtBytesV)
    Set evt2 = EventParser.ReadEvent(deltaTime:=50, _
                                     absoluteTime:=100, _
                                     midiStatus:=&HFF, _
                                     trackPosition:=0, _
                                     prevStatusChan:=0, _
                                     trkChunkBytes:=evtBytes)
    Debug.Assert IsEqual(Mock.GetTestMetaEvent.ToBytes(), evt2.ToBytes())
End Sub

Sub Test_EventParser_ReadEvent_SyxEvent()
    Dim evtBytes() As Byte
    Dim evtBytesV() As Variant
    Dim evt2 As SystemExclusiveEvent
    evtBytesV = Array(&HF0, 2, 10, &HF7)
    evtBytes = ToByteArray(evtBytesV)
    Set evt2 = EventParser.ReadEvent(deltaTime:=50, _
                                     absoluteTime:=100, _
                                     midiStatus:=&HF0, _
                                     trackPosition:=0, _
                                     prevStatusChan:=0, _
                                     trkChunkBytes:=evtBytes)
    Debug.Assert IsEqual(Mock.GetTestSystemExclusiveEvent.ToBytes(), evt2.ToBytes())
End Sub

Sub Test_EventParser_ReadEvent_InvalidStatus()
    Dim evtBytes() As Byte
    Dim evtBytesV() As Variant
    Dim evt2 As SystemExclusiveEvent
    evtBytesV = Array(&HF1, 2, 10, &HF7)
    evtBytes = ToByteArray(evtBytesV)
    Set evt2 = EventParser.ReadEvent(deltaTime:=50, _
                                     absoluteTime:=100, _
                                     midiStatus:=&HFF, _
                                     trackPosition:=0, _
                                     prevStatusChan:=0, _
                                     trkChunkBytes:=evtBytes)
    Debug.Assert IsEqual(Mock.GetTestSystemExclusiveEvent.ToBytes(), evt2.ToBytes())
End Sub

Private Function IsEqual(byteArray1() As Byte, byteArray2() As Byte) As Boolean
    Dim b As Variant
    Dim i As Long
    i = 0
    For Each b In byteArray1
        IsEqual = (byteArray1(i) = byteArray2(i))
        Debug.Assert IsEqual
        If Not IsEqual Then
            Debug.Print "IsEqual(array1, array2): the value at index " & i & " differs."
            Debug.Print "array1(" & i & ") = " & byteArray1(i) & _
            ", array2(" & i & ") = " & byteArray2(i)
        End If
        i = i + 1
    Next b
End Function

Private Function IsEqualExitsOnFirstDiff(byteArray1() As Byte, byteArray2() As Byte) As Boolean
    Dim b As Variant
    Dim i As Long
    i = 0
    For Each b In byteArray1
        IsEqualExitsOnFirstDiff = (byteArray1(i) = byteArray2(i))
        If Not IsEqualExitsOnFirstDiff Then
            Exit For
        End If
        i = i + 1
    Next b
End Function
   
  
    
