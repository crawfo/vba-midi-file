Attribute VB_Name = "Midi"
Option Explicit

Public Enum SystemExclusiveType
    etNormal = 0
    etDivided = 1
    etAuthorization = 2
End Enum

Public Enum StatusChannelEvent
    ceNoteOff = &H8
    ceNoteOn = &H9
    ceNoteAftertouch = &HA
    ceController = &HB
    ceProgramChange = &HC
    ceChannelAftertouch = &HD
    cePitchBend = &HE
End Enum

Public Enum StatusNonChannelEvent
    msMetaEvent = &HFF
End Enum

Public Enum MetaEventType
    meSequenceNumber = &H0
    meTextEvent = &H1
    meCopyrightNotice = &H2
    meSequenceTrackName = &H3
    meInstrumentName = &H4
    meLyrics = &H5
    meMarker = &H6
    meCuePoint = &H7
    meMidiChannelPrefix = &H20
    meMidiPort = &H21
    meEndOfTrack = &H2F
    meSetTempo = &H51
    meSmpteOffset = &H54
    meTimeSignature = &H58
    meKeySignature = &H59
    meSequencerSpecific = &H7F
End Enum

Public Enum ContinuousControllerType
    ccBankSelectMSB = 0
    ccModulationMSB = 1
    ccBreathControllerMSB = 2
    ccFootControllerMSB = 4
    ccPortamentoTimeMSB = 5
    ccDataEntryMSB = 6
    ccMainVolumeMSB = 7
    ccBalanceMSB = 8
    ccPanMSB = 10
    ccExpressionControllerMSB = 11
    ccEffectControl1MSB = 12
    ccEffectControl2MSB = 13
    
    ccContinuousController14 = 14
    ccContinuousController15 = 15
    
    ccGeneralPurposeController1 = 16
    ccGeneralPurposeController2 = 17
    ccGeneralPurposeController3 = 18
    ccGeneralPurposeController4 = 19
    
    ccContinuousController20 = 20
    ccContinuousController21 = 21
    ccContinuousController22 = 22
    ccContinuousController23 = 23
    ccContinuousController24 = 24
    ccContinuousController25 = 25
    ccContinuousController26 = 26
    ccContinuousController27 = 27
    ccContinuousController28 = 28
    ccContinuousController29 = 29
    ccContinuousController30 = 30
    ccContinuousController31 = 31
    
    ccBankSelectLSB = 32
    ccModulationLSB = 33
    ccBreathControllerLSB = 34
    ccFootControllerLSB = 36
    ccPortamentoTimeLSB = 37
    ccDataEntryLSB = 38
    ccMainVolumeLSB = 39
    ccBalanceLSB = 40
    ccPanLSB = 42
    ccExpressionControllerLSB = 43
    ccEffectControl1LSB = 44
    ccEffectControl2LSB = 45
    ccController14LSB = 46
    ccController15LSB = 47
    ccController16LSB = 48
    ccController17LSB = 49
    ccController18LSB = 50
    ccController19LSB = 51
    ccController20LSB = 52
    ccController21LSB = 53
    ccController22LSB = 54
    ccController23LSB = 55
    ccController24LSB = 56
    ccController25LSB = 57
    ccController26LSB = 58
    ccController27LSB = 59
    ccController28LSB = 60
    ccController29LSB = 61
    ccController30LSB = 62
    ccController31LSB = 63

    ccSustainPedal = 64
    ccPortamento = 65
    ccSostenuto = 66
    ccSoftPedal = 67
    ccLegatoFootswitch = 68
    ccHold2Pedal = 69

    ccSoundVariation = 70
    ccSoundResonance = 71
    ccSoundReleaseTime = 72
    ccSoundAttackTime = 73
    ccSoundFrequencyCutoff = 74
    ccSoundController6 = 75
    ccSoundController7 = 76
    ccSoundController8 = 77
    ccSoundController9 = 78
    ccSoundController10 = 79
    
    ccGeneralPurposeController5 = 80
    ccGeneralPurposeController6 = 81
    ccGeneralPurposeController7 = 82
    ccGeneralPurposeController8 = 83
    
    ccPortamentoControl = 84
    
    ccContinuousController85 = 85
    ccContinuousController86 = 86
    ccContinuousController87 = 87
    ccContinuousController88 = 88
    ccContinuousController89 = 89
    ccContinuousController90 = 90

    ccEffects1DepthExternalEffectsDepth = 91
    ccEffects2DepthTremoloDepth = 92
    ccEffects3DepthChorusDepth = 93
    ccEffects4DepthCelesteDetune = 94
    ccEffects5DepthPhaserDepth = 95
    
    ccDataIncrement = 96
    ccDataDecrement = 97
    ccNonRegisteredParameterNumberLSB = 98
    ccNonRegisteredParameterNumberMSB = 99
    ccRegisteredParameterNumberLSB = 100
    ccRegisteredParameterNumberMSB = 101

End Enum

Public Enum KeySignatureKeyValue
    ksCMajor = 0
    ksGMajor = 1
    ksDMajor = 2
    ksAMajor = 3
    ksEMajor = 4
    ksBMajor = 5
    ksFsMajor = 6
    ksCsMajor = 7
    ksAMinor = ksCMajor
    ksEMinor = ksGMajor
    ksBMinor = ksDMajor
    ksFsMinor = ksAMajor
    ksCsMinor = ksEMajor
    ksGsMinor = ksBMajor
    ksDsMinor = ksFsMajor
    ksAsMinor = ksCsMajor
    ksFMajor = 255
    ksBbMajor = 254
    ksEbMajor = 253
    ksAbMajor = 252
    ksDbMajor = 251
    ksGbMajor = 250
    ksCbMajor = 249
    ksDMinor = ksFMajor
    ksGMinor = ksBbMajor
    ksCMinor = ksEbMajor
    ksFMinor = ksAbMajor
    ksBbMinor = ksDbMajor
    ksEbMinor = ksGbMajor
    ksAbMinor = ksCbMajor
End Enum

Public Enum KeySignatureModeValue
    kmMajor = 0
    kmMinor = 1
End Enum

Public Enum KeySignatureFull
    fkCMajor
    fkGMajor
    fkDMajor
    fkAMajor
    fkEMajor
    fkBMajor
    fkFsMajor
    fkCsMajor
    fkFMajor
    fkBbMajor
    fkEbMajor
    fkAbMajor
    fkDbMajor
    fkGbMajor
    fkCbMajor
    fkAMinor
    fkEMinor
    fkBMinor
    fkFsMinor
    fkCsMinor
    fkGsMinor
    fkDsMinor
    fkAsMinor
    fkDMinor
    fkGMinor
    fkCMinor
    fkFMinor
    fkBbMinor
    fkEbMinor
    fkAbMinor
    fkUnknown
End Enum

Public Const NOT_IN_RANGE_0_127 = "Value must be in range 0-127. Error raised by ChannelEvent Property Let Data2. "
Public Const INACTIVE_PROPERTY = "Property is inactive when ChannelEventType is ceTwoByte"

Public Const THREE_BYTE_CHAN_EVT_LEN = 3
Public Const TWO_BYTE_CHAN_EVT_LEN = 2
Public Const STATUS_SYSEX_START = &HF0
Public Const STATUS_SYSEX_CONTINUE = &HF7
Public Const TRACK_HEADER_LENGTH = 8

Public Const SYSTEM_EXCLUSIVE_START_NORMAL = &HF0
Public Const SYSTEM_EXCLUSIVE_START_DIVIDED = &HF7
Public Const SYSTEM_EXCLUSIVE_END = &HF7

Public Const MAX_MIDI_VALUE = &HFFFFFFF '(268435455)
Public Const MICROSECONDS_PER_MINUTE = 60000000

Private Function ParseTrack(trackBytes() As Byte) As Collection
     Dim prevStatusChan As Byte 'used for running status
     Dim absoluteTime As Long
     Dim deltaTime As Long, i As Long
     Dim trackPosition As Long
     Dim midiStatus As Byte
     Dim midiEvents As Collection, vlvBytes As Collection
     Dim upperBound As Long
     
     Set midiEvents = New Collection
     trackPosition = 0
     upperBound = UBound(trackBytes)
     
     'loop thru track parsing msgs
     Do While trackPosition < upperBound
        'read vlv (variable length value)
        Set vlvBytes = GetVLVBytes(trackBytes, trackPosition)
        deltaTime = DecodeVLV(vlvBytes)
        absoluteTime = absoluteTime + deltaTime
        'incr trackPosition to start of next event
        trackPosition = trackPosition + vlvBytes.Count
        'get event
        'read status byte
        midiStatus = trackBytes(trackPosition)
        'eval status byte and adjust trackPosition to start of next event
        If IsChannelEvent(midiStatus) Then
            'for a ChannelEvent the status byte also contains channel info (lower nibble)
            midiEvents.Add ReadChannelEvent(deltaTime, absoluteTime, trackBytes, trackPosition)
            'handle 2 and 3 byte msgs
            If IsTwoByteChannelEvent(midiStatus) Then
                trackPosition = trackPosition + TWO_BYTE_CHAN_EVT_LEN
            Else 'is 3 byte msg
                trackPosition = trackPosition + THREE_BYTE_CHAN_EVT_LEN
            End If
            'store status/chan for running status
            prevStatusChan = midiStatus
        ElseIf IsRunningStatus(midiStatus) Then
            midiEvents.Add ReadChannelEventRunningStatus(deltaTime, absoluteTime, trackBytes, trackPosition, prevStatusChan)
            'incr trk pos
            If IsTwoByteChannelEvent(prevStatusChan) Then
                trackPosition = trackPosition + TWO_BYTE_CHAN_EVT_LEN - 1
            Else
                trackPosition = trackPosition + TWO_BYTE_CHAN_EVT_LEN
            End If
        ElseIf IsMetaEvent(midiStatus) Then
            midiEvents.Add ReadMetaEvent(deltaTime, absoluteTime, trackBytes, trackPosition)
            trackPosition = trackPosition + 1
        ElseIf IsSysExEvent(midiStatus) Then
            midiEvents.Add ReadSystemExclusiveEvent(deltaTime, absoluteTime, trackBytes, trackPosition)
            trackPosition = trackPosition + 1
        Else
            Stop
        End If
    Loop
    
    Set ParseTrack = midiEvents
End Function

Private Function ReadMidiFile(ByVal fileNameFullyQualified As String) As Byte()
    Dim bytes() As Byte
    Open fileNameFullyQualified For Binary As #1
    ReDim bytes(LOF(1) - 1)
    Get #1, , bytes
    Close #1
    ReadMidiFile = bytes
End Function

Public Function ParseMidiFile(ByVal fileNameFullyQualified As String) As Collection
    Dim midiFileBytes() As Byte
    Dim parsedTracks As Collection, i As Long, trackCount As Long
    Dim rawTracks As Collection
   
    midiFileBytes = ReadMidiFile(fileNameFullyQualified)
    Set rawTracks = GetTracks(midiFileBytes, GetTrackDimensions(midiFileBytes))
    Set parsedTracks = New Collection
    trackCount = rawTracks.Count
        
    For i = 1 To trackCount
        parsedTracks.Add ParseTrack(rawTracks(i))
    Next i
    
    Set ParseMidiFile = parsedTracks
End Function

Private Function GetFileInfo(midiFileBytes() As Byte) As MidiFileInfo
    Dim midiFileFormat As Integer, trackCount As Long, TimeDivision As Long
    Dim isPpqTime As Boolean, isFpsTime As Boolean, fpsByte1 As Byte, fpsByte2 As Byte
    Dim bitmaskByte1 As Byte, bitmaskByte2 As Byte
    Dim smpteFrames As Byte 'the number of SMPTE frames can be 24, 25, 29 (for 29.97 fps) or 30
    Dim ticksPerFrame As Byte
    
    isPpqTime = midiFileBytes(12) <= &H7F
    midiFileFormat = midiFileBytes(9)
    trackCount = JoinTwoBytes(midiFileBytes(10), midiFileBytes(11))
    bitmaskByte1 = &H7F
    bitmaskByte2 = &HFF
    
    If isPpqTime Then
        TimeDivision = JoinTwoBytes(midiFileBytes(12), midiFileBytes(13))
    Else
        'is SMPTE frames/sec. TODO: not fully implemented
        fpsByte1 = bitmaskByte1 And midiFileBytes(12)
        fpsByte2 = bitmaskByte2 And midiFileBytes(13)
    End If
    
    Set GetFileInfo = Factory.CreateNewMidiFileInfo(midiFileFormat, trackCount, TimeDivision)
End Function

Private Function ReadMetaEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, trackBytes() As Byte, ByRef eventStartPosition As Long) As MetaEvent
    'side effect: mutates parameter eventStartPosition
    
    Dim vlvStartPosition As Long, currentPosition As Long, eventEndPosition As Long, i As Long
    Dim eventLength As Long, vlvBytes As Collection
    Dim midiStatus As Byte, midiMetaType As Byte, eventData As Collection
    
    currentPosition = eventStartPosition
    Set eventData = New Collection
    
    'status
    midiStatus = trackBytes(currentPosition)
    currentPosition = currentPosition + 1
    'meta type
    midiMetaType = trackBytes(currentPosition)
    currentPosition = currentPosition + 1
    
    'length vlv
    vlvStartPosition = currentPosition
    Set vlvBytes = GetVLVBytes(trackBytes, vlvStartPosition)
    eventLength = DecodeVLV(vlvBytes)
    currentPosition = currentPosition + vlvBytes.Count
    
    'data
    eventEndPosition = currentPosition + eventLength - 1
    For i = currentPosition To eventEndPosition
        eventData.Add trackBytes(i)
    Next i
    
    'return
    Set ReadMetaEvent = Factory.CreateNewMetaEvent(deltaTime, absoluteTime, midiMetaType, eventData)
    eventStartPosition = eventEndPosition
End Function

Private Function ReadChannelEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, trackBytes() As Byte, ByVal eventStartPosition As Long) As ChannelEvent
    'reads channel event bytes
    Dim midiStatusNibble As Byte, midiChannelNibble As Byte, dataByte1  As Byte, dataByte2 As Byte
    Dim statusChannelByte As Byte, isThreeByteChanEvt As Boolean

    'status/channel
    statusChannelByte = trackBytes(eventStartPosition)
    midiStatusNibble = GetNibbleHigh(statusChannelByte)
    midiChannelNibble = GetNibbleLow(statusChannelByte)
    isThreeByteChanEvt = IsThreeByteChannelEvent(statusChannelByte)
    
    'data
    dataByte1 = trackBytes(eventStartPosition + 1)
    If isThreeByteChanEvt Then
        dataByte2 = trackBytes(eventStartPosition + 2)
    End If
    
    'return
    If isThreeByteChanEvt Then
        Set ReadChannelEvent = Factory.CreateNewChannelEvent(deltaTime, absoluteTime, midiStatusNibble, midiChannelNibble, dataByte1, dataByte2)
    Else
        Set ReadChannelEvent = Factory.CreateNewChannelEvent(deltaTime, absoluteTime, midiStatusNibble, midiChannelNibble, dataByte1)
    End If
End Function

Private Function ReadChannelEventRunningStatus(ByVal deltaTime As Long, ByVal absoluteTime As Long, trackBytes() As Byte, ByVal eventStartPosition As Long, bytPrevStatusChan As Byte) As ChannelEvent
    'reads channel event running status bytes
    Dim midiStatus As Byte, Channel As Byte, dataByte1  As Byte, dataByte2 As Byte
    Dim statusChannelByte As Byte, isThreeByteChanEvt As Boolean
    
    'status/channel
    midiStatus = GetNibbleHigh(bytPrevStatusChan)
    Channel = GetNibbleLow(bytPrevStatusChan)
    isThreeByteChanEvt = IsThreeByteChannelEvent(bytPrevStatusChan)
    
    'data
    dataByte1 = trackBytes(eventStartPosition)
    If isThreeByteChanEvt Then
        'dataByte2 = clnTrack(eventStartPosition + 1)
        dataByte2 = trackBytes(eventStartPosition + 1)
    End If
        
    'return chanEvt
    If isThreeByteChanEvt Then
        Set ReadChannelEventRunningStatus = Factory.CreateNewChannelEvent(deltaTime, absoluteTime, midiStatus, Channel, dataByte1, dataByte2)
    Else
        Set ReadChannelEventRunningStatus = Factory.CreateNewChannelEvent(deltaTime, absoluteTime, midiStatus, Channel, dataByte1)
    End If
End Function

Private Function ReadSystemExclusiveEvent(ByVal deltaTime As Long, ByVal absoluteTime As Long, trackBytes() As Byte, ByRef eventStartPosition As Long) As SystemExclusiveEvent
    'reads SystemExclusive bytes
    Dim midiStatus As Byte, eventLength As Long, vlvByteCount As Long
    Dim vlvStartPosition As Long, currentPosition As Long, eventEndPosition As Long, i As Long
    Dim eventData As Collection, vlvBytes As Collection, systemExType As SystemExclusiveType
    
    Set eventData = New Collection
    currentPosition = eventStartPosition
    
    'status
    midiStatus = trackBytes(currentPosition)
    currentPosition = currentPosition + 1
    
    'length of vlv
    vlvStartPosition = currentPosition
    Set vlvBytes = GetVLVBytes(trackBytes, vlvStartPosition)
    vlvByteCount = vlvBytes.Count
    currentPosition = currentPosition + vlvByteCount
    
    'length of data
    eventLength = DecodeVLV(vlvBytes)
    
    'data
    eventEndPosition = currentPosition + eventLength - 1
    For i = currentPosition To eventEndPosition
        eventData.Add trackBytes(i)
    Next i
    
    'type
    If midiStatus = SYSTEM_EXCLUSIVE_START_NORMAL Then
        systemExType = etNormal
    Else
        systemExType = etDivided
    End If
    
    'return
    Set ReadSystemExclusiveEvent = Factory.CreateNewSystemExclusiveEvent(deltaTime, absoluteTime, midiStatus, eventData, systemExType)
    eventStartPosition = eventEndPosition
   
End Function

Private Function GetTrack(midiFileBytes() As Byte, ByVal trackDataStartPosition As Long, ByVal trackDataEndPosition As Long) As Byte()
    Dim i As Long
    Dim trackBytes() As Byte
    
    ReDim trackBytes(trackDataEndPosition - trackDataStartPosition + 1)
    
    For i = trackDataStartPosition To trackDataEndPosition
        trackBytes(i - trackDataStartPosition) = midiFileBytes(i)
    Next i
    
    GetTrack = trackBytes
End Function

Private Function GetTracks(midiFileBytes() As Byte, midiTrackDimensions As TrackDimensions) As Collection
    Dim i As Long, trackCount As Long
    Dim clnTracks As Collection
    
    Set clnTracks = New Collection
    trackCount = midiTrackDimensions.DataSize.Count 'num elements in each TrackDimensions array = num tracks
    
    For i = 1 To trackCount
        clnTracks.Add GetTrack(midiFileBytes, midiTrackDimensions.DataStart(i), midiTrackDimensions.DataEnd(i))
    Next i
    
    Set GetTracks = clnTracks
End Function

Private Function GetTrackDimensions(midiFileBytes() As Byte) As TrackDimensions
    Dim i As Long, j As Long, upperBound As Long
    Dim dataStartPosition As Long, dataEndPosition As Long, trackLength As Long
    Dim dataStartPositions As Collection, dataEndPositions As Collection, trackLengths As Collection
   
    Set dataStartPositions = New Collection
    Set dataEndPositions = New Collection
    Set trackLengths = New Collection
    i = 14 'pos of 1st byte after file hdr, ie start pos of 1st trk
    upperBound = UBound(midiFileBytes)
    
    'iterate through file bytes to find all tracks and their len, pos.
    Do While i <= upperBound
        If IsTrackChunk(i, midiFileBytes) Then
            dataStartPosition = i + 8
            i = i + 4 'move to length bytes
            trackLength = JoinFourBytes(midiFileBytes(i), midiFileBytes(i + 1), midiFileBytes(i + 2), midiFileBytes(i + 3))
            dataStartPositions.Add dataStartPosition
            i = dataStartPosition + trackLength 'increment to start of next track chunk
            trackLengths.Add trackLength
            dataEndPosition = i - 1
            dataEndPositions.Add dataEndPosition
            trackLength = 0 'init for new track
        Else
            i = i + 1
        End If
     Loop
     
     Set GetTrackDimensions = Factory.CreateNewTrackDimensions(dataStartPositions, dataEndPositions, trackLengths)
End Function

Private Function IsTrackChunk(ByVal currentPosition As Long, midiFileBytes() As Byte) As Boolean
    If currentPosition < UBound(midiFileBytes) - 3 Then 'prevent index out of bounds error
        IsTrackChunk = (midiFileBytes(currentPosition) = 77 And midiFileBytes(currentPosition + 1) = 84 _
        And midiFileBytes(currentPosition + 2) = 114 And midiFileBytes(currentPosition + 3) = 107)
    End If
End Function

Private Function GetNibbleHigh(ByVal b As Byte) As Byte
    GetNibbleHigh = (b \ 16) And &HF
End Function

Private Function GetNibbleLow(ByVal b As Byte) As Byte
    GetNibbleLow = b And &HF
End Function

Public Function JoinTwoNibbles(ByVal nibbleHigh As Byte, ByVal nibbleLow As Byte) As Byte
    'nibbleHigh * 16 is equivalent to nibbleHigh << 4
    JoinTwoNibbles = (nibbleHigh * 16) Or nibbleLow
End Function

Private Function JoinTwoBytes(ByVal byteHigh As Byte, ByVal byteLow As Byte) As Long
    'byteHigh * 256 is equivalent to byteHigh << 8 (shifting 8 bits to the left).
    JoinTwoBytes = (byteHigh * 256) Or byteLow
End Function

Private Function JoinFourBytes(ByVal byte1 As Byte, ByVal byte2 As Byte, ByVal byte3 As Byte, ByVal byte4 As Byte) As Long
    'byte1 is high and byte4 is low
    Dim hexString As String
    
    If byte1 < 16 Then
        hexString = hexString & "0" & Hex(byte1)
    Else
        hexString = hexString & Hex(byte1)
    End If
    
    If byte2 < 16 Then
        hexString = hexString & "0" & Hex(byte2)
    Else
        hexString = hexString & Hex(byte2)
    End If
    
    If byte3 < 16 Then
        hexString = hexString & "0" & Hex(byte3)
    Else
        hexString = hexString & Hex(byte3)
    End If
    
    If byte4 < 16 Then
        hexString = hexString & "0" & Hex(byte4)
    Else
        hexString = hexString & Hex(byte4)
    End If

    JoinFourBytes = CLng("&H" & hexString)
End Function

Function IsTwoByteChannelEvent(ByVal statusByte As Byte) As Boolean
   IsTwoByteChannelEvent = (statusByte >= &HC0 And statusByte <= &HDF)
End Function

Private Function IsChannelEvent(ByVal statusByte As Byte) As Boolean
   IsChannelEvent = statusByte >= &H80 And statusByte <= &HEF
End Function

Public Function IsThreeByteChannelEvent(ByVal statusByte As Byte) As Boolean
'   IsThreeByteChannelEvent = (IsChannelEvent(statusByte) = True) And (IsTwoByteChannelEvent(statusByte) = False)
    IsThreeByteChannelEvent = IsChannelEvent(statusByte) And Not IsTwoByteChannelEvent(statusByte)
End Function

Private Function IsMetaEvent(ByVal statusByte As Byte) As Boolean
   IsMetaEvent = statusByte = &HFF
End Function

Private Function IsSysExEvent(ByVal statusByte As Byte) As Boolean
   IsSysExEvent = statusByte = &HF0 Or statusByte = &HF7
End Function

Private Function IsRunningStatus(ByVal statusByte As Byte) As Boolean
   IsRunningStatus = statusByte < &H80
End Function

Public Function ToByteFromHex(ByVal hexString As String) As Byte
    ToByteFromHex = CByte("&H" & hexString)
End Function

Private Function GetVLVBytes(bytes() As Byte, ByVal startPosition As Long) As Collection
    Dim i As Long, upperBound As Long
    Dim clnVLV As Collection
    
    Set clnVLV = New Collection
    i = startPosition
    upperBound = UBound(bytes)
    
    'collect vlv bytes
    Do While IsSetMsb(bytes(i)) And i <= upperBound
        clnVLV.Add bytes(i)
        i = i + 1
    Loop
     
    'add last value but no need to check msb as if last value then msb must be 0
    clnVLV.Add bytes(i)
    
    Set GetVLVBytes = clnVLV
End Function

Private Function DecodeVLV(ByVal bytes As Collection) As Long
    Dim result As Long, n As Long, currentByte As Byte, i As Long
    n = bytes.Count
    For i = 1 To n
        'apply msb bit mask
        currentByte = bytes(i) And &HFF
        'shifts bits 7 places to left,     ,masks first 7 bits
        result = (result * 128) Or (currentByte And &H7F)
        'if last value (ie msb cleared)
        If (currentByte And &H80) = 0 Then
            Exit For
        End If
    Next i
    DecodeVLV = result
End Function

Public Function EncodeVLV(ByVal value As Long) As Byte()
    'encodes long into array of variable length value bytes
    'e.g. 32768 into array containing bytes 130,128,0
    'max value is &H0FFFFFFF (268435455)
    Dim vlvBytes() As Byte
    Dim i As Long, exponent As Integer
    ReDim vlvBytes(3) 'max value is 0FFFFFFF so max 28 bits (4 bytes) needed
    
    'split into bytes with msb (bit7) set correctly:
    '(value \ (2 ^ exponent)) shifts bits right by exponent number of bits to position lower 7 bits
    'And &H7F preserves only lower 7 bits
    'Or &H80 sets msb and preserves lower bits
    i = 0
    For exponent = 21 To 7 Step -7
        If value \ (2 ^ exponent) = 0 Then
            ReDim Preserve vlvBytes(UBound(vlvBytes) - 1)
        Else
            vlvBytes(i) = (value \ (2 ^ exponent)) And &H7F Or &H80
            i = i + 1
        End If
    Next exponent
    
    'last byte so must have msb clear not set
    vlvBytes(i) = (value \ (2 ^ exponent)) And &H7F
    
    EncodeVLV = vlvBytes
End Function

Private Function IsSetMsb(ByVal bytByte As Byte) As Boolean
    'returns true when msb is set, otherwise false
    IsSetMsb = bytByte > &H7F
End Function

Private Function GetNotes(ByVal track As Collection) As Variant()
    Dim midiEvent As Variant
    Dim i As Long
    Dim notes() As Variant
    'protect against bad values of array index
    If track.Count > 0 Then ReDim notes(track.Count - 1)
    i = 0
    
    For Each midiEvent In track
        If midiEvent.Status = ceNoteOn Or midiEvent.Status = ceNoteOff Then
            Set notes(i) = midiEvent
            i = i + 1
        End If
    Next midiEvent
    
    'protect against bad values of array index i
    If i > 0 Then
        ReDim Preserve notes(i - 1)
    Else
        ReDim Preserve notes(0)
    End If
    
    GetNotes = notes
End Function
