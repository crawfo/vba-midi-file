VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MidiFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TRACK_HEADER_ID_LENGTH = 4
Private Const END_OF_TRACK_EVENT_DATA = 0
Private pHeaderChunk As HeaderChunk
Private pTrackList As EventTracks
Private pFileBytes() As Byte

Public Property Get HeaderChunk() As HeaderChunk
    Set HeaderChunk = pHeaderChunk
End Property

Public Property Get TrackList() As EventTracks
    Set TrackList = pTrackList
End Property

Public Property Set TrackList(ByVal eventTrks As EventTracks)
    Set pTrackList = eventTrks
End Property

Public Property Get FileBytes() As Byte()
    FileBytes = pFileBytes
End Property

Public Sub InitializeFromEventTracks(ByVal hdrChunk As HeaderChunk, _
                                     ByVal eventTrks As EventTracks)
    Set pTrackList = eventTrks
    Set pHeaderChunk = hdrChunk
    ToBytes
End Sub

Public Sub Initialize(midiFileBytes() As Byte)
    Dim midiTrkChunks As TrackChunks
    Dim eventTrks As EventTracks
    
    pFileBytes = midiFileBytes
    Set pHeaderChunk = GetHeaderChunk()
    Set midiTrkChunks = GetTracks()

    Set eventTrks = TrackParser.ParseTracks(midiTrkChunks)
    InitializeFromEventTracks pHeaderChunk, eventTrks
End Sub


Public Sub AddTrack(ByVal eventTrk As EventTrack)
    pTrackList.Add eventTrk
    'update num tracks
    Set pHeaderChunk = Factory.CreateNewHeaderChunk(HeaderChunk.FileFormat, _
                                                    TrackList.Count, _
                                                    HeaderChunk.TimeDivision)
End Sub

Private Sub ToBytes()
    Dim trackBytesCollection As Collection
    
    Set trackBytesCollection = pTrackList.ToBytesCollection()
    AddFileHeaderChunkBytes trackBytesCollection.Count
    AddAllTrackBytesToFileBytes trackBytesCollection
End Sub

Private Sub AddFileHeaderChunkBytes(ByVal trackBytesCollectionSize As Long)
    Dim headerBytes() As Byte
    Dim headerByte As Variant
    Dim i As Long
    
    i = 0
    headerBytes = pHeaderChunk.ChunkBytes
    ReDim pFileBytes(trackBytesCollectionSize + UBound(headerBytes))
    'put file header chunk bytes into an array
    For Each headerByte In headerBytes
        pFileBytes(i) = headerByte
        i = i + 1
    Next headerByte
End Sub

Private Sub AddAllTrackBytesToFileBytes(ByVal trackBytesCollection As Collection)
    Dim trkByte As Variant
    Dim i As Long

    i = Midi.FILE_HEADER_LENGTH
    For Each trkByte In trackBytesCollection
        pFileBytes(i) = trkByte
        i = i + 1
    Next trkByte
End Sub

Public Function GetTracks() As TrackChunks
    Dim midiTrackDimensions As TrackDimensions
    Dim i As Long
    Dim trackCount As Long
    Dim trks As TrackChunks
    
    Set midiTrackDimensions = GetTrackDimensions()
    Set trks = Factory.CreateNewMidiTrackChunks()
    'num elements in each TrackDimensions array = num tracks
    trackCount = midiTrackDimensions.TrackStart.Count
    
    For i = 1 To trackCount
        trks.Add GetTrack(midiTrackDimensions.TrackStart(i), midiTrackDimensions.TrackEnd(i))
    Next i
    
    Set GetTracks = trks
End Function

Private Function GetTrack(ByVal trackDataStartPosition As Long, _
                          ByVal trackDataEndPosition As Long) As TrackChunk
    'returns a track as a TrackChunk
    Dim i As Long
    Dim trackBytes() As Byte
    
    ReDim trackBytes(trackDataEndPosition - trackDataStartPosition)
    trackBytes = ListUtils.ByteArraySlice(pFileBytes, _
                                          trackDataStartPosition, _
                                          trackDataEndPosition)
    
    Set GetTrack = Factory.CreateNewMidiTrackChunk(trackBytes)
End Function

Private Function IsValidTrackLength(ByVal dataEndPosition As Long, midiFileBytes() As Byte) As Boolean
    If dataEndPosition > UBound(midiFileBytes) Then Exit Function 'guard clause for array index
    IsValidTrackLength = (midiFileBytes(dataEndPosition) = END_OF_TRACK_EVENT_DATA And _
                          midiFileBytes(dataEndPosition - 1) = MetaEventTypeEnum.END_OF_TRACK And _
                          midiFileBytes(dataEndPosition - 2) = StatusEnum.META_EVENT)
End Function

Public Function GetTrackDimensions() As TrackDimensions
    'returns the length/start/end positions of each track in midi file
    Const TRACK_HEADER_LENGTH_LENGTH = 4
    Dim currentPosition As Long
    Dim upperBound As Long
    Dim trackStartPosition As Long
    Dim trackEndPosition As Long
    Dim trackDataStartPosition As Long
    Dim trackLen As Long
    Dim trackStartPositions As Collection
    Dim trackEndPositions As Collection
    Dim trackLengths As Collection
   
    Set trackStartPositions = New Collection
    Set trackEndPositions = New Collection
    Set trackLengths = New Collection
    upperBound = UBound(pFileBytes)
    'pos of 1st byte after file hdr, ie start pos of 1st trk
    currentPosition = Midi.FILE_HEADER_LENGTH
    
    'iterate file bytes to find all tracks and their len, pos.
    Do While currentPosition <= upperBound
        If IsTrackChunk(currentPosition, pFileBytes) Then
            trackStartPosition = currentPosition
            currentPosition = currentPosition + TRACK_HEADER_ID_LENGTH
            trackLen = JoinFourBytes(pFileBytes(currentPosition), _
                                     pFileBytes(currentPosition + 1), _
                                     pFileBytes(currentPosition + 2), _
                                     pFileBytes(currentPosition + 3))
            trackDataStartPosition = currentPosition + TRACK_HEADER_LENGTH_LENGTH
            If Not IsCorrectTrackLength(trackDataStartPosition, trackLen) Then
                trackLen = GetTrackLengthIteratively(trackDataStartPosition)
            End If
            trackStartPositions.Add trackStartPosition
            currentPosition = trackDataStartPosition + trackLen
            trackLengths.Add trackLen
            trackEndPosition = currentPosition - 1
            trackEndPositions.Add trackEndPosition
            trackLen = 0
        Else
            currentPosition = currentPosition + 1
        End If
    Loop
     
    Set GetTrackDimensions = Factory.CreateNewTrackDimensions(trackStartPositions, _
                                                              trackEndPositions, _
                                                              trackLengths)
End Function

Private Function GetTrackLengthIteratively(dataStartPosition As Long) As Long
    Const END_OF_TRACK_EVENT_LENGTH = 3
        Dim i As Long
    Dim isEndOfTrackMetaEvent As Boolean
    Dim currentPos As Long
    Dim isWithinArrayBounds As Boolean
    Dim upperBound As Long
    
    i = dataStartPosition
    currentPos = dataStartPosition
    upperBound = UBound(pFileBytes)
    isEndOfTrackMetaEvent = pFileBytes(dataStartPosition) = StatusEnum.META_EVENT And _
                            pFileBytes(dataStartPosition + 1) = MetaEventTypeEnum.END_OF_TRACK And _
                            pFileBytes(dataStartPosition + 2) = END_OF_TRACK_EVENT_DATA
    isWithinArrayBounds = upperBound >= i + END_OF_TRACK_EVENT_LENGTH
    
    Do While Not isEndOfTrackMetaEvent And isWithinArrayBounds
        i = i + 1
        isEndOfTrackMetaEvent = pFileBytes(i) = StatusEnum.META_EVENT And _
                                pFileBytes(i + 1) = MetaEventTypeEnum.END_OF_TRACK And _
                                pFileBytes(i + 2) = END_OF_TRACK_EVENT_DATA
        isWithinArrayBounds = upperBound >= i + END_OF_TRACK_EVENT_LENGTH
    Loop
    
    GetTrackLengthIteratively = i + END_OF_TRACK_EVENT_LENGTH - dataStartPosition
End Function

Private Function IsCorrectTrackLength(ByVal currentPosition As Long, _
                                      ByVal reportedTrackLength As Long) As Boolean
    'checks if there is an "end of file" meta event ending right before the reportedTrackLength
    'ie last 3 bytes are FF 2F 00
    IsCorrectTrackLength = pFileBytes(currentPosition + reportedTrackLength - 1) = END_OF_TRACK_EVENT_DATA _
                           And pFileBytes(currentPosition + reportedTrackLength - 2) = MetaEventTypeEnum.END_OF_TRACK _
                           And pFileBytes(currentPosition + reportedTrackLength - 3) = StatusEnum.META_EVENT
End Function

Private Function IsTrackChunk(ByVal currentPosition As Long, midiFileBytes() As Byte) As Boolean
    'Checks if the first 4 bytes from the current position in an array of
    'midi file bytes indicate a track chunk header.
    If currentPosition < UBound(midiFileBytes) - 3 Then 'prevent index out of bounds error
        IsTrackChunk = (midiFileBytes(currentPosition) = TrackChunkTypeEnum.BYTE_1 _
                        And midiFileBytes(currentPosition + 1) = TrackChunkTypeEnum.BYTE_2 _
                        And midiFileBytes(currentPosition + 2) = TrackChunkTypeEnum.BYTE_3 _
                        And midiFileBytes(currentPosition + 3) = TrackChunkTypeEnum.BYTE_4)
    End If
End Function

Private Function JoinFourBytes(ByVal byte1 As Byte, _
                               ByVal byte2 As Byte, _
                               ByVal byte3 As Byte, _
                               ByVal byte4 As Byte) As Long
    'byte1 is high and byte4 is low
    Dim bitShiftedByte1 As Long
    Dim bitShiftedByte2 As Long
    Dim bitShiftedByte3 As Long

    If byte1 <> 0 Then
        bitShiftedByte1 = BitUtils.ShiftBitsLeft(byte1, numBits:=24)
    Else
        bitShiftedByte1 = 0
    End If
    If byte2 <> 0 Then
        bitShiftedByte2 = BitUtils.ShiftBitsLeft(byte2, numBits:=16)
    Else
        bitShiftedByte2 = 0
    End If
    If byte3 <> 0 Then
        bitShiftedByte3 = BitUtils.ShiftBitsLeft(byte3, numBits:=8)
    Else
        bitShiftedByte3 = 0
    End If
    
    JoinFourBytes = bitShiftedByte1 Or bitShiftedByte2 Or bitShiftedByte3 Or byte4
End Function

Function GetHeaderChunk() As HeaderChunk
    'returns midi file header
    Const BIT_MASK_BYTE_1 = &H7F
    Const BIT_MASK_BYTE_2 = &HFF
    Dim midiFileFormat As Integer
    Dim trackCount As Long
    Dim timeDiv As Long
    Dim isPpqTime As Boolean
    Dim isFpsTime As Boolean
    Dim smpteFrames As Byte 'the number of SMPTE frames can be 24, 25, 29 (for 29.97 fps) or 30
    Dim ticksPerFrame As Byte
    Dim smpteFormat As Byte
    Dim timeDivType As TimeDivisionType

    isPpqTime = pFileBytes(12) <= &H7F
    midiFileFormat = pFileBytes(9)
    trackCount = JoinTwoBytes(pFileBytes(10), pFileBytes(11))
    
    If isPpqTime Then
        timeDivType = PULSES_PER_QUARTER_NOTE
        timeDiv = JoinTwoBytes(pFileBytes(12), pFileBytes(13))
    Else
        'is SMPTE frames/sec
        timeDivType = SMPTE_MIDI_TIME_CODE
        smpteFormat = DecodeTwosComplementNumber(pFileBytes(12))
        ticksPerFrame = pFileBytes(13)
    End If
    Set GetHeaderChunk = Factory.CreateNewHeaderChunk(midiFileFormat, _
                                                      trackCount, _
                                                      timeDivType, _
                                                      timeDiv)
End Function

Private Function DecodeTwosComplementNumber(ByVal twosComplementNumber As Byte) As Byte
    Const FLIP_BITS_BITMASK = 255 'binary 1111 1111
    Const SEVEN_BIT_NUMBER_MAX = 127
    Dim isPositive As Boolean
    
    isPositive = (twosComplementNumber <= SEVEN_BIT_NUMBER_MAX)
    If isPositive Then
        DecodeTwosComplementNumber = twosComplementNumber
    Else
        'complement (flip) bits
        DecodeTwosComplementNumber = (twosComplementNumber Xor FLIP_BITS_BITMASK) + 1
    End If
End Function

Private Function JoinTwoBytes(ByVal byteHigh As Byte, ByVal byteLow As Byte) As Long
    JoinTwoBytes = BitUtils.ShiftBitsLeft(byteHigh, numBits:=8) Or byteLow
End Function


