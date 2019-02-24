VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeaderChunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TimeDivisionType
    PULSES_PER_QUARTER_NOTE
    SMPTE_MIDI_TIME_CODE
End Enum

Private pChunkBytes() As Byte
Private pFileFormat As Integer
Private pNumberTracks As Long
Private pTimeDivisionType As TimeDivisionType
Private pTimeDivision As Long

Public Property Get ChunkBytes() As Byte()
    ChunkBytes = pChunkBytes
End Property

Public Property Get FileFormat() As Integer
    FileFormat = pFileFormat
End Property

Public Property Get NumberTracks() As Long
    NumberTracks = pNumberTracks
End Property

Public Property Let NumberTracks(numTracks As Long)
    pNumberTracks = numTracks
    're-initialize obj on change
    Initialize FileFormat, numTracks, TimeDivision
End Property

Public Property Get TimeDivisionType() As Long
    TimeDivisionType = pTimeDivisionType
End Property

Public Property Get TimeDivision() As Long
    TimeDivision = pTimeDivision
End Property

Public Function GetChunkByte(ByVal i As Long) As Byte
    'accessor for array element at index i
    GetChunkByte = pChunkBytes(i)
End Function

Public Sub Initialize(ByVal midiFileFormat As Integer, _
                      ByVal numTracks As Long, _
                      ByVal timeDivType As TimeDivisionType, _
                      ByVal timeDiv As Long)
    Const HEADER_CHUNK_LENGTH = 14
    Const MAX_TIME_DIVISION = 32768 '&H8000
    Dim numTracksBytes() As Byte
    Dim timeDivisionBytes() As Byte
    Dim isTicksPerQuarterNote As Boolean
    Dim smpteFormat As Byte
    Dim frameResolution As Byte
    
    
    pFileFormat = midiFileFormat
    pNumberTracks = numTracks
    pTimeDivisionType = timeDivType
    pTimeDivision = timeDiv
    ReDim pChunkBytes(HEADER_CHUNK_LENGTH - 1)
    
    'midi file id ('MThd')
    pChunkBytes(0) = HeaderChunkEnum.CHUNK_TYPE_BYTE_1 ' &H4D
    pChunkBytes(1) = HeaderChunkEnum.CHUNK_TYPE_BYTE_2 ' &H54
    pChunkBytes(2) = HeaderChunkEnum.CHUNK_TYPE_BYTE_3 ' &H68
    pChunkBytes(3) = HeaderChunkEnum.CHUNK_TYPE_BYTE_4 ' &H64
    'chunk size
    pChunkBytes(4) = HeaderChunkEnum.CHUNK_LENGTH_BYTE_1 ' &H0
    pChunkBytes(5) = HeaderChunkEnum.CHUNK_LENGTH_BYTE_2 ' &H0
    pChunkBytes(6) = HeaderChunkEnum.CHUNK_LENGTH_BYTE_3 ' &H0
    pChunkBytes(7) = HeaderChunkEnum.CHUNK_LENGTH_BYTE_4 ' &H6
    'format type
    pChunkBytes(8) = &H0
    pChunkBytes(9) = pFileFormat
    'num tracks 1-65,535
    numTracksBytes = ToTwoBytesFromLong(pNumberTracks)
    pChunkBytes(10) = numTracksBytes(0)
    pChunkBytes(11) = numTracksBytes(1)
    'time division
    timeDivisionBytes = ToTwoBytesFromLong(pTimeDivision)
    isTicksPerQuarterNote = timeDiv < MAX_TIME_DIVISION '32768 &H8000
    If isTicksPerQuarterNote Then
        pChunkBytes(12) = timeDivisionBytes(0)
        pChunkBytes(13) = timeDivisionBytes(1)
    Else
        smpteFormat = timeDivisionBytes(0)
        frameResolution = timeDivisionBytes(1)
        pChunkBytes(12) = smpteFormat
        pChunkBytes(13) = frameResolution
    End If
End Sub

Private Function ToTwoBytesFromLong(ByVal numTracks As Long) As Byte()
    'numTracks must be <= &HFFFF (65,535) TODO:guard clause?
    Dim bytes(1) As Byte
    
    bytes(0) = BitUtils.ShiftBitsRight(numTracks And 65280, numBits:=8) '1111111100000000 = 65280: Note that &HFF00 evaluates to -256 instead of 65280
    bytes(1) = numTracks And &HFF '0000000011111111 = &HFF = 255
    
    ToTwoBytesFromLong = bytes
End Function


