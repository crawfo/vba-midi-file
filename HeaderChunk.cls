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

Private pChunkBytes() As Byte
Private pFileFormat As Integer
Private pNumberTracks As Long
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

Public Property Get TimeDivision() As Long
    TimeDivision = pTimeDivision
End Property

Public Sub Initialize(ByVal midiFileFormat As Integer, ByVal numTracks As Long, ByVal timeDiv As Long)
    'side effect: mutates object properties
    Dim abytNumTracks() As Byte, timeDivisionBytes() As Byte, isTicksPerQuarterNote As Boolean
    
    pFileFormat = midiFileFormat
    pNumberTracks = numTracks
    pTimeDivision = timeDiv
    ReDim pChunkBytes(13)
    
    'midi file id ('MThd')
    pChunkBytes(0) = &H4D
    pChunkBytes(1) = &H54
    pChunkBytes(2) = &H68
    pChunkBytes(3) = &H64

    'chunk size
    pChunkBytes(4) = &H0
    pChunkBytes(5) = &H0
    pChunkBytes(6) = &H0
    pChunkBytes(7) = &H6

    'format type
    pChunkBytes(8) = &H0
    pChunkBytes(9) = pFileFormat

    'num tracks 1-65,535
    abytNumTracks = ToBytesFromLong(pNumberTracks)
    pChunkBytes(10) = abytNumTracks(0)
    pChunkBytes(11) = abytNumTracks(1)

    'time division
    timeDivisionBytes = ToBytesFromLong(pTimeDivision)
    isTicksPerQuarterNote = timeDiv < 32768 '&H8000
    If isTicksPerQuarterNote Then
        pChunkBytes(12) = timeDivisionBytes(0)
        pChunkBytes(13) = timeDivisionBytes(1)
    Else
        'is SMPTE frames/sec
    End If
End Sub

Public Function GetChunkByte(ByVal i As Long) As Byte
    'accessor for array element at index i
    GetChunkByte = pChunkBytes(i)
End Function

Private Function ToBytesFromLong(ByVal numTracks As Long) As Byte()
    'numTracks must be <= &HFFFF (65,535)
    Dim numTracksHex As String, numHexDigits As Integer, bytes(1) As Byte
    Dim isDoubleByte As Boolean, isSingleByte As Boolean
    numTracksHex = Hex(numTracks)
    numHexDigits = Len(numTracksHex)
    isSingleByte = numHexDigits < 3
    
    If isSingleByte Then
        bytes(0) = 0
        bytes(1) = numTracks
    Else
        If numHexDigits = 3 Then
            bytes(0) = ToByteFromHex(Left(numTracksHex, 1))
        Else
            bytes(0) = ToByteFromHex(Left(numTracksHex, 2))
        End If
        bytes(1) = ToByteFromHex(Right(numTracksHex, 2))
    End If
    ToBytesFromLong = bytes
End Function
