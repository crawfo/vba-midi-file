VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackChunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pChunkBytes() As Byte
Private pLength As Long
Private pTrackEvents As Collection 'of MetaEvent, ChanbelEvent, SystemExclusiveEvent

Public Property Get ChunkBytes() As Byte()
    ChunkBytes = pChunkBytes
End Property

Public Property Get Length() As Long
    Length = pLength
End Property

Public Property Get TrackEvents() As Collection
    Set TrackEvents = pTrackEvents
End Property

Public Sub Initialize(ByVal events As Collection)
    'side effect: mutates object properties
    Dim trackLengths() As Byte, e As Variant, b As Variant
    Dim trackBytes As Collection, bytes As Variant, i As Long
    
    Set pTrackEvents = events
    Set trackBytes = New Collection

    For Each e In pTrackEvents
        bytes = e.ToBytes() 'all midi events have a ToBytes() method
        For Each b In bytes
            trackBytes.Add b
         Next b
    Next e
    
    'add end of track meta event bytes (if it doesn't exist)
    If pTrackEvents(pTrackEvents.Count).Status <> msMetaEvent Then
        'add delta
        trackBytes.Add 0
        'add event
        trackBytes.Add 255
        trackBytes.Add 47
        trackBytes.Add 0
    ElseIf pTrackEvents(pTrackEvents.Count).Status = msMetaEvent And _
       pTrackEvents(pTrackEvents.Count).MetaType <> meEndOfTrack Then
        'add delta
        trackBytes.Add 0
        'add event
        trackBytes.Add 255
        trackBytes.Add 47
        trackBytes.Add 0
    End If
    

    pLength = trackBytes.Count
    ReDim pChunkBytes(7 + Length)
    
    'track chunk id ('MTrk')
    pChunkBytes(0) = &H4D
    pChunkBytes(1) = &H54
    pChunkBytes(2) = &H72
    pChunkBytes(3) = &H6B
    
    'chunk size
    trackLengths = ToBytesFromLong(Length)
    pChunkBytes(4) = trackLengths(0)
    pChunkBytes(5) = trackLengths(1)
    pChunkBytes(6) = trackLengths(2)
    pChunkBytes(7) = trackLengths(3)
    
    'event bytes
    i = 8
    For Each e In trackBytes
        pChunkBytes(i) = e
        i = i + 1
    Next e
End Sub

Public Function GetChunkByte(ByVal i As Long) As Byte
    'accessor for array element at index i
    GetChunkByte = pChunkBytes(i)
End Function

Private Function ToBytesFromLong(ByVal trackLength As Long) As Byte()
    'trackLength must be <= &H0FFFFFFF (268435455)
    Dim numTracksHex As String, numberHexDigits As Integer, bytes(3) As Byte
    Dim isDoubleByte As Boolean, isSingleByte As Boolean, isTripleByte As Boolean
    numTracksHex = Hex(trackLength)
    numberHexDigits = Len(numTracksHex)
    isSingleByte = numberHexDigits < 3
    isDoubleByte = numberHexDigits < 5
    isTripleByte = numberHexDigits < 7
    
    If isSingleByte Then
        bytes(0) = 0
        bytes(1) = 0
        bytes(2) = 0
        bytes(3) = trackLength
    ElseIf isDoubleByte Then
        'bytes: FF FF
        'pos:   12 34
        bytes(0) = 0
        bytes(1) = 0
        If numberHexDigits = 3 Then
            '0F FF
            ' 1 23
            bytes(2) = ToByteFromHex(Left(numTracksHex, 1))
        Else
            'FF FF
            '12 34
            bytes(2) = ToByteFromHex(Left(numTracksHex, 2))
        End If
        bytes(3) = ToByteFromHex(Right(numTracksHex, 2))
    ElseIf isTripleByte Then
        '00 FF FF FF
        '12 34 56
        bytes(0) = 0
        If numberHexDigits = 5 Then
            'F FF FF
            '1 23 45
            bytes(1) = ToByteFromHex(Left(numTracksHex, 1))
            bytes(2) = ToByteFromHex(Mid(numTracksHex, 2, 2))
        Else
            'FF FF FF
            '12 34 56
            bytes(1) = ToByteFromHex(Left(numTracksHex, 2))
            bytes(2) = ToByteFromHex(Mid(numTracksHex, 3, 2))
        End If
        
        bytes(3) = ToByteFromHex(Right(numTracksHex, 2))
    Else
        'isQuadByte
        '0F FF FF FF (max midi value)
        ' 1 23 45 67
        If numberHexDigits = 7 Then
            'F FF FF FF
            '1 23 45 67
            bytes(0) = ToByteFromHex(Left(numTracksHex, 1))
            bytes(1) = ToByteFromHex(Mid(numTracksHex, 2, 2))
            bytes(2) = ToByteFromHex(Mid(numTracksHex, 4, 2))
            bytes(3) = ToByteFromHex(Right(numTracksHex, 2))
        End If
    End If
    
    ToBytesFromLong = bytes
End Function



