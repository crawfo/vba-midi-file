Attribute VB_Name = "Convert"
Option Explicit

Public Function ToFourBytesFromLong(ByVal trackDataLength As Long) As Byte()
    Dim bytes(3) As Byte
    bytes(0) = BitUtils.ShiftBitsRight(trackDataLength And &HFF000000, numBits:=24) '11111111000000000000000000000000 = &HFF000000
    bytes(1) = BitUtils.ShiftBitsRight(trackDataLength And &HFF0000, numBits:=16)   '00000000111111110000000000000000 = &HFF0000
    bytes(2) = BitUtils.ShiftBitsRight(trackDataLength And 65280, numBits:=8)       '00000000000000001111111100000000 = 65280 NB &HFF00 evaluates to -256 instead of 65280
    bytes(3) = trackDataLength And &HFF                                             '00000000000000000000000011111111
    ToFourBytesFromLong = bytes
End Function

Public Function ToLongFromFourBytes(fourBytes() As Byte) As Long
    ToLongFromFourBytes = BitUtils.ShiftBitsLeft(fourBytes(0), 24) And _
                          BitUtils.ShiftBitsLeft(fourBytes(1), 16) And _
                          BitUtils.ShiftBitsLeft(fourBytes(2), 8) And fourBytes(3)
End Function

'encodes long into collection of variable length value bytes
'e.g. 32768 into array containing bytes 130,128,00
'max value is &H0FFFFFFF (268435455)
Public Function EncodeVLV(ByVal value As Long) As Byte()
    Dim valueBitShiftedRight
    Dim vlvBytes() As Byte
    Dim i As Long
    Dim numBits As Integer
        
    Const BIT_MASK_PRESERVE_BITS_0_TO_7 = &H7F ' 0 1111111
    Const BIT_MASK_SET_BIT_8_PRESERVE_BITS_0_TO_7 = &H80 '1 0000000
    
    If value > &HFFFFFFF Or value < 0 Then 'TODO:MAGIC NUM
        ReDim vlvBytes(0)
        EncodeVLV = vlvBytes
        Exit Function
    End If
    
    ReDim vlvBytes(3)
    'max value is 0FFFFFFF so max 28 bits (4 bytes) needed
    'split into bytes with msb (bit7) set correctly:
    'shift bits right by number of bits to position lower 7 bits
    'And &H7F preserves only lower 7 bits
    'Or &H80 sets msb and preserves lower bits
    i = 0
    'starts with highest byte (MSB)(left)
    For numBits = 21 To 7 Step -7
        valueBitShiftedRight = BitUtils.ShiftBitsRight(value, numBits)
        If valueBitShiftedRight = 0 Then
            ReDim Preserve vlvBytes(UBound(vlvBytes) - 1)
        Else
            vlvBytes(i) = valueBitShiftedRight And BIT_MASK_PRESERVE_BITS_0_TO_7 Or _
                          BIT_MASK_SET_BIT_8_PRESERVE_BITS_0_TO_7
            i = i + 1
        End If
    Next numBits
    'last byte so must have msb clear not set
    vlvBytes(i) = BitUtils.ShiftBitsRight(value, numBits) And &H7F
    EncodeVLV = vlvBytes
End Function

Public Function DecodeVLV(ByVal bytes As Collection) As Long
    'TODO: magic nums
    'decodes a collection of numeric vlv bytes to long value
    Const BIT_MASK_MSB = &HFF
    Const BIT_MASK_MSB_CLEARED = &H80
    Dim decodedVlv As Long
    Dim byteCount As Long
    Dim currentByte As Byte
    Dim i As Long
    
    byteCount = bytes.Count
    For i = 1 To byteCount
        'apply msb bit mask
        currentByte = bytes(i) And BIT_MASK_MSB '&HFF
        'shifts bits 7 places to left,     ,masks first 7 bits
        decodedVlv = BitUtils.ShiftBitsLeft(decodedVlv, numBits:=7) Or (currentByte And &H7F)
        'if last value (ie msb cleared)
        If (currentByte And BIT_MASK_MSB_CLEARED) = 0 Then
            Exit For
        End If
    Next i
    DecodeVLV = decodedVlv
End Function

Public Function GetVLVBytes(bytes() As Byte, ByVal startPosition As Long) As Collection
    'returns vlv byte collection from an array
    Dim i As Long
    Dim upperBound As Long
    Dim vlvBytes As Collection

    Set vlvBytes = New Collection
    i = startPosition
    upperBound = UBound(bytes)

    'collect vlv bytes
    Do While HasMsbSet(bytes(i)) And i <= upperBound
    'Do While bytes(i) > &H7F And i <= upperBound 'slightly faster
        vlvBytes.Add bytes(i)
        i = i + 1
    Loop

    'add last value - no need to check msb as if last value then msb must be 0
    vlvBytes.Add bytes(i)

    Set GetVLVBytes = vlvBytes
End Function

Public Function GetVLVBytesTrackChunkVersion(ByVal trkChunk As TrackChunk, _
                                             ByVal startPosition As Long) As Collection
    'returns vlv byte collection from an array
    Dim i As Long
    Dim upperBound As Long
    Dim vlvBytes As Collection

    Set vlvBytes = New Collection
    i = startPosition
    upperBound = UBound(trkChunk)

    'collect vlv bytes
    Do While HasMsbSet(trkChunk(i)) And i <= upperBound
    'Do While bytes(i) > &H7F And i <= upperBound 'slightly faster
        vlvBytes.Add trkChunk(i)
        i = i + 1
    Loop

    'add last value - no need to check msb as if last value then msb must be 0
    vlvBytes.Add trkChunk(i)

    Set GetVLVBytes = vlvBytes
End Function

Private Function HasMsbSet(ByVal byteToCheck As Byte) As Boolean
    'returns true when msb is set, otherwise false
    Const SEVEN_BIT_VALUE_MAX = &H7F
    HasMsbSet = byteToCheck > SEVEN_BIT_VALUE_MAX
End Function

Public Function ToArrayFromCollection(byteCollection As Collection) As Byte()
    Dim bytesArray() As Byte
    Dim collectionByte As Variant
    Dim i As Long
    Dim upperBound As Long

    upperBound = byteCollection.Count - 1
    ReDim bytesArray(upperBound)
    i = 0

    For Each collectionByte In byteCollection
        bytesArray(i) = collectionByte
        i = i + 1
    Next collectionByte
    
    ToArrayFromCollection = bytesArray
End Function
