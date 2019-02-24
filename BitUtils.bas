Attribute VB_Name = "BitUtils"
Option Explicit

Public Function GetNibbleHigh(ByVal byteToNibblize As Byte) As Byte
    GetNibbleHigh = ShiftBitsRight(byteToNibblize, numBits:=4) And &HF
End Function

Public Function GetNibbleLow(ByVal byteToNibblize As Byte) As Byte
    Const BIT_MASK_LOW_NIBBLE = &HF
    GetNibbleLow = byteToNibblize And BIT_MASK_LOW_NIBBLE
End Function

Public Function ShiftBitsLeft(ByVal value As Long, ByVal numBits As Long) As Long
    ShiftBitsLeft = value * (2 ^ numBits)
End Function

Public Function ShiftBitsRight(ByVal value As Long, ByVal numBits As Long) As Long
    ShiftBitsRight = value \ (2 ^ numBits)
End Function
