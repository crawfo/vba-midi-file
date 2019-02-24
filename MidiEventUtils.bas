Attribute VB_Name = "MidiEventUtils"
Option Explicit

Public Function IsChannelEvent(ByVal statusByte As Byte) As Boolean
    Const CHANNEL_EVENT_MIN = &H80
    Const CHANNEL_EVENT_MAX = &HEF
    IsChannelEvent = (statusByte >= CHANNEL_EVENT_MIN And statusByte <= CHANNEL_EVENT_MAX)
End Function

Public Function IsTwoByteChannelEvent(ByVal statusByte As Byte) As Boolean
    Const TWO_BYTE_CHANNEL_EVENT_MIN = &HC0
    Const TWO_BYTE_CHANNEL_EVENT_MAX = &HDF
    IsTwoByteChannelEvent = (statusByte >= TWO_BYTE_CHANNEL_EVENT_MIN And _
                             statusByte <= TWO_BYTE_CHANNEL_EVENT_MAX)
End Function

Public Function IsThreeByteChannelEvent(ByVal statusByte As Byte) As Boolean
    IsThreeByteChannelEvent = IsChannelEvent(statusByte) And _
                              Not IsTwoByteChannelEvent(statusByte)
End Function

Public Function IsMetaEvent(ByVal statusByte As Byte) As Boolean
    IsMetaEvent = (statusByte = StatusEnum.META_EVENT)
End Function

Public Function IsSysExEvent(ByVal statusByte As Byte) As Boolean
    IsSysExEvent = (statusByte = SYSTEM_EXCLUSIVE_START Or _
                    statusByte = SYSTEM_EXCLUSIVE_CONTINUE)
End Function

Public Function IsRunningStatus(ByVal statusByte As Byte) As Boolean
    Const STATUS_MIN = &H80
    IsRunningStatus = statusByte < STATUS_MIN
End Function


