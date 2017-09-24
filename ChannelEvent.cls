VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChannelEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'uses Midi.JoinTwoNibbles
Public Enum ChannelEventType
    ceTwoByte
    ceThreeByte
End Enum

Private pDelta As Long
Private pTimeStamp As Long
Private pStatus As Byte
Private pStatusName As String
Private pChannel As Byte
Private pData1 As Byte
Private pData2 As Byte
Private pChannelEventType As ChannelEventType

Public Property Get Delta() As Long
    Delta = pDelta
End Property


Public Property Get TimeStamp() As Long
    TimeStamp = pTimeStamp
End Property

Public Property Get Status() As Byte
    Status = pStatus
End Property


Public Property Get StatusName() As String
    StatusName = pStatusName
End Property


Public Property Get Channel() As Byte
    Channel = pChannel
End Property

Public Property Get Data1() As Byte
    Data1 = pData1
End Property

Public Property Get Data2() As Byte
    'must only be accessed when ChannelEventType = ceThreeByte
    If ChannelEventType = ceThreeByte Then
        Data2 = pData2
    Else
        'inaccessible property when ChannelEventType = ceTwoByte
        Stop
    End If
End Property

Public Property Get ChannelEventType() As ChannelEventType
    ChannelEventType = pChannelEventType
End Property

Public Property Let ChannelEventType(enmChEventType As ChannelEventType)
    pChannelEventType = enmChEventType
End Property

Public Sub Initialize(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiStatus As Byte, ByVal Channel As Byte, ByVal dataByte1 As Byte, Optional ByVal dataByte2 As Variant)
    'side effect: mutates object properties
    pDelta = deltaTime
    pTimeStamp = absoluteTime
    pStatus = midiStatus
    pStatusName = ToStatusName(midiStatus)
    pChannel = Channel
    If IsValidData(dataByte1) Then
        pData1 = dataByte1
    Else
        Stop
    End If
    If IsMissing(dataByte2) Then
        ChannelEventType = ceTwoByte
    Else
        pChannelEventType = ceThreeByte
        If IsValidData(dataByte2) Then
            pData2 = dataByte2
        Else
            Stop
        End If
    End If
End Sub

Public Function ToBytes() As Byte()
    Dim eventBytes() As Byte, deltaBytes() As Byte, i As Long, eventUpperBound As Long
    Dim deltaUpperBound As Long, eventLength As Integer, isTwoByteEvt As Boolean
    
    isTwoByteEvt = ChannelEventType = ceTwoByte
    If isTwoByteEvt Then
        eventLength = 2
    Else
        eventLength = 3
    End If
    
    deltaBytes = EncodeVLV(pDelta)
    deltaUpperBound = UBound(deltaBytes)
    eventUpperBound = deltaUpperBound + eventLength
    ReDim eventBytes(eventUpperBound)
    For i = 0 To deltaUpperBound
        eventBytes(i) = deltaBytes(i)
    Next i
        
    eventBytes(deltaUpperBound + 1) = JoinTwoNibbles(pStatus, pChannel)
    eventBytes(deltaUpperBound + 2) = Data1
    If isTwoByteEvt = False Then
        eventBytes(deltaUpperBound + 3) = Data2
    End If

    ToBytes = eventBytes
End Function

Private Function IsValidData(ByVal sevenBitValue As Byte) As Boolean
    IsValidData = sevenBitValue >= 0 And sevenBitValue <= 127
End Function

Private Function ToStatusName(midiStatus As Byte) As String
    'converts status num to status name: e.g. 8 --> "NoteOff"
    Select Case midiStatus
        Case &H8
            ToStatusName = "NoteOff"
        Case &H9
            ToStatusName = "NoteOn"
        Case &HA
            ToStatusName = "NoteAftertouch"
        Case &HB
            ToStatusName = "Controller"
        Case &HC
            ToStatusName = "ProgramChange"
        Case &HD
            ToStatusName = "ChannelAftertouch"
        Case &HE
            ToStatusName = "PitchBend"
        Case Else
            ToStatusName = "Invalid status byte"
            Stop
    End Select
End Function

