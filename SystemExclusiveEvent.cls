VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemExclusiveEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDelta As Long
Private pTimeStamp As Long
Private pEvtType As SystemExclusiveType
Private pStatus As Byte
Private pLength As Long
Private pData As Collection

Public Property Get Delta() As Long
    Delta = pDelta
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pTimeStamp
End Property

Public Property Get EvtType() As SystemExclusiveType
    EvtType = pEvtType
End Property

Public Property Get Status() As Byte
    Status = pStatus
End Property

Public Property Get Length() As Long
    Length = pData.Count
End Property

Public Property Get Data() As Collection
    Set Data = pData
End Property

Public Sub Initialize(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal midiStatus As Byte, ByVal eventData As Collection, ByVal syxType As SystemExclusiveType)
    'side effect: mutates object properties
    pDelta = deltaTime
    pTimeStamp = absoluteTime
    pStatus = midiStatus
    Set pData = eventData
    pEvtType = syxType
    pLength = pData.Count
End Sub

Public Function ToBytes() As Byte()
    Dim eventString As String, i As Long, vlvBytes() As Byte, deltaBytes() As Byte
    Dim eventStrings() As String, eventBytes() As Byte
    
    'delta
    deltaBytes = EncodeVLV(pDelta)
    For i = 0 To UBound(deltaBytes)
        eventString = eventString & deltaBytes(i) & ","
    Next i
    
    'status
    eventString = eventString & pStatus & ","
       
    'data len
    vlvBytes = EncodeVLV(pLength)
    For i = 0 To UBound(vlvBytes)
        eventString = eventString & vlvBytes(i) & ","
    Next i
    
    'data in bytes
    For i = 1 To pLength
        eventString = eventString & pData(i) & ","
    Next i
    
    'syx evt
    eventString = Left(eventString, Len(eventString) - 1) 'remove last comma
    eventStrings = Split(eventString, ",")
    ReDim eventBytes(UBound(eventStrings))
    For i = 0 To UBound(eventStrings)
        eventBytes(i) = CByte(eventStrings(i))
    Next i
    
    ToBytes = eventBytes
End Function


