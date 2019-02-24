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

Private pEvtType As SystemExclusiveType
Private pCoreEvent As CoreEvent

Public Property Get Delta() As Long
    Delta = pCoreEvent.Delta
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pCoreEvent.TimeStamp
End Property

Public Property Get EvtType() As SystemExclusiveType
    EvtType = pEvtType
End Property

Public Property Get Status() As Byte
    Status = pCoreEvent.Status
End Property

Public Property Get DataLength() As Long
    DataLength = pCoreEvent.DataLength
End Property

Public Property Get Data() As Collection
    Set Data = pCoreEvent.Data
End Property

Public Property Get EventLength() As Long
    EventLength = pCoreEvent.EventLength
End Property

Public Property Get TrackEventBytes() As Collection
    Set TrackEventBytes = pCoreEvent.TrackEventBytes
End Property

Public Sub Initialize(ByVal deltaTime As Long, _
                      ByVal absoluteTime As Long, _
                      ByVal midiStatus As Byte, _
                      ByVal eventData As Collection, _
                      ByVal syxType As SystemExclusiveType)
    Const MAX_7_BIT_VALUE = 127
    Dim eventCoreLength As Long
    Dim lengthVLV() As Byte
    
    pEvtType = syxType
    If eventData.Count <= MAX_7_BIT_VALUE Then
        eventCoreLength = 2
    Else
        lengthVLV = Convert.EncodeVLV(eventData.Count)
        eventCoreLength = UBound(lengthVLV)
    End If
    Set pCoreEvent = CreateNewCoreEvent(deltaTime, _
                                        absoluteTime, _
                                        midiStatus, _
                                        eventData, _
                                        eventCoreLength)
End Sub

Public Function ToBytes() As Byte()
    ToBytes = pCoreEvent.ToBytes(Me)
End Function

