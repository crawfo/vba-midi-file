VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDelta As Long
Private pTimeStamp As Long
Private pStatus As Byte
Private pDataLength As Long
Private pData As Collection
Private pEventLength As Long
Private pEventBytes As Collection

Private Sub Class_Initialize()
    Set pData = New Collection
    Set pEventBytes = New Collection
End Sub

Private Sub Class_Terminate()
    Set pData = Nothing
    Set pEventBytes = Nothing
End Sub

Public Property Get Delta() As Long
    Delta = pDelta
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pTimeStamp
End Property

Public Property Get Status() As Byte
    Status = pStatus
End Property

Public Property Get DataLength() As Long
    DataLength = pDataLength
End Property

Public Property Get Data() As Collection
    Set Data = pData
End Property

Public Property Get EventLength() As Long
    EventLength = pEventLength
End Property

Public Property Get TrackEventBytes() As Collection
    Set TrackEventBytes = pEventBytes
End Property

Public Sub Initialize(ByVal deltaTime As Long, _
                      ByVal absoluteTime As Long, _
                      ByVal statusByte As Byte, _
                      ByVal evtData As Collection, _
                      ByVal eventCoreLength As Long)
    InitializeCoreProperties deltaTime, absoluteTime, statusByte, evtData
    pEventLength = eventCoreLength + pDataLength
End Sub

Private Sub InitializeCoreProperties(ByVal deltaTime As Long, _
                                     ByVal absoluteTime As Long, _
                                     ByVal statusByte As Byte, _
                                     ByVal evtData As Collection)
    pDelta = deltaTime
    pTimeStamp = absoluteTime
    pStatus = statusByte
    Set pData = evtData
    pDataLength = pData.Count
End Sub

Public Function ToBytes(ByVal trkEvent As Object) As Byte()
    Dim eventBytes() As Byte
    Dim deltaBytes() As Byte
    Dim vlvBytes() As Byte '*
    Dim i As Long
    Dim eventArrayUpperBound As Long
    Dim deltaArrayUpperBound As Long
    Dim vlvArrayUpperBound As Long '*
    Dim evtArrayPosition As Long '*
    Dim dataByte As Variant
    Dim isMetaEvt As Boolean
    Dim isSyxEvt As Boolean
    Dim isChannelEvt As Boolean
    Dim isTwoByteEvt As Boolean

    'add delta bytes
    deltaBytes = Convert.EncodeVLV(trkEvent.Delta)
    deltaArrayUpperBound = UBound(deltaBytes)
    eventArrayUpperBound = deltaArrayUpperBound + trkEvent.EventLength
    ReDim eventBytes(eventArrayUpperBound)
    For i = 0 To deltaArrayUpperBound
        eventBytes(i) = deltaBytes(i)
    Next i
    evtArrayPosition = deltaArrayUpperBound + 1
    
    'add event bytes
    isMetaEvt = (pStatus = StatusEnum.META_EVENT)
    isSyxEvt = (pStatus = StatusEnum.SYSTEM_EXCLUSIVE_START Or _
                pStatus = StatusEnum.SYSTEM_EXCLUSIVE_CONTINUE)
    isChannelEvt = Not (isMetaEvt Or isSyxEvt)
        
    If isChannelEvt Then
        isTwoByteEvt = (trkEvent.ChannelEventType = ceTwoByte) '**'
        If trkEvent.IsRunningStatus Then
            'data
            eventBytes(deltaArrayUpperBound + 1) = trkEvent.Data1
            If Not isTwoByteEvt Then
                eventBytes(deltaArrayUpperBound + 2) = trkEvent.Data2
            End If
        Else
            'status
            eventBytes(deltaArrayUpperBound + 1) = trkEvent.JoinTwoNibbles(pStatus, trkEvent.Channel)
            'data
            eventBytes(deltaArrayUpperBound + 2) = trkEvent.Data1
            If Not isTwoByteEvt Then
                eventBytes(deltaArrayUpperBound + 3) = trkEvent.Data2
            End If
        End If
    Else
        'status
        eventBytes(evtArrayPosition) = trkEvent.Status
        evtArrayPosition = evtArrayPosition + 1
        'meta type
        If isMetaEvt Then
            eventBytes(evtArrayPosition) = trkEvent.MetaType
            evtArrayPosition = evtArrayPosition + 1
        End If
        'data len
        vlvBytes = Convert.EncodeVLV(trkEvent.DataLength)
        vlvArrayUpperBound = UBound(vlvBytes)
        For i = 0 To vlvArrayUpperBound
            eventBytes(evtArrayPosition) = vlvBytes(i)
            evtArrayPosition = evtArrayPosition + 1
        Next i
        'data
        For Each dataByte In trkEvent.Data
            eventBytes(evtArrayPosition) = dataByte
            evtArrayPosition = evtArrayPosition + 1
        Next dataByte
    End If
    
    ToBytes = eventBytes
End Function

Private Function ToBytesEventOnly(ByVal trkEvent As Object) As Collection
    'no delta included
    Const DATA_BYTE_1_INDEX = 1
    Const DATA_BYTE_2_INDEX = 2
    Dim eventBytes As Collection
    Dim vlvBytes() As Byte
    Dim vlvArrayUpperBound As Long
    Dim i As Long
    Dim dataByte As Variant
    Dim isMetaEvt As Boolean
    Dim isSyxEvt As Boolean
    Dim isChannelEvt As Boolean
    Dim isTwoByteEvt As Boolean
    
    Set eventBytes = New Collection
    isMetaEvt = (pStatus = StatusEnum.META_EVENT)
    isSyxEvt = (pStatus = StatusEnum.SYSTEM_EXCLUSIVE_START Or pStatus = StatusEnum.SYSTEM_EXCLUSIVE_CONTINUE)
    isChannelEvt = Not (isMetaEvt Or isSyxEvt)
    
    'add event bytes
    If isChannelEvt Then
        isTwoByteEvt = (trkEvent.ChannelEventType = ceTwoByte) '**'
        If trkEvent.IsRunningStatus Then
            'data
            eventBytes.Add pData(DATA_BYTE_1_INDEX)
            If Not isTwoByteEvt Then
                eventBytes.Add pData(DATA_BYTE_2_INDEX)
            End If
        Else
            'status
            eventBytes.Add trkEvent.JoinTwoNibbles(pStatus, trkEvent.Channel)
            'data
            eventBytes.Add pData(DATA_BYTE_1_INDEX)
            If Not isTwoByteEvt Then
                eventBytes.Add pData(DATA_BYTE_2_INDEX)
            End If
        End If
    Else
        'status
        eventBytes.Add pStatus
        'meta type
        If isMetaEvt Then
            eventBytes.Add trkEvent.MetaType
        End If
        'data len
        vlvBytes = Convert.EncodeVLV(pDataLength)
        vlvArrayUpperBound = UBound(vlvBytes)
        For i = 0 To vlvArrayUpperBound
            eventBytes.Add vlvBytes(i)
        Next i
        'data
        For Each dataByte In pData
            eventBytes.Add dataByte
        Next dataByte
    End If
    
    Set ToBytesEventOnly = eventBytes
End Function


