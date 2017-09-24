VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetaEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDelta As Long
Private pTimeStamp As Long
Private pStatus As Byte
Private pMetaType As Byte
Private pMetaTypeName As String
Private pLength As Long
Private pData As Collection

Public Property Get Delta() As Long
    Delta = pDelta
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pTimeStamp
End Property

Public Property Get Status() As Byte
    Status = pStatus
End Property

Public Property Get MetaType() As Byte
    MetaType = pMetaType
End Property

Public Property Get MetaTypeName() As String
    MetaTypeName = pMetaTypeName
End Property

Public Property Get Length() As Long
    Length = pData.Count
End Property

Public Property Get Data() As Collection
    Set Data = pData
End Property

Public Sub Initialize(ByVal deltaTime As Long, ByVal absoluteTime As Long, ByVal metaTypeByte As Byte, ByVal eventData As Collection)
    'side effect: mutates object properties
    pDelta = deltaTime
    pTimeStamp = absoluteTime
    pStatus = &HFF
    pMetaType = metaTypeByte
    pMetaTypeName = ToMetaTypeName(metaTypeByte)
    Set pData = eventData
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
    
    'status, type
    eventString = eventString & pStatus & ","
    eventString = eventString & pMetaType & ","
   
    'data len
    vlvBytes = EncodeVLV(pLength)
    For i = 0 To UBound(vlvBytes)
        eventString = eventString & vlvBytes(i) & ","
    Next i
    
    'data in bytes
    For i = 1 To pLength
        eventString = eventString & pData(i) & ","
    Next i
    
    'meta evt
    eventString = Left(eventString, Len(eventString) - 1) 'remove last comma
    eventStrings = Split(eventString, ",")
    ReDim eventBytes(UBound(eventStrings))
    For i = 0 To UBound(eventStrings)
        eventBytes(i) = CByte(eventStrings(i))
    Next i
    
    ToBytes = eventBytes
End Function

Private Function ToMetaTypeName(ByVal metaTypeByte As Byte) As String
    'converts MetaType to name: e.g. 7 --> "CuePoint"
    Select Case metaTypeByte
        Case &H0
            ToMetaTypeName = "SequenceNumber"
        Case &H1
            ToMetaTypeName = "TextEvent"
        Case &H2
            ToMetaTypeName = "CopyrightNotice"
        Case &H3
            ToMetaTypeName = "SequenceTrackName"
        Case &H4
            ToMetaTypeName = "InstrumentName"
        Case &H5
            ToMetaTypeName = "Lyrics"
        Case &H6
            ToMetaTypeName = "Marker"
        Case &H7
            ToMetaTypeName = "CuePoint"
        Case &H20
            ToMetaTypeName = "MidiChannelPrefix"
        Case &H21
            ToMetaTypeName = "MidiPort"
        Case &H2F
            ToMetaTypeName = "EndOfTrack"
        Case &H51
            ToMetaTypeName = "SetTempo"
        Case &H54
            ToMetaTypeName = "SmpteOffset"
        Case &H58
            ToMetaTypeName = "TimeSignature"
        Case &H59
            ToMetaTypeName = "KeySignature"
        Case &H7F
            ToMetaTypeName = "SequencerSpecific"
        Case Else
            ToMetaTypeName = "UnknownMetaEvent"
    End Select
End Function

