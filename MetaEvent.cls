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

Private pMetaType As Byte
Private pMetaTypeName As String
Private pCoreEvent As CoreEvent

Public Property Get Delta() As Long
    Delta = pCoreEvent.Delta
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pCoreEvent.TimeStamp
End Property

Public Property Get Status() As Byte
    Status = pCoreEvent.Status
End Property

Public Property Get MetaType() As Byte
    MetaType = pMetaType
End Property

Public Property Get MetaTypeName() As String
    MetaTypeName = pMetaTypeName
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
                      ByVal metaTypeByte As Byte, _
                      ByVal eventData As Collection)
    Const MAX_7_BIT_VALUE = 127
    Dim eventCoreLength As Long
    Dim midiStatus As Byte
    Dim lengthVLV() As Byte
    
    midiStatus = StatusEnum.META_EVENT
    pMetaType = metaTypeByte
    pMetaTypeName = ToMetaTypeName(metaTypeByte)
    If eventData.Count <= MAX_7_BIT_VALUE Then
        eventCoreLength = 3
    Else
        lengthVLV = Convert.EncodeVLV(eventData.Count)
        eventCoreLength = UBound(lengthVLV)
    End If
    Set pCoreEvent = Factory.CreateNewCoreEvent(deltaTime, _
                                                absoluteTime, _
                                                midiStatus, _
                                                eventData, _
                                                eventCoreLength)
End Sub

Public Function ToBytes() As Byte()
    ToBytes = pCoreEvent.ToBytes(Me)
End Function

Private Function ToMetaTypeName(ByVal metaTypeByte As Byte) As String
    Select Case metaTypeByte
        Case MetaEventTypeEnum.SEQUENCE_NUMBER
            ToMetaTypeName = "SequenceNumber"
        Case MetaEventTypeEnum.TEXT_EVENT
            ToMetaTypeName = "TextEvent"
        Case MetaEventTypeEnum.COPYRIGHT_NOTICE
            ToMetaTypeName = "CopyrightNotice"
        Case MetaEventTypeEnum.SEQUENCE_TRACK_NAME
            ToMetaTypeName = "SequenceTrackName"
        Case MetaEventTypeEnum.INSTRUMENT_NAME
            ToMetaTypeName = "InstrumentName"
        Case MetaEventTypeEnum.LYRICS
            ToMetaTypeName = "Lyrics"
        Case MetaEventTypeEnum.MARKER
            ToMetaTypeName = "Marker"
        Case MetaEventTypeEnum.CUE_POINT
            ToMetaTypeName = "CuePoint"
        Case MetaEventTypeEnum.MIDI_CHANNEL_PREFIX
            ToMetaTypeName = "MidiChannelPrefix"
        Case MetaEventTypeEnum.MIDI_PORT
            ToMetaTypeName = "MidiPort"
        Case MetaEventTypeEnum.END_OF_TRACK
            ToMetaTypeName = "EndOfTrack"
        Case MetaEventTypeEnum.SET_TEMPO
            ToMetaTypeName = "SetTempo"
        Case MetaEventTypeEnum.SMPTE_OFFSET
            ToMetaTypeName = "SmpteOffset"
        Case MetaEventTypeEnum.TIME_SIGNATURE
            ToMetaTypeName = "TimeSignature"
        Case MetaEventTypeEnum.KEY_SIGNATURE
            ToMetaTypeName = "KeySignature"
        Case MetaEventTypeEnum.SEQUENCER_SPECIFIC
            ToMetaTypeName = "SequencerSpecific"
        Case Else
            ToMetaTypeName = "UnknownMetaEvent"
    End Select
End Function





