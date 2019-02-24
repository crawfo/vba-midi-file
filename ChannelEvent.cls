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

Public Enum ChannelEventType
    ceTwoByte
    ceThreeByte
End Enum

Private pStatusName As String
Private pChannel As Byte
Private pData1 As Byte
Private pData2 As Byte
Private pChannelEventType As ChannelEventType
Private pIsRunningStatus As Boolean
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

Public Property Get DataLength() As Long
    DataLength = pCoreEvent.DataLength
End Property

Public Property Get Data() As Collection
    Set Data = pCoreEvent.Data
End Property

Public Property Get Data1() As Byte
    Data1 = pCoreEvent.Data(1)
End Property

Public Property Get Data2() As Byte
    If ChannelEventType = ceThreeByte Then
        Data2 = pCoreEvent.Data(2)
    Else
        Stop
    End If
End Property

Public Property Get EventLength() As Long
    EventLength = pCoreEvent.EventLength
End Property

Public Property Get TrackEventBytes() As Collection
    Set TrackEventBytes = pCoreEvent.TrackEventBytes
End Property

Public Property Get StatusName() As String
    StatusName = pStatusName
End Property

Public Property Get Channel() As Byte
    Channel = pChannel
End Property

Public Property Get ChannelEventType() As ChannelEventType
    ChannelEventType = pChannelEventType
End Property

Public Property Get IsRunningStatus() As Boolean
    IsRunningStatus = pIsRunningStatus
End Property

Public Sub Initialize(ByVal isRunStatus As Boolean, _
                      ByVal deltaTime As Long, _
                      ByVal absoluteTime As Long, _
                      ByVal statusNibble As Byte, _
                      ByVal channelNibble As Byte, _
                      ByVal dataByte1 As Byte, _
                      Optional ByVal dataByte2 As Variant)
    Dim eventCoreLength As Long
    Dim evtData As Collection
    
    Set evtData = New Collection
    pIsRunningStatus = isRunStatus
    If pIsRunningStatus Then
        eventCoreLength = 0
    Else
        eventCoreLength = 1
    End If
    pChannel = channelNibble
    If Is7BitValue(dataByte1) Then
        pData1 = dataByte1
        evtData.Add dataByte1
    Else
        Stop
    End If
    If IsMissing(dataByte2) Then
        pChannelEventType = ceTwoByte
    Else
        pChannelEventType = ceThreeByte
        If Is7BitValue(dataByte2) Then
            pData2 = dataByte2
            evtData.Add dataByte2
        Else
            Stop
        End If
    End If
    
    Set pCoreEvent = Factory.CreateNewCoreEvent(deltaTime, absoluteTime, statusNibble, evtData, eventCoreLength)
    
    pStatusName = ToStatusName(statusNibble)
    If statusNibble = StatusEnum.CONTROLLER Then
        pStatusName = pStatusName & ":" & ToControllerName(dataByte1)
    End If
End Sub

Public Function ToBytes() As Byte()
    ToBytes = pCoreEvent.ToBytes(Me)
End Function

Public Function JoinTwoNibbles(ByVal nibbleHigh As Byte, ByVal nibbleLow As Byte) As Byte
    JoinTwoNibbles = BitUtils.ShiftBitsLeft(nibbleHigh, numBits:=4) Or nibbleLow
End Function

Private Function Is7BitValue(ByVal valueToTest As Byte) As Boolean
    Is7BitValue = valueToTest >= 0 And valueToTest <= 127
End Function

Private Function ToStatusName(ByVal statusNibble As Byte) As String
    'converts status nibble to name: e.g. 8 --> "NoteOff"
    Select Case statusNibble
        Case StatusEnum.NOTE_OFF
            ToStatusName = "NoteOff"
        Case StatusEnum.NOTE_ON
            ToStatusName = "NoteOn"
        Case StatusEnum.NOTE_AFTERTOUCH
            ToStatusName = "NoteAftertouch"
        Case StatusEnum.CONTROLLER
            ToStatusName = "Controller"
        Case StatusEnum.PROGRAM_CHANGE
            ToStatusName = "ProgramChange"
        Case StatusEnum.CHANNEL_AFTERTOUCH
            ToStatusName = "ChannelAftertouch"
        Case StatusEnum.PITCH_BEND
            ToStatusName = "PitchBend"
        Case Else
            ToStatusName = "Invalid status byte"
            Stop
    End Select
End Function

Private Function ToControllerName(ByVal controllerNumber As Byte) As String
    Select Case controllerNumber
        Case ContinuousControllerType.BANK_SELECT_MSB
            ToControllerName = "BankSelectMSB"
        Case ContinuousControllerType.MODULATION_MSB
            ToControllerName = "ModulationMSB"
        Case ContinuousControllerType.BREATH_CONTROLLER_MSB
            ToControllerName = "BreathControllerMSB"
        Case ContinuousControllerType.FOOT_CONTROLLER_MSB
            ToControllerName = "FootControllerMSB"
        Case ContinuousControllerType.PORTAMENTO_TIME_MSB
            ToControllerName = "PortamentoTimeMSB"
        Case ContinuousControllerType.DATA_ENTRY_MSB
            ToControllerName = "DataEntryMSB"
        Case ContinuousControllerType.CHANNEL_VOLUME_MSB
            ToControllerName = "ChannelVolumeMSB"
        Case ContinuousControllerType.BALANCE_MSB
            ToControllerName = "BalanceMSB"
        Case ContinuousControllerType.PAN_MSB
            ToControllerName = "PanMSB"
        Case ContinuousControllerType.EXPRESSION_CONTROLLER_MSB
            ToControllerName = "ExpressionControllerMSB"
        Case ContinuousControllerType.EFFECT_CONTROL_1_MSB
            ToControllerName = "EffectControl1MSB"
        Case ContinuousControllerType.EFFECT_CONTROL_2_MSB
            ToControllerName = "EffectControl2MSB"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_14
            ToControllerName = "ContinuousController14"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_15
            ToControllerName = "ContinuousController15"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_1
            ToControllerName = "GeneralPurposeController1"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_2
            ToControllerName = "GeneralPurposeController2"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_3
            ToControllerName = "GeneralPurposeController3"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_4
            ToControllerName = "GeneralPurposeController4"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_20
            ToControllerName = "ContinuousController20"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_21
            ToControllerName = "ContinuousController21"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_22
            ToControllerName = "ContinuousController22"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_23
            ToControllerName = "ContinuousController23"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_24
            ToControllerName = "ContinuousController24"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_25
            ToControllerName = "ContinuousController25"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_26
            ToControllerName = "ContinuousController26"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_27
            ToControllerName = "ContinuousController27"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_28
            ToControllerName = "ContinuousController28"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_29
            ToControllerName = "ContinuousController29"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_30
            ToControllerName = "ContinuousController30"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_31
            ToControllerName = "ContinuousController31"
        Case ContinuousControllerType.BALANCE_LSB
            ToControllerName = "BankSelectLSB"
        Case ContinuousControllerType.MODULATION_LSB
            ToControllerName = "ModulationLSB"
        Case ContinuousControllerType.BREATH_CONTROLLER_LSB
            ToControllerName = "BreathControllerLSB"
        Case ContinuousControllerType.FOOT_CONTROLLER_LSB
            ToControllerName = "FootControllerLSB"
        Case ContinuousControllerType.PORTAMENTO_TIME_LSB
            ToControllerName = "PortamentoTimeLSB"
        Case ContinuousControllerType.DATA_ENTRY_LSB
            ToControllerName = "DataEntryLSB"
        Case ContinuousControllerType.CHANNEL_VOLUME_LSB
            ToControllerName = "ChannelVolumeLSB"
        Case ContinuousControllerType.BALANCE_LSB
            ToControllerName = "BalanceLSB"
        Case ContinuousControllerType.PAN_LSB
            ToControllerName = "PanLSB"
        Case ContinuousControllerType.EXPRESSION_CONTROLLER_LSB
            ToControllerName = "ExpressionControllerLSB"
        Case ContinuousControllerType.EFFECT_CONTROL_1_LSB
            ToControllerName = "EffectControl1LSB"
        Case ContinuousControllerType.EFFECT_CONTROL_2_LSB
            ToControllerName = "EffectControl2LSB"
        Case ContinuousControllerType.CONTROLLER_14_LSB
            ToControllerName = "Controller14LSB"
        Case ContinuousControllerType.CONTROLLER_15_LSB
            ToControllerName = "Controller15LSB"
        Case ContinuousControllerType.CONTROLLER_16_LSB
            ToControllerName = "Controller16LSB"
        Case ContinuousControllerType.CONTROLLER_17_LSB
            ToControllerName = "Controller17LSB"
        Case ContinuousControllerType.CONTROLLER_18_LSB
            ToControllerName = "Controller18LSB"
        Case ContinuousControllerType.CONTROLLER_19_LSB
            ToControllerName = "Controller19LSB"
        Case ContinuousControllerType.CONTROLLER_20_LSB
            ToControllerName = "Controller20LSB"
        Case ContinuousControllerType.CONTROLLER_21_LSB
            ToControllerName = "Controller21LSB"
        Case ContinuousControllerType.CONTROLLER_22_LSB
            ToControllerName = "Controller22LSB"
        Case ContinuousControllerType.CONTROLLER_23_LSB
            ToControllerName = "Controller23LSB"
        Case ContinuousControllerType.CONTROLLER_24_LSB
            ToControllerName = "Controller24LSB"
        Case ContinuousControllerType.CONTROLLER_25_LSB
            ToControllerName = "Controller25LSB"
        Case ContinuousControllerType.CONTROLLER_26_LSB
            ToControllerName = "Controller26LSB"
        Case ContinuousControllerType.CONTROLLER_27_LSB
            ToControllerName = "Controller27LSB"
        Case ContinuousControllerType.CONTROLLER_28_LSB
            ToControllerName = "Controller28LSB"
        Case ContinuousControllerType.CONTROLLER_29_LSB
            ToControllerName = "Controller29LSB"
        Case ContinuousControllerType.CONTROLLER_30_LSB
            ToControllerName = "Controller30LSB"
        Case ContinuousControllerType.CONTROLLER_31_LSB
            ToControllerName = "Controller31LSB"
        Case ContinuousControllerType.SUSTAIN_PEDAL
            ToControllerName = "SustainPedal"
        Case ContinuousControllerType.PORTAMENTO
            ToControllerName = "Portamento"
        Case ContinuousControllerType.SOSTENUTO
            ToControllerName = "Sostenuto"
        Case ContinuousControllerType.SOFT_PEDAL
            ToControllerName = "SoftPedal"
        Case ContinuousControllerType.LEGATO_FOOTSWITCH
            ToControllerName = "LegatoFootswitch"
        Case ContinuousControllerType.HOLD_2_PEDAL
            ToControllerName = "Hold2Pedal"
        Case ContinuousControllerType.SOUND_VARIATION
            ToControllerName = "SoundVariation"
        Case ContinuousControllerType.SOUND_RESONANCE
            ToControllerName = "SoundResonance"
        Case ContinuousControllerType.SOUND_RELEASE_TIME
            ToControllerName = "SoundReleaseTime"
        Case ContinuousControllerType.SOUND_ATTACK_TIME
            ToControllerName = "SoundAttackTime"
        Case ContinuousControllerType.SOUND_FREQUENCY_CUTOFF
            ToControllerName = "SoundFrequencyCutoff"
        Case ContinuousControllerType.SOUND_CONTROLLER_6
            ToControllerName = "SoundController6"
        Case ContinuousControllerType.SOUND_CONTROLLER_7
            ToControllerName = "SoundController7"
        Case ContinuousControllerType.SOUND_CONTROLLER_8
            ToControllerName = "SoundController8"
        Case ContinuousControllerType.SOUND_CONTROLLER_9
            ToControllerName = "SoundController9"
        Case ContinuousControllerType.SOUND_CONTROLLER_10
            ToControllerName = "SoundController10"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_5
            ToControllerName = "GeneralPurposeController5"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_6
            ToControllerName = "GeneralPurposeController6"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_7
            ToControllerName = "GeneralPurposeController7"
        Case ContinuousControllerType.GENERAL_PURPOSE_CONTROLLER_8
            ToControllerName = "GeneralPurposeController8"
        Case ContinuousControllerType.PORTAMENTO_CONTROL
            ToControllerName = "PortamentoControl"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_85
            ToControllerName = "ContinuousController85"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_86
            ToControllerName = "ContinuousController86"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_87
            ToControllerName = "ContinuousController87"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_88
            ToControllerName = "ContinuousController88"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_89
            ToControllerName = "ContinuousController89"
        Case ContinuousControllerType.CONTINUOUS_CONTROLLER_90
            ToControllerName = "ContinuousController90"
        Case ContinuousControllerType.EFFECTS_DEPTH_1_EXTERNAL_EFFECTS_DEPTH
            ToControllerName = "Effects1Depth(ExternalEffectsDepth)"
        Case ContinuousControllerType.EFFECTS_DEPTH_2_TREMOLO_DEPTH
            ToControllerName = "Effects2Depth(TremoloDepth)"
        Case ContinuousControllerType.EFFECTS_DEPTH_3_CHORUS_DEPTH
            ToControllerName = "Effects3Depth(ChorusDepth)"
        Case ContinuousControllerType.EFFECTS_DEPTH_4_CELESTE_DETUNE
            ToControllerName = "Effects4Depth(CelesteDetune)"
        Case ContinuousControllerType.EFFECTS_DEPTH_5_PHASER_DEPTH
            ToControllerName = "Effects5Depth(PhaserDepth)"
        Case ContinuousControllerType.DATA_INCREMENT
            ToControllerName = "DataIncrement"
        Case ContinuousControllerType.DATA_DECREMENT
            ToControllerName = "DataDecrement"
        Case ContinuousControllerType.NON_REGISTERED_PARAMETER_NUMBER_LSB
            ToControllerName = "NonRegisteredParameterNumberLSB"
        Case ContinuousControllerType.NON_REGISTERED_PARAMETER_NUMBER_MSB
            ToControllerName = "NonRegisteredParameterNumberMSB"
        Case ContinuousControllerType.REGISTERED_PARAMETER_NUMBER_LSB
            ToControllerName = "RegisteredParameterNumberLSB"
        Case ContinuousControllerType.REGISTERED_PARAMETER_NUMBER_MSB
            ToControllerName = "RegisteredParameterNumberMSB"
        Case ContinuousControllerType.ALL_SOUND_OFF
            ToControllerName = "AllSoundsOff"
        Case ContinuousControllerType.RESET_ALL_CONTROLLERS
            ToControllerName = "ResetAllControllers"
        Case ContinuousControllerType.LOCAL_CONTROL_ON_OFF
            ToControllerName = "LocalContorlOnOff"
        Case ContinuousControllerType.ALL_NOTES_OFF
            ToControllerName = "AllNotesOff"
        Case ContinuousControllerType.OMNI_MODE_OFF
            ToControllerName = "OmniModeOff"
        Case ContinuousControllerType.OMNI_MODE_ON
            ToControllerName = "OmniModeOn"
        Case ContinuousControllerType.POLY_MODE_ON_OFF
            ToControllerName = "PolyModeOnOff"
        Case ContinuousControllerType.POLY_MODE_ON
            ToControllerName = "PolyModeOn"
        Case Else
            ToControllerName = "Undefined"
    End Select
End Function


