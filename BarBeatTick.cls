VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BarBeatTick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pBar As Long
Private pBeat As Long
Private pTick As Long

Public Property Get Bar() As Long
    Bar = pBar
End Property

Public Property Get Beat() As Long
    Beat = pBeat
End Property

Public Property Get Tick() As Long
    Tick = pTick
End Property

Public Sub Initialize(ByVal positionInTicks As Long, ByVal ticksPerQuarterNote As Long)
    Dim ticksLeft As Long, ticksPerBeat As Long, ticksPerBar As Long
    ticksPerBar = ticksPerQuarterNote * 4
    pBar = (positionInTicks \ ticksPerBar) + 1 'add 1 because first bar is 1 not 0
    ticksLeft = positionInTicks Mod ticksPerBar
    pBeat = (ticksLeft \ ticksPerQuarterNote) + 1 'add 1 because first beat is 1 not 0
    pTick = ticksLeft Mod ticksPerQuarterNote
End Sub

Public Function ToTick(ByVal bbt As BarBeatTick, ByVal ticksPerQuarterNote As Long) As Long
    Dim ticksPerBar As Long
    ticksPerBar = ticksPerQuarterNote * 4
    'subtract 1 from bar, beat because first bar, beat is 1 not 0
    ToTick = (bbt.Bar - 1) * ticksPerBar + (bbt.Beat - 1) * ticksPerQuarterNote + bbt.Tick
End Function
