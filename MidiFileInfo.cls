VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MidiFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pFileFormat As Integer
Private pNumberTracks As Long
Private pTimeDivision As Long

Public Property Get FileFormat() As Integer
    FileFormat = pFileFormat
End Property

Public Property Get NumberTracks() As Long
    NumberTracks = pNumberTracks
End Property

Public Property Get TimeDivision() As Long
    TimeDivision = pTimeDivision
End Property

Public Sub Initialize(ByVal midiFileFormat As Integer, ByVal numTracks As Long, ByVal timeDiv As Long)
    'side effect: mutates object properties
    pFileFormat = midiFileFormat
    pNumberTracks = numTracks
    pTimeDivision = timeDiv
End Sub





