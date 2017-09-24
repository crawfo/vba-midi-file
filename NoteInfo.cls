VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pNoteNumber As Integer
Private pTimeStamp As Long
Private pTrackNumber As Long
Private pNoteLength As Long

Public Property Get NoteNumber() As Long
    NoteNumber = pNoteNumber
End Property

Public Property Get TimeStamp() As Long
    TimeStamp = pTimeStamp
End Property

Public Property Get TrackNumber() As Long
    TrackNumber = pTrackNumber
End Property

Public Property Get NoteLength() As Long
    NoteLength = pNoteLength
End Property

Public Sub Initialize(ByVal noteNum As Long, ByVal absoluteTime As Long, ByVal noteLen As Long, ByVal trackNum As Long)
    pNoteNumber = noteNum
    pTimeStamp = absoluteTime
    pTrackNumber = trackNum
    pNoteLength = noteLen
End Sub
