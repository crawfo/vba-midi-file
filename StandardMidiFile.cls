VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StandardMidiFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pHeaderChunk As HeaderChunk
Private pTrackList As TrackCollection
Private pFileBytes() As Byte

Public Property Get HeaderChunk() As HeaderChunk
    Set HeaderChunk = pHeaderChunk
End Property

Public Property Get TrackList() As TrackCollection
    Set TrackList = pTrackList
End Property

Public Property Set TrackList(tracks As TrackCollection)
    Set pTrackList = tracks
End Property

Public Property Get FileBytes() As Byte()
    FileBytes = pFileBytes
End Property

Public Property Let FileBytes(bytes() As Byte)
    pFileBytes = bytes
End Property

Public Sub Initialize(ByVal midiFileFormat As Integer, ByVal timeDiv As Long)
    'side effect: mutates object properties
    Const NUM_TRACKS = 0
    Set pTrackList = New TrackCollection
    Set pHeaderChunk = Factory.CreateNewHeaderChunk(midiFileFormat, NUM_TRACKS, timeDiv)
End Sub

Public Sub InitA(ByVal midiFileFormat As Integer, ByVal timeDiv As Long, ByVal tracks As TrackCollection)
    'side effect: mutates object properties
    Set pTrackList = tracks
    Set pHeaderChunk = Factory.CreateNewHeaderChunk(midiFileFormat, tracks.Count, timeDiv)
    ToBytes
End Sub

Public Sub AddTrack(ByVal track As TrackChunk)
    'side effect: mutates pTrackList, pHeaderChunk
    pTrackList.Add track
    'update num tracks
    Set pHeaderChunk = Factory.CreateNewHeaderChunk(HeaderChunk.FileFormat, TrackList.Count, HeaderChunk.timeDiv)
End Sub

Public Sub ToBytes()
    'side effect: mutates pFileBytes
    Dim headerBytes() As Byte, trackBytes() As Byte
    Dim midiTrack As Variant, trackByte As Variant, headerByte As Variant
    Dim trackBytesCollection As Collection, trackByteCount As Long, i As Long
    
    headerBytes = pHeaderChunk.ChunkBytes
    Set trackBytesCollection = New Collection
    'put all bytes from each track into a collection
    For Each midiTrack In pTrackList
        trackBytes = midiTrack.ChunkBytes
        For Each trackByte In trackBytes
            trackBytesCollection.Add trackByte
        Next trackByte
    Next midiTrack

    trackByteCount = trackBytesCollection.Count
    i = 0
    ReDim pFileBytes(trackBytesCollection.Count + UBound(headerBytes))
    
    'put header chunk bytes into an array
    For Each headerByte In headerBytes
        pFileBytes(i) = headerByte
        i = i + 1
    Next headerByte
    'put all bytes from all tracks into an array
    For Each trackByte In trackBytesCollection
        pFileBytes(i) = trackByte
        i = i + 1
    Next trackByte
End Sub

Public Sub WriteToDisk(ByVal fileNameFullyQualified As String)
    Open fileNameFullyQualified For Binary As #1
    Put #1, , pFileBytes
    Close #1
End Sub

Private Function ToTrackCollection(ByVal collectionOfCollections As Collection) As TrackCollection
    Dim bytes As Collection
    Set ToTrackCollection = New TrackCollection
    For Each bytes In collectionOfCollections
        ToTrackCollection.Add Factory.CreateNewTrackChunk(bytes)
    Next bytes
End Function
















