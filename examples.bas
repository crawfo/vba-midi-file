Sub ExampleReadMidiFileIntoDataStructure()
    Dim parsedTracks As Collection
    Dim fileNameFullyQualified As String

    fileNameFullyQualified = "C:\exampleFileName.mid" 'place the correct fully qualified filename here
    Set parsedTracks = ParseMidiFile(fileNameFullyQualified)
    Stop
    'Examine the parsedTracks collection in the View > Locals window.
    'Each element represents a track of the midi file...
    '...and contains MetaEvent, ChannelEvent, or SystemExclusiveEvent objects.
End Sub
