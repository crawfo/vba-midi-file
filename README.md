# vba-midi
A set of classes, functions, and methods for reading and writing MIDI files from Excel written in VBA.

A factory module is used for the creation of objects. Most objects are immutable.

To parse a MIDI file, call the following factory functions which return the parsed MIDI bytes as a MidiFile object:
Factory.CreateNewMidiFileFromArray(midiFileBytes() As Byte)
Factory.CreateNewMidiFileFromFile(ByVal fileNameFullyQualified As String)

A MidiFile object contains EventTracks objects which contain EventTrack objects which contain ChannelEvent, MetaEvent, and SystemExclusiveEvent objects. The original file bytes can be accessed with the FileBytes property, or regenerated with the ToBytes function.

To create a MIDI file, call the following factory function:
Factory.CreateNewMidiFile(ByVal hdrChunk As HeaderChunk, ByVal eventTrks As EventTracks).

To create a MIDI file from scratch:
1) Create any MIDI event with:
   a) the basic event constructors:
      Factory.CreateNewChannelEvent, 
      Factory.CreateNewMetaEvent, 
      Factory.CreateNewSystemExclusiveEvent.
   and 
   b) the convenience constructors:
      Factory.CreateNewNoteOnEvent,
      Factory.CreateNewNoteOffEvent,
      Factory.CreateNewNoteAftertouchEvent,
      Factory.CreateNewControlChangeEvent,
      Factory.CreateNewChannelAftertouchEvent,
      Factory.CreateNewPitchBendEvent,
      Factory.CreateNewProgramChangeEvent,
      Factory.CreateNewTimeSignatureMetaEvent,
      Factory.CreateNewKeySignatureMetaEvent.
   
2) Add all events to a collection.
3) Create an EventTrack object with Factory.CreateNewEventTrack(ByVal trkEvents As Collection).
4) Create an EventTracks object with Factory.CreateNewEventTracks(ByVal eventTrks As Collection).
5) Create a HeaderChunk object with Factory.CreateNewHeaderChunk
6) Create an MidiFile object with Factory.CreateNewMidiFile(ByVal hdrChunk As HeaderChunk, ByVal eventTrks As EventTracks).
7) To save a MidiFile object as a standard midi file use the FileUtils.WriteToDisk(bytes() As Byte, ByVal fileNameFullyQualified As String) function, passing MidiFile.FileBytes as the bytes parameter.
