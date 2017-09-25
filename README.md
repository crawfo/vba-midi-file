# vba-midi
A set of classes, functions, and methods for reading and writing MIDI files from Excel written in VBA.

Valid MIDI files are assumed.

A Factory module is provided for the safe creation of all MIDI related objects. Most objects are immutable once created.

To parse a MIDI file, call the Midi.ParseMidiFile function which will return a collection of tracks each containing MetaEvent, ChannelEvent, or SystemExclusiveEvent objects. 

Creating a MIDI file is left to the implementor to ensure validity and requires the creation of a TrackCollection object of TrackChunks. The TrackCollection is then passed to the Factory.CreateStandardMidiFile function to create a StandardMidiFile object.
StandardMidiFile objects contain a Write method which will write the object to disk as a midi file when invoked.

The examples.bas module currently provides an example of usage for the Midi.ParseMidiFile function, called ExampleReadMidiFileIntoDataStructure.
