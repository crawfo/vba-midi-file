# vba-midi
A set of classes, functions, and methods for reading and writing MIDI files from Excel written in VBA.

To parse a MIDI file, call the Midi.ParseMidiFile function which will return a collection of tracks each containing MetaEvent, ChannelEvent, or SystemExclusiveEvent objects.

A Factory module is provided for the safe creation of all MIDI related objects. Most objects are immutable once created.

Creating a MIDI file is left to the implementor to ensure validity (if desired) and requires the creation of StandardMidiFile objects and TrackChunk objects. StandardMidiFile objects contain a Write method which will write the object to disk as a midi file.
