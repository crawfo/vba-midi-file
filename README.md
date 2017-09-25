# vba-midi
A set of classes, functions, and methods for reading and writing MIDI files from Excel written in VBA.

To parse a MIDI file, call the ParseMidiFile function which will return a collection of tracks each containing MetaEvent, ChannelEvent, or SystemExclusiveEvent objects.
