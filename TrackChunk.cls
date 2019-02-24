VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackChunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pChunkBytes() As Byte

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator Property"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pChunkBytes.[_NewEnum]
End Property

Public Property Get ChunkType() As Byte()
    Const CHUNK_TYPE_START_INDEX = 0
    Const CHUNK_TYPE_END_INDEX = 3
    ChunkType = ListUtils.ByteArraySlice(pChunkBytes, _
                                         sliceStart:=CHUNK_TYPE_START_INDEX, _
                                         sliceEnd:=CHUNK_TYPE_END_INDEX)
End Property

Public Property Get ChunkSizeBytes() As Byte()
    Const CHUNK_SIZE_START_INDEX = 4
    Const CHUNK_SIZE_END_INDEX = 7
    ChunkSizeBytes = ListUtils.ByteArraySlice(pChunkBytes, _
                                              sliceStart:=CHUNK_SIZE_START_INDEX, _
                                              sliceEnd:=CHUNK_SIZE_END_INDEX)
End Property

Public Property Get ChunkBytes() As Byte()
    ChunkBytes = pChunkBytes
End Property

Public Property Get Item(arrayIndex As Long) As Byte
Attribute Item.VB_Description = "Default Property"
Attribute Item.VB_UserMemId = 0
    Item = pChunkBytes(arrayIndex)
End Property

Public Sub Initialize(trackBytes() As Byte)
    pChunkBytes = trackBytes
End Sub

Public Function ToBytes() As Byte()
    ToBytes = pChunkBytes
End Function

