Attribute VB_Name = "ListUtils"
Option Explicit

Public Function ByteArraySlice(sourceArray() As Byte, _
                               ByVal sliceStart As Long, _
                               ByVal sliceEnd As Long) As Byte()
    Dim sourceIndex As Long
    Dim sliceIndex As Long
    Dim upperBound As Long
    Dim slice() As Byte
    
    upperBound = sliceEnd - sliceStart
    ReDim slice(upperBound)
    sliceIndex = 0
    
    For sourceIndex = sliceStart To sliceEnd
        slice(sliceIndex) = sourceArray(sourceIndex)
        sliceIndex = sliceIndex + 1
    Next sourceIndex
    
    ByteArraySlice = slice
End Function

Public Function CollectionSlice(ByVal sourceCollection As Collection, _
                                ByVal sliceStart As Long, _
                                ByVal sliceEnd As Long) As Collection
    Dim sourceIndex As Long
    Dim slice As Collection
    
    Set slice = New Collection
    
    For sourceIndex = sliceStart To sliceEnd
        slice.Add sourceCollection(sourceIndex)
    Next sourceIndex
    
    Set CollectionSlice = slice
End Function

Public Function CollectionSliceFromArray(sourceArray() As Byte, _
                                         ByVal sliceStart As Long, _
                                         ByVal sliceEnd As Long) As Collection
    Dim sourceIndex As Long
    Dim slice As Collection
    
    Set slice = New Collection
    
    For sourceIndex = sliceStart To sliceEnd
        slice.Add sourceArray(sourceIndex)
    Next sourceIndex
    
    Set CollectionSliceFromArray = slice
End Function

Public Function AppendCollectionToCollection(ByVal collection1 As Collection, _
                                             ByVal collectionToAppend As Collection) As Collection
    Dim e As Variant
    Dim newCollection As Collection
    
    Set newCollection = collection1
    
    For Each e In collectionToAppend
        newCollection.Add e
    Next e
    
    Set AppendCollectionToCollection = newCollection
End Function

Public Function ToByteArray(variantArray As Variant) As Byte()
    Dim e As Variant
    Dim i As Long
    Dim byteArray() As Byte
    
    ReDim byteArray(UBound(variantArray))
    i = 0
    For Each e In variantArray
        byteArray(i) = CByte(e)
        i = i + 1
    Next e
    
    ToByteArray = byteArray
End Function

Public Function ToCollectionFromByteArray(byteArray() As Byte) As Collection
    Dim e As Variant
    Dim bytes As Collection
    
    Set bytes = New Collection
    For Each e In byteArray
        bytes.Add CByte(e)
    Next e
    
    Set ToCollectionFromByteArray = bytes
End Function

