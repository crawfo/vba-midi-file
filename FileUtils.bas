Attribute VB_Name = "FileUtils"
Option Explicit

Public Function ReadFile(ByVal fileNameFullyQualified As String) As Byte()
    Dim bytes() As Byte
    Open fileNameFullyQualified For Binary As #1
    ReDim bytes(LOF(1) - 1)
    Get #1, , bytes
    Close #1
    ReadFile = bytes
End Function

Public Sub WriteToDisk(bytes() As Byte, ByVal fileNameFullyQualified As String)
    Open fileNameFullyQualified For Binary As #1
    Put #1, , bytes
    Close #1
End Sub

Public Sub ImportToSheet(ByVal fileNameFullyQualified As String, ByVal sheetName As String)
    Const DIMENSION_2_UPPER_BOUND = 0
    Const DIMENSION_2_INDEX = 0
    Dim bytes As Variant
    Dim i As Long
    Dim e As Variant
    Dim bytesV() As Variant
    
    bytes = FileUtils.ReadFile(fileNameFullyQualified)
    'must convert to 2D variant array before assigning to a range
    ReDim bytesV(UBound(bytes), DIMENSION_2_UPPER_BOUND)
    For Each e In bytes
        bytesV(i, DIMENSION_2_INDEX) = e
        i = i + 1
    Next e
    
    Const START_CELL_ROW_INDEX = 1
    Const COLUMN_INDEX = 1
    Dim endCellRowIndex As Long
    Dim dataSheet As Worksheet
    
    Set dataSheet = Worksheets(sheetName)
    endCellRowIndex = i
    dataSheet.Range(Cells(START_CELL_ROW_INDEX, COLUMN_INDEX), _
                    Cells(endCellRowIndex, COLUMN_INDEX)) = bytesV
End Sub    
