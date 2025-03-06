Option Explicit

Sub Main()
    On Error GoTo ErrorHandler

    Set wsMAWBConfig = ThisWorkbook.Sheets("MAWB Config")
    Set wsMAWB = ThisWorkbook.Sheets("MAWB")
    Set wsSHP = ThisWorkbook.Sheets("SHP")

    ' Define the MAWB every details into an array var.
    Dim rng As Range
    Dim arr As Variant
    Dim startRow As Long, endRow As Long
    Dim currentRow As Long

    ' Check if a range is selected
    If TypeName(Selection) = "Range" Then
        startRow = Selection.Row
        endRow = Selection.Row + Selection.Rows.Count - 1
    Else
        MsgBox "Please select a range first."
        Exit Sub
    End If

    ' Loop through each row in the selected range
    For currentRow = startRow To endRow
        ' Define the range for the current row from column A to Y
        Set rng = wsMAWBConfig.Range(wsMAWBConfig.Cells(currentRow, 1), wsMAWBConfig.Cells(currentRow, 25))
        arr = rng.Value

        ' Main procedure for the current row
        Call Read_SetMAWBnum(arr, currentRow)
        Call AirlineName
        Call Shipper(arr, currentRow)
    Next currentRow

    ' Clean up.
    Set wsMAWBConfig = Nothing
    Set wsMAWB = Nothing
    Set wsSHP = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ' Clean up in case of error
    Set wsMAWBConfig = Nothing
    Set wsMAWB = Nothing
    Set wsSHP = Nothing
End Sub