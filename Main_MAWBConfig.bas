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

    ' Check if a range is selected
    If TypeName(Selection) = "Range" Then
        startRow = Selection.Row
        endRow = Selection.Row + Selection.Rows.Count - 1
        
        ' Define the range from column A to Y for the selected rows
        Set rng = wsMAWBConfig.Range(wsMAWBConfig.Cells(startRow, 1), wsMAWBConfig.Cells(endRow, 25))
        arr = rng.Value
    Else
        MsgBox "Please select a range first."
        Exit Sub
    End If

    ' Main procedure.
    Call Read_SetMAWBnum(arr, startRow)
    Call AirlineName
    Call Shipper(arr, startRow)

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