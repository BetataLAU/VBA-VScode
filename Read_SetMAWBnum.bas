Option Explicit

Sub Read_SetMAWBnum(ByRef arr As Variant, startRow As Long)
    ' Ensure wsMAWB is correctly referenced
    If wsMAWB Is Nothing Then
        MsgBox "Worksheet 'MAWB' is not set."
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not ValidateAWB(CStr(arr(i, 1))) Then
            MsgBox "AWB number is invalid at row " & (startRow + i - 1)
            wsMAWBConfig.Cells(startRow + i - 1, 1).Select
            Exit Sub
        End If
    Next i

    ' Set the MAWB number in the MAWB sheet
    ' Extract the suffix (last 8 digits)
    Dim suffix As String
    suffix = Right(arr(1, 1), 8)

    ' Extract the prefix (remaining digits)
    Dim prefix As String
    prefix = Left(arr(1, 1), Len(arr(1, 1)) - 8)

    ' Ensure the prefix is 3 digits by adding "0" to the left if necessary
    Do While Len(prefix) < 3
        prefix = "0" & prefix
    Loop

    ' Set the values in the MAWB sheet
    With wsMAWB
        .Range("A1").Value = prefix
        .Range("E1").Value = suffix
        .Range("C1").Value = "HKG"
        .Range("AH1").Value = "=A1"
        .Range("AJ1").Value = "=E1"
        .Range("AF62").Value = "=AH1"
        .Range("AH62").Value = "HKG"
        .Range("AJ62").Value = "=AJ1"
    End With
End Sub