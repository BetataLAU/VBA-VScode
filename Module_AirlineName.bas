Option Explicit

Sub AirlineName(ByRef arr As Variant, startRow As Long)
    Dim Airline As String
    Dim Prefix As String
    Dim found As Range
    
    ' Get the prefix from cell B3
    Prefix = GetAWBPrefix(CStr(arr(1, 1)))
    
    ' Find the corresponding prefix in column A of "Airline Info and Remark" sheet
    Set found = wsAirlineInfo.Columns("A").Find(What:=Prefix, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        ' Get the airline details from column B
        Airline = found.Offset(0, 1).Value
        ' Put the airline details in wsMAWB cell A9
        wsMAWB.Cells(3, LetterToNumber("Z")).Value = Airline
    Else
        MsgBox "Airline details not found: " & Prefix
    End If
 
End Sub