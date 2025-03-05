Option explicit

Sub AirlineName()
    Dim Airline As String
    Airline = wsMAWBConfig.Cells(3, 2).Value    '=Range("B3")
    wsMAWB.Cells(3, LetterToNumber("Z")).Value = Airline
End Sub