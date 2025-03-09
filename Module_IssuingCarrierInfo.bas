Option Explicit

Sub IssuingCarrierInfo()
    Dim IssuingCarrier As String
    Dim AgentIATACode As String
    Dim AgentAccountCode As String

    ' Clear the content of the target cells
    With wsMAWB.Range("A15").MergeArea
        .ClearContents
    End With
    With wsMAWB.Range("A19").MergeArea
        .ClearContents
    End With
    With wsMAWB.Range("K19").MergeArea
        .ClearContents
    End With

    ' Get the value from wsMAWBConfig cells and put them into wsMAWB cells
    IssuingCarrier = wsMAWBConfig.Cells(5, 2).Value
    wsMAWB.Cells(15, 1).Value = IssuingCarrier & " / HKG"

    AgentIATACode = wsMAWBConfig.Cells(6, 2).Value
    wsMAWB.Cells(19, 1).Value = AgentIATACode

    AgentAccountCode = wsMAWBConfig.Cells(7, 2).Value
    wsMAWB.Cells(19, 11).Value = AgentAccountCode
End Sub
