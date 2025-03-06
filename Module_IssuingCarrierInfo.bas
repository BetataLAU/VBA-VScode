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

    ' Get the value from wsMAWBConfig cell B5
    IssuingCarrier = wsMAWBConfig.Cells(5, 2).Value
    ' Put the value into wsMAWB cell A15 with additional text
    wsMAWB.Cells(15, 1).Value = IssuingCarrier & " / HKG"

    ' Get the value from wsMAWBConfig cell B6
    AgentIATACode = wsMAWBConfig.Cells(6, 2).Value
    ' Put the value into wsMAWB cell A19
    wsMAWB.Cells(19, 1).Value = AgentIATACode

    ' Get the value from wsMAWBConfig cell B7
    AgentAccountCode = wsMAWBConfig.Cells(7, 2).Value
    ' Put the value into wsMAWB cell K19
    wsMAWB.Cells(19, 11).Value = AgentAccountCode
End Sub
