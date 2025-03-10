Option Explicit

Sub AccountingInfo(ByRef arr As Variant, startRow As Long)
    Dim accountingCode As String
    Dim accountingDetails As String
    Dim found As Range

    ' Read the accounting code from the "S" column (19th column in the array)
    accountingCode = arr(1, 19)

    ' Clear the content of the merged cell U14 before assigning the accounting details
    With wsMAWB.Range("U14").MergeArea
        .ClearContents
    End With

    ' Find the accounting details in the wsACC sheet
    Set found = wsACC.Columns("A").Find(What:=accountingCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        ' Get the accounting details from column B
        accountingDetails = found.Offset(0, 1).Value
        ' Put the accounting details in the merged cell U14
        wsMAWB.Range("U14").Value = vbNewLine & accountingDetails
    Else
        wsMAWB.Range("U14").Value = vbNewLine & "FREIGHT PREPAID"
    End If
End Sub