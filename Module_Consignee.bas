Option Explicit

Sub Consignee(ByRef arr As Variant, startRow As Long)
    Dim consigneeCode As String
    Dim consigneeDetails As String
    Dim found As Range

    ' Read the consignee code from the "Q" column (17th column in the array)
    consigneeCode = arr(1, 17)

    ' Find the consignee details in the wsSHP sheet
    Set found = wsCNE.Columns("A").Find(What:=consigneeCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        ' Get the consignee details from column B
        consigneeDetails = found.Offset(0, 1).Value
        ' Put the consignee details in wsMAWB cell A9
        wsMAWB.Cells(9, 1).Value = consigneeDetails
    Else
        MsgBox "Consignee code not found: " & consigneeCode
    End If
End Sub