Option Explicit

Sub Shipper(ByRef arr As Variant, startRow As Long)
    Dim shipperCode As String
    Dim shipperDetails As String
    Dim found As Range

    ' Read the shipper code from the "P" column (16th column in the array)
    shipperCode = arr(1, 16)

    ' Find the shipper details in the wsSHP sheet
    Set found = wsSHP.Columns("A").Find(What:=shipperCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        ' Get the shipper details from column B
        shipperDetails = found.Offset(0, 1).Value
        ' Put the shipper details in wsMAWB cell A3
        wsMAWB.Cells(3, 1).Value = shipperDetails
    Else
        MsgBox "Shipper code not found: " & shipperCode
    End If
End Sub