Option Explicit

Sub Notify(ByRef arr As Variant, startRow As Long)
    Dim notifyCode As String
    Dim notifyDetails As String
    Dim found As Range

    ' Read the notify code from the "R" column (18th column in the array)
    notifyCode = arr(1, 18)

    ' Clear the content of the merged cell A40 before assigning the notify details
    wsMAWB.Range("A40").ClearContents

    ' Find the notify details in the wsNTY sheet
    Set found = wsNTY.Columns("A").Find(What:=notifyCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        ' Get the notify details from column B
        notifyDetails = found.Offset(0, 1).Value
        ' Put the notify details in wsMAWB cell A40
        wsMAWB.Cells(40, 1).Value = notifyDetails
    Else
        MsgBox "Notify code not found: " & notifyCode
    End If
End Sub