Option Explicit

Sub SetPort(arr As Variant, currentRow As Long)
    ' Set wsMAWB cell A21 to "HONG KONG"
    wsMAWB.Cells(21, LetterToNumber("A")).Value = "HONG KONG"

    ' Define the RoutingPort and CarrierCode variables
    Dim RoutingPort As String
    Dim CarrierCode As String
    Dim cell As Range

    ' Find the cell containing "<Routing>"
    Set cell = wsMAWBConfig.Cells.Find(What:="<Routing>", LookIn:=xlValues, LookAt:=xlPart)
    If Not cell Is Nothing Then
        ' Assign the value to the right of "<Routing>" to RoutingPort and convert to upper case
        RoutingPort = UCase(cell.Offset(0, 1).Value)
    Else
        MsgBox "Routing not found."
        Exit Sub
    End If

    ' Find the cell containing "<Carrier Code>"
    Set cell = wsMAWBConfig.Cells.Find(What:="<Carrier Code>", LookIn:=xlValues, LookAt:=xlPart)
    If Not cell Is Nothing Then
        ' Assign the value to the right of "<Carrier Code>" to CarrierCode and convert to upper case
        CarrierCode = UCase(cell.Offset(0, 1).Value)
    Else
        MsgBox "Carrier Code not found."
        Exit Sub
    End If

    ' Check if wsMAWBConfig cell H10 is not empty
    If RoutingPort <> "" Then
        ' Set wsMAWB cell A23 to the routing port
        wsMAWB.Cells(23, LetterToNumber("A")).Value = RoutingPort
        ' Set wsMAWB cell D23 to the carrier code
        wsMAWB.Cells(23, LetterToNumber("D")).Value = CarrierCode
        ' Set wsMAWB cell M23 to the port code from arr column "B"
        wsMAWB.Cells(23, LetterToNumber("M")).Value = arr(1, 2)
        ' Set the right cell of M23 to the carrier code
        wsMAWB.Cells(23, LetterToNumber("O")).Value = CarrierCode
    Else
        ' Set wsMAWB cell A23 to the port code from arr column "B"
        wsMAWB.Cells(23, LetterToNumber("A")).Value = arr(1, 2)
        ' Set the right cell of A23 to the carrier code
        wsMAWB.Cells(23, LetterToNumber("D")).Value = CarrierCode
        
        wsMAWB.Cells(23, LetterToNumber("M")).Value = ""
        wsMAWB.Cells(23, LetterToNumber("O")).Value = ""
    End If

    ' Load the sheet DEST-IATA rate and set to a variable
    Dim wsDESTIATARate As Worksheet
    Set wsDESTIATARate = ThisWorkbook.Sheets("DEST-IATA rate")

    ' Search for the value in arr(1, 2) from this sheet
    Dim portCode As String
    portCode = UCase(arr(1, 2)) ' Convert port code to upper case
    Set cell = wsDESTIATARate.Cells.Find(What:=portCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not cell Is Nothing Then
        ' Return the Complete Port Name and assign it to wsMAWB cell A25
        wsMAWB.Cells(25, LetterToNumber("A")).Value = cell.Offset(0, 1).Value
    Else
        MsgBox "Port code not found in DEST-IATA rate sheet."
    End If
End Sub