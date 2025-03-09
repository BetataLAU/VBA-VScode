Option Explicit

Function ValidateAWB(AWBNumber As String) As Boolean
    Dim suffix As String
    Dim suffixFirst7 As String
    Dim suffixLast1 As String
    Dim modResult As Long
    
    ' Ensure the AWB number is 9-11 digits
    If Len(AWBNumber) < 9 Or Len(AWBNumber) > 11 Then
        ValidateAWB = False
        Exit Function
    End If
    
    ' Extract the suffix (last 8 digits)
    suffix = Right(AWBNumber, 8)
    
    ' Extract the first 7 digits of the suffix and the last digit
    suffixFirst7 = Left(suffix, 7)
    suffixLast1 = Right(suffix, 1)
    
    ' Calculate the modulus of the first 7 digits by 7
    modResult = CLng(suffixFirst7) Mod 7
    
    ' Check if the mod result equals the last digit of the suffix
    ValidateAWB = (modResult = CInt(suffixLast1))
End Function

Function LetterToNumber(letter As String) As Integer
    ' Convert the letter to uppercase to handle both uppercase and lowercase letters
    letter = UCase(letter)
    
    ' Calculate the corresponding number
    LetterToNumber = Asc(letter) - Asc("A") + 1
End Function

Function GetAWBPrefix(AWBNumber As String) As String
    Dim prefix As String
    
    ' Ensure the AWB number is 9-11 digits
    If Len(AWBNumber) < 9 Or Len(AWBNumber) > 11 Then
        GetAWBPrefix = ""
        Exit Function
    End If
    
    ' Extract the prefix (first part of the AWB number before the last 8 digits)
    prefix = Left(AWBNumber, Len(AWBNumber) - 8)
    
    ' Ensure the prefix length is 3 by adding "0" to the left side if necessary
    Do While Len(prefix) < 3
        prefix = "0" & prefix
    Loop
    
    GetAWBPrefix = prefix
End Function

Function GetAWBSuffix(AWBNumber As String) As String
    Dim suffix As String
    
    ' Ensure the AWB number is 9-11 digits
    If Len(AWBNumber) < 9 Or Len(AWBNumber) > 11 Then
        GetAWBSuffix = ""
        Exit Function
    End If
    
    ' Extract the suffix (last 8 digits)
    suffix = Right(AWBNumber, 8)
    
    GetAWBSuffix = suffix
End Function