Dim word As String = "строк"
dim num
num=(WScript.Arguments(0))
num2=CInt(num)
If num2 < 21 Then
        word = "строк"
ElseIf num2 = 21 Then
        word = "строка"
ElseIf num2 > 21 And num2 < 25 Then
        word = "строки"
End If
WScript.OutArgument(word)

