Dim word 
dim num
word="strok"
num=(WScript.Arguments(0))
num2=CInt(num)
If num2 < 21 Then
        word = "strok"
        ElseIf num2 = 21 Then
        word = "stroka"
        ElseIf num2 > 21 And num2 < 25 Then
        word = "stroki"
		ElseIf num2 > 25 Then
        word = "strok"
 End If
WScript.Echo word

