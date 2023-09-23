Public Function bilance(a, b, c)
bilance = a + b - c
End Function

Public Function atlaide(a, b)
If a = 0 Then
atlaide = 0
ElseIf a >= 1 Then
atlaide = b * a * 3 / 100
ElseIf a = 2 Then
atlaide = b * a * 6 / 100
ElseIf a = 3 Then
atlaide = b * a * 9 / 100
ElseIf a = 4 Then
atlaide = b * a * 12 / 100
ElseIf a >= 5 Then
atlaide = b * a * 15 / 100
Else
atlaide = "Kluda!"
End If
End Function
Public Function palidzibas_d(a, b)
If a > 15 And b = True Then
palidzibas_d = 250
Else
palidzibas_d = 0
End If
End Function
Public Function jauna_bil(a, b, c, d, e)
jauna_bil = a - b - c - d + e
End Function
