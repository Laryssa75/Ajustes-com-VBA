Function Abreviar3(Texto As String) As String
    Dim Palavras() As String
    Dim i As Integer
    Dim Resultado As String
    
    Palavras = Split(Texto, " ")
    For i = LBound(Palavras) To UBound(Palavras)
        If Len(Palavras(i)) >= 3 Then
            Resultado = Resultado & Left(Palavras(i), 3) & " "
        Else
            Resultado = Resultado & Palavras(i) & " "
        End If
    Next i
    
    Abreviar3 = Trim(Resultado)
End Function

' para ler dentro do excel, basta colocar na primeira
' coluna as frases que se quer abreviar e na segunda 
' essa formula =Abreviar3(A1)