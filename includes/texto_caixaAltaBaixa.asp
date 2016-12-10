<%
Function caixaAltaBaixa(tipo, texto)
	If Len(texto) > 0 Then
		If tipo = "caixa_altabaixa" Then
			novo_texto = ""
			palavras = Split(texto)
			For pal = 0 to Ubound(palavras)
				If LCase(palavras(pal)) = "and" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "da" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "de" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "do" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "para" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "e" Then 
					novo_texto = novo_texto & LCase(palavras(pal)) & " "
				ElseIf LCase(palavras(pal)) = "me" Then 
					novo_texto = novo_texto & UCase(palavras(pal)) & " "
				Else
					If Len(palavras(pal)) > 2 Then
						novo_texto = novo_texto & Ucase(Left(palavras(pal),1)) & LCase(Right(palavras(pal), Len(palavras(pal)) -1)) & " "
					Else
						novo_texto = novo_texto & LCase(palavras(pal)) & " "
					End If
				End If
			Next
			caixaAltaBaixa = novo_texto
		End If
		If tipo = "minusculas" Then
			caixaAltaBaixa = Lcase(texto)
		End If
		If tipo = "maiusculas" Then
			caixaAltaBaixa = Ucase(texto)
		End If
	End If
End Function
Function preencher_zeros (qtde, numero) 
	preencher_zeros = numero
	For i = Len(numero)+1 To qtde
		preencher_zeros = "0" & preencher_zeros
	Next
End Function
%>