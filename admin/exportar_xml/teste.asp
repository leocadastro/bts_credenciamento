<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
	
	'======================================================
	SQL_4_Telefones = 	"Select " &_
						"	* " &_
						"From Vw_Telefones " &_
						"Where  " &_
						"	ID_Visitante = 190 "' & IDs(x)(3)
		
	Set RS_4_Telefones = Server.CreateObject("ADODB.RecordSet")
	RS_4_Telefones.CursorLocation = 3
	RS_4_Telefones.CursorType = 3
	RS_4_Telefones.LockType = 1
	RS_4_Telefones.Open SQL_4_Telefones, Conexao
	'======================================================
	
	'======================================================
	If not RS_4_Telefones.BOF or not RS_4_Telefones.EOF Then
		i = 0
		While not RS_4_Telefones.EOF
			i = i + 1
			RS_4_Telefones.MoveNext
		Wend
		Redim Visitante_Telefones(i)
		RS_4_Telefones.MoveFirst
	
		i = 0
		While not RS_4_Telefones.EOF
			Visitante_Tipo_Telefone = RS_4_Telefones.Fields("Tipo_Telefone").Value
			Visitante_DDI 			= RS_4_Telefones.Fields("DDI").Value
			Visitante_DDD 			= RS_4_Telefones.Fields("DDD").Value
			Visitante_Numero		= RS_4_Telefones.Fields("Numero").Value
			Visitante_Ramal 		= RS_4_Telefones.Fields("Ramal").Value
			Visitante_SMS 			= RS_4_Telefones.Fields("SMS").Value
			
			Visitante_Telefones(i) = Array(Visitante_Tipo_Telefone, Visitante_DDI, Visitante_DDD, Visitante_Numero, Visitante_Ramal, Visitante_SMS)
			RS_4_Telefones.MoveNext
			i = i + 1
		Wend					
		RS_4_Telefones.Close
		Set RS_4_Telefones = Nothing
	End If
	'======================================================

	'======================================================
	SQL_11_Perguntas = 	"Select " &_
						"	* " &_
						"From Vw_Respostas_Perguntas " &_
						"Where  " &_
						"	ID_Relacionamento_Cadastro = 270 " ' & IDs(x)(0)
		
	Set RS_11_Perguntas = Server.CreateObject("ADODB.RecordSet")
	RS_11_Perguntas.CursorLocation = 3
	RS_11_Perguntas.CursorType = 3
	RS_11_Perguntas.LockType = 1
	RS_11_Perguntas.Open SQL_11_Perguntas, Conexao
	'======================================================
	
	'======================================================
	If not RS_11_Perguntas.BOF or not RS_11_Perguntas.EOF Then
		Pergunta_OLD = ""
		qtde_perguntas = 0
		While not RS_11_Perguntas.EOF
			Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
			If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
				Pergunta_OLD = Pergunta_Atual
				qtde_perguntas = qtde_perguntas + 1
			End If
			RS_11_Perguntas.MoveNext
		Wend
		RS_11_Perguntas.MoveFirst

		Redim Perguntas_e_Respostas(qtde_perguntas)
		Pergunta_OLD = ""
		Todas_Respostas = ""
		i = 0
		While not RS_11_Perguntas.EOF
			Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
			Resposta		= RS_11_Perguntas.Fields("Resposta").Value

			If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
				Pergunta_OLD = Pergunta_Atual
				i = i + 1
				Todas_Respostas = ""
				Todas_Respostas = Todas_Respostas & Resposta
				Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
			ElseIf Trim(Pergunta_OLD) = Trim(Pergunta_Atual) Then 
				Todas_Respostas = Todas_Respostas & "; " & Resposta
				Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
			End If
			RS_11_Perguntas.MoveNext
		Wend
		RS_11_Perguntas.Close
		Set RS_11_Perguntas = Nothing
	End If
	'======================================================
	

Conexao.Close
Set Conexao = Nothing
%>
L <%=Lbound(Perguntas_e_Respostas)%><br>
U <%=Ubound(Perguntas_e_Respostas)%><br>
<hr>
<%
For x = Lbound(Perguntas_e_Respostas)+1 To Ubound(Perguntas_e_Respostas)
	response.write(x & " - " & Perguntas_e_Respostas(x)(0) & " / R.: " & Perguntas_e_Respostas(x)(1) & "<br>")
Next
%>
L <%=Lbound(Visitante_Telefones)%><br>
U <%=Ubound(Visitante_Telefones)%><br>
<hr>
<%
For x = Lbound(Visitante_Telefones) To Ubound(Visitante_Telefones)-1
	response.write(x & " - " & Visitante_Telefones(x)(0) & " / R.: " & Visitante_Telefones(x)(1) & Visitante_Telefones(x)(2) & Visitante_Telefones(x)(3) & " R " & Visitante_Telefones(x)(4) & " SMS " & Visitante_Telefones(x)(5) & "<br>")
Next
%>
