<% Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache"
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
Id = Limpar_Texto(Request("id"))
Acao = Limpar_Texto(Request("acao"))
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_edicao"
		id_evento 	= Limpar_Texto(Request("evento"))
		ano 		= Limpar_Texto(Request("ano"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		ativo 		= Limpar_Texto(Request("ativo"))

		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio 	= "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"

		diaf = Left(data_fim, 2)
		mesf = Mid(data_fim, 4, 2)
		anof = Right(data_fim, 4)
		fim 	= "'" & anof & "-" & mesf & "-" & diaf & " " & hora_fim & ":01.000'"

		SQL_Verificar =	"Select id_evento " &_
						"From Eventos_Edicoes " &_
						"Where id_evento = '" & id_evento & "' and ano = " & ano
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Eventos_Edicoes " &_
							"(id_evento, ano, data_inicio_feira, data_fim_feira, ativo) " &_
							"Values " &_
							"(" & id_evento & "," & ano & "," & inicio & "," & fim & "," & ativo & ")"

							response.write(SQL_Inserir)

			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao

			response.Redirect("default.asp?msg=add_ok")
			response.write(SQL_Inserir)
			response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
		Else
			RS_Verificar.Close
			response.Redirect("default.asp?msg=add_erro_existe")
		End If
	'==================================================
	Case "upd_edicao"
		' Campos POST
		id_evento 	= Limpar_Texto(Request("evento"))
		ano 		= Limpar_Texto(Request("ano"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		ativo 		= Limpar_Texto(Request("ativo"))

		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio 	= "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"

		diaf = Left(data_fim, 2)
		mesf = Mid(data_fim, 4, 2)
		anof = Right(data_fim, 4)
		fim 	= "'" & anof & "-" & mesf & "-" & diaf & " " & hora_fim & ":01.000'"

		SQL_Update = 	"Update Eventos_Edicoes " &_
						"Set " &_
						"	id_evento = '" & id_evento & "', " &_
						"	ano = '" & ano & "', " &_
						"	data_inicio_feira = " & inicio & ", " &_
						"	data_fim_feira = " & fim & ", " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_edicao = " & id

		response.write("<hr>" & SQL_Update & "<hr>")

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao

		response.Redirect("default.asp?msg=upd_ok")
	'==================================================
	Case "desativar"
		id = Limpar_Texto(Request("id"))

		SQL_Update =	"Update Eventos_Edicoes " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_edicao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao

		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))

		SQL_Update =	"Update Eventos_Edicoes " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_edicao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao

		response.Redirect("default.asp?msg=atv_ok")
	'==================================================

	Case "updateLote"

	id_lote 	= Limpar_Texto(Request("id-lote"))
	data_ini 	= Limpar_Texto(Request("data_ini_lote"))
	data_fim 	= Limpar_Texto(Request("data_fim_lote"))
	valor_lote 	= Limpar_Texto(Replace(Request("val_lote"),",","."))
	serie_lote 	= Limpar_Texto(Request("serie_lote"))

	dia = Left(data_ini, 2)
	mes = Mid(data_ini, 4, 2)
	ano = Right(data_ini, 4)
	inicio 	= "'" & ano & "-" & mes & "-" & dia & "'"

	diaf = Left(data_fim, 2)
	mesf = Mid(data_fim, 4, 2)
	anof = Right(data_fim, 4)
	fim 	= "'" & anof & "-" & mesf & "-" & diaf & "'"


	SQL_Update = 	"Update Edicoes_Lote " &_
					"Set " &_
					"	Valor = " & valor_lote & ", " &_
					"	Data_Inicio = " & inicio & ", " &_
					"	Data_Fim = " & fim & ", " &_
					"	Nome = '" & serie_lote & "', " &_
					"	Ativo = 1 " &_
					"Where ID_Lote_Edicao = " & id_lote

	response.write("<hr>Lote alterado com sucesso<hr>")

	Set RS_Update = Server.CreateObject("ADODB.Recordset")
	RS_Update.Open SQL_Update, Conexao

	%>

	<script type="text/javascript">

		setTimeout(function(){
			window.location.href = '/admin/edicoes/editar.asp?id=<%=Id%>';
		},1000)

	</script>
	<%
'==================================================



	'==================================================

	Case "removeLote"

	id_lote 	= Limpar_Texto(Request("id-lote"))

	SQL_Update = 	"Update Edicoes_Lote " &_
					"Set " &_
					"	Ativo = 0 " &_
					"Where ID_Lote_Edicao = " & id_lote

	response.write("<hr>Lote removido com sucesso<hr>")

	Set RS_Update = Server.CreateObject("ADODB.Recordset")
	RS_Update.Open SQL_Update, Conexao

	%>

	<script type="text/javascript">

		setTimeout(function(){
			window.location.href = '/admin/edicoes/editar.asp?id=<%=Id%>';
		},1000)

	</script>
	<%

	'==================================================

	Case "insertLote"


	data_ini 	= Limpar_Texto(Request("data_ini_lote"))
	data_fim 	= Limpar_Texto(Request("data_fim_lote"))
	valor_lote 	= Limpar_Texto(Replace(Request("val_lote"),",","."))
	serie_lote 	= Limpar_Texto(Request("serie_lote"))

	dia = Left(data_ini, 2)
	mes = Mid(data_ini, 4, 2)
	ano = Right(data_ini, 4)
	inicio 	= "'" & ano & "-" & mes & "-" & dia & "'"

	diaf = Left(data_fim, 2)
	mesf = Mid(data_fim, 4, 2)
	anof = Right(data_fim, 4)
	fim 	= "'" & anof & "-" & mesf & "-" & diaf & "'"

	SQL_Inserir = 	"Insert Into Edicoes_Lote " &_
					"(ID_Edicao, Data_Inicio, Data_Fim, Ativo, Valor, Nome,  Data_Cadastro) " &_
					"Values " &_
					"(" & Id & "," & inicio & "," & fim & ", 1," & valor_lote & ",'" & serie_lote & "', GETDATE())"

	response.write("<hr>Lote adicionado com sucesso<hr>")

	Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
	RS_Inserir.Open SQL_Inserir, Conexao


	%>

	<script type="text/javascript">

		setTimeout(function(){
			window.location.href = '/admin/edicoes/editar.asp?id=<%=Id%>';
		},1000)

	</script>
	<%

	'==================================================



End Select

Conexao.Close
%>
