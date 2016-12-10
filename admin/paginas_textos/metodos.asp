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
Function Limpar_Texto_Esp(campo)
	limpar = campo
	limpar = Replace(limpar, "'", "&rsquo;")
	limpar = Replace(limpar, ",", "&sbquo;")
	limpar = Replace(limpar, """", "")
	Limpar_Texto_Esp = limpar
End Function

	For Each item In Request.Form
		Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
	Next

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_texto"
		id_pagina 		= Limpar_Texto(Request("id_pagina"))
		id_tipo 		= Limpar_Texto(Request("id_tipo"))
		ordem 			= Limpar_Texto(Request("ordem"))
		identificacao 	= Limpar_Texto(Request("identificacao"))
		texto_ptb 		= Limpar_Texto_Esp(Request("texto_ptb"))
		texto_eng 		= Limpar_Texto_Esp(Request("texto_eng"))
		texto_esp 		= Limpar_Texto_Esp(Request("texto_esp"))
		img_ptb			= Limpar_Texto(Request("img_ptb"))
		img_eng			= Limpar_Texto(Request("img_eng"))
		img_esp			= Limpar_Texto(Request("img_esp"))
			
		SQL_Verificar =	"Select ordem " &_
						"From Paginas_Textos " &_
						"Where " &_
						"	ordem = '" & ordem & "' " &_
						"	and id_pagina = '" & id_pagina & "' "
						
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			'Inserir em PT
			'Tratar Vazio
			If texto_ptb = "" Then texto_ptb = "NULL" Else texto_ptb = "'" & texto_ptb & "'" End If
			If img_ptb = "" or img_ptb = "/img/" Then img_ptb = "NULL" Else img_ptb = "'" & img_ptb & "'" End If
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(1," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_ptb & "," & img_ptb & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			'Inserir em EN
			'Tratar Vazio
			If texto_eng = "" Then texto_eng = "NULL" Else texto_eng = "'" & texto_eng & "'" End If
			If img_eng = "" or img_eng = "/img/" Then img_eng = "NULL" Else img_eng = "'" & img_eng & "'" End If
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(3," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_eng & "," & img_eng & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			'Inserir em ESP
			'Tratar Vazio
			If texto_esp = "" Then texto_esp = "NULL" Else texto_esp = "'" & texto_esp & "'" End If
			If img_esp = "" or img_esp = "/img/" Then img_esp = "NULL" Else img_esp = "'" & img_esp & "'" End If
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(2," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_esp & "," & img_esp & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			response.write("<br><a href='default.asp?msg=add_ok&id=" & id_pagina & "'>Voltar</a>")
			response.Redirect("default.asp?msg=add_ok&id=" & id_pagina)
		Else
			RS_Verificar.Close
			response.Redirect("default.asp?msg=add_erro_existe&id=" & id_pagina)
		End If
	'==================================================
	Case "upd_texto"
		' Campos POST 
		id_pagina 		= Limpar_Texto(Request("id_pagina"))
		id_texto_ptb	= Limpar_Texto(Request("id_texto_ptb"))
		id_texto_eng	= Limpar_Texto(Request("id_texto_eng"))
		id_texto_esp	= Limpar_Texto(Request("id_texto_esp"))
		id_tipo 		= Limpar_Texto(Request("id_tipo"))
		ordem	 		= Limpar_Texto(Request("ordem"))
		identificacao 	= Limpar_Texto(Request("identificacao"))
		texto_ptb 		= Limpar_Texto_Esp(Request("texto_ptb"))
		texto_eng 		= Limpar_Texto_Esp(Request("texto_eng"))
		texto_esp 		= Limpar_Texto_Esp(Request("texto_esp"))
		img_ptb			= Limpar_Texto(Request("img_ptb"))
		img_eng			= Limpar_Texto(Request("img_eng"))
		img_esp			= Limpar_Texto(Request("img_esp"))

		'Tratar Vazio
		If texto_ptb = "" Then texto_ptb = "NULL" Else texto_ptb = "'" & texto_ptb & "'" End If
		If img_ptb = "" or img_ptb = "/img/" Then img_ptb = "NULL" Else img_ptb = "'" & img_ptb & "'" End If

		' Se existir ID em PT atualizar
		If id_texto_ptb <> "" Then
			'Atualizar em PT
			
			SQL_Update = 	"Update Paginas_Textos " &_
							"Set " &_
							"	id_tipo = '" & id_tipo & "', " &_
							"	identificacao = '" & identificacao & "', " &_
							"	texto = " & texto_ptb & ", " &_
							"	url_imagem = " & img_ptb & ", " &_
							"	id_admin = '" & Session("admin_id_usuario") & "', " &_
							"	data_atualizacao = getDate() " &_
							"Where  " &_
							"	id_texto = " & id_texto_ptb 
			
			response.write("<hr>" & SQL_Update & "<hr>")
			
			Set RS_Update = Server.CreateObject("ADODB.Recordset")
			RS_Update.Open SQL_Update, Conexao
		' Se nao inserir
		Else
			'Inserir em ptb
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(1," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_ptb & "," & img_ptb & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
		End If
'=================================================================
		'Tratar Vazio
		If texto_eng = "" Then texto_eng = "NULL" Else texto_eng = "'" & texto_eng & "'" End If
		If img_eng = "" or img_eng = "/img/" Then img_eng = "NULL" Else img_eng = "'" & img_eng & "'" End If

		' Se existir ID em eng atualizar
		If id_texto_eng <> "" Then
			'Atualizar em eng
			
			SQL_Update = 	"Update Paginas_Textos " &_
							"Set " &_
							"	id_tipo = '" & id_tipo & "', " &_
							"	identificacao = '" & identificacao & "', " &_
							"	texto = " & texto_eng & ", " &_
							"	url_imagem = " & img_eng & ", " &_
							"	id_admin = '" & Session("admin_id_usuario") & "', " &_
							"	data_atualizacao = getDate() " &_
							"Where  " &_
							"	id_texto = " & id_texto_eng 
			
			response.write("<hr>" & SQL_Update & "<hr>")
			
			Set RS_Update = Server.CreateObject("ADODB.Recordset")
			RS_Update.Open SQL_Update, Conexao
		' Se nao inserir
		Else
			'Inserir em eng
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(3," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_eng & "," & img_eng & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
		End If
'=================================================================
		'Tratar Vazio
		If texto_esp = "" Then texto_esp = "NULL" Else texto_esp = "'" & texto_esp & "'" End If
		If img_esp = "" or img_esp = "/img/" Then img_esp = "NULL" Else img_esp = "'" & img_esp & "'" End If

		' Se existir ID em ESP atualizar
		If id_texto_esp <> "" Then
			'Atualizar em ESP
			
			SQL_Update = 	"Update Paginas_Textos " &_
							"Set " &_
							"	id_tipo = '" & id_tipo & "', " &_
							"	identificacao = '" & identificacao & "', " &_
							"	texto = " & texto_esp & ", " &_
							"	url_imagem = " & img_esp & ", " &_
							"	id_admin = '" & Session("admin_id_usuario") & "', " &_
							"	data_atualizacao = getDate() " &_
							"Where  " &_
							"	id_texto = " & id_texto_esp
			
			response.write("<hr>" & SQL_Update & "<hr>")
			
			Set RS_Update = Server.CreateObject("ADODB.Recordset")
			RS_Update.Open SQL_Update, Conexao
		' Se nao inserir
		Else
			'Inserir em ESP
			
			SQL_Inserir = 	"Insert Into Paginas_Textos " &_
							"(id_idioma, id_pagina, id_tipo, ordem, identificacao, texto, url_imagem, id_admin) " &_
							"Values " &_
							"(2," & id_pagina & "," & id_tipo & ",'" & ordem & "','" & identificacao & "'," & texto_esp & "," & img_esp & "," & Session("admin_id_usuario") & ")"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
		End If
		
		response.write("<br><a href='default.asp?msg=upd_ok&id=" & id_pagina & "'>Voltar</a>")
		response.Redirect("default.asp?msg=upd_ok&id=" & id_pagina )
	'==================================================	
	
End Select

Conexao.Close
%>