<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%

Server.ScriptTimeout=9999

Function RemoverAcentuacao(texto)
	limpar = Ucase(texto)
	If Len(limpar) <= 0 Then
	Else
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "À", "A")
		limpar = Replace(limpar, "Â", "A")
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "Ã", "A")
		limpar = Replace(limpar, "É", "E")
		limpar = Replace(limpar, "È", "E")
		limpar = Replace(limpar, "Ê", "E")
		limpar = Replace(limpar, "Í", "I")
		limpar = Replace(limpar, "Ì", "I")
		limpar = Replace(limpar, "Î", "I")
		limpar = Replace(limpar, "Ó", "O")
		limpar = Replace(limpar, "Ò", "O")
		limpar = Replace(limpar, "Ô", "O")
		limpar = Replace(limpar, "Õ", "O")
		limpar = Replace(limpar, "Ú", "U")
		limpar = Replace(limpar, "Ù", "U")
		limpar = Replace(limpar, "Û", "U")
		limpar = Replace(limpar, "Ç", "Ç")
		limpar = Replace(limpar, "&", "E")
		limpar = Replace(limpar, " ", "E")
	End If
	RemoverAcentuacao = limpar
End Function

Function getWeek(mes,ano)
	If mes < 10 Then
		mes = "0" & mes
	End If
	dt = fncLastDay(mes,ano) & "-" & mes & "-" & ano
	ret = Datepart("ww",dt)
	getWeek = ret
End Function

Function fncLastDay(intMonth,intYear)
	Dim intDay
	Select Case intMonth
		Case 1, 3, 5, 7, 8, 10, 12
			intDay = 31
		Case 4, 6, 9, 11
			intDay = 30
		Case 2
			If intYear mod 4 = 0 Then
				If intYear mod 100 = 0 AND intYear mod 400 <> 0 Then
					intDay = 28
				Else
					intDay = 29
				End If
			Else
				intDay = 28
			End If
	End Select
	fncLastDay = intDay
End Function

Function fcDescMes(s)
	Select Case s
		Case "1"
			s = "Jan"
		Case "2"
			s = "Fev"
		Case "3"
			s = "Mar"
		Case "4"
			s = "Abr"
		Case "5"
			s = "Mai"
		Case "6"
			s = "Jun"
		Case "7"
			s = "Jul"
		Case "8"
			s = "Ago"
		Case "9"
			s = "Set"
		Case "10"
			s = "Out"
		Case "11"
			s = "Nov"
		Case "12"
			s = "Dez"
	End Select
	fcDescMes = s
End Function

Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

  SQL_Eventos =   "Select " &_
          " Ee.ID_Evento, " &_
          " Ee.ID_Edicao, " &_
          " E.Nome_PTB as Evento, " &_
          " Ee.Ano " &_
          "From Eventos_Edicoes as Ee " &_
          "Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
          "Order by Ano DESC, Evento"
  
Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
RS_Eventos.Open SQL_Eventos, Conexao

' * Request BUSCA
id_edicao         = Limpar_Texto(Request("id_edicao"))

If id_edicao <> "" and id_edicao <> "-" Then


	selEdicao 	= id_edicao
	arquivo 	= "pre_credenciados_semanal"
	extensao 	= ".XLS"

	SQL_File	=	"SELECT (Nome_PTB + '_' + Convert(varchar(4),Ano)) as arquivo " &_
					"FROM Eventos_Edicoes EE " &_
					"INNER JOIN Eventos EV " &_
					"	ON EE.ID_Evento=EV.ID_Evento " &_
					"WHERE ID_Edicao=" & selEdicao
	Set RS_File = Server.CreateObject("ADODB.Recordset")
	RS_File.Open SQL_File, Conexao

	If Not RS_File.Eof Then
		feira=RS_File("arquivo")
	End If

	Filename = arquivo & "_" & RemoverAcentuacao(feira)  & extensao

	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

	' Create a filesystem object
	Dim FSO
	set FSO = server.createObject("Scripting.FileSystemObject")

	' Map the logical path to the physical system path
	Dim Filepath
	Filepath = Server.MapPath("/admin/exportar_xml/arquivos_2013/" & Filename)

	Set oFiletxt= FSO.CreateTextFile(Filepath, True)
	sPath 		= FSO.GetAbsolutePathName(Filepath)
	sFilename 	= FSO.GetFileName(sPath)

	SQL_Dados 	=	"SELECT " &_
					"	DATEPART(wk, Data_Cadastro) as Semana " &_
					"	,ID_Tipo_Credenciamento " &_
					"	,Data_Cadastro " &_
					"FROM " &_
					"	Relacionamento_Cadastro " &_
					"WHERE " &_
					"	ID_Edicao = " & selEdicao & " " &_
					"	ORDER BY Semana "
	'response.write(SQL_Dados)

	Set RS_Dados = Server.CreateObject("ADODB.Recordset")
	RS_Dados.Open SQL_Dados, Conexao,3,3

	If not RS_Dados.BOF or not RS_Dados.EOF Then

		' Selecionar Anos e Semanas
		SQL_CDados ="SELECT " &_
					"	Min(DATEPART(wk, Data_Cadastro)) 	as WeekMin " &_
					"	,Max(DATEPART(wk, Data_Cadastro)) 	as WeekMax " &_
					"	,Min(DATEPART(yy, Data_Cadastro)) 	as YearMin " &_
					"	,Max(DATEPART(yy, Data_Cadastro)) 	as YearMax " &_
					"FROM " &_
					"	Relacionamento_Cadastro " &_
					"WHERE " &_
					"	ID_Edicao = " & selEdicao
		'response.write(SQL_CDados)

		Set RS_CDados = Server.CreateObject("ADODB.Recordset")
		RS_CDados.Open SQL_CDados, Conexao,3,3

		If not RS_CDados.BOF or not RS_CDados.EOF Then

			YearMin = RS_CDados("YearMin")
			YearMax = RS_CDados("YearMax")
			WeekMin = RS_CDados("WeekMin")
			WeekMax = RS_CDados("WeekMax")

		End If
		RS_CDados.Close                
		
		If YearMin = YearMax Then
			'response.write("anos iguais")
		Else
			'response.write("anos #")
		End If


		oFiletxt.WriteLine("<table border='1' cellpadding='2' cellspacing='2'>")
		oFiletxt.WriteLine("<tr>")
		oFiletxt.WriteLine("<td>Mes</td>")
		oFiletxt.WriteLine("<td>Semana/Ano</td>")
		oFiletxt.WriteLine("<td>EmpresaPTB</td>")
		oFiletxt.WriteLine("<td>EmpresaESP</td>")
		oFiletxt.WriteLine("<td>EmpresaENG</td>")
		oFiletxt.WriteLine("<td>EntidadePTB</td>")
		oFiletxt.WriteLine("<td>EntidadeESP</td>")
		oFiletxt.WriteLine("<td>EntidadeENG</td>")
		oFiletxt.WriteLine("<td>EImprensaPTB</td>")
		oFiletxt.WriteLine("<td>EImprensaESP</td>")
		oFiletxt.WriteLine("<td>EImprensaENG</td>")
		oFiletxt.WriteLine("<td>PFisicaPTB</td>")
		oFiletxt.WriteLine("<td>PFisicaESP</td>")
		oFiletxt.WriteLine("<td>PFisicaENG</td>")
		oFiletxt.WriteLine("<td>Universidade</td>")
		oFiletxt.WriteLine("<td>Aluno</td>")
		oFiletxt.WriteLine("<td>Total Semana</td>")
		oFiletxt.WriteLine("</tr>")

		' Fazer looping de acordo com as Semanas
		For i = WeekMin to WeekMax Step 1 
			
			' Contadores de Tipo de Credenciamento
			CtEmpresaPTB	= 0
			CtEmpresaESP	= 0
			CtEmpresaENG	= 0
			CtEntidadePTB	= 0
			CtEntidadeESP	= 0
			CtEntidadeENG	= 0
			CtEImprensaPTB	= 0
			CtEImprensaESP	= 0
			CtEImprensaENG	= 0
			CtPFisicaPTB	= 0
			CtPFisicaESP	= 0
			CtPFisicaENG	= 0
			CtUniversidade 	= 0
			CtAluno			= 0
			CtTotalWeek		= 0

			SQL_WDados =	"SELECT " &_
							"	DATEPART(wk, Data_Cadastro) as Semana " &_
							"	,DATEPART(mm, Data_Cadastro) as Mes " &_
							"	,ID_Tipo_Credenciamento " &_
							"	,Data_Cadastro " &_
							"FROM " &_
							"	Relacionamento_Cadastro " &_
							"WHERE " &_
							"	ID_Edicao = " & selEdicao & " " &_
							"	AND DATEPART(wk, Data_Cadastro) = " & i
			'response.write(SQL_WDados)

			Set RS_WDados = Server.CreateObject("ADODB.Recordset")
			RS_WDados.Open SQL_WDados, Conexao,3,3

			' Fazer a contagem das Semanas
			While Not RS_WDados.EOF

				Select Case Cint(RS_WDados("ID_Tipo_Credenciamento"))
					Case 1
						CtEmpresaPTB 	= CtEmpresaPTB + 1
					Case 2
						CtEmpresaESP 	= CtEmpresaESP + 1
					Case 3
						CtEmpresaENG 	= CtEmpresaENG + 1
					Case 4
						CtEntidadePTB 	= CtEntidadePTB + 1
					Case 5
						CtEntidadeESP 	= CtEntidadeESP + 1
					Case 6
						CtEntidadeENG 	= CtEntidadeENG + 1
					Case 7
						CtEImprensaPTB 	= CtEImprensaPTB + 1
					Case 8
						CtEImprensaESP 	= CtEImprensaESP + 1
					Case 9
						CtEImprensaENG 	= CtEImprensaENG + 1
					Case 10
						CtPFisicaPTB 	= CtPFisicaPTB + 1
					Case 11
						CtPFisicaESP 	= CtPFisicaESP + 1
					Case 12
						CtPFisicaENG 	= CtPFisicaENG + 1
					Case 13
						CtUniversidade 	= CtUniversidade + 1
					Case 14
						CtAluno 		= CtAluno + 1
				End Select

				Mes 		= RS_WDados("Mes")

				RS_WDados.Movenext
			Wend
				' Total da Semana
				CtTotalWeek 		= CtEmpresaPTB + CtEmpresaESP + CtEmpresaENG + CtEntidadePTB + CtEntidadeESP + CtEntidadeENG + CtEImprensaPTB + CtEImprensaESP + CtEImprensaENG + CtPFisicaPTB + CtPFisicaESP + CtPFisicaENG + CtUniversidade + CtAluno
				CtEmpresaMesPTB 	= CtEmpresaMesPTB + CtEmpresaPTB
				CtEmpresaMesESP		= CtEmpresaMesESP + CtEmpresaESP
				CtEmpresaMesENG		= CtEmpresaMesENG + CtEmpresaENG
				CtEntidadeMesPTB	= CtEntidadeMesPTB + CtEntidadePTB
				CtEntidadeMesESP	= CtEntidadeMesESP + CtEntidadeESP
				CtEntidadeMesENG	= CtEntidadeMesENG + CtEntidadeENG
				CtEImprensaMesPTB	= CtEImprensaMesPTB + CtEImprensaPTB
				CtEImprensaMesESP	= CtEImprensaMesESP + CtEImprensaESP
				CtEImprensaMesENG	= CtEImprensaMesENG + CtEImprensaENG
				CtPFisicaMesPTB		= CtPFisicaMesPTB + CtPFisicaPTB
				CtPFisicaMesESP		= CtPFisicaMesESP + CtPFisicaESP
				CtPFisicaMesENG		= CtPFisicaMesENG + CtPFisicaENG
				CtUniversidadeMes 	= CtUniversidadeMes + CtUniversidade
				CtAlunoMes			= CtAlunoMes + CtAluno

			RS_WDados.Close()

				oFiletxt.WriteLine("<tr>")
				oFiletxt.WriteLine("<td>" & Mes & "</td>")
				oFiletxt.WriteLine("<td>" & i & "</td>")
				oFiletxt.WriteLine("<td>" & CtEmpresaPTB & "</td>")
				oFiletxt.WriteLine("<td>" & CtEmpresaESP & "</td>")
				oFiletxt.WriteLine("<td>" & CtEmpresaENG & "</td>")
				oFiletxt.WriteLine("<td>" & CtEntidadePTB & "</td>")
				oFiletxt.WriteLine("<td>" & CtEntidadeESP & "</td>")
				oFiletxt.WriteLine("<td>" & CtEntidadeENG & "</td>")
				oFiletxt.WriteLine("<td>" & CtEImprensaPTB & "</td>")
				oFiletxt.WriteLine("<td>" & CtEImprensaESP & "</td>")
				oFiletxt.WriteLine("<td>" & CtEImprensaENG & "</td>")
				oFiletxt.WriteLine("<td>" & CtPFisicaPTB & "</td>")
				oFiletxt.WriteLine("<td>" & CtPFisicaESP & "</td>")
				oFiletxt.WriteLine("<td>" & CtPFisicaENG & "</td>")
				oFiletxt.WriteLine("<td>" & CtUniversidade & "</td>")
				oFiletxt.WriteLine("<td>" & CtAluno & "</td>")
				oFiletxt.WriteLine("<td>" & CtTotalWeek & "</td>")
				oFiletxt.WriteLine("</tr>")

	Next
	RS_Dados.Close

	CtTotalGeral = CtEmpresaMesPTB + CtEmpresaMesESP + CtEmpresaMesENG + CtEntidadeMesPTB + CtEntidadeMesESP + CtEntidadeMesENG + CtEImprensaMesPTB + CtEImprensaMesESP + CtEImprensaMesENG + CtPFisicaMesPTB + CtPFisicaMesESP + CtPFisicaMesENG + CtUniversidadeMes + CtAlunoMes

	oFiletxt.WriteLine("<tr>")
	oFiletxt.WriteLine("<td>&nbsp;</td>")
	oFiletxt.WriteLine("<td>&nbsp;</td>")
	oFiletxt.WriteLine("<td>" & CtEmpresaMesPTB & "</td>")
	oFiletxt.WriteLine("<td>" & CtEmpresaMesESP & "</td>")
	oFiletxt.WriteLine("<td>" & CtEmpresaMesENG & "</td>")
	oFiletxt.WriteLine("<td>" & CtEntidadeMesPTB & "</td>")
	oFiletxt.WriteLine("<td>" & CtEntidadeMesESP & "</td>")
	oFiletxt.WriteLine("<td>" & CtEntidadeMesENG & "</td>")
	oFiletxt.WriteLine("<td>" & CtEImprensaMesPTB & "</td>")
	oFiletxt.WriteLine("<td>" & CtEImprensaMesESP & "</td>")
	oFiletxt.WriteLine("<td>" & CtEImprensaMesENG & "</td>")
	oFiletxt.WriteLine("<td>" & CtPFisicaMesPTB & "</td>")
	oFiletxt.WriteLine("<td>" & CtPFisicaMesESP & "</td>")
	oFiletxt.WriteLine("<td>" & CtPFisicaMesENG & "</td>")
	oFiletxt.WriteLine("<td>" & CtUniversidadeMes & "</td>")
	oFiletxt.WriteLine("<td>" & CtAlunoMes & "</td>")
	oFiletxt.WriteLine("<td>" & CtTotalGeral &"</td>")
	oFiletxt.WriteLine("</tr>")
	oFiletxt.WriteLine("</table>")

	Else
		response.write("nao tem conteudo")
	End If

End If


%>
<html>
<head>
    <meta content="text/html; charset=iso-8859-1" http-equiv="Content-type">
    <title>Administração Cred. 2012</title>
    <link type="text/css" rel="stylesheet" href="/admin/css/bts.css">
    <link type="text/css" rel="stylesheet" href="/admin/css/admin.css">
    <link media="screen" type="text/css" rel="stylesheet" href="/css/calendar.css">
    <script src="/js/jquery-1.3.2.min.js" language="javascript"></script>
    <script src="/admin/js/validar_forms.js" language="javascript"></script>
    <script src="/js/colorpicker/colorpicker.js" language="javascript"></script>
    <script src="/js/Calendario/calendar.js" language="javascript"></script>
    <script src="/js/jquery.maskedinput-1.2.2.min.js" language="javascript"></script>
</head>
<body>
	<!--#include virtual="/admin/inc/menu_top.asp"-->
	<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td><img src="/admin/images/img_tabela_branca_top.jpg" width="955" height="15" /></td>
        </tr>
    </table>
    <table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
            <td align="center" bgcolor="#FFFFFF">
                <table width="900" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="100" height="50">&nbsp;</td>
                        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center">
                    <span style="color: #B01D22">Listagem de Pr&eacute;-Credenciados - Semanal</span></td>
                        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
                            <a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td></td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;">&nbsp;</td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;">&nbsp;</td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="buscar" name="buscar" method="post">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Buscar Pr&eacute;-Credenciado</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts"> Edi&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login" onChange="document.location='?id_edicao='+this.value;">
                <option value="-">-- Selecione --</option>
                <%
                  If not RS_Eventos.BOF or not RS_Eventos.EOF Then
                    While not RS_Eventos.EOF
                      selecionado = ""
                      If Cstr(id_edicao) = Cstr(RS_Eventos("ID_Edicao")) Then
                        selecionado = " selected "
                      End If
                    %><option <%=selecionado%> value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("ano")%> - <%=RS_Eventos("Evento")%></option><%
                      RS_Eventos.MoveNext
                    Wend
                    RS_Eventos.Close
                  End If
                  %>
                </select>
                </td>
              </tr>
            </form>
          </table>
          </td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;">&nbsp;</td>
        </tr>
        <%
        	' Mostra quando gerar o arquivo
        	If id_edicao <> "" and id_edicao <> "-" Then

        %>
        <tr>
            <td style="background-color:#FFF; text-align:center;">
            	<table style="width:600px; height:90px; margin:0 auto;" class="bordaRedonda">
                	<tr>
                    	<td style="text-align:center;">
                            Arquivo <B><%=Filename%></B> criado com sucesso<br><br>
                            <span style="background-color:#FF0;">
                                <a href="/admin/exportar_xml/arquivos_2013/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo.</a>
                            </span>&nbsp;* Botão direito > Salvar Como
						</td>
					</tr>
				</table>
			</td>
        </tr>
        <%

        	End If

        %>
        <tr>
            <td style="background-color:#FFF;">&nbsp;</td>
        </tr>
    </table>

</body>
</html>