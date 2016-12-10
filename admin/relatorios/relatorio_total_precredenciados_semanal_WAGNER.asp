<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%

Server.ScriptTimeout = 9999

sub rw(s,b)
	response.Write("<textarea>"&s&"</textarea>")
	if b then
		response.End()
	end if
end sub

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

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
				If intYear Mod 100 = 0 AND intYear Mod 400 <> 0 Then
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
	fcDescMes=s
End Function

strURL = "relatorio_precredenciados.asp?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 20
	qtde = 20
End If

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

id_edicao         = Limpar_Texto(Request("id_edicao"))

If id_edicao <> "" And id_edicao <> "-" Then
	SQL_Listar =  	"Select " &_
					" Count(TC.ID_Formulario) as Total " &_
					"    ,F.ID_Formulario as ID_Formulario " &_
					"    ,F.Nome as Formulario " &_
					"   ,I.Nome as Idioma " &_
					"   ,RC.ID_Idioma as ID_Idioma " &_
					"  From  " &_
					"    Relacionamento_Cadastro as RC  " &_
					"  Inner Join  " &_
					"    Tipo_Credenciamento as TC  " &_
					"    ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento  " &_
					"  Inner Join  " &_
					"    Formularios as F  " &_
					"    ON F.ID_Formulario = TC.ID_Formulario  " &_
					"  Inner Join  " &_
					"    Eventos_Edicoes as EE  " &_
					"    ON EE.ID_Edicao = RC.ID_Edicao " &_
					"  Inner Join  " &_
					"    Eventos as EV  " &_
					"    ON EV.ID_Evento = EE.ID_Evento  " &_
					"  Inner Join  " &_
					"    Visitantes as V  " &_
					"    ON V.ID_Visitante = RC.ID_Visitante  " &_
					"  Inner Join " &_
					"    Idiomas as I " &_
					"    ON I.ID_Idioma = RC.ID_Idioma " &_
					"  Left Join  " &_
					"    Empresas as E  " &_
					"    ON E.ID_Empresa = RC.ID_Empresa  " &_
					" Where RC.ID_Edicao = " & id_edicao & " " &_
					"   Group By F.Nome, F.ID_Formulario, I.Nome, RC.ID_Idioma " &_
					"   Order by Idioma "
	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao,3,3
End If
	
%>
<html>
<head>
    <meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
    <title>Administração Cred. 2012</title>
    <link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
    <link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
    <link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">
    <script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
    <script language="javascript" src="/admin/js/validar_forms.js"></script>
    <script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
    <script language="javascript" src="/js/Calendario/calendar.js"></script>
    <script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
</head>
<script language="javascript">

function fcSubmit()
{
	$('#frm').submit();
}
function fcExcel()
{
	var frm = document.getElementById("frm");
	frm.action="relatorio_total_precredenciados_semanal_excel.asp";
	frm.method="post";
	frm.submit();
}
</script>
<body>
<!--#include virtual="/admin/inc/menu_top.asp"-->
<form id="frm" method="post" action="relatorio_total_precredenciados_semanal.asp">
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
                </table></td>
        </tr>
    </table>
    <table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td></td>
        </tr>
		<tr>
        	<td style="background-color:#FFF; text-align:center;">
            	<table style="width:500px; margin:0 auto; height:85px;" cellpadding="0" cellspacing="0">
                	<tr>
                    	<td class="bordaRedonda" style="text-align:center;">
                        	<table>
                            	<tr>
                                	<td style="text-align:center; width:98%; font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
	                                    Buscar Pr&eacute;-Credenciado
                                    </td>
                                    <td>
                                    	<img src="/admin/images/ico_excel.gif" border="0" style="cursor:pointer;" onClick="fcExcel();" title="Exportar" alt="Exportar">
                                    </td>
                                </tr>
                            </table>
                            <br>
                            <span class="titulo_menu_site_bts">Edição</span>
                            <%
                                SQL_DropDown =	"SELECT	 EE.ID_Edicao " &_
                                                "		,(Convert(varchar(10),EE.Ano)) + ' - ' + EV.Nome_PTB evento " &_
                                                "		,EE.Ano " &_
                                                "FROM Eventos EV " &_
                                                "INNER JOIN Eventos_Edicoes EE " &_
                                                "	ON EE.ID_Evento=EV.ID_Evento " &_
                                                "WHERE EV.Ativo=1 " &_
                                                "ORDER BY EE.Ano DESC,EV.Nome_PTB ASC"
                                Set RS_DropDown = Server.CreateObject("ADODB.Recordset")
                                RS_DropDown.Open SQL_DropDown, Conexao
                                selEdicao = Limpar_Texto(Request("selEdicao"))
                            %>
                            <select name="selEdicao" id="selEdicao" onChange="fcSubmit();">
                                <%
                                If Not RS_DropDown.Eof Then
                                    If Len(selEdicao) = 0 Then
                                        selectedFirst = " selected='selected' "
                                    End If
                                    %>
                                        <option <%=selectedFirst%> value="">Selecione...</option>
                                    <%
                                    While Not RS_DropDown.Eof
                                        bdSelEdicao = cInt(RS_DropDown("ID_Edicao"))
										If Len(selEdicao) > 0 Then
											nSelEdicao=cInt(selEdicao)
										End If
                                        If nselEdicao = bdSelEdicao Then
                                            selctedBd = " selected='selected' "
                                        Else
                                            selctedBd = ""
                                        End If
                                        %>
                                        <option <%=selctedBd%> value="<%=RS_DropDown("ID_Edicao")%>">
                                            <%=RS_DropDown("evento")%>
                                        </option>
                                        <%
                                    RS_DropDown.Movenext
                                    Wend
                                End If
                                %>
                            </select>
						</td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;">&nbsp;</td>
        </tr>
        <%
		If Len(selEdicao) > 0 Then
		%>
            <tr>
            	<td style="background-color:#FFF;">&nbsp;</td>
            </tr>
            <tr>
                <td style="background-color:#FFF; text-align:center;">
                    <table style="width:900px; margin:0 auto;" class="bordaRedonda">
                        <tr>
                            <td>
                                <table style="width:900px;" cellpadding="0" cellspacing="0" class="arredondamento fs11px t_arial">
                                	<tr>
                                    	<td colspan="2" class="borda_dir linha_16px" style="text-align:center; background-color:#CCC;">Período</td>
                                        <td colspan="3" class="borda_dir linha_16px" style="text-align:center; background-color:#E9E9E9">Empresa</td>
                                        <td colspan="3" class="borda_dir linha_16px" style="text-align:center; background-color:#CCC;">Entidade</td>
                                        <td colspan="3" class="borda_dir linha_16px" style="text-align:center; background-color:#E9E9E9;">Pessoa Física</td>
                                        <td class="borda_dir linha_16px" style="text-align:center; background-color:#66FFFF;">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#CCC;">Mês</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#CCC;">Semana</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Português</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Inglês</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Espanhol</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#CCC;">Português</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#CCC;">Inglês</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#CCC;">Espanhol</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Português</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Inglês</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#E9E9E9">Espanhol</td>
                                        <td class="borda_dir linha_16px cursor" style="text-align:center; background-color:#66FFFF">Totais</td>
                                    </tr>
                                    <%
									SQL_Dados =	"SELECT	 DISTINCT(datepart(ww,RC.Data_Cadastro)) As Semana " &_
												"		,Convert(int,MONTH(RC.Data_Cadastro)) As Mes " &_
												"		,YEAR(RC.Data_Cadastro) As Ano " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=1  " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Empresa_PT " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=2  " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Empresa_ESP " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=3  " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Empresa_EN " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=4  " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Entidade_PT " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=5 " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Entidade_ES " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=6 " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Entidade_EN " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=10 " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Fisica_PT " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=11 " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Fisica_ES " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento=12 " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Fisica_EN " &_
												"		,( " &_
												"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
												"			FROM Relacionamento_Cadastro EMP  " &_
												"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
												"			AND EMP.ID_Tipo_Credenciamento IN (1,2,3,4,5,6,10,11,12) " &_
												"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
												"		) AS Total " &_
												"FROM Relacionamento_Cadastro RC " &_
												"Inner Join Tipo_Credenciamento as TC " &_
												"	ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
												"Where RC.ID_Edicao =  " & selEdicao & " " &_
												"And RC.ID_Tipo_Credenciamento IN (1,2,3,4,5,6,10,11,12)"
									'response.write(SQL_Dados)
									Set RS_Dados = Server.CreateObject("ADODB.Recordset")
									RS_Dados.Open SQL_Dados, Conexao
									bFstMes = true
									If Not RS_Dados.Eof Then
										MesAnt=RS_Dados("Mes")
										corAnt = "#EAEAEA"
										While Not RS_Dados.Eof
											If corAnt="#EAEAEA" Then
												cor="#FFF"
												corB="#FFF"
											Else
												cor="#EAEAEA"
												corB="#66FFFF"
											End If
											Semana			=RS_Dados("Semana")
											Mes				=Cint(RS_Dados("Mes"))
											Ano				=RS_Dados("Ano")
											If MesAnt<>Mes Then
												i = 1
												bPrimeiraMes=false
											Else
												If bFstMes Then
													If MesAnt > 1 Then
														nUltimaSemana = getWeek((MesAnt-1),ano)
														i = (Semana - nUltimaSemana)
													else
														i=Semana
													End if
													bFstMes=False
												End If
											End If											
											Empresa_PT		=RS_Dados("Empresa_PT")
											Empresa_ESP		=RS_Dados("Empresa_ESP")
											Empresa_EN		=RS_Dados("Empresa_EN")
											Entidade_PT		=RS_Dados("Entidade_PT")
											Entidade_ES		=RS_Dados("Entidade_ES")
											Entidade_EN		=RS_Dados("Entidade_EN")
											Fisica_PT		=RS_Dados("Fisica_PT")
											Fisica_ES		=RS_Dados("Fisica_ES")
											Fisica_EN		=RS_Dados("Fisica_EN")
											DescMes 		=fcDescMes(Cint(RS_Dados("Mes")))
											Total			=RS_Dados("Total")
											

											%>
												<tr>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=DescMes%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;" nowrap><%=i%>&ordf; Semana</td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Empresa_PT%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Empresa_ESP%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Empresa_EN%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Entidade_PT%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Entidade_ES%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Entidade_EN%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Fisica_PT%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Fisica_ES%></td>
													<td class="borda_dir linha_16px" style="background-color:<%=cor%>; text-align:center;"><%=Fisica_EN%></td>
                                                    <td class="borda_dir linha_16px" style="background-color:<%=corB%>; text-align:center;"><%=Total%></td>
                                                    
												</tr>
											<%
											i=i+1
											MesAnt=Mes
											corAnt=cor
											RS_Dados.Movenext
										Wend
									End If
									%>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
            	<td style="background-color:#FFF;">&nbsp;</td>
            </tr>
		<%
		end if
		%>
    </table>
</form>
</body>
</html>