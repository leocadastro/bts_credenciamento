<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

If IsNumeric(id) = false Then response.Redirect("default.asp?msg=erro_nao_encontrado")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Evento_Autorizado = "Select " &_
							"	ID,  " &_
							"	ID_Edicao,  " &_
							"	ID_Admin,  " &_
							"	Data  " &_
							"From Administradores_Edicoes  " &_
							"Where  ID_Admin = " & session("admin_id_usuario") & " " &_
							"AND ID_Edicao = " & id

	Set RS_Evento_Autorizado = Server.CreateObject("ADODB.Recordset")
	RS_Evento_Autorizado.Open SQL_Evento_Autorizado, Conexao				

	If RS_Evento_Autorizado.BOF or RS_Evento_Autorizado.EOF Then
		response.Redirect("default.asp?msg=erro_nao_autorizado")
	Else
		RS_Evento_Autorizado.Close
	End If

	
	
'	SQL_Listar = 	"Select " &_
'					"	nome_ptb " &_
'					"From Eventos as E " &_
'					"Where ID_Evento = " & id
	
	SQL_Listar = 	"Select	  E.Nome_PTB,EE.Ano" & vbcrlf & _
					"From Eventos as E " & vbcrlf & _
					"Inner Join Eventos_Edicoes as EE" & vbcrlf & _ 
					"	ON EE.ID_Evento = E.ID_Evento " & vbcrlf & _
					"WHERE  E.Ativo='1'" & vbcrlf & _
					"AND EE.ID_Edicao = " & id
	'	response.write("<hr>" & SQL_Listar & "<hr>")
	
	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		Ano_Pasta = RS_Listar("Ano")
		nome_ptb = RS_Listar("nome_ptb")
		RS_Listar.Close
	End If
	
	SQL_Status_Exportados = 	"Select " &_
								"	Count(TC.ID_Formulario) as Total " &_
								"	,RC.Exportado " &_
								"From " &_
								"	Relacionamento_Cadastro as RC " &_
								"Inner Join " &_
								"	Tipo_Credenciamento as TC " &_
								"	ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
								"Inner Join " &_
								"	Formularios as F " &_
								"	ON F.ID_Formulario = TC.ID_Formulario " &_
								"Inner Join " &_
								"	Eventos_Edicoes as EE " &_
								"	ON EE.ID_Edicao = RC.ID_Edicao " &_
								"Inner Join " &_
								"	Eventos as EV " &_
								"	ON EV.ID_Evento = EE.ID_Evento " &_
								"Inner Join " &_
								"	Visitantes as V " &_
								"	ON V.ID_Visitante = RC.ID_Visitante " &_
								"Inner Join " &_
								"	Idiomas as I " &_  
								"	ON I.ID_Idioma = RC.ID_Idioma " &_
								"Left Join " &_
								"	Empresas as E " &_
								"	ON E.ID_Empresa = RC.ID_Empresa " &_
								"Where RC.ID_Edicao = " & ID & " " &_
								"Group By RC.Exportado"
	Set RS_Status_Exportados = Server.CreateObject("ADODB.Recordset")
	RS_Status_Exportados.Open SQL_Status_Exportados, Conexao
	
'	response.write(SQL_Status_Exportados)
  
	novos 		= 0
	exportados 	= 0

	If not RS_Status_Exportados.BOF or not RS_Status_Exportados.EOF Then
  		While not RS_Status_Exportados.EOF
			Exportado = RS_Status_Exportados("Exportado")
			If Exportado = True 	Then exportados	= RS_Status_Exportados("Total")
			If Exportado = False 	Then novos		= RS_Status_Exportados("Total")
			RS_Status_Exportados.MoveNext
		Wend
		RS_Status_Exportados.Close
	End If

	If ID = 1 Then exportados = Cint(exportados) + 288
	
	SQL_Arquivos = 	"Select " &_
					"	Arquivo " &_ 
					"	,Total " &_ 
					"	,Data_Cadastro " &_
					"From Arquivos_XML " &_ 
					"Where  " &_ 
					"	Ativo = 1 " &_ 
					"	AND ID_Edicao = " & ID & " " &_
					"	AND CHARINDEX('.xml', arquivo) > 0 " &_
					"Order by Data_Cadastro DESC "
	Set RS_Arquivos = Server.CreateObject("ADODB.Recordset")
	RS_Arquivos.Open SQL_Arquivos, Conexao
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração Cred. 2012</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#protheus').select().focus();
	$('#aviso').hide();
	<% 
	'msg = Request("msg")'
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
		%>
		$('#aviso_conteudo').html('Página não permitida !');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>
});


function voltar() {
		document.location = 'default.asp';	
}
</script>

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
    <td align="center" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100" height="50">
			<%
			ssql=	"Select	  E.ID_Evento " & VBCRLF & _
					"		, E.Nome_PTB  " & VBCRLF & _
					"		, E.Ativo  " & VBCRLF & _
					"		, EE.Ano  " & VBCRLF & _
					"		, EC.Logo_Box  " & VBCRLF & _
					"From Eventos as E  " & VBCRLF & _
					"Inner Join Eventos_Edicoes as EE  " & VBCRLF & _
					"	ON EE.ID_Evento = E.ID_Evento  " & VBCRLF & _
					"Inner Join Edicoes_Configuracao as EC  " & VBCRLF & _
					"	ON EE.ID_Edicao = EC.ID_Edicao  " & VBCRLF & _
					"where EE.ID_Edicao = " & ID
			Set RS_Imgs = Conexao.execute(ssql)
			if not RS_Imgs.eof then
			%>
		        <img src="<%=RS_Imgs("Logo_Box")%>" title="<%=RS_Imgs("Nome_PTB")%> - <%=RS_Imgs("Ano")%>" alt="<%=RS_Imgs("Nome_PTB")%> - <%=RS_Imgs("Ano")%>" border="0">
            <%
			end if
			%>
        </td>
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Evento: <%=nome_ptb%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>
      <table width="750" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td style="font-size: 12px; font-family: Arial, Helvetica">
            <strong style="font-size: 15px;">Informações:</strong><hr>
            <div style="float: left;"><%=novos%> - Registros Novos</div>
            <div style="float: right;"><%=exportados%> - Registros Exportados</div>
            <br><br>

            <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
            <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco "><a href="default.asp" style="color: #fff">Listar Eventos</a></div>
            <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>


            <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left; margin-left: 10px"></div>
            <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco "><a href="exportar.asp?id=<%=id%>" style="color: #fff">Gerar Diferencial</a></div>
            <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>

          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td style="font-size: 12px; font-family: Arial, Helvetica">
          	<hr>
            <br>
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999" class="conteudo_site">
            <%
              If RS_Arquivos.BOF or RS_Arquivos.EOF Then
                %>
                <tr>
                    <td colspan="3" bgcolor="#FFFFFF" align="center">Não foram encontrados registros !</td>
                </tr>
                <%
              ElseIf not RS_Arquivos.BOF or not RS_Arquivos.EOF Then
                %>
                  <tr>
                    <td width="50" align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>N&ordm;</strong></td>
                    <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Registros</strong></td>
                    <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Arquivo</strong></td>
                    <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Data</strong></td>
                  </tr>
                <%
                n = 0
                While not RS_Arquivos.EOF
                    n = n + 1
                    %>
                    <tr>
                        <td bgcolor="#FFFFFF" align="center"><%=n%></td>
                        <td bgcolor="#FFFFFF" align="center"><%=RS_Arquivos("total")%></td>
                        <td bgcolor="#FFFFFF" align="center"><a href="arquivos_<%=Ano_Pasta%>/<%=RS_Arquivos("arquivo")%>" target="_blank"><%=RS_Arquivos("arquivo")%></a></td>
                        <td bgcolor="#FFFFFF" align="center"><%=FormatDateTime(RS_Arquivos("data_cadastro"),2)%></td>
                    </tr>
                    <%
                    RS_Arquivos.MoveNext
                    response.Flush()
                Wend
                RS_Arquivos.Close
              End If
              %>
            </table>
          
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12" /></td>
        </tr>
    </table></td>
  </tr>
</table>
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/admin/images/img_tabela_branca_inferior.jpg" width="955" height="15" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>