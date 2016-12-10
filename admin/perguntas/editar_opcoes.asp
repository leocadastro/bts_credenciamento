<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id_op	       = Limpar_Texto(Request("id"))
id_pergunta  = Limpar_Texto(Request("id_pergunta"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

' Listar Op��es
SQL_Listar =   "Select " &_
                      " ID_Opcoes " &_
                      " ,ID_Perguntas " &_
                      " ,Opcao_PTB " &_
                      " ,Opcao_ENG " &_
                      " ,Opcao_ESP " &_
                      " ,Ordem " &_
                      " ,Ativo " &_
                      " ,Data_Cadastro " &_
                      "From Perguntas_Opcoes " &_
                      "Where ID_Opcoes = " & id_op
	'response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("editar.asp?id=" & id_pergunta & "&msg=erro_nao_encontrado")
	Else
    id_perguntas  = RS_Listar("id_perguntas")
    opcoes_PTB    = RS_Listar("opcao_PTB")
		opcoes_ENG    = RS_Listar("opcao_ENG")
		opcoes_ESP    = RS_Listar("opcao_ESP")
		Ordem         = RS_Listar("Ordem")
		Ativo 	      = RS_Listar("Ativo")
		RS_Listar.Close
	End If
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administra��o Cred. 2012</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
<script language="javascript" src="/js/Calendario/calendar.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#hora_ini').mask("99:99",{placeholder:"_"});
	$('#aviso').hide();
	<% 
	
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
		%>
		$('#aviso_conteudo').html('P�gina n�o permitida !');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>
});

function Enviar() {
	var erros = 0;
	$('select:enabled').each(function(i) {
		// Se n�o for obrigat�rio
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	$('input:enabled').each(function(i) {
		// Se n�o for obrigat�rio
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	if (erros == 0) {
		document.cad.submit();	
	} else {
		$('#aviso_conteudo').html('Favor preencher corretamente os campos em destaque.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
	}
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
        <td width="100" height="50">&nbsp;</td>
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Pergunta ID: <%=id%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>
      <!-- Cadastrar e lisstar Op��es -->
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="cad" name="cad" method="post" action="metodos.asp">
            <input type="hidden" id="acao" name="acao" value="upd_opcoes">
            <input type="hidden" id="id_op" name="id_op" value="<%=id_op%>">
            <input type="hidden" id="id_pergunta" name="id_pergunta" value="<%=id_pergunta%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar Op��es</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Op��es - PTB</td>
                <td class="titulo_noticias_home"><input name="opcoes_ptb" id="opcoes_ptb" type="text" class="admin_txtfield_login" size="30" value="<%=opcoes_ptb%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Op��es - ENG</td>
                <td class="titulo_noticias_home"><input name="opcoes_eng" id="opcoes_eng" type="text" class="admin_txtfield_login" size="30" value="<%=opcoes_eng%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Op��es - ESP</td>
                <td class="titulo_noticias_home"><input name="opcoes_esp" id="opcoes_esp" type="text" class="admin_txtfield_login" size="30" value="<%=opcoes_esp%>" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Ordem</td>
                <td class="titulo_noticias_home"><input name="ordem" id="ordem" type="text" class="admin_txtfield_login" size="30" value="<%=ordem%>" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Cadastrar </div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
                </tr>
            </form>
          </table>
          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12" /></td>
        </tr>
      </table>
    <!-- Listar as Op��es Cadastradas -->
    </td>
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