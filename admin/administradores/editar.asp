<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Perfil = 	"Select " &_
					"	ID_Perfil, " &_
					"	Perfil " &_
					"From Administradores_Perfis " &_
					"Order by ID_Perfil"
	Set RS_Perfil = Server.CreateObject("ADODB.Recordset")
	RS_Perfil.Open SQL_Perfil, Conexao

	SQL_Listar = 	"Select " &_
					"	* " &_
					"From Administradores " &_
					"Where ID_Admin = " & id

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		id_perfil = RS_Listar("id_perfil")
		nome = RS_Listar("nome")
		departamento = RS_Listar("departamento")
		email = RS_Listar("email")
		ativo = RS_Listar("ativo")
		RS_Listar.Close
	End If
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração CSC - BTS Informa</title>
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
	'msg = Request("msg")
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

function Enviar() {
	var erros = 0;
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	$('input:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case 'nova_senha':
				break;
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
function voltar() {
	confirmacao = confirm("Os dados não foram salvos, deseja sair ?")
	if (confirmacao) {
		document.location = 'default.asp';	
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Administrador</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>
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
            <input type="hidden" id="acao" name="acao" value="upd_admin">
            <input type="hidden" id="url" name="url" value="default.asp">
            <input type="hidden" id="id" name="id" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar Usu&aacute;rio</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Perfil</td>
                <td class="titulo_noticias_home">
                <select id="perfil" name="perfil" class="admin_txtfield_login">
                <%
				If not RS_Perfil.BOF or not RS_Perfil.EOF Then
					While not RS_Perfil.EOF
						selecionado = ""
						If Cint(id_perfil) = Cint(RS_Perfil("id_perfil")) Then selecionado = " selected "
						%><option value="<%=RS_Perfil("id_perfil")%>" <%=selecionado%> ><%=RS_Perfil("perfil")%></option><%
						RS_Perfil.MoveNext
					Wend
					RS_Perfil.Close
				End If
				%>
                </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Nome</td>
                <td class="titulo_noticias_home"><input name="nome" id="nome" type="text" class="admin_txtfield_login" size="30" value="<%=nome%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Departamento</td>
                <td class="titulo_noticias_home"><input name="departamento" id="departamento" type="text" class="admin_txtfield_login" size="30" value="<%=departamento%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Email / Login</td>
                <td class="titulo_noticias_home"><input name="email" id="email" type="text" class="admin_txtfield_login" size="30" value="<%=email%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Nova Senha:</td>
                <td class="titulo_noticias_home"><input name="nova_senha" id="nova_senha" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ativo</td>
                <td class="titulo_noticias_home"><select id="ativo" name="ativo" class="admin_txtfield_login">
                  <option value="0" <% If ativo = false Then response.write("selected") %>>N&atilde;o</option>
                  <option value="1" <% If ativo = true Then response.write("selected") %>>Sim</option>
                </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_tec">Definir Acesso &agrave;s Edi&ccedil;&otilde;es:</td>
                <td class="titulo_noticias_home"><div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="document.location='relacionar_edicoes/?id=<%=id%>'">Relacionar Edi&ccedil;&otilde;es</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Atualizar Usu&aacute;rio</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
              </tr>
            </form>
          </table></td>
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