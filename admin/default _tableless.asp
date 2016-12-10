<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<%
acao = Limpar_Texto(Request("acao"))
login = Limpar_Texto(Request("txt_login"))
senha = Limpar_Texto(Request("txt_senha"))

'==================================================
If acao <> "" then
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open Application("cnn")
End If 
'==================================================
' METODOS
Function limpar_sessao_admin () 
	Session("admin_id_usuario") = ""
	Session("admin_id_perfil") = ""
	Session("admin_txt_nome") = ""
	Session("admin_txt_email") = ""
	Session("admin_id_empresa") = ""
	Session("admin_dt_inclusao") = ""
End Function
Select Case Acao
	'==============================================
	Case "login"
		SQL_Verifica = 	"SELECT " &_
						"	ID_Admin, " &_
						"	ID_Perfil, " &_
						"	Nome, " &_
						"	Email, " &_
						"	ativo " &_
						"FROM Administradores " &_
						"Where  " &_
						"	email = '" & login & "' " &_
						"	AND senha = '" & senha & "' " 
		Set RS_Verifica = Server.CreateObject("ADODB.Recordset")
		RS_Verifica.Open SQL_Verifica, Conexao
		

'response.write(SQL_Verifica & "<hr>")		

		If RS_Verifica.BOF or RS_Verifica.EOF Then
			'response.write("<p>Nao encontrado</p>")
			Session("admin_logado") = false
			msg = "login_invalido"
			limpar_sessao_admin()
		Else
			'response.write("<p>Encontrado</p>")
			If RS_Verifica("ativo") = false Then
				'response.write("<p>Desativado</p>")
				Session("admin_logado") = false
				msg = "desativado"
				limpar_sessao_admin()
			ElseIf RS_Verifica("ativo") = true Then
				'response.write("<p>Ativo</p>")
			
				SQL_Update = 	"Update Administradores " &_
								"Set visitas = visitas + 1, " &_
								"data_ultima_visita = GetDate() " &_
								"Where ID_Admin = " & RS_Verifica("ID_Admin")
				Set RS_Update = Server.CreateObject("ADODB.Recordset")
				RS_Update.Open SQL_Update, Conexao 	
			
				nome = RS_Verifica("nome")
				Session("admin_logado") = true
				Session("admin_txt_nome") = RS_Verifica("nome")
				Session("admin_txt_email") = RS_Verifica("email")
				Session("admin_id_usuario") = RS_Verifica("ID_Admin")
				Session("admin_id_perfil") = RS_Verifica("ID_Perfil")
				Session("admin_dt_inclusao") = ""
				
				If Session("admin_url") <> "" Then
					url = Session("admin_url")
					response.Redirect(url)
				Else
					response.Redirect("menu.asp")
				End If
			End If
		End IF
	'==============================================
End Select
'==================================================
%>

<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração CSC - Brazil Trade Shows</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/validar_forms.js"></script>
<noscript>
	<meta http-equiv="refresh" content="0;url=javascript_desabilitado.html" />
</noscript>
</head>

<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	<%
	If Session("admin_url") <> "" Then
		%>$('#url').show();<%
	Else
		%>$('#url').hide();<%
	End If
	%>
	$('#txt_login').focus();
	<% 
	'msg = Request("msg")
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "login_invalido"
		%>
		$('#aviso_conteudo').html('Login e/ou Senha inválidos!');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
		Case "desativado"
		%>
		$('#aviso_conteudo').html('Seu Login foi desativado, entre em contato com a empresa.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
		Case "novo_login"
		%>
		$('#aviso_conteudo').html('Por favor logue-se no sistema.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>
});
function validar() {
	var erros = 0;
	
	// Regra e Campo a ser validado
	if ($('#txt_login').val().length < 3 ) { erros++; mudar_aviso('txt_login','x',false); }
	if ($('#txt_senha').val().length < 3 ) { erros++; mudar_aviso('txt_senha','x',false); }
	if (erros > 0) {
		$('#aviso_conteudo').html('Preencha os campos em destaque');
		$('#aviso').fadeIn().animate({opacity: '+=0'}, 2000).fadeOut();
		return false;
	} else {
		return true;
	}
}
</script>

<%
'Session.Abandon()
%>

<body>
<!--#include virtual="/admin/inc/topo.asp"-->
<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="363" bgcolor="#FFFFFF"><table width="955" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="middle" align="center">
        <!--[if IE]>
        <span class="fs12px t_verdana c_vermelho bold">Para melhor visualização utilize Firefox / Chrome ou Safari</span>
        <![endif]-->      
          <table width="390" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="15" colspan="2"><table width="370" height="103" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><img src="/admin/images/img_cadeado_login.gif" width="45" height="65" /></td>
                  <td width="317" background="/admin/images/img_tabela_login.gif">
                    <form id="form" name="form" method="post" action="" onSubmit="if (!validar()) { return false }">
                      <input type="hidden" name="acao" value="login" />
                      <br />
                      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="20%" class="admin_tela_login">Login:</td>
                          <td><input name="txt_login" type="text" class="admin_txtfield_login" id="txt_login" size="30" maxlength="40" onBlur="verificar(this.id, false);" value="<%=login%>" /></td>
                          </tr>
                        <tr>
                          <td class="admin_tela_login">Senha:</td>
                          <td><input name="txt_senha" type="password" class="admin_txtfield_login" id="txt_senha" size="15" maxlength="15" onBlur="verificar(this.id, false);" value="<%=senha%>" />
                            <input type="image" src="/admin/images/bt_entrar_login.gif" width="50" height="17" border="0" align="absmiddle" /></td>
                          </tr>
                        </table>
                      </form>
                    </td>
                  </tr>
                </table></td>
              </tr>
            </table>
          <center>
            <div id="aviso" style="background-color:#FFFF00; width:400px; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso - Ocorreu um erro</span> </div>
            <br>
            <div id="url" style="background-color:#FFFF00; width:400px; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Irá retornar à página solicitada</span> </div>
            </center>
        </td>
        </tr>
    </table></td>
  </tr>
</table>
<!--#include virtual="/admin/inc/rodape.asp"-->
</body>
</html>