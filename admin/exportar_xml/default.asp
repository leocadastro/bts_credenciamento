<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))

strURL = "default.asp?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 20
	qtde = 20
End If
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar = 	"Select " &_
					"	* " &_
					"From Eventos " &_
					"Order by ID_Evento"
	'response.write(SQL_Listar)

          'response.write session("admin_id_usuario")

  SQL_Contagem = "Select ID_Admin, count(ID_Admin) as Valor From Administradores_Edicoes where ID_Admin = " & session("admin_id_usuario") & " Group By ID_Admin"
  'response.Write(SQL_Contagem)
  Set RS_Conta = Server.CreateObject("ADODB.Recordset")
  RS_Conta.Open SQL_Contagem, Conexao



  SQL_Menus = "Select ID, ID_Edicao, ID_Admin, Data From Administradores_Edicoes where ID_Admin = " & session("admin_id_usuario")
  'response.Write(SQL_Menus)
  Set RS_Menu = Server.CreateObject("ADODB.Recordset")
  RS_Menu.Open SQL_Menus, Conexao

	'response.write("<hr>" & SQL_Contagem & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
		RS_Listar.CursorLocation = 3
		RS_Listar.CacheSize = qtde_itens
		RS_Listar.PageSize = qtde_itens
	RS_Listar.Open SQL_Listar, Conexao, 3, 3
	
	If not RS_Listar.BOF and not RS_Listar.EOF Then
		TotalPaginas = RS_Listar.PageCount	
	Else 
		TotalPaginas = 0 
	End IF
'==================================================
intPageCount = TotalPaginas
Select Case acao_pagina
		Case "I" ' inicio
			intpage = 1
		Case "a" ' anterior
			intpage = pagina - 1
			if intpage < 1 then intpage = 1
		Case "p" ' proxima
			intpage = pagina + 1
			IF intpage > intPageCount Then intpage = IntPageCount
		Case "U" 'ultima
			intpage = intPageCount
		Case "n" ' numero X
			intpage = pagina
		Case Else
			intpage = 1
End Select
'==================================================
PaginaAtual = intpage
'==================================================
If paginacao = "" then
	If Request.QueryString("pg") <> "" Then
		PaginaAtual = Cint(Request.QueryString("pg"))
		intpage = PaginaAtual
	End If
End If
'==================================================
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
	$('#aviso').hide();
	<% 
	msg = Request("msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
			%>
			$('#aviso_conteudo').html('Página não permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Evento adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Evento informado já existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "erro_nao_encontrado"
			%>
			$('#aviso_conteudo').html('Erro - Evento não encontrado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "erro_nao_autorizado"
			%>
			$('#aviso_conteudo').html('Erro - Você não está autorizado à acessar o evento informado  !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Evento atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Evento ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Evento desativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
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
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	if (erros == 0) {
		document.cad.submit();	
	} else {
		$('#aviso_conteudo').html('Favor preencher corretamente os campos em destaque.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
	}
}
</script>


<%
' Níveis
' 1 - Admin
' 2 - Marketing
' 3 - Editor
erro = 0
if not RS_Conta.eof then
  if RS_Conta("Valor") <> "" and RS_Conta("Valor") > 0 then
    Quant_Menus = RS_Conta("Valor")-1
    erro = 0
  else
    Quant_Menus = 0
    erro = 1
  end if
else
  Quant_Menus = 0
  erro = 1
end if

'Dim menu()
ReDim menu(Quant_Menus)
'* Na ordem de exibição 
'  ( 0 - 1 ) 
'  ( 2 - 3 ) ...
'MODELO - Array (titulo, icone, link, permissao)
if erro = 0 then

  if not RS_Menu.eof then

    Valor_Menus = 0

    while not RS_Menu.Eof
      SQL_Evento =  "Select " &_
                    "  E.ID_Evento " &_
                    "  , E.Nome_PTB " &_
                    "  , E.Ativo  " &_
                    "  , EE.Ano " &_
                    "  , EC.Logo_Box " &_
					"  , EE.ID_Edicao " &_
                    "From Eventos as E  " &_
                    "Inner Join Eventos_Edicoes as EE ON EE.ID_Evento = E.ID_Evento " &_
                    "Inner Join Edicoes_Configuracao as EC ON EE.ID_Edicao = EC.ID_Edicao " &_
                    "where EE.ID_Edicao = " & RS_Menu("ID_Edicao")
      'response.write ("<hr>" & SQL_Evento & "<hr>")
      Set RS_EventoMenu = Server.CreateObject("ADODB.Recordset")
      RS_EventoMenu.Open SQL_Evento, Conexao

      nome_evento = ""
      If not RS_EventoMenu.BOF or not RS_EventoMenu.EOF Then 
        nome_evento = RS_EventoMenu("Ano") & " - " & RS_EventoMenu("Nome_PTB")
        logo_evento = RS_EventoMenu("logo_box")
		id_evento	= RS_EventoMenu("ID_Edicao")
      End If

      menu(Valor_Menus) = Array(nome_evento, logo_evento, "document.location='evento.asp?id=" & id_evento & "';", "1,2,4")

      RS_Menu.MoveNext
      If not RS_Menu.EOF Then Valor_Menus = Valor_Menus + 1
    wend
    RS_Menu.Close
  end if

else
  menu(0) = Array("Atualizando", "spacer", "document.location='/admin/menu.asp';", "1,2,4")
end if

%>

<script language="javascript">

// INICIO - Configuração de navegar por TECLAS
var links = new Array();
<%
For i = Lbound(menu) to Ubound(menu)
  valor = Replace(menu(i)(2),"document.location=","")
  valor = Replace(valor, "window.open","")
  valor = Replace(valor, "(","")
  valor = Replace(valor, ")","")
  %>links[<%=i+1%>] = <%=valor%><%
Next
%>
var tecla = '';
function showKeyPress(evt)
{
  tecla += String.fromCharCode(evt.charCode);
  if (tecla.length >= 2) {
    verificar_tecla();
  } else {
    setTimeout(function() {
      verificar_tecla();
    }, 400);
  }
}
function verificar_tecla() {
  for (i = 1; i < links.length; i++) {
    if (tecla.toString() == i.toString()) {
      document.location = links[i];
    } 
  }
  tecla = ''; 
}
// FIM - Configuração de navegar por TECLAS
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Eventos</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span> </div>
      <br>
      <% 'Loop Menu %>
      <table width="560" border="0" cellspacing="0" cellpadding="0">
        <% 
  colunas = 0
  For i = LBound(menu) to Ubound(menu)
    exibir = false
    'Array (titulo, icone, link, permissao)
    'Array (0     , 1    , 2   , 3        )
    'response.Write(menu(i)(0) & " / " & menu(i)(1) & " / " & menu(i)(2) & " / " & menu(i)(3) & "<br>")
    permissao_item = Split(menu(i)(3), ",")
    For p = LBound(permissao_item) to Ubound(permissao_item)
      If Cstr(Session("admin_id_perfil")) = Cstr(permissao_item(p)) Then
        exibir = true
      End If
    Next
    'response.write(colunas & "-")
    If colunas = 0 Then 
      'response.write(colunas & "-")
      colunas = colunas + 1
      %>
        <tr valign="top">
          <%
    End If

    If exibir = true Then 
      If colunas = 1 or colunas = 3 Then
      'response.write(colunas & "-")
      colunas = colunas + 1
      %>
          <td><table width="271" border="0" cellspacing="0" cellpadding="0" background="/admin/images/bts/fundo_bts_menu.gif" class="cursor" onClick="<%=menu(i)(2)%>">
            <tr>
              <td width="74" class="c_vermelho fs22px t_arial bold" align="center"><%=i + 1%></td>            
              <% If menu(i)(1) <> "spacer" Then%>
                <td height="54" class="bt_menu_titulo_home fs12px" align="center">
                  <img src="<%=menu(i)(1)%>" title="<%=menu(i)(0)%>" alt="<%=menu(i)(0)%>" border="0">
                </td>
              <% Else %>
                <td height="54" class="bt_menu_titulo_home fs12px" style="padding-right:4px;"><%=menu(i)(0)%></td>
              <% End If %>
            </tr>
          </table></td>
          <%
      End If
    End If
    If colunas = 2 Then
      'response.write(colunas & "-")
      colunas = colunas + 1
    %>
          <td width="20" height="70">&nbsp;</td>
          <%
    ElseIf colunas = 4 Then 
      colunas = 0
    %>
        </tr>
        <%
    End If
    Next
  %>
      </table>
      <!-- ************** CONTEUDO ************** -->

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