<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%

ID = Request("ID")

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar = 	"Select " &_
                "     * " &_
                "From Relacionamento_Cadastro as RC " &_
                "Where ID_Relacionamento_cadastro = " & ID & " "
	'response.write("<hr>" & SQL_Listar & "<hr>")
	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao

	If not RS_Listar.BOF or not RS_Listar.EOF Then
	
    id_tipo_credenciamento = RS_Listar("ID_Tipo_Credenciamento")
    id_edicao              = RS_Listar("ID_Edicao")
    id_idioma              = RS_Listar("ID_Idioma")
    id_visitante			     = RS_Listar("ID_Visitante")
    id_empresa 				     = RS_Listar("ID_Empresa")

    ' Verificando o ID_Formulario
    SQL_Formulario =  "Select " &_
                      " ID_Formulario " &_
                      "From " &_
                      " Tipo_Credenciamento " &_
                      "Where  " &_
                      " ID_Tipo_Credenciamento =  " & id_tipo_credenciamento
    'response.write("<hr>" & SQL_Formulario & "<hr>")
    Set RS_Formulario = Server.CreateObject("ADODB.Recordset")
    RS_Formulario.Open SQL_Formulario, Conexao

    If not RS_Formulario.BOF or not RS_Formulario.EOF Then

        ' **** Resultado
        ' ID_Formulario - Nome
        ' 1 - Empresa
        ' 2 - Entidades
        ' 3 - Imprensa
        ' 4 - Pessoa F�sica
        ' 5 - Universidades
        ' 6 - Alunos  

		'response.write("<b>Formulário:</b> " & RS_Formulario("ID_Formulario") & "<hr>")
		tipo_formulario = RS_Formulario("ID_Formulario")
        ' METODOS
        Select Case tipo_formulario
            ' =================================================
            Case "1" 
            %>
              <!--#include virtual="/admin/inc/metodo_busca_empresa.asp"-->
              <!--#include virtual="/admin/inc/metodo_busca_visitante.asp"-->
            <%
            '==================================================
            Case "2" 
            %>
              <!--#include virtual="/admin/inc/metodo_busca_empresa.asp"-->
              <!--#include virtual="/admin/inc/metodo_busca_visitante.asp"-->
            <%
            '==================================================
            Case "3" 
            %>
              <!-- Ainda nao ha Caso -->
            <%
            '==================================================
            Case "4" 
            %>
              <!--#include virtual="/admin/inc/metodo_busca_visitante.asp"-->
            <%
            '==================================================
            Case "5" 
            %>
              <!--#include virtual="/admin/inc/metodo_busca_empresa.asp"-->
              <!--#include virtual="/admin/inc/metodo_busca_visitante.asp"-->
            <%
            '==================================================
        End Select
      
      End If 

  End If		
%>
<html>
<head>
  <meta http-equiv="Content-type" content="text/html; charset=UTF-8" />
  <title>Administra&ccedil;&atilde;o Cred. 2012</title>
  <link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
  <link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
  <link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">
  <script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
  <script language="javascript" src="/admin/js/validar_forms.js"></script>
  <script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
  <script language="javascript" src="/js/Calendario/calendar.js"></script>
  <script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
</head>

<style type="text/css">
.borda {
	border-right:#666 solid 1px;
	border-bottom:#666 solid 1px;
}
</style>

<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	$('#hora_ini').mask("99:99",{placeholder:"_"});

	<% 
	msg = Request("msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
			%>
			$('#aviso_conteudo').html('P�gina n�o permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Item adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "erro_nao_encontrado"
			%>
			$('#aviso_conteudo').html('Erro - N�o foi encontrado nenhum registro !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Item informado j� existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Item atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Item ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Item desativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
	End Select
	%>
});

function EnviarConfirmacao() {

  // id_edicao, id_idioma, ID_Formulario, Email, Novo_ID_Rel_Cadastro, CPF, Nome, Cargo, Depto, CNPJ, Razao
  var answer = confirm("Tem certeza que Reenviar o E-mail de confirmação?")  
    if (answer) {
      alert("sim")
        window.location = "processar.asp?id_edicao=<%=id_edicao%>&id_idioma=<%=id_idioma%>&ID_Formulario=<%=tipo_formulario%>&email=<%=Lcase(Email)%>&id=<%=ID%>&cpf=<%=CPF%>&nome=<%=Nome_Completo%>&cargo=<%=Cargo%>&depto=<%=Depto%>&cnpj=<%=CNPJ%>&razao=<%=Razao%>"

    } else {
      alert("nao")
    }
}

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
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Detalhes de Pré-Credenciados</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/relatorios/relatorio_precredenciados.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
<div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>
<table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td>
          <% If tipo_formulario <> "4" Then %>
          <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
            <tr>
              <td width="418"> <strong>CNPJ</strong></td>
              <td>&nbsp;</td>
              <td width="418">&nbsp;</td>
            </tr>
            <tr>
              <td class="borda"><%=CNPJMask%></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><strong>Raz&atilde;o Social:</strong></td>
              <td>&nbsp;</td>
              <td><strong>Nome Fantasia</strong></td>
            </tr>
            <tr>
              <td class="borda"><%=Razao%></td>
              <td>&nbsp;</td>
              <td class="borda"><%=Fantasia%></td>
            </tr>
            <tr>
              <td><strong>Ramo de Atividade</strong></td>
              <td>&nbsp;</td>
              <td><strong>Atividade Econ&ocirc;mica</strong></td>
            </tr>
            <tr>
              <td class="borda"><%=Ramo%></td>
              <td>&nbsp;</td>
              <td class="borda"><%=Atividade%></td>
            </tr>
            <tr>
              <td><strong>Principal Produto</strong></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="borda"><%=Produto%></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><strong>Telefone</strong></td>
              <td>&nbsp;</td>
              <td class="borda"><strong><%=TituloSMS%></strong></td>
            </tr>
            <tr>
              <td class="borda"><%=TelefoneEmpresa%></td>
              <td>&nbsp;</td>
              <td class="borda"><%=RecebeSMSEmpresa%></td>
            </tr>
            <tr>
              <td><strong>Endere&ccedil;o</strong></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="borda"><%=Endereco%></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><strong>Site</strong></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="borda"><a href="<%=Lcase("http://" & Site)%>" target="_blank"><%=Site%></a></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><strong>Número de Funcionários</strong></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="borda"><%=Funcionarios_Qtde%></td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <br/><br/>
          <% End If %>
          <%'="ID_Visitante: " & ID_Visitante %>
          <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
              <tr>
                <td width="418"><strong>CPF</strong></td>
                <td>&nbsp;</td>
                <td width="418">&nbsp;</td>
              </tr>
              <tr>
                <td class="borda"><%=CPFMask%></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><strong>Nome Completo</strong></td>
                <td>&nbsp;</td>
                <td><strong>Credencial</strong></td>
              </tr>
              <tr>
                <td class="borda"><%=Nome_Completo%></td>
                <td>&nbsp;</td>
                <td class="borda"><%=Nome_Credencial%></td>
              </tr>
              <tr>
                <td><strong>Data de Nascimento</strong></td>
                <td>&nbsp;</td>
                <td><strong>Sexo</strong></td>
              </tr>
              <tr>
                <td class="borda"><%=DataMask%></td>
                <td>&nbsp;</td>
                <td class="borda"><%=Sexo%></td>
              </tr>
              <tr>
                <td><strong>Cargo</strong></td>
                <td>&nbsp;</td>
                <td><strong>Sub-Cargo</strong></td>
              </tr>
              <tr>
                <td class="borda"><%=Cargo%></td>
                <td>&nbsp;</td>
                <td class="borda"><%=SubCargo%></td>
              </tr>
              <tr>
                <td><strong>Departamento</strong></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="borda"><%=Depto%></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><strong>Telefone</strong></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="borda"><%=TelefoneVisitante%></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <% If Endereco_Visitante <> "" Then %>
              <tr>
                <td><strong>Endereço</strong></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="borda"><%=Endereco_Visitante%></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <% End If %>
              <tr>
                <td><strong>E-mail</strong></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="borda"><%=Email%></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <br/>
              <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
              <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="EnviarConfirmacao();">Reenvio de E-mail de Confirmação</div>
              <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
            <br/>
          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12"></td>
        </tr>
      </table>
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