<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
'if session("admin_id_usuario") <> "21" then
'	response.Redirect("/admin/relatorios/default.asp")
'end if
' * Dados Paginação
evento = Request("evento")

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
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
$(document).ready(function(){
	$('#aviso').hide();
	$('#hora_ini').mask("99:99",{placeholder:"_"});

	<% 
	msg = Request("msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
			%>
			$('#aviso_conteudo').html('Página não permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Item adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "erro_nao_encontrado"
			%>
			$('#aviso_conteudo').html('Erro - Não foi encontrado nenhum registro !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Item informado já existe !');
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

function Enviar() {
	var erros = 0;
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case "id_idioma":
				break;
			case "id_tipo":
				break;
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	if (erros == 0) {
		document.buscar.submit();	
	} else {
		$('#aviso_conteudo').html('Favor preencher corretamente os campos em destaque.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
	}
}

function ExportarExcel(value){
	
	$('#box_preloader').css('display','block');
	
	$.ajax({
		url: "/admin/relatorios/relatorio_alunos_gerar.asp?v=2&evento=" + value,
		success: function(data) {
			if(data=="ERRO"){
				$('#preloader').css('display','none');
				$('#resultado').css('display','none');
				$('#erro_resultado').css('display','block');
			}else{
				$('#preloader').css('display','none');
				$('#erro_resultado').css('display','none');
				$('#resultado').css('display','block');
				$('#nome_arquivo').html(data);
				$('#link_arquivo').attr('href','/admin/relatorios/excel/' + data);
			}
		}
	});
}
</script>

<body>



<div id="box_preloader" style="background: url(/admin/images/fundo.png) repeat; height: 100%; width: 100%; position: fixed; z-index: 999999999; font-family: Arial, Helvetica, sans-serif; display: none">

    <div style="background: #FFFFFF; border-radius: 8px 8px 8px 8px; width: 250; padding: 15px; text-align: center; position: absolute; left: 40%; top: 150px;">
    
        <div id="preloader">
            <img src="/admin/images/pre-loader.gif" width="128"><br><br>
            Aguarde...<br>Gerando arquivo <strong>EXCEL</strong>!
            <br><br><br>
            <a href="#fechar" onClick="$('#box_preloader').css('display','none'); $('#preloader').css('display','block'); $('#erro_resultado').css('display','block');" style="color: #000000"><em><strong>[FECHAR]</strong></em></a>
        </div>
        
        <div id="resultado" style="font-size: 12px; text-align: left; line-height: 150%; display: none">
        	O arquivo foi GERADO com sucesso e está SALVO em nosso servidor.<br><br>
            <strong>Nome do Arquivo:</strong><br> <span id="nome_arquivo"></span><br><br>
        	<a href="" id="link_arquivo" style="color: #900; font-size: 14px;"><strong>CLIQUE AQUI</strong></a> para fazer o download do arquivo.
            <br><br><br>
            <a href="#fechar" onClick="$('#box_preloader').css('display','none'); $('#preloader').css('display','block'); $('#erro_resultado').css('display','block');" style="color: #000000"><em><strong>[FECHAR]</strong></em></a>
        </div>
        
        <div id="erro_resultado" style="font-size: 12px; text-align: left; line-height: 150%; display: none">
        	Não há Alunos cadastrados para este Evento.
            <br><br><br>
            <a href="#fechar" onClick="$('#box_preloader').css('display','none'); $('#preloader').css('display','block'); $('#erro_resultado').css('display','block');" style="color: #000000"><em><strong>[FECHAR]</strong></em></a>
        </div>
    
    </div>

</div>

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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Listagem de Ingressos</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
<div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>

      <table width="90%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td>
          
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="buscar" name="buscar" method="get" action="relatorio_alunos_gerar.asp">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Relatório de Alunos - Por Evento</span></td>
              </tr>
              <tr>
                <td width="120" height="30" class="titulo_menu_site_bts" style="width: 120px;">Eventos:</td>
                <td class="titulo_noticias_home">
                    <select id="evento" name="evento" class="admin_txtfield_login" style="width: 200px">
                        <option value="">-- Selecione um Evento --</option>
                        <%
						SQL_Eventos = 	"Select " &_
										"	Ee.ID_Edicao " &_
										"	,Ee.ID_Evento " &_
										"	,Ee.Ano " &_
										"	,Ev.Nome_PTB " &_
										"From Eventos_Edicoes As Ee " &_
										"Inner Join Eventos As Ev " &_
										"On Ee.ID_Evento = Ev.ID_Evento " &_
										"Order By Ee.Ano Desc, Ev.Nome_PTB Asc"
						Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
						RS_Eventos.Open SQL_Eventos, Conexao
					  
						If Not RS_Eventos.EOF Then
						
							While Not RS_Eventos.Eof
							
							%><option value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("Ano")%> - <%=RS_Eventos("Nome_PTB")%></option><%
							
							RS_Eventos.MoveNext
							Wend
						
						End If
						RS_Eventos.Close
						%>
                    </select>
                </td>
              </tr>
              <tr>
                <td width="120" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="ExportarExcel($('#evento').val())">Gerar Relatório</div>
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