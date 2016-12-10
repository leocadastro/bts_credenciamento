// ================================================================
var idioma_atual = 1;
var edicao = '';
var tela = 'logos';
var id_configuracao_atual = '';

$(document).ready(function(){
	$("#faixa_selecionada").hide();
	$("#faixa").hide().fadeIn('slow');
	$("#feiras").hide().fadeIn('slow');
	
	idioma(1);
});
function efeito_divs (efeito) {
	eval('$("#faixa_selecionada").' + efeito + '();');
	eval('$("#faixa").' + efeito + '();');
	eval('$("#feiras").' + efeito + '();');
}
// ================================================================
function show_loading(top) {
	if (top == undefined) {
		var top = 0;
	}
	$("#loading").height( $(document).height() ).css( {"background-color" : "#ccc", "top" : top}).addClass('transparent').show();	
}
// ================================================================
function proxima_tela(id_configuracao, id_edicao, teste) {
	var avisos = 'idioma: ' + idiomas[idioma_atual];
		avisos += '\nfeira: ' + id_configuracao;
	//alert(avisos);

	// Verificar TIPOS
	show_loading();
	var timeout = setTimeout( 
		function (){
			alert('Ocorreu um erro no processamento !');	
			//\nAutomaticamente foi enviado um email com o relatorio para:\nti@btsmedia.biz;
		}
	, 5000);
	$.getJSON('/scripts/default_listar_tipos.asp?i=' + idioma_atual + '&e=' + id_edicao, function(data, textStatus) {
		
		$("#loading").fadeOut();
		// verificar erro de retorno
		if (textStatus == 'success') {
			clearTimeout(timeout);
		}
		var exibir_aviso = false;
		// Sem Idioma e Edicao
		if (data.msg == 'ids nao recebidos') {
			var msg = 'Recarregue a p&aacute;gina, dados n&atilde;o enviados.';
			exibir_aviso = true;
		}
		// Nenhum Tipo Configurado para esse Idioma e Edicao
		if (data.msg == 'sem tipos disponiveis') {
			switch (idioma_atual) {
				case 1:
					var msg = 'No momento n&atilde;o existem formul&aacute;rios dispon&iacute;veis para acesso nesta feira no idioma selecionado.';
					break;
				case 2:
					var msg = 'Actualmente no hay formularios disponibles para el llenado en la feria en el idioma seleccionado.';
					break;
				case 3:
					var msg = 'There are currently no forms available for filling at the trade show in the selected language.';
					break;
			}
			exibir_aviso = true;
			if (tela == 'tipos' || tela == 'novo_idioma') { voltar(); }
		}
		// Apos validação do Contrato
		if (data.msg == 'ok') {
			edicao = id_edicao;
			id_configuracao_atual = id_configuracao;
			
			// Exibir faixa da FEIRA
			selecionar_feira(id_configuracao);
			
			if (tela == 'novo_idioma') {
				preencher_tipos(data.itens);
				$('#tipos').css( {"z-index": 2 }).show("slide", { direction: "right" }, 500);
			} else {
				preencher_tipos(data.itens,teste);
				$('#feiras').css( {"z-index": 2 }).hide("slide", { direction: "left" }, 1000);
				$('#faixa_selecionada').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	
				// Exibir Tipos
				setTimeout(function() {
					$('#tipos').css( {"z-index": 2 }).show("slide", { direction: "right" }, 1000);
				}, 500);
			}
			tela = 'tipos';
		}
		if (exibir_aviso == true) { 
			$("#loading").fadeOut();
			jAlert(msg, 'Aviso');
			/*
			setTimeout(function() {
				jAlert(msg, 'Aviso', voltar);
			}, 1500);
			*/
		}
	});
}
// ================================================================
function voltar() {
	$('#tipos').css( {"z-index": 2 }).hide("slide", { direction: "right" }, 1000);
	$('#faixa_selecionada').css( {"z-index": 10 }).hide("slide", { direction: "right" }, 1000);
	setTimeout(function() {
		$('#feiras').css( {"z-index": 2 }).show("slide", { direction: "left" }, 1000);
	}, 500);
	// Limpar TIPOS
	setTimeout(function() {
		$('#tipos').html('');
	}, 1000);
	tela = 'logos';
}
// ================================================================
// Arrays de Objetos e Idiomas;
var idiomas = new Array();
idiomas[1] = 'pt';
idiomas[2] = 'es';
idiomas[3] = 'en';
// ================================================================
var objetos = new Array();
// [i] 			= 0 - [id] / 1 - [tipo] 
//			  	= 0 - [conteúdo] / 1 - [title] 
// Como usar = objetos[0][0][1]);
objetos[0] 	 	= Array('img_faixa_esq','img');
objetos[0][2] 	= Array('/img/geral/tipos/faixa_visitantes.gif','Visitantes');
objetos[0][3] 	= Array('/img/geral/tipos/faixa_visitantes.gif','Visitantes');
objetos[0][4] 	= Array('/img/geral/tipos/faixa_visitantes_eng.gif','Visitors');

objetos[1] 	 	= Array('img_bt_ajuda','img');
objetos[1][2] 	= Array('/img/botoes/ajuda.gif','Ajuda');
objetos[1][3] 	= Array('/img/botoes/ajuda_esp.gif','Ayuda');
objetos[1][4] 	= Array('/img/botoes/ajuda_eng.gif','Help');

objetos[2] 	 	= Array('img_tit_cabecalho','img');
objetos[2][2] 	= Array('/img/geral/nome_cabecalho.gif','Credenciamento');
objetos[2][3] 	= Array('/img/geral/nome_cabecalho_esp.gif','Acreditac&iacute;on de los Eventos');
objetos[2][4] 	= Array('/img/geral/nome_cabecalho_eng.gif','Events Registration');

objetos[3] 	 	= Array('txt_1','txt');
objetos[3][2] 	= Array('Escolha a feira em que voc&ecirc; deseja se Credenciar');
objetos[3][3] 	= Array('Elija el programa que desea acreditar');
objetos[3][4] 	= Array('Choose the trade show you want to register');

objetos[4] 	 	= Array('txt_2','txt');
objetos[4][2] 	= Array('Qual o tipo de Credencial desejada ?');
objetos[4][3] 	= Array('&iquest;Qu&eacute; tipo de credencial que desea?');
objetos[4][4] 	= Array('What type of Credential you want?');

objetos[5] 	 	= Array('tit_inicio','txt');
objetos[5][2] 	= Array('In&iacute;cio da Feira:');
objetos[5][3] 	= Array('Fecha de inicio:');
objetos[5][4] 	= Array('Start Date:');

// ================================================================
function idioma(qual) {
	var n = qual + 1;
	for (i = 0; i < objetos.length; i++) {
		// esconder objetos
		$('#' + objetos[i][0]).hide();
		if (objetos[i][1] == 'img') {
			$('#' + objetos[i][0]).attr('src',objetos[i][n][0]);
			$('#' + objetos[i][0]).attr('title',objetos[i][n][1]);
			$('#' + objetos[i][0]).attr('alt',objetos[i][n][1]);
		}
		if (objetos[i][1] == 'txt') {
			// regra para box de feiras
			if (objetos[i][0] == 'tit_inicio') {
				for (x = 0; x <= total_feiras; x ++) {
					$('#' + objetos[i][0] + x).html(objetos[i][n][0]);	
				}
			// regra geral
			} else {
				$('#' + objetos[i][0]).html(objetos[i][n][0]);	
			}
		}
		$('#' + objetos[i][0]).fadeIn();
	}
	idioma_atual = qual;
	if (tela == 'tipos') { 
		//voltar(); 
		//tela = 'logos'; 
		tela = 'novo_idioma';
		proxima_tela(id_configuracao_atual, edicao);
	}
}
// ================================================================
var configuracao_feira = new Array();
// [i] 			= 0 -[cor] / 1 - [logo] / 2 - [fundo] / 3 - [logo] 
//configuracao_feira[0] = Array('#ad2924','/img/geral/logos/formobile.gif','/img/geral/faixa_feiras/fm_fundo.gif','/img/geral/faixa_feiras/fm_logo.gif');
// ================================================================
function selecionar_feira(qual) {
	$('#img_fundo_selecionado').attr('background',configuracao_feira[qual][2]);
	$('#img_logo_selecionado').attr('src',configuracao_feira[qual][3]);
	$('#faixa_dir').css("background-image", "url(" + configuracao_feira[qual][2] + ")"); 
	
	cor_muito_clara(configuracao_feira[qual][0],'txt_2');
}
// ================================================================
function preencher_tipos(itens,teste) {
	
	var inicio 	= 	'<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">' + 
					'  <tr>' +
					'    <td height="150" colspan="3">&nbsp;</td>' +
					'  </tr>';
	var divisao	=	'  <tr>' +
					'    <td height="50" colspan="3">&nbsp;</td>' +
					'  </tr>';
	var tr	 	= 	'  <tr>';
	var td	 	= 	'    <td align="center">'
	var tipo	= 	'';
	var fim_td 	=	'    </td>';
	var fim_tr	=	'  </tr>';
	var fim 	=	'</table>';
	
	var conteudo = '';
	var lista 	= '';
	var colunas = 0;
	
	var qtde_colunas = 3;
	conteudo = inicio + tr + td;
	for (x = 0; x < itens.length; x ++ ) {	
		colunas ++;
		if (itens[x].url.substring(0,9) != '/empresa/' && itens[x].url.substring(0,10) != '/entidade/' && itens[x].url.substring(0,4) != '/pf/' && itens[x].url.substring(0,15) != '/universidades/' && itens[x].url.substring(0,8) != '/alunos/') {
			//.substring(0,1) != '/' || itens[x].url.substring(0,7) != 'http://' || itens[x].url.substring(0,3) != 'www'
			var aviso = itens[x].url;
			aviso = aviso.replace("|", "'");
			aviso = aviso.replace("|", "'");
			aviso = aviso.replace("|", "'");
			aviso = aviso.replace("|", "'");
			aviso = aviso.replace("'", "");

			aviso = "jAlert('" + aviso + "','Aviso');"
			tipo =	'		<table width="195" border="0" cellspacing="0" cellpadding="0" class="bg_feira cursor" onClick="' + aviso + '">' +
					'          <tr>' +
					'            <td height="4" bgcolor="#5a5a5a"><img src="/img/geral/spacer.gif" width="110" height="4"></td>' +
					'          </tr>' +
					'          <tr>' +
					'            <td height="55" align="center"><img src="' + itens[x].img + '" title="' + itens[x].nome + '" alt="' + itens[x].nome + '" border="0"></td>' +
					'          </tr>' +
					'          <tr>' +
					'            <td height="4" bgcolor="#5a5a5a"><img src="/img/geral/spacer.gif" width="110" height="4"></td>' +
					'         </tr>' +
					'        </table>';	
		} else {
			tipo =	'		<table width="195" border="0" cellspacing="0" cellpadding="0" class="bg_feira cursor" onClick="definir(' + idioma_atual + ',' + edicao + ',' + itens[x].tipo + ',' + itens[x].url + ',' + itens[x].formulario + ');" title="' + itens[x].nome + '" alt="' + itens[x].nome + '">' +
					'          <tr>' +
					'            <td height="4" bgcolor="#5a5a5a"><img src="/img/geral/spacer.gif" width="110" height="4"></td>' +
					'          </tr>' +
					'          <tr>' +
					'            <td height="55" align="center"><img src="' + itens[x].img + '" title="' + itens[x].nome + '" alt="' + itens[x].nome + '" border="0"></td>' +
					'          </tr>' +
					'          <tr>' +
					'            <td height="4" bgcolor="#5a5a5a"><img src="/img/geral/spacer.gif" width="110" height="4"></td>' +
					'         </tr>' +
					'        </table>';
		}
		conteudo += tipo;
		if (x < itens.length) { 
			if (colunas < qtde_colunas) { 
				conteudo += fim_td + td;
			} else if (colunas == qtde_colunas) {
				colunas = 0;
				conteudo += fim_td + fim_tr + divisao + tr + td;
			} else {
				conteudo += fim_td + fim_tr;
			}
		}
	}
	//conteudo += fim_tr + fim;
	
	conteudo += fim_tr;

	/*
	if (edicao == '46') {
		 // Mensagem de venda de Tickets ABF 2013
		conteudo += '<table class="div_parceria confirmacao cursor" align="center" style="padding: 5px;">';
		conteudo += '	<tbody>';
		conteudo += '		<tr>';
		conteudo += '			<td onClick="definir(' + idioma_atual + ',' + edicao + ',10,/tickets/,4);">';
		conteudo += '				<br><br><img src="/img/geral/1155_765x85.gif" width="765" height="85">';
		conteudo += '			</td>';
		conteudo += '		</tr>';
		conteudo += '	</tbody>';
		conteudo += '</table>';

	}
	*/
	
		//conteudo += fim;
	
	$('#tipos').html(conteudo);
}
// ================================================================

function definir (idioma, edicao, tipo, url, formulario) {
	aviso = 'idioma: ' + idiomas[idioma] +
			'\nedicao: ' + edicao +
			'\ntipo: ' + tipo +
			'\nurl: ' + url +
			'\nformulario: ' + formulario;
	
	$('#idioma').val(idioma);
	$('#edicao').val(edicao);
	$('#tipo').val(tipo);
	$('#url').val(url);
	$('#formulario').val(formulario);
	
	$('#tipos').css( {"z-index": 2 }).hide("slide", { direction: "left" }, 1000);
	show_loading();
	setTimeout (function() {
		document.continuar.submit();
	}, 200);
}