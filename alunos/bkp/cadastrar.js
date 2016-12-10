var timeout = '';
$(document).ready(function(){
	$('#atualizar-form').hide();
	$('#atualizar-linha').hide();
	$('#aviso').hide();
	$('#aviso_topo').hide();
	$('#duvida').hide();
	$('#duvida_invertida').hide();

	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	$("#loading").fadeOut();
	
	setTimeout(function() {
		$('#loading_iframe').remove();
		$('#iframe_content').append('<iframe scrolling="auto" src="credenciais_listar.asp" frameborder="0" marginheight="0" marginwidth="0" id="lista" name="lista" width="790" height="300"style="height:300px;"></iframe>');
	}, 1000);
	$('#bt_adicionar').click( function () {
		// Chegar se ao menos 1 produto possui quantidade
		var erros = 0;
		var msg = '<b>Aten&ccedil;&atilde;o</b>:<br>';
		erros = verificar('frmTipo',false);
		erros = verificar('frmNome',false);
		erros = verificar('frmCurso',false);
		if (erros > 0) {
			msg += '<br>Corrija os campos em destaque.';
			mostrar_aviso (msg, this.id, 7000);
		} else {
			show_loading();
			cadastrar();
			timeout = setTimeout( 
				function (){
					alert('Ocorreu um erro durante o processamento.');	
				}
			, 5000);
		}
	});
	$('#bt_cancelar').click( function () {
		fechar_atualizacao();
	});
	$('#bt_atualizar').click( function () {
		// Chegar se ao menos 1 produto possui quantidade
		var erros = 0;
		var msg = '<b>Aten&ccedil;&atilde;o</b>:<br>';
		erros = verificar('frmTipo2',false);
		erros = verificar('frmEditarNome', false);
		erros = verificar('frmEditarCurso', false);
		if (erros > 0) {
			msg += '<br>Corrija os campos em destaque.';
			mostrar_aviso (msg, this.id, 7000);
		} else {
			show_loading();
			atualizar();
			timeout = setTimeout( 
				function (){
					alert('Ocorreu um erro durante o processamento.');	
				}
			, 5000);
		}
	});
});

alert_msg = new Array();
alert_msg[0] = 'Credencial Cadastrada!<br>Qtde restante à preencher: ';
alert_msg[1] = 'Credencial atualizada!';
alert_msg[2] = 'Deseja remover essa credencial?';
alert_msg[3] = 'Credencial já cadastrada anteriormente.';
alert_msg[4] = 'O limite de credenciais à preencher foi atingida!';

function fechar_atualizacao () {
	$('#editar').hide();
	$('#frmID').val('');
	$('#frmTipo2').val('-');
	$('#frmEditarNome').val('');
	$('#frmEditarCurso').val('');
}
function atualizar_credencial (id, nome, curso, id_tipo) {
	$('#editar').fadeIn();
	$('#frmID').val(id);
	$('#frmTipo2').val(id_tipo);
	$('#frmEditarNome').val(nome);
	$('#frmEditarCurso').val(curso);
}
function cadastrar () {
	$.getJSON('credencial_cadastrar.asp?id_tipo=' + $('#frmTipo').val() + '&nome=' + $('#frmNome').val() + '&curso=' + $('#frmCurso').val(), function(data, textStatus) {
		$("#loading").fadeOut().height(452).css( {"z-index": 1}).hide();
		// verificar erro de retorno
		if (textStatus == 'success') {
			clearTimeout(timeout);
		}
		var msg = '';
		var msg_duvida = '<b>Aten&ccedil;&atilde;o</b>:<br>';
		var objeto = 'bt_adicionar';
		switch (data.msg) {
			case 'tipo invalido':
				msg_duvida += 'Tipo não selecionado.'
//				objeto = 'nome';
				break;
			case 'nome curto':
				msg_duvida += 'Nome muito curto.'
//				objeto = 'nome';
				break;
			case 'curso curto':
				msg_duvida += 'Curso muito curto.'
//				objeto = 'curso';
				break;
			case 'sessao expirou':
				msg = 'Sua Sess&atilde;o expirou, entre novamente no sistema';
				break;
			case 'cadastrada':
				msg = alert_msg[0] + ' ' + data.qtde_restante;
				if (data.qtde_restante == 0) {
					$('#qtde_preenchida_bg').attr('bgcolor','#990000');
				} else {
					$('#qtde_preenchida_bg').attr('bgcolor','#00CC66');					
				}
				$('#qtde_preenchida').html(data.qtde_preenchida);
				//alert(data.qtde_preenchida);
				//$('#qtde_preenchida_bg').fadeOut().fadeIn();
				top.lista.location.reload();
				$('#frmTipo').val('-');
				$('#frmNome').val('');
				$('#frmCurso').val('');
				//ok();
				break;
			case 'qtde esgotou':
				msg = alert_msg[4];
				$('#qtde_preenchida_bg').attr('bgcolor','#990000');
				break;
			case 'duplicada':
				msg = alert_msg[3];
				break;	
		}
		if (msg != '') {
			jAlert(msg, 'Aviso');
		} else if (msg_duvida != '') {
			mostrar_aviso(msg_duvida, objeto, 5000);
		}
	});	
}
function atualizar () {
	$.getJSON('credencial_atualizar.asp?id=' + $('#frmID').val() + '&id_tipo=' + $('#frmTipo2').val() + '&nome=' + $('#frmEditarNome').val() + '&curso=' + $('#frmEditarCurso').val(), function(data, textStatus) {
		$("#loading").fadeOut().height(452).css( {"z-index": 1}).hide();
		// verificar erro de retorno
		if (textStatus == 'success') {
			clearTimeout(timeout);
		}
		var msg = '';
		var msg_duvida = '<b>Aten&ccedil;&atilde;o</b>:<br>';

		var objeto = 'bt_adicionar';
		switch (data.msg) {
			case 'tipo invalido':
				msg_duvida += 'Tipo não selecionado.'
				objeto = 'frmTipo2';
				break;
			case 'nome curto':
				msg_duvida += 'Nome muito curto.'
				objeto = 'frmEditarNome';
				break;
			case 'curso curto':
				msg_duvida += 'Curso muito curto.'
				objeto = 'frmEditarCurso';
				break;
			case 'sessao expirou':
				msg = 'Sua Sess&atilde;o expirou, entre novamente no sistema';
				break;
			case 'atualizada':
				msg = 'Credencial atualizada !';
				fechar_atualizacao();
				top.lista.location.reload();
				//ok();
				break;
		}
		if (msg != '') {
			jAlert(msg, 'Aviso');
		} else if (msg_duvida != '') {
			mostrar_aviso(msg_duvida, objeto, 5000);
		}
	});	
}
function remover_credencial(id, nome, curso) {
	jConfirm(alert_msg[2] + '\n- ' + nome + '\n- ' + curso, 'Remover Credencial', function(r) {
		if (r == true) {
			show_loading();
			var timeout = setTimeout( 
				function (){
					alert('Tempo de resposta de 15 seg. excedido.\n\nFavor tentar novamente ou reiniciar seu processo.\n\nti@btsmedia.biz');	
				}
			, 15000);
			$.getJSON('credencial_remover.asp?id=' + id, function(data, textStatus) {
				var erros = 0;
				var mensagem = '';
				$("#loading").fadeOut();
				// verificar erro de retorno
				if (textStatus == 'success') {
					clearTimeout(timeout);
				}
				// retornos
				msg = alert_msg[0] + ' ' + data.qtde_restante;
				$('#qtde_preenchida').html(data.qtde_preenchida);
				//$('#qtde_preenchida_bg').fadeOut().fadeIn();
				//alert( $('#qtde_preenchida_bg').attr('bgcolor') );
				$('#qtde_preenchida_bg').attr('bgcolor','#00CC66');
				top.lista.location.reload();
			});
		}
	});
}
timeout_duvida = 0;
function show_loading(top) {
	if (top == undefined) {
		top = 0;
	}
	$("#loading").height( $(document).height() ).css( {"background-color" : "#ccc", "top" : top}).addClass('transparent').show();	
}
function mostrar_aviso (msg, objeto, tempo) {
	clearTimeout(timeout_duvida);
	hide_duvidas();
	xleft = $("#" + objeto).offset().left;
	xtop = $("#" + objeto).offset().top;
	exec_duvida(objeto, { 'left': xleft, 'top': xtop } , msg, 'false', 'texto');
	timeout_duvida = setTimeout (function () {
		$('#duvida').fadeOut();
		$('#duvida_invertida').fadeOut();
	}, tempo);
}

function Enviar() {
	var qtde = $('#qtde_preenchida').html();

	if (qtde == '0') {
		jConfirm('Você não cadastrou nenhum aluno !\nDeseja continuar mesmo assim ?', 'Aviso', function(r) {
			if (r == true) {
				$('#confirmacao').submit();
			}
		});
	} else {
		jConfirm('Você cadastrou ' + qtde + ' credencial(is) !\nDeseja continuar mesmo assim ?', 'Aviso', function(r) {
			if (r == true) {
				$('#confirmacao').submit();
			}
		});
	}
}