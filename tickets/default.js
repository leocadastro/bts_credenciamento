$(document).ready(function(){
	cor_muito_clara(cor_fundo,'txt_1');
	
	// Ao terminar de carregar o documento executar:
	$('#aviso_topo').hide();
	$('#aviso').hide();
	$('#acao_remover').hide();
	
	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	$("#loading").fadeOut();
		
});

function voltar () {
	$("#conteudo").css( {"overflow" : "hidden"} );
	$("#loading").height(452).css( {"z-index": 1}).show();
	$('#' + item_atual).css( {"z-index": 2 }).hide("slide", { direction: "right" }, 700);
	$("#faixa").css( {"left" : 0} ).hide("slide", { direction: "right" }, 700);
	setTimeout(function() {
			document.location = '/status.asp?v=s';
	}, 800);
}

function show_loading(top) {
	if (top == undefined) {
		var top = 0;
	}
	$("#loading").height( $(document).height() ).css( {"background-color" : "#ccc", "top" : top}).addClass('transparent').show();	
}



function Enviar() {
	var erros = 0;
	// Limpar os Desativados
	$('select:disabled').each(function(i) { } );
	$('input:disabled').each(function(i)  { } );
	// Verificar os Ativos
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			// Regra Padrão		
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	
	$('input:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case 'acao':
				break;
			case 'frmLoginRecuperar':
				break;	
			// Regra Padrão
			default:
				if ( $('#' + this.id).attr('type') != 'checkbox' || $('#' + this.id).attr('type') != 'radio' ) {
					erros += verificar(this.id, false);
				}
				break;
		}
	});

	if (erros == 0) {
		$('#acSubmit').fadeOut();
		document.prcAcessoTicket.submit();
	} else {
		//$('#aviso_topo').hide().fadeIn().fadeOut().fadeIn();
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
//		alert('com erros');	
	}		
}

function senha() {
	var erros = 0;

	erros += verificar('frmLoginRecuperar', false);
	
	if (erros == 0) {
		show_loading();
		var timeout = setTimeout( 
			function (){
				alert('Tempo de resposta de 15 seg. excedido.\n\nFavor tentar novamente ou reiniciar seu processo.\n\nti@btsmedia.biz');
			}
		, 15000);
		$.getJSON('recuperar_senha.asp?login=' + $('#frmLoginRecuperar').val(), function(data, textStatus) {
			$("#loading").fadeOut();
			// verificar erro de retorno
			if (textStatus == 'success') {
				clearTimeout(timeout);
			}
			var msg = '';
			var msg_duvida = '<b>Aten&ccedil;&atilde;o</b>:<br>';
			var objeto = 'bt_adicionar';
			switch (data.retorno) {
				case 'login invalido':
					msg += 'login Inválido'
					break;
				case 'login nao cadastrado':
					msg += 'Login não cadastrado.'
					break;
				case 'email enviado':
					msg = data.nome  + '\nSeu C&oacute;digo de identifica&ccedil;&atilde;o foi enviado com sucesso para o e-mail: ' + data.email;
					//ok();
					break;
			}
			if (msg != '') {
				jAlert(msg, 'Aviso');
			}
		});	
	} else {
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
	}
}