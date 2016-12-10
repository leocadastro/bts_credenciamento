$(document).ready(function(){
	cor_muito_clara(cor_fundo,'txt_1');
	
	// Ao terminar de carregar o documento executar:
	$('#aviso_topo').hide();
	$('#aviso').hide();
	
	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	$("#loading").fadeOut();
		
	// Modelo Mascara
	$("#frmCNPJ").mask("99.999.999/9999-99",{placeholder:"_"});
	setTimeout(function() {
		$("#frmCNPJ").focus();
	},1500);
});

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
			// Regra Padrão
			case 'frmCNPJ':
				var valor_cnpj = $('#' + this.id).val();
					cnpj_incorreto = false;
					if (valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99' || valor_cnpj == '__.___.___/____-__') {
						cnpj_incorreto = true;
						alert("CNPJ inválido");
					}
					if (cnpj_incorreto || ValidaCNPJ(valor_cnpj) == false) {
						cnpj_incorreto = true;
					}
					if (cnpj_incorreto) {
						mudar_aviso('frmCNPJ', 'x', false);
						erros ++;
					} else {
						mudar_aviso('frmCNPJ', 'ok', false);
					}
				break;
			default:
				if ( $('#' + this.id).attr('type') != 'checkbox' || $('#' + this.id).attr('type') != 'radio' ) {
					erros += verificar(this.id, false);
				}
				break;
		}
	});

	if (erros == 0) {
		$('#acSubmit').fadeOut();
		document.prcCadEmpresa.submit();
	} else {
		//$('#aviso_topo').hide().fadeIn().fadeOut().fadeIn();
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
//		alert('com erros');	
	}		
}

function senha() {
	var erros = 0;
	var valor_cnpj = $('#frmCNPJ').val();
	var cnpj_incorreto = false;
	if (valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99' || valor_cnpj == '__.___.___/____-__') {
		cnpj_incorreto = true;
	}
	if (cnpj_incorreto || ValidaCNPJ(valor_cnpj) == false) {
		cnpj_incorreto = true;
	}
	if (cnpj_incorreto) {
		mudar_aviso('frmCNPJ', 'x', false);
		erros ++;
	} else {
		mudar_aviso('frmCNPJ', 'ok', false);
	}
	
	if (erros == 0) {
		show_loading();
		var timeout = setTimeout( 
			function (){
				alert('Tempo de resposta de 15 seg. excedido.\n\nFavor tentar novamente ou reiniciar seu processo.\n\nti@btsmedia.biz');	
			}
		, 15000);
		$.getJSON('recuperar_senha.asp?cnpj=' + $('#frmCNPJ').val(), function(data, textStatus) {
			$("#loading").fadeOut();
			// verificar erro de retorno
			if (textStatus == 'success') {
				clearTimeout(timeout);
			}
			var msg = '';
			var msg_duvida = '<b>Aten&ccedil;&atilde;o</b>:<br>';
			var objeto = 'bt_adicionar';
			switch (data.retorno) {
				case 'cnpj invalido':
					msg += 'CNPJ Inválido'
					break;
				case 'cnpj nao cadastrado':
					msg += 'CNPJ não cadastrado como Universidade no evento selecionado.'
					break;
				case 'email enviado':
					msg = data.razao  + '\nE-mail enviado com sucesso para: ' + data.email;
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