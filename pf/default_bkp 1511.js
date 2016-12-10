$(document).ready(function(){
	cor_muito_clara(cor_fundo,'txt_1');
	// Ao terminar de carregar o documento executar:
	$('#aviso').hide();
	$('#aviso_topo').hide();
	$('#RecebeSmS1').hide();
	$('#RecebeSmS2').hide();
	$('#RamoAtividadeOutros').hide();
	$('#RamoOutros').hide();
	$("#SubCargo").hide();
	
	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	
	// Modelo Mascara
	if (idioma_atual == 1){		
		$("#frmCEP").mask("99999-999",{placeholder:"_"});
		$("#frmCPF").mask("999.999.999-99",{placeholder:"_"});
	}		
	$("#frmDtNasc").mask("99/99/9999",{placeholder:"_"});	
	
	// CheckBox Personalizado
	$('input:checkbox').screwDefaultButtons({
		checked: "url(/img/forms/checkbox_on.png)",
		unchecked: "url(/img/forms/checkbox_off.png)",
		width: 21,
		height: 21
	});
	
	// Radio Personalizado
	$('input:radio').screwDefaultButtons({
		checked: "url(/img/forms/radio_on.png)",
		unchecked: "url(/img/forms/radio_off.png)",
		width: 21,
		height: 21
	});

	$("#loading").fadeOut();
	
	// verificar alterações via JS
	TipoTelefone( $('#frmTipo').val() );
	TipoTelefone2( $('#frmTipo2').val() );

 	// so aceitar numeros
 	$("#frmDDI").keypress(verificaNumero);
 	$("#frmDDD").keypress(verificaNumero);
 	$("#frmDDI2").keypress(verificaNumero);
 	$("#frmDDD2").keypress(verificaNumero);
 	$("#frmRamal").keypress(verificaNumero);
 	$("#frmRamal2").keypress(verificaNumero);
	
	// verifica o click em ramos pra verificar OUTROS
	$('input[name=frmOptRamo]').change(function() {
		// se for OUTROS
		if (this.value == -1) {
			ramoOutros(this.checked);
		}
	});
	
	// verifica se o item outros estava clicado qdo re-carregou a página
	$('input[name=frmOptRamo]').each(function() {
		// se for OUTROS
		if (this.value == -1) {
			ramoOutros(this.checked);
		}
	});

	if ($('#frmCPF').val() != '') {
		getCadastroCPF();
	}

	$("#frmCPF").keypress(function(e){
		if(e.which==13){
			getCadastroCPF();
		}
	});
	$("#frmCEP").keypress(function(e){
		if(e.which==13){
			getEndereco();
		}
	});
});
// Verificando Tipo do Telefone
function TipoTelefone(Tipo) {
	switch (Tipo) {
		case '1':
			$('#Ramal1').fadeIn();
			$('#RecebeSmS1').hide();
			break;
		case '3':
			$('#Ramal1').hide();
			$('#RecebeSmS1').fadeIn();
			break;
		default:
			$('#Ramal1').hide();
			$('#RecebeSmS1').hide();
			break;
	}
}
// Verificando Tipo do Telefone
function TipoTelefone2(Tipo) {
	switch (Tipo) {
		case '1':
			$('#Ramal2').fadeIn();
			$('#RecebeSmS2').hide();
			break;
		case '3':
			$('#Ramal2').hide();
			$('#RecebeSmS2').fadeIn();
			break;
		default:
			$('#Ramal2').hide();
			$('#RecebeSmS2').hide();
			break;
	}
}

function show_loading(top) {
	if (top == undefined) {
		var top = 0;
	}
	$("#loading").height( $(document).height() ).css( {"background-color" : "#ccc", "top" : top}).addClass('transparent').show();	
}

function Enviar(teste) {
	var erros = 0;
	var itens_com_erro = '';
	// Limpar os Desativados
	$('select:disabled').each(function(i) { } );
	$('input:disabled').each(function(i)  { } );
	// Verificar os Ativos
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case 'target':
				mudar_aviso(this.id, 'ok', false);
				break;
			case 'frmRamo':
				erros += verificar(this.id, false);
				if (verificar(this.id, false) == 0) {
					// retornou 0 está validado
					// se o valor for -1 
					texto = $("select[name=frmRamo] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'OTROS') {
						$('#RamoOutros').fadeIn();
						// validar outros
						verificar('frmOptRamoOutros', false);
					} 
				}
				break;
			case 'frmCargo':
				erros += verificar(this.id, false);
				if (verificar(this.id, false) == 0) {
					// retornou 0 está validado
					texto = $("select[name=frmCargo] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'OTROS') {
						// validar outros
						verificar('frmCargoOutros', false);
					}
				}
				break;
			case 'frmSubCargo':
				if (subcargo_existe == 'sim') {
					if ( $("#frmSubCargo").val() == '-' ) {
						erros ++; itens_com_erro += this.id + '; '	
						mudar_aviso(this.id, 'x', false);
					} else {
						mudar_aviso(this.id, 'ok', false);
					}

					erros += verificar(this.id, false);
					if (verificar(this.id, false) == 0) {
						// retornou 0 está validado
						texto = $("select[name=frmSubCargo] option:selected").text();
						nome = jQuery.trim(texto);	
						if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'OTROS') {
							// validar outros
							verificar('frmSubCargoOutros', false);
						}
					}
				}
				break;
			// validar depto
			case 'frmDepto':
				erros += verificar(this.id, false);
				if (verificar(this.id, false) == 0) {
					// retornou 0 está validado
					texto = $("select[name=frmDepto] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'OTROS') {
						// validar depto outros
						verificar('frmDeptoOutros', false);
					}
				}
				break;
			// validar Atividade
			case 'frmAtividade':
				erros += verificar(this.id, false);
				if (verificar(this.id, false) == 0) {
					// retornou 0 está validado
					texto = $("select[name=frmAtividade] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'OTROS') {
						// validar depto outros
						verificar('frmfrmAtividadeOutros', false);
					}
				}
				break;
			case 'frmTipo2':
	 			break;	
			// Regra Padrão		
			default:
				erros += verificar(this.id, false);
				break;
		}
	});

	// Desabilitar para validar no for EACH
	for (i = 1; i <= qtde_perguntas; i++) {

		//alert($('input[name=frmPergunta_'+ i + ']')[0]);
		
		$('#frmPergunta_1').removeAttr('style');
		$('#frmPergunta_1').removeAttr('display');
		$('#frmPergunta_1').addClass('noneCheckbox');
		
		//alert($('input[name=frmPergunta_'+ i + ']'));
		//$("#frmPergunta_' + i + '").attr('display','none');
		//$("input[name=frmPergunta_" + i + "]").css({'display':'none'});
		//$("input[name=frmPergunta_" + i + "]").attr('display','none');
	}

	//alert($('#frmTipo2').val());
	// Regra para verificar Segundo Telefone caso preenchido
	if ($('#frmDDI2').val() != '' || $('#frmDDD2').val() != '' || $('#frmTelefone2').val() != '' || $('#frmTipo2').val() != '-') {
		erros += verificar('frmDDI2', false);
		erros += verificar('frmDDD2', false);
		erros += verificar('frmTelefone2', false);
		erros += verificar('frmTipo2', false);
	} else {
		mudar_aviso('frmDDI2', 'ok', false);
		mudar_aviso('frmDDD2', 'ok', false);
		mudar_aviso('frmTelefone2', 'ok', false);
		mudar_aviso('frmTipo2', 'ok', false);
	}
	
	$('input:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case 'acao':
				break;
			case 'id_edicao':
				break;
			case 'id_idioma':
				break;
			case 'id_tipo':
				break;
			case 'origem_cnpj':
				break;
			case 'origem_cpf':
				break;
			case 'id_empresa':
				break;
			case 'id_visitante':
				break;
			case 'frmOptRamo':
				retorno = $('input[name="frmOptRamo"]:checked').val();
				if (retorno == undefined) {
					$('#divOptRamo').addClass("div_alerta");
					erros ++; itens_com_erro += this.id + '; '
				} else {
					$('#divOptRamo').removeClass("div_alerta");
				}
				break;	
			case 'frmInteresse':
				retorno = $('input[name="frmInteresse"]:checked').val();
				if (retorno == undefined) {
					$('#divInteresse').addClass("div_alerta");
					erros ++; itens_com_erro += this.id + '; '
				} else {
					$('#divInteresse').removeClass("div_alerta");
				}
				break;
			// nao validar ja esta validado acima
			case 'frmOptRamoOutros':
				break;
			// nao validar ja esta validado acima
			case 'frmCargoOutros':
				break;
			// nao validar ja esta validado acima
			case 'frmSubCargoOutros':
				break;		
			// nao validar ja esta validado acima
			case 'frmDeptoOutros':
				break;
			case 'frmTelefone':
				var t = $('#' + this.id).val();

				var val_t 	= t.length;
				var novo_t 	= t.substring(0,1);
				var i = 1;
				var val_t_recebido = '';

				for (i = 1; i <= val_t; i++) {
					val_t_recebido = val_t_recebido + '' + novo_t;
				};

				if (t == val_t_recebido) {
					mudar_aviso(this.id, 'x', false);
					erros ++; itens_com_erro += this.id + '; '
				} else {

					if (idioma_atual == 1){
						if (t.length < 7 || t.length > 9 || isNumber(t) == false) {
							mudar_aviso(this.id, 'x', false);
							erros ++; itens_com_erro += this.id + '; '
						} else {
							mudar_aviso(this.id, 'ok', false);
						}	
					} else {
						if (t.length < 7 || t.length > 11 || isNumber(t) == false) {
							mudar_aviso(this.id, 'x', false);
							erros ++; itens_com_erro += this.id + '; '
						} else {
							mudar_aviso(this.id, 'ok', false);
						}	
					}

				}
				break;
			case 'frmTelefone2':
				var t = $('#' + this.id).val();

				var val_t 	= t.length;
				var novo_t 	= t.substring(0,1);
				var i = 1;
				var val_t_recebido = '';

				for (i = 1; i <= val_t; i++) {
					val_t_recebido = val_t_recebido + '' + novo_t;
				};

				if (t == val_t_recebido && val_t_recebido != '') {
					mudar_aviso(this.id, 'x', false);
					erros ++; itens_com_erro += this.id + '; '
				} else {

					if (t != ""){
						if (idioma_atual == 1){
							if (t.length < 7 || t.length > 9 || isNumber(t) == false) {
								mudar_aviso(this.id, 'x', false);
								erros ++; itens_com_erro += this.id + '; '
							} else {
								mudar_aviso(this.id, 'ok', false);
							}	
						} else {
							if (t.length < 7 || t.length > 11 || isNumber(t) == false) {
								mudar_aviso(this.id, 'x', false);
								erros ++; itens_com_erro += this.id + '; '
							} else {
								mudar_aviso(this.id, 'ok', false);
							}	
						}
					}

				}
				break;
			case 'frmTelefoneEmpresa':
				var t = $('#' + this.id).val();

				var val_t 	= t.length;
				var novo_t 	= t.substring(0,1);
				var i = 1;
				var val_t_recebido = '';

				for (i = 1; i <= val_t; i++) {
					val_t_recebido = val_t_recebido + '' + novo_t;
				};

				if (t == val_t_recebido && val_t_recebido != '') {
					mudar_aviso(this.id, 'x', false);
					erros ++; itens_com_erro += this.id + '; '
				} else {

					if (t.length < 7 || t.length > 11 || isNumber(t) == false) {
						mudar_aviso(this.id, 'x', false);
						erros ++; itens_com_erro += this.id + '; '
					} else {
						mudar_aviso(this.id, 'ok', false);
					}

				}
				break;
			case 'frmRamal':
				break;
			case 'frmRamal2':
				break;
			case 'frmComplemento':
				break;
			case 'frmSite':
				break;
			case 'frmDDI2':
 				break;	
	 		case 'frmDDD2': 
	 			break;
			case 'frmDtNasc':
				var str = $("#frmDtNasc").val();
					dia = str.substring(0,2);
					mes = str.substring(3,5);
					ano = str.substring(6,10);
					
					if ((dia > 31) || (mes >12) || (ano < 1900) || (ano > 1996) ) {
						mudar_aviso('frmDtNasc', 'x', false);
						erros ++; itens_com_erro += this.id + '; '
					} else {
						mudar_aviso('frmDtNasc', 'ok', false);
					}
				break;
			case 'frmEmail':
				if (verificar(this.id, false) == 0) {
					retorno_email = valida_email_novo('frmEmail');
					//alert(retorno_email);
					if (retorno_email != '') {
						mudar_aviso('frmEmail', 'x', false);
						erros ++; itens_com_erro += this.id + '; '
					} else {
						mudar_aviso('frmEmail', 'ok', false);
					}
				}
				break;
			// Regra Padrão
			default:
				if ( $('#' + this.id).attr('type') != 'checkbox' || $('#' + this.id).attr('type') != 'radio' ) {
					erros += verificar(this.id, false);
				}
				break;
		}
	});
	//alert('qtde_perguntas: ' + qtde_perguntas);

	// Habilitar para validar
	//for (i = 1; i <= qtde_perguntas; i++) {
	//	$("frmPergunta_' + i + '").attr('disabled');
	//}
	
	
	$('#frmNewsletter').removeAttr('style');
	$('#frmNewsletter').css({'display':'none'});
	
	$('#frmSMS').removeAttr('style');
	$('#frmSMS').css({'display':'none'});
	
	$('#frmSMS2').removeAttr('style');
	$('#frmSMS2').css({'display':'none'});
		
	if (verificar('frmEmail', false) == 0) {	
		if ($('#frmEmail').val() != $('#frmEmailConf').val()) {
			$('#frmEmailConf').addClass("formulario_alerta");
				erros ++; itens_com_erro += this.id + '; '
		} else {
			$('#frmEmailConf').removeClass("formulario_alerta");
		}
	}
	for (i = 1; i <= qtde_perguntas; i++) {
		retorno = $('input[name="frmPergunta_' + i + '"]:checked').val();
			$('#frmPergunta_' + i).removeAttr('style');
			$('#frmPergunta_' + i).css({'display':'none'});
			
		if (retorno == undefined) {
			$('#divPergunta_' + i).addClass("div_alerta");
			erros ++; itens_com_erro += this.id + '; '
		} else {
			$('#divPergunta_' + i).removeClass("div_alerta");
		}
	}

	if (teste == 'teste') {
		alert(erros + '\n\n' + itens_com_erro);
	}

	if (erros == 0) {
		$('#acSubmit').fadeOut();
		document.prcCadEmpresa.submit();
	} else {
		$('#aviso_topo').hide().fadeIn().fadeOut().fadeIn();
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
//		alert('com erros');	
	}		
}