
var cnpj_validado = false;
var id_empresa = '';

$(document).ready(function(){

	cor_muito_clara(cor_fundo,'txt_1');

	// Ao terminar de carregar o documento executar:
	$('#aviso').hide();
	$('#aviso_topo').hide();
	$('#RecebeSmS1').hide();
	$('#RecebeSmS2').hide();
	$('#RecebeSmSEmpresa').hide();
	$('#RamoAtividadeOutros').hide();
//	$('#Atividade').hide();
	$('#RamoOutros').hide();
	$("#SubCargo").hide();
	// $("#TitulosListaProdutosCadastrados").hide();
	// $("#ListaProdutosCadastrados").hide();

	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	
	// Modelo Mascara
	$("#frmCNPJ").mask("99.999.999/9999-99",{placeholder:"_"});
	$("#frmDtNasc").mask("99/99/9999",{placeholder:"_"});	

	if (idioma_atual == 1){
		$("#frmCPF").mask("999.999.999-99",{placeholder:"_"});
		$("#frmCEP").mask("99999-999",{placeholder:"_"});
	}

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
	TipoTelefoneEmpresa( $('#frmTipoEmpresa').val() );

 	// so aceitar numeros
 	$("#frmDDI").keypress(verificaNumero);
 	$("#frmDDD").keypress(verificaNumero);
 	$("#frmTelefone").keypress(verificaNumero);
 	$("#frmRamal").keypress(verificaNumero);
 	$("#frmDDI2").keypress(verificaNumero);
 	$("#frmDDD2").keypress(verificaNumero);
 	$("#frmTelefone2").keypress(verificaNumero);
 	$("#frmRamal2").keypress(verificaNumero);
 	$("#frmDDIEmpresa").keypress(verificaNumero);
 	$("#frmDDDEmpresa").keypress(verificaNumero);
 	$("#frmTelefoneEmpresa").keypress(verificaNumero);
 	$("#frmRamalEmpresa").keypress(verificaNumero);

	$("#frmCNPJ").keypress(function(e){
		if(e.which==13){
			getCadastroCNPJ();
		}	
	});
	// Ao Recarregar verificar CNPJ
	
	// Ao Recarregar verificar se o Ramo possui Complemento
	RamoComplemento();
	
	// Botão Interrogação do Produtos
	$("#bt_faq").click(function() {
		switch(idioma_atual) {
			case '1':
				var msg = "Utilize este campo para informar com quais produtos/serviços sua empresa trabalha.<br><br>Os produtos aqui inseridos serão listados abaixo e caso seja necessário retirar algum, basta nos informar <i><b><a href='#_' onclick='aviso_retirar_produto();'>clicando aqui</a></b></i> que entraremos em contato."
				var titulo = 'Informação';
				break;
			case '3':
				var msg = "Use this field to inform all your company's products and services.<br><br>All the products will be listed here, in case you need to change any of them, please let us know by <i><b><a href='#_' onclick='aviso_retirar_produto();'>clicking here</a></b></i>."
				var titulo = 'Information';
				break;
			case '2':
				var msg = "Utilice este campo para introducir los produtos/servicios que su empresa trabaja.<br><br>Los produtos incluídos aqui seran presentados abajo, si es necesario quitar alguno, por favor, <i><b><a href='#_' onclick='aviso_retirar_produto();'>haz clic aqui</a></b></i> para ponerse en contacto."
				var titulo = 'Information';
				break;
		}
		jAlert(msg, titulo);
	});
	
	if (idioma_atual == 1){
		if ($('#frmCNPJ').val() != '') {
			getCadastroCNPJ();
		}
	
		$('#frmCNPJ').change(function() {
			getCadastroCNPJ();
		});
		
		/* Corrigido para nao buscar duas vezes o CPF e Subcargo*/
		/* Removid Homero */
		if ($('#frmCPF').val() != '') {
			getCadastroCPF();
		}
		
		$('#frmCPF').change(function() {
			getCadastroCPF();
		});
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
	
	$("#frmFantasia").change(function() {
		$('#nome_empresa').html($("#frmFantasia").val());
	});
	
	
	// Limpar ao carregar
	$("#produtos_inserir").val('');
	$("#bt_busca_produto").click(function(){
		var conteudo_produto = $.trim($("#frmPriProdut").val());
		//alert(conteudo_produto);
		
		switch (validar_novo_produto(conteudo_produto)) {
			case true:
				//$(this).html()=$(this).html() + \"<img src=/img/forms/delete.png>\"
				//$("#ListaProdutos").prepend("<li id='produtos_adicionados_" + contador_produto + "' onMouseOut='$(this).html(li_temporario);' onMouseOver='li_temporario = $(this).html(); $(this).html($(this).html() + botao_delete)'; class='cursor'>" + conteudo_produto + "</li>").slideDown('slow');
				$("#ListaProdutos").append("<li id='produtos_adicionados_" + contador_produto + "' onclick='RemoverProduto(this.id)'>" + conteudo_produto + " <img src='/img/forms/delete.png' width='14' class='produtos_adicionados_remover'/></li>").hide().slideDown('slow');
				
				// limpando o campo de Produtos
				$("#frmPriProdut").val('');
	
				contador_produto = contador_produto + 1;
				atualizar_produtos();
			break;
			
			case 'vazio':
				jAlert('O campo esta em branco!','Cadastro de Produtos');
			break;
			
			case 'existe_antes':
				jAlert('O produto já foi cadastrado anteriormente!','Cadastro de Produtos');
			break;
			
			case 'existe_agora':
				jAlert('Você já inseriu esse produto!','Cadastro de Produtos');
			break;
		}


	});

	
});
function aviso_retirar_produto() {
	$('#produtos_alterar').val('alterar');
	switch(idioma_atual) {
		case '1':
			var msg = "Solicitação registrada.<br><br>Após finalizar seu pré-credenciamento nossa equipe receberá sua informação."
			var titulo = 'Informação';
			break;
		case '2':
			var msg = "Request registered. After finishing your pre-registration our team will receive your information."
			var titulo = 'Information';
			break;
		case '3':
			var msg = "Hemos registrado su pedido. Al finalicar su pre-subscipción nuestro equipo recebirá su information."
			var titulo = 'Information';
			break;
	}
	jAlert(msg, titulo);
}

function validar_novo_produto(produto) {
	var produto_valido = true;
	
	// Se for vazio
	if (produto.length == 0) { produto_valido = 'vazio' }
	// Conferir se já existe nos previamente cadsatrados
	for (i = 0; i < produtos_cadastrados.length; i ++) {
		if (produtos_cadastrados[i] == produto) { produto_valido = 'existe_antes' };
	}
	// Conferir se já existe nos produtos à gravar
	$('#ListaProdutos').children().each(function(index) {
		//alert( index + " : " + $(this).text() );
		if ($.trim($(this).text()) == produto) { produto_valido = 'existe_agora' };
	});
	
	return produto_valido;
}


var produtos_cadastrados	= new Array(); // Criando Array de Produtos
var ramos_cadastrados 		= new Array(); // Criando Array de Ramos

var contador_produto = 1;

function RemoverProduto(ID) {
	$('#' + ID).remove();
	atualizar_produtos();
}

function atualizar_produtos () {
	var produtos = '';
	$('#ListaProdutos').children().each(function(index) {
		//alert( index + " : " + $(this).text() );
		produtos += $.trim($(this).text()) + '; ';
	});
	
	$('#produtos_inserir').val(produtos);	
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
				if (ramos_cadastrados.length == 0) {
					if (verificar(this.id, false) == 0) {
						mudar_aviso(this.id, 'ok', false);
					
						var complemento = $('select[name="frmRamo"] option:selected').attr('complemento').toLowerCase();
						if (complemento == 'true') {
							$('#frmOptRamoComplemento').fadeIn();
							// validar outros
							if (verificar('frmOptRamoComplemento', false) == 0) {
								mudar_aviso('frmOptRamoComplemento', 'ok', false);	
							} else {
								erros ++; itens_com_erro += 'frmOptRamoComplemento; '	
								mudar_aviso('frmOptRamoComplemento', 'x', false);	
							}
						}
					} else {
						erros ++; itens_com_erro += this.id + '; '	
						mudar_aviso(this.id, 'x', false);
					}
				} else {
					mudar_aviso(this.id, 'ok', false);
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

	// Desabilitar para validar no for EACH
	if (qtde_perguntas != '') {
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
			case 'frmPriProdut':
				break;
			case 'frmOptRamoComplemento':
				break;
			case 'produtos_alterar':
				break;
			case 'produtos_inserir':
				// Se não houver produtos cadastrados anteriormente
				if (produtos_cadastrados.length == 0) {
					if ( $('#produtos_inserir').val() == '' ) {
						mudar_aviso('frmPriProdut', 'x', false);
						erros ++; itens_com_erro += this.id + '; '
					} else {
						mudar_aviso('frmPriProdut', 'ok', false);
					}
				} else {
					mudar_aviso('frmPriProdut', 'ok', false);
				}
				break;
			case 'frmCPF':
				var cpf = $.trim($("#frmCPF").val());
				cpf = cpf.replace('_','');
				if (validarCPF(cpf) == false) {
					mudar_aviso(this.id, 'x', false);
					jAlert(aviso_msg_cpf_err,aviso_titulo);
					erros ++; itens_com_erro += this.id + '; '
				} else {
					mudar_aviso(this.id, 'ok', false);
				}
				break;
			case 'frmCNPJ':
				var valor_cnpj = $('#' + this.id).val();
					cnpj_incorreto = false;
					if (valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99' || valor_cnpj == '__.___.___/____-__') {
						cnpj_incorreto = true;
					}
					if (cnpj_incorreto || ValidaCNPJ(valor_cnpj) == false) {
						cnpj_incorreto = true;
					}
					if (cnpj_incorreto) {
						mudar_aviso('frmCNPJ', 'x', false);
						jAlert(aviso_msg_err,aviso_titulo);
						erros ++; itens_com_erro += this.id + '; '
					} else {
						mudar_aviso('frmCNPJ', 'ok', false);
					}
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
			case 'frmInteresseFeira':
				retorno = $('input[name="frmInteresseFeira"]:checked').val();
				if (retorno == undefined) {
					$('#parcFeira').addClass("div_alerta");
					$('#parcFeira').removeClass("div_parceria");
					erros ++; itens_com_erro += this.id + '; '
				} else {
					$('#parcFeira').removeClass("div_alerta");
					$('#parcFeira').addClass("div_parceria");
				}
				break;	
			case 'frmInteresseAnuncio':
				retorno = $('input[name="frmInteresseAnuncio"]:checked').val();
				if (retorno == undefined) {
					$('#parcAnuncio').addClass("div_alerta");
					$('#parcAnuncio').removeClass("div_parceria");
					erros ++; itens_com_erro += this.id + '; '
				} else {
					$('#parcAnuncio').removeClass("div_alerta");
					$('#parcAnuncio').addClass("div_parceria");
				}
				break;
			case 'frmInteresseAssinatura':
				retorno = $('input[name="frmInteresseAssinatura"]:checked').val();
				if (retorno == undefined) {
					$('#parcAssis').addClass("div_alerta");
					$('#parcAssis').removeClass("div_parceria");
					erros ++; itens_com_erro += this.id + '; '
				} else {
					$('#parcAssis').removeClass("div_alerta");
					$('#parcAssis').addClass("div_parceria");
				}
				break;
			// nao validar ja esta validado acima
			case 'frmAtividadeOutros':
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

				if (t == val_t_recebido) {
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
			case 'frmRamalEmpresa':
				break;
			case 'frmRamal':
				break;
			case 'frmRamal2':
				break;
			case 'frmBairro':
				break;	
			case 'frmComplemento':
				break;	
			case 'frmSite':
				break;	
 			case 'frmDDI2':
 				break;	
	 		case 'frmDDD2': 
	 			break;
	 		case 'frmTotPerguntas': 
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
			// nao validar
			case 'frmCodigoConvite':
				break;
			// Regra Padrão
			default:
				if ( $('#' + this.id).attr('type') != 'checkbox' || $('#' + this.id).attr('type') != 'radio' ) {
					erros += verificar(this.id, false);
					}
				break;
		}
	});
	
	$('#frmNewsletter').removeAttr('style');
	$('#frmNewsletter').css({'display':'none'});
	
	$('#frmSMS').removeAttr('style');
	$('#frmSMS').css({'display':'none'});
	
	$('#frmSMS2').removeAttr('style');
	$('#frmSMS2').css({'display':'none'});
		
	if (verificar('frmEmail', false) == 0) {	
		if ($('#frmEmail').val() != $('#frmEmailConf').val()) {
			$('#frmEmailConf').addClass("formulario_alerta");
				erros ++; itens_com_erro += 'frmEmailConf; '
		} else {
			$('#frmEmailConf').removeClass("formulario_alerta");
		}
	}
	
	// Se for ABF
	var checar_senha = false;
	switch ($('#id_edicao').val()) {
		case '5': // ABF 2012
			checar_senha = true;
			break;
		case '22': // ABF 2013
			checar_senha = true;
			break;
	}
	
	// Checar Senha
	if (checar_senha) {
		//alert('checar_senha');
		if (verificar('frmSenha', false) == 0) {	
			if ($('#frmSenha').val() != $('#frmSenhaConf').val()) {
				$('#frmSenhaConf').addClass("formulario_alerta");
					erros ++; itens_com_erro += 'frmSenhaConf ; '
			} else {
				$('#frmSenhaConf').removeClass("formulario_alerta");
			}
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
		$('#aviso').hide();
		$('#aviso_topo').hide();
		show_loading();
		document.prcCadEmpresa.submit();
	} else {
		$('#aviso_topo').hide().fadeIn().fadeOut().fadeIn();
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
//		alert('com erros');	
	}
}