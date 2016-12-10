// $(window).unload( function () { alert("Bye now!"); } );
// jQuery(window).bind("beforeunload", function(){return confirm("Do you really want to close?") })


$(document).ready(function(){

	fechar('Grupo1e2');

	TipoAtividadeOn( $('#frmAtividade').val() );
	TipoCargoOn( $('#frmCargo').val() );
	TipoDeptoOn( $('#frmDepto').val() );

	 // verifica o click em ramos pra verificar OUTROS
	$('#frmRamo').change(function() {
		// se for OUTROS
		// alert ( $('select[name="frmRamo"] option:selected').text() );
		RamoComplemento();
	});
	
	// verifica se o item outros estava clicado qdo re-carregou a página
	/*$('select[name=frmRamo]').each(function() {
		// se for OUTROS
		texto = $("select[name=frmRamo] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoRamoOn(true);
		} else {
			var complemento = $('select[name="frmRamo"] option:selected').attr('complemento').toLowerCase();
			if (complemento == 'true') { TipoRamoOn(true) } else { TipoRamoOn(false) }
		}
	});
	*/


	// verifica o click em Cargo pra verificar OUTROS
	$('#frmCargo').change(function() {
		// se for OUTROS
		//alert ($("select[name=frmCargo] option:selected").text());
		texto = $("select[name=frmCargo] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoCargoOn(true);
			
		} else {
			TipoCargoOn(false);
			var valor = $("select[name=frmCargo] option:selected").val();
			if (valor != '-') {
				getSubCargo(valor);
			}
		}
	});
	// verifica se o item outros estava clicado qdo re-carregou a página
	$('select[name=frmCargo]').each(function() {
		// se for OUTROS
		texto = $("select[name=frmCargo] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoCargoOn(true);
		} else {
			var valor = $("select[name=frmCargo] option:selected").val();
			if (valor != '-') {
				getSubCargo(valor);
			}
			TipoCargoOn(false);
		}
	});

	// verifica o click em SubCargos pra verificar OUTROS
	$('#frmSubCargo').change(function() {
		// se for OUTROS
		texto = $("select[name=frmSubCargo] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoSubCargoOn(true);
		} else {
			TipoSubCargoOn(false);
		}
	});
	// verifica se o item outros estava clicado qdo re-carregou a página
	$('select[name=frmSubCargo]').each(function() {
		// se for OUTROS
		texto = $("select[name=frmSubCargo] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoSubCargoOn(true);
		} else {
			TipoSubCargoOn(false);
		}
	});

	// verifica o click em Depto pra verificar OUTROS
	$('#frmDepto').change(function() {
		// se for OUTROS
		//alert ($("select[name=frmCargo] option:selected").text());
		texto = $("select[name=frmDepto] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoDeptoOn(true);
		} else {
			TipoDeptoOn(false);
		}
	});
	// verifica se o item outros estava clicado qdo re-carregou a página
	$('select[name=frmDepto]').each(function() {
		// se for OUTROS
		texto = $("select[name=frmDepto] option:selected").text();
		nome = jQuery.trim(texto);	
		if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
			TipoDeptoOn(true);
		} else {
			TipoDeptoOn(false);
		}
	});


if (idioma_atual != "1") {
	// Verifica Pais
	$('#frmPais').change(function() {
		// se for Brasil
		texto = $("select[name=frmPais] option:selected").text();
		nome = jQuery.trim(texto);	

		if (nome == 'Brasil' || nome == 'Brazil') {

			$.getJSON("/scripts/metodo_busca_estados.asp", function(data,textStatus){   
				if (data.Resultado == '1') {
					var listItems = '<option value="-">-- ' + select + ' --</option>';
				
					for (var i = 0; i < data.Estados.length; i++) {
						listItems += "<option sigla='" + data.Estados[i].sigla + "' value='" + data.Estados[i].id + "'>" + data.Estados[i].nome + "</option>";
					}

					$("#frmEstado").hide();
					$("#frmEstado").fadeIn();
					$("#frmEstado").html(listItems);
				} 
			}); 
		} else {
			// Quando nao for Brasil Voltar para Exterior
			var listItems = '<option value="28" selected>' + msg_estado + '</option>';

			$("#frmEstado").hide();
			$("#frmEstado").fadeIn();
			$("#frmEstado").html(listItems);			
		}		
	});
}

if (idioma_atual == "1") {
	// Verifica Pais
	$('#frmPais').change(function() {
		// se for Brasil
		texto = $("select[name=frmPais] option:selected").text();
		nome = jQuery.trim(texto);	

		if (nome != 'Brasil') {

			var listItems = '<option value="28">Exterior</option>';

			$("#frmEstado").hide();
			$("#frmEstado").fadeIn();
			$("#frmEstado").html(listItems);

		} else {

			$.getJSON("/scripts/metodo_busca_estados.asp", function(data,textStatus){   
				if (data.Resultado == '1') {
					var listItems = '<option value="-">-- ' + select + ' --</option>';
				
					for (var i = 0; i < data.Estados.length; i++) {
						listItems += "<option sigla='" + data.Estados[i].sigla + "' value='" + data.Estados[i].id + "'>" + data.Estados[i].nome + "</option>";
					}

					$("#frmEstado").hide();
					$("#frmEstado").fadeIn();
					$("#frmEstado").html(listItems);
				} 
			}); 

		}
	});

}


});
// Checar Ramo se Possui Complemento
function RamoComplemento () {
	texto = $("select[name=frmRamo] option:selected").text();
	nome = jQuery.trim(texto);	
	if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
		TipoRamoOn(true);
	} else {
		var complemento = $('select[name="frmRamo"] option:selected').attr('complemento').toLowerCase();
		if (complemento == 'true') { TipoRamoOn(true) } else { TipoRamoOn(false) }
	}	
}

// Verificando campos para aceitarem somente numeros
function verificaNumero(e) {
	if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {
	return false;
	}
}

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

// Verificando Tipo do Telefone
function TipoTelefoneEmpresa(Tipo) {
	switch (Tipo) {
		case '1':
			$('#RamalEmpresa').fadeIn();
			$('#RecebeSmSEmpresa').hide();
			break;
		case '3':
			$('#RamalEmpresa').hide();
			$('#RecebeSmSEmpresa').fadeIn();
			break;
		default:
			$('#RamalEmpresa').hide();
			$('#RecebeSmSEmpresa').hide();
			break;
	}
}

// Buscar CNPJ
function getCadastroCNPJ() {  
	// Se o campo CNPJ não estiver vazio  
	var valor_cnpj = $("#frmCNPJ").val();
	$('#loading').fadeIn();

	cnpj_incorreto = false;
	if (valor_cnpj == '' || valor_cnpj == '__.___.___/____-__' || valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99' || valor_cnpj == '__.___.___/____-__') {
		cnpj_incorreto = true;
		//alert("CNPJ inválido");
		jAlert(aviso_msg_err,aviso_titulo);
	} else if (cnpj_incorreto || ValidaCNPJ(valor_cnpj) == false) {
		cnpj_incorreto = true;
		//alert("CNPJ inválido");
		jAlert(aviso_msg_err,aviso_titulo);
	} else if (cnpj_incorreto) {
		mudar_aviso('frmCNPJ', 'x', false);
		//alert("CNPJ inválido");
		jAlert(aviso_msg_err,aviso_titulo);
	} else {
		mudar_aviso('frmCNPJ', 'ok', false);

		var valor_cnpj_consultado = valor_cnpj;

		//$('#bt_busca_cnpj').fadeOut();
		$.getJSON("/scripts/metodo_busca_cnpj.asp?cnpj=" + $("#frmCNPJ").val(), function(data,textStatus){  
			//Se o resultado for igual a 1  

			if (data.Resultado == '0') {
				$("#id_empresa").val('');
				$("#origem_cnpj").val('');
				$("#frmRazao").val('');		
				$("#frmFantasia").val('');
				$("#frmSigla").val('');
				$("#frmResp").val('');
				$("#frmPriProdut").val('');
				$("#frmCEP").val('');
				$("#frmEndereco").val('');
				$("#frmNumero").val('');
				$("#frmComplemento").val('');
				$("#frmBairro").val('');
				$("#frmCidade").val('');
				$("select[name=frmEstado] option[sigla='-']").attr("selected","selected");
				$("#frmSite").val('');
				$("#frmDDIEmpresa").val('');
				$("#frmDDDEmpresa").val('');
				$("#frmTelefoneEmpresa").val('');
				$("select[name=frmTipoEmpresa] option[sigla='-']").attr("selected","selected");
				$("select[name=frmRamo] option[sigla='-']").attr("selected","selected");
				$("#frmEmail").val('');
				$("#frmSenha").val('');
				//alert(unescape(data.ResultadoTXT));
				//alert(aviso_doc_pf);
				jAlert(aviso_doc_pf,aviso_titulo);
				
				/* ***********************************************************************************/
				// Apagar todos LI se Existirem Primeiro
				$('#BoxListaRamos').html('');
				// Inserir nova UL
				$('#BoxListaRamos').append('<ul id="ListaRamosCadastrados"></ul>');
				/* ***********************************************************************************/
				// Apagar todos LI se Existirem Primeiro
				$('#TitulosListaProdutosCadastrados').hide();
				/* ***********************************************************************************/
			}

			// Banco 2012 >
			if (data.Resultado == '1') {
				// Dados de ORIGEM do Registro
				$("#id_empresa").val(unescape(data.ID_Empresa));
				$("#origem_cnpj").val(unescape(data.Banco));
				
				// Banco novo mas não se cadastrou como universidade
				if (data.Banco == 'New' && data.Empresa_em_universidade == 'false') {
					jAlert(aviso_msg,aviso_titulo);
				}
				// Banco e está em universidade
				if (data.Banco == 'New' && data.Empresa_em_universidade == 'true') {
					function retorno(data) {
						if (data) {
							window.top.document.location = '/alunos/';
						}
					}
					jAlert("CNPJ cadastrado em <b>Universidades nesta Edição</b>! <br><br>Você será redirecionado à página para cadastro de alunos", aviso_titulo, retorno);
					//window.top.document.location = '/alunos/';
				}
				
				// troca o valor dos elementos  		
				$("#frmRazao").val(unescape(data.Razao));		
				$("#frmFantasia").val(unescape(data.Fantasia));
				$('#nome_empresa').html($("#frmFantasia").val());
				
				$("#frmSigla").val(unescape(data.Fantasia));
				if (data.Sigla != undefined) { $("#frmResp").val(unescape(data.Sigla)); }

				$('#nome_empresa').html($("#frmSigla").val());
				
				$("#frmCEP").val(unescape(data.CEP));
				$("#frmEndereco").val(unescape(data.Endereco));
				$("#frmNumero").val(unescape(data.Numero));
				$("#frmComplemento").val(unescape(data.Complemento));
				$("#frmBairro").val(unescape(data.Bairro));
				$("#frmCidade").val(unescape(data.Cidade));
				$("select[name=frmEstado] option[sigla="+data.UF+"]").attr("selected","selected");
				if (data.Pais != ''){ $("select[name=frmPais] option[sigla="+data.Pais+"]").attr("selected","selected"); }
				$("#frmSite").val(unescape(data.Site));
							
				if (data.DDI1 != undefined) { $("#frmDDIEmpresa").val(unescape(data.DDI1)); }
				if (data.DDD1 != undefined) { $("#frmDDDEmpresa").val(unescape(data.DDD1)); }
				if (data.Fone1 != undefined) { $("#frmTelefoneEmpresa").val(unescape(data.Fone1)); }
				if (data.Email != undefined) { $("#frmEmail").val(unescape(data.Email)); }	
				//alert(unescape(data.ResultadoTXT));
				
				// =======================================================================================================
				//if (data.Produto != undefined) { $("#frmPriProdut").val(unescape(data.Produto)); }

				// alert(tp_formulario);
				if (tp_formulario == 1){

					if (data.Produtos != undefined) {
						// Zerar Array antes de começar
						if (produtos_cadastrados.length >= 0) {
							produtos_cadastrados.splice(0, produtos_cadastrados.length);
						}
						// Exibir produtos Cadastrados Anteriormente para Esta Empresa
						for (var y = 0; y < data.Produtos.length; y ++ ) {
							produtos_cadastrados[produtos_cadastrados.length ++] = data.Produtos[y].Produto;
						}
						//if (data.Produto != undefined) { produtos_cadastrados[produtos_cadastrados.length ++] = data.Produto; }
						
						// Exibe produtos cadastrados previamente
						if (produtos_cadastrados.length >= 0) {
							
							// Apagar todos LI se Existirem Primeiro
							$('#ListaProdutosCadastrados').remove();
							// Inserir nova UL
							$('#TitulosListaProdutosCadastrados').append('<ul id="ListaProdutosCadastrados"></ul>');
							$('#TitulosListaProdutosCadastrados').show();
							
							// Adicionar os novos
							for (var x = 0; x < produtos_cadastrados.length ; x ++) {
								$("#ListaProdutosCadastrados").append("<li>" + produtos_cadastrados[x] + "</li>");
							}
						}
					// Se não retornar produtos esconder o label
					} else {
						$('#TitulosListaProdutosCadastrados').hide();
					}
					
				}
				// =======================================================================================================
				
				if (tp_formulario == 1 || tp_formulario == 2) {
					if (data.Ramos != undefined) {
						// Zerar Array antes de começar
						//var ramos_cadastrados = new Array();
						if (ramos_cadastrados.length >= 0) {
							ramos_cadastrados.splice(0, ramos_cadastrados.length);
						}
						// Exibir produtos Cadastrados Anteriormente para Esta Empresa
						for (var x = 0; x < data.Ramos.length; x ++ ) {
							var complemento = $.trim(data.Ramos[x].Complemento);
							if (complemento.length > 0) {
								ramos_cadastrados[ramos_cadastrados.length ++] = data.Ramos[x].Ramo + ' - ' + data.Ramos[x].Complemento;
							} else {
								ramos_cadastrados[ramos_cadastrados.length ++] = data.Ramos[x].Ramo;
							}
						}
						// =========================================
						// Exibe produtos cadastrados previamente
						if (ramos_cadastrados.length >= 0) {
							
							// Apagar todos LI se Existirem Primeiro
							$('#BoxListaRamos').html('');
							// Inserir nova UL
							$('#BoxListaRamos').append('<ul id="ListaRamosCadastrados"></ul>');
							
							// Adicionar os novos
							for (var x = 0; x < ramos_cadastrados.length ; x ++) {
								$("#ListaRamosCadastrados").append("<li>" + ramos_cadastrados[x] + "</li>");
							}
						}
					} else {
						// Apagar todos LI se Existirem Primeiro
						$('#BoxListaRamos').html('');
						// Inserir nova UL
						$('#BoxListaRamos').append('<ul id="ListaRamosCadastrados"></ul>');
					}
				}
				// =======================================================================================================
				
			} 

			$('#frmRazao').focus().select();
			exibir('GrupoEmpresa');
		});
		cnpj_validado = true;
		$('#produtos_alterar').val('');
	}
	$('#loading').fadeOut();
} 

function exibir(grupo) {
	if (idioma_atual == 1){
		// Muda a quantidade por tipo de formulário
		switch (tp_formulario) {
			case '1': // empresa
				var total = 13;
				var empresa = 8;
				break;
			case '2': // entidade
				var total = 12;
				var empresa = 7;
				break;
			case '4': // pf
				var total = 6;
				var empresa = 0;
				break;
			case '5': // universidade
				var total = 10;
				var empresa = 6;
				break;
		}
		
		if (grupo == 'Grupo1e2') {
			for (i = 1; i <= total; i ++) {
				$('#grupo' + i).show();
			}
		} else if (grupo == 'GrupoEmpresa') {
			for (i = 1; i <= empresa; i ++) {
				$('#grupo' + i).show();
			}
		} else if (grupo == 'GrupoVisitante') {
			for (i = empresa; i <= total; i ++) {
				$('#grupo' + i).show();
			}
		}
	}
}
function fechar(grupo) {
	
	if (idioma_atual == 1){
		// Muda a quantidade por tipo de formulário
		switch (tp_formulario) {
			case '1': // empresa
				var total = 13;
				var empresa = 8;
				break;
			case '2': // entidade
				var total = 12;
				var empresa = 7;
				break;
			case '4': // pf
				var total = 6;
				var empresa = 6;
				break;
			case '5': // universidade
				var total = 10;
				var empresa = 6;
				break;
		}

		if (grupo == 'Grupo1e2') {
			for (i = 1; i <= total; i ++) {
				$('#grupo' + i).hide();
			}
		} else if (grupo == 'GrupoEmpresa') {
			for (i = 1; i <= empresa; i ++) {
				$('#grupo' + i).hide();
			}
		} else if (grupo == 'GrupoVisitante') {
			for (i = 7; i <= empresa; i ++) {
				$('#grupo' + i).hide();
			}
		}
	}
}

// Buscar CPF
function getCadastroCPF() {  
	// Se o campo CPF não estiver vazio  
	var cpf = $.trim($("#frmCPF").val());
	var cpf = cpf.replace('_','');
	
	if (idioma_atual == 1) {
		var cpf_validado = validarCPF(cpf);
	} else {
		var cpf_validado = true;	
	}
	
	if (cpf_validado) {
		//alert("/scripts/metodo_busca_cpf.asp?cpf="+$("#frmCPF").val());
		$('#bt_busca_cpf').fadeOut();
		$.getJSON("/scripts/metodo_busca_cpf.asp?cpf=" + cpf, function(data,textStatus){  
			// Se o resultado for igual a 1  
			if (data.Resultado == '1') {
				// Dados de ORIGEM do Registro
				$("#id_visitante").val(unescape(data.ID_Visitante));
				$("#origem_cpf").val(unescape(data.Banco));
				
				if (data.Banco == 'New') {
					jAlert(aviso_msg,aviso_titulo);
				}
				
				// troca o valor dos elementos  
				$("#frmNome").val(unescape(data.NomeF));
				$("#frmNmCracha").val(unescape(data.NomeCredencialF));
				$("#frmDtNasc").val(unescape(data.DTNasc));
				
				$("select[name=frmCargo] option[sigla="+data.Cargo+"]").attr("selected","selected");
				$("select[name=frmDepto] option[sigla="+data.Departamento+"]").attr("selected","selected");
				
				if (data.DepartamentoOutros) {
					$("#frmDeptoOutros").val(unescape(data.DepartamentoOutros));	
				}
				if (data.CargoOutros) {
					$("#frmCargoOutros").val(unescape(data.CargoOutros));
				}

				// verifica se o item outros estava clicado qdo re-carregou a página
				$('select[name=frmCargo]').each(function() {
					// se for OUTROS
					texto = $("select[name=frmCargo] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
						TipoCargoOn(true);
					} else {
						var valor = $("select[name=frmCargo] option:selected").val();
						if (valor != '-') {
							getSubCargo(valor);
						}
						TipoCargoOn(false);
					}
				});

				// verifica se o item outros estava clicado qdo re-carregou a página
				$('select[name=frmDepto]').each(function() {
					// se for OUTROS
					texto = $("select[name=frmDepto] option:selected").text();
					nome = jQuery.trim(texto);	
					if (nome == 'Outros' || nome == 'Others' || nome == 'Otros' || nome == 'Otras Posiciones' || nome == 'OUTROS' || nome == 'OTHERS' || nome == 'Other Positions' || nome == 'OTROS') {
						TipoDeptoOn(true);
					} else {
						TipoDeptoOn(false);
					}
				});

				$("#frmDDI").val(unescape(data.DDI1));
				$("#frmDDD").val(unescape(data.DDD1));
				$("#frmTelefone").val(unescape(data.Fone1));
				/* Trazer o TIPO e verificar */
				//$("select[name=frmTipo] option[sigla="+data.ID_Tipo_Telefone1+"]").attr("selected","selected");
				//TipoTelefone(data.ID_Tipo_Telefone1);
				$("#frmRamal").val(unescape(data.Ramal1));
				if (data.DDI2) {
					$("#frmDDI2").val(unescape(data.DDI2));
				}
				if (data.DDD2) {
					$("#frmDDD2").val(unescape(data.DDD2));
				}
				if (data.Fone2) {
					$("#frmTelefone2").val(unescape(data.Fone2));
				}
				/* Trazer o TIPO e verificar */
				//$("select[name=frmTipo2] option[sigla="+data.ID_Tipo_Telefone2+"]").attr("selected","selected");
				//TipoTelefone2(data.ID_Tipo_Telefone2);
				if (data.Ramal2) {
					$("#frmRamal2").val(unescape(data.Ramal2));
				}	
				$("#frmEmail").val(unescape(data.Email));
				$("#frmEmailConf").val('');
				$("select[name=frmSexo] option[sigla="+data.Sexo+"]").attr("selected","selected");
				//alert(unescape(data.ResultadoTXT));
				
				$('#texto_final').fadeIn();
				
				exibir('GrupoVisitante');
			} else if (data.Resultado == '2') {
				$("#frmCPF").val('');
				$("#frmNome").val('');
				$("#frmNmCracha").val('');
				$("#frmDtNasc").val('');
				$("select[name=frmCargo] option[sigla='-']").attr("selected","selected");
				$("select[name=frmDepto] option[sigla='-']").attr("selected","selected");
				$("#frmDDI").val('');
				$("#frmDDD").val('');
				$("#frmTelefone").val('');
				$("#frmDDI2").val('');
				$("#frmDDD2").val('');
				$("#frmTelefone2").val('');
				$("#frmEmail").val('');
				$("#frmEmailConf").val('');
				//alert(unescape(data.ResultadoTXT));
				//alert(aviso_doc_pf);
				jAlert(aviso_doc_existe,aviso_titulo_existe);
				//exibir('GrupoVisitante');
			} else {  
				$("#frmNome").val('');
				$("#frmNmCracha").val('');
				$("#frmDtNasc").val('');
				$("select[name=frmCargo] option[sigla='-']").attr("selected","selected");
				$("select[name=frmDepto] option[sigla='-']").attr("selected","selected");
				$("#frmDDI").val('');
				$("#frmDDD").val('');
				$("#frmTelefone").val('');
				$("#frmDDI2").val('');
				$("#frmDDD2").val('');
				$("#frmTelefone2").val('');
				$("#frmEmail").val('');
				$("#frmEmailConf").val('');
				//alert(unescape(data.ResultadoTXT));
				//alert(aviso_doc_pf);
				jAlert(aviso_doc_pf,aviso_titulo);
				exibir('GrupoVisitante');
			}
			$('#bt_busca_cpf').fadeIn();
			
		});  
	}  else {
		jAlert(aviso_msg_cpf_err,'Atenção');
	}
} 

// Buscar CEP
function getEndereco() {  
	// Se o campo CEP não estiver vazio  
	if($.trim($("#frmCEP").val()) != ""){ 
		
		$.getScript("http://cep.republicavirtual.com.br/web_cep.php?formato=javascript&cep="+$("#frmCEP").val(), function(){  
			// o getScript dá um eval no script, então é só ler!  
			//Se o resultado for igual a 1  
			if(resultadoCEP["resultado"] == 1){  
				// troca o valor dos elementos  
				$("#frmEndereco").val(unescape(resultadoCEP["tipo_logradouro"])+" "+unescape(resultadoCEP["logradouro"]));  
				$("#frmNumero").val('');
				$("#frmComplemento").val('');
				$("#frmBairro").val(unescape(resultadoCEP["bairro"]));  
				$("#frmCidade").val(unescape(resultadoCEP["cidade"]));  
				$("select[name=frmEstado] option[sigla="+resultadoCEP["uf"]+"]").attr("selected","selected");
				//alert(unescape(resultadoCEP["resultado_txt"]))
			}else{  
				$("#frmEndereco").val('');  
				$("#frmNumero").val('');
				$("#frmComplemento").val('');
				$("#frmBairro").val('');  
				$("#frmCidade").val('');  
				$("select[name=frmEstado] option[sigla='-']").attr("selected","selected");
				jAlert("Endereço não encontrado","Atenção");  
			}  
		});  
	}  
} 


function TipoRamoOn(selecionado) {
	if (selecionado == true) {
		//alert(selecionado)
		$('#Complemento').hide().fadeIn();
		$('#frmOptRamoComplemento').show();
		$('#frmOptRamoComplemento').focus();
		$('#FecharRamoComplemento').show();
	} else {
		$('#Complemento').hide();
		$('#FecharRamoComplemento').hide();
		$('#frmOptRamoComplemento').hide();
		$('#frmOptRamoComplemento').val('');
	}
}
function TipoRamoOff() {
	$('#Complemento').hide();
	$('#FecharRamoComplemento').hide();
	$('#frmOptRamoComplemento').hide();
	$('#frmOptRamoComplemento').val('');
	$("select[name=frmRamo] option[sigla='-']").attr("selected","selected");
} 

// Buscar Atividade
var atividade_existe = 'nao';
function getAtividade(ID_Ramo) {  
	$.getJSON("/scripts/metodo_busca_atividade.asp?ID_Ramo=" + ID_Ramo, function(data,textStatus){  
		if (data.Resultado == '1') {
			var listItems = '<option value="-" sigla="-">-- ' + select + ' --</option>';
			for (var i = 0; i < data.Atividades.length; i++){
				listItems += "<option value='" + data.Atividades[i].id + "'>" + data.Atividades[i].nome + "</option>";
			}
			$("#Atividade").hide();
			$("#Atividade").fadeIn();
			$("#frmAtividade").html(listItems);
			$("#frmAtividadeOutros").val('');
			$("#frmOptRamoOutros").val('');
			atividade_existe = 'sim';
		}else{  
			$("#Atividade").hide();
			$("#frmAtividade").html('');
			$("#frmAtividadeOutros").val('');
			$("#frmOptRamoOutros").val('');
			atividade_existe = 'nao';
		} 
	});  
}

function TipoAtividadeOn(selecionado) {
	if (selecionado == true) {
		//alert(selecionado)
		$('#frmAtividade').hide();
		$('#frmAtividadeOutros').show();
		$('#frmAtividadeOutros').focus();
		$('#FecharAtividadeOutros').show();
	} else {
		$('#FecharAtividadeOutros').hide();
		$('#frmAtividadeOutros').hide();
		$('#frmAtividade').show();	
		$("#frmAtividadeOutros").val('');
		$("#frmOptRamoOutros").val('');
	}
}
function TipoAtividadeOff() {
	$('#FecharAtividadeOutros').hide();
	$('#frmAtividadeOutros').hide();
	$('#frmAtividade').show();
	$("select[name=frmAtividade] option[sigla='-']").attr("selected","selected");
} 

// Buscar SubCargo
var subcargo_existe = 'nao';
function getSubCargo(ID_Cargo) {  
	$.getJSON("/scripts/metodo_busca_subcargo.asp?ID_Cargo=" + ID_Cargo, function(data,textStatus){  
		if (data.Resultado == '1') {
			var listItems = '<option value="-" sigla="-">-- ' + select + ' --</option>';
			for (var i = 0; i < data.SubCargos.length; i++){
				listItems += "<option value='" + data.SubCargos[i].id + "' sigla='" + data.SubCargos[i].id + "'>" + data.SubCargos[i].nome + "</option>";
			}
			$("#SubCargo").hide();
			$("#SubCargo").fadeIn();
			$("#frmSubCargo").html(listItems);
			$("#frmCargoOutros").val('');
			$("#frmSubCargoOutros").val('');
			subcargo_existe = 'sim';
		}else{  
			$("#SubCargo").hide();
			$("#frmSubCargo").html('');
			subcargo_existe = 'nao';
		} 
	});  
} 

function TipoCargoOn(selecionado) {
	if (selecionado == true) {
		//alert(selecionado)
		getSubCargo("-")
		$('#frmCargo').hide();
		$('#frmCargoOutros').show();
		$('#frmCargoOutros').focus();
		$('#FecharCargoOutros').show();
		$("#frmCargoOutros").val('');
		$("#frmSubCargoOutros").val('');
	} else {
		$('#FecharCargoOutros').hide();
		$('#frmCargoOutros').hide();
		$('#frmCargo').show();	
		$("#frmCargoOutros").val('');
		$("#frmSubCargoOutros").val('');
	}
}
function TipoCargoOff() {
	$('#FecharCargoOutros').hide();
	$('#frmCargoOutros').hide();
	$('#frmCargo').show();
	$("select[name=frmCargo] option[sigla='-']").attr("selected","selected");
} 

function TipoSubCargoOn(selecionado) {
	if (selecionado == true) {
		//alert(selecionado)
		$('#frmSubCargo').hide();
		$('#frmSubCargoOutros').show();
		$('#frmSubCargoOutros').focus();
		$('#FecharSubCargoOutros').show();
	} else {
		$('#FecharSubCargoOutros').hide();
		$('#frmSubCargoOutros').val('');
		$('#frmSubCargoOutros').hide();
		$('#frmSubCargo').show();	
	}
}
function TipoSubCargoOff() {
	$('#FecharSubCargoOutros').hide();
	$('#frmSubCargoOutros').val('');
	$('#frmSubCargoOutros').hide();
	$('#frmSubCargo').show();
	$("select[name=frmSubCargo] option[sigla='-']").attr("selected","selected");
} 
	
function TipoDeptoOn(selecionado) {
	if (selecionado == true) {
		$('#frmDepto').hide();
		$('#frmDeptoOutros').show();
		$('#frmDeptoOutros').focus();
		$('#FecharDeptoOutros').show();
	} else {
		$('#FecharDeptoOutros').hide();
		$('#frmDeptoOutros').hide();
		$('#frmDepto').show();		
	}
}
function TipoDeptoOff() {
	$('#FecharDeptoOutros').hide();
	$('#frmDeptoOutros').hide();
	$('#frmDepto').show();
	$("select[name=frmDepto] option[sigla='-']").attr("selected","selected");
} 

function exec_duvida(qual, posicao, texto, efeito, tipo, largura) {
	var exibir = texto;
	if (largura == undefined) {
		width = 200;
	} else {
		width = largura
	}
	var larg_pag = $('#conteudo').innerWidth();
	var larg_relacionada = (larg_pag-870)/2;
	if (posicao.left + width < larg_pag) {
		$('#tabela_duvida').css('width', width);
		$('#duvida').css('width', width);
		$('#duvida').css('top', posicao.top - 20);
		$('#duvida').css('left', posicao.left + 20);
		$('#duvida').css('z-index', 99);
		$('#texto').html(exibir);
		if (efeito == undefined) { 
			$('#duvida').show();
		} else {
			$('#duvida').fadeIn(); 
		}
	} else {
		$('#tabela_duvida_invertida').css('width', width);
		$('#duvida_invertida').css('width', width);
		$('#duvida_invertida').css('top', posicao.top - 20);
		$('#duvida_invertida').css('left', posicao.left - width);
		$('#duvida').css('z-index', 99);
		$('#texto_invertido').html(exibir);
		if (efeito == undefined) { 
			$('#duvida_invertida').show(); 
		} else {
			$('#duvida_invertida').fadeIn(); 
		}
	}
}

function hide_duvidas () {
	$('#duvida').hide();
	$('#duvida_invertida').hide();
}
function menu (qual, acao) {
	switch(acao) {
		case 'over':
//			$(qual).css( {'background-color':'#d6d6d6' } );
			$(qual).css( {'background-color':'#e4e5e6' } );
			break;
		case 'out':
//			$(qual).css( {'background-color':'#e4e5e6' } );
			$(qual).css( {'background-color':'#ffffff' } );			
			break;
	}
}
function homelink() {
	var confirmacao = confirm('Voce esta deixando a sessao atual !\n\nDeseja continuar ?');
	if (confirmacao) {
			document.location = 'http://www.mbxeventos.net/aol3abf2016';
	}
}
function sair() {
	var msg = '';
	var titulo = '';
	if (typeof idioma_atual != "undefined") {
		if (idioma_atual == '1') {
			titulo = 'Sair';
			msg = 'Deseja fechar sua sess&atilde;o?';		
		} else if (idioma_atual == '2') {
			titulo = 'Sa&iacute;da';
			msg = '&iquest;Desea cerrar la sesi&oacute;n?';
		} else if (idioma_atual == '3') {
			titulo = 'Exit';
			msg = 'Do you want to close your session?';
		}
	} else {
		titulo = 'Sair';
		msg = 'Deseja fechar sua sess&atilde;o?';	
	}
	jConfirm(msg, titulo, function(r) {
		switch (r) {
			case true:
				//show_loading();
				document.location = 'http://www.mbxeventos.net/aol3abf2016';
				break;
			case false:
				break;
		}
	});
}
function cor_muito_clara(hexcolor, objeto){
	hexcolor = hexcolor.replace('#','');
	var r = parseInt(hexcolor.substr(0,2),16);
	var g = parseInt(hexcolor.substr(2,2),16);
	var b = parseInt(hexcolor.substr(4,2),16);
	var yiq = ((r*299)+(g*587)+(b*114))/1000;
	if (yiq >= 170) {
		$("#" + objeto).addClass('cor_cinza2');
	} else {
		$("#" + objeto).removeClass('cor_cinza2');	
	}
}

function tabClose() {
  var tab = window.open(window.location,"_self");
  tab.close();
}

function fecharJanela2(){

  window.opener='X';
  window.open("blank.asp", "_parent");
  window.close();

}
function CloseWindow() 
{	window.location='http://www.mbxeventos.net/aol3abf2016';
}
function checar(id, nome, form) {
	if (nome != undefined) {
		// Verificar se foi o item "nao" CLICADO
		if (nome.substr(0,3) == "Não" || nome.substr(0,12) == 'We will no' || nome.substr(0,10) == 'We are not' || nome.substr(0,10) == 'No debemos' || nome.substr(0,10) == 'No tenemos' || nome.substr(0,8) == 'No tengo' || nome.substr(0,8) == 'I am not'  ) {
 			for (x = 0; x < $('input[name="' + form + '"]').length; x ++) {
 				// se nao for o item NAO
				if (x != id) {
					// se estiver marcado desmarque
					if ($('input[name="' + form + '"]')[x].checked) {
						$('input[name="' + form + '"]')[x].click();
					}
				}
			}
		// Se nao foi o item "nao" CLICADO
		} else { 
			//alert(id + ', ' + nome + ', ' + form + ', ' + $('input[name="' + form + '"]').length);
			for (x = 0; x < $('input[name="' + form + '"]').length; x ++) {
				var titulo = $('#' + form + '_' + x).html();
				while (titulo.indexOf('&nbsp;') >= 0) {
					titulo = titulo.replace('&nbsp;','');
				}
				//alert(titulo);
				if (titulo.substr(0,3) == "Não" || titulo.substr(0,12) == 'We will no' || titulo.substr(0,10) == 'We are not' || titulo.substr(0,10) == 'No debemos' || titulo.substr(0,10) == 'No tenemos' || titulo.substr(0,8) == 'No tengo' || titulo.substr(0,8) == 'I am not') {
					// se estiver marcado desmarque
					if ($('input[name="' + form + '"]')[x].checked == true) {
						$('input[name="' + form + '"]')[x].click();
					}
				}
			}
		}
	}
}