$(document).ready(function(){
	// Ao terminar de carregar o documento executar:
	$('#aviso').hide();
	$('#aviso_topo').hide();
	$('#RecebeSmS1').hide();
	$('#RecebeSmS2').hide();
	$('#parcFeira').hide();
	$('#parcAnuncio').hide();
	$('#parcAssis').hide();

	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);

	// Modelo Mascara
	$("#frmCNPJ").mask("99.999.999/9999-99",{placeholder:"_"});
	$("#frmCEP").mask("99999-999",{placeholder:"_"});
	$("#frmTelefone").mask("9999-9999",{placeholder:"_"});	
	$("#frmTelefone2").mask("9999-9999",{placeholder:"_"});	

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
	
});

function verificar_interesse() {
		
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

// Buscar CNPJ
function getCadastroCNPJ() {  
	// Se o campo CNPJ não estiver vazio  
	if($.trim($("#frmCNPJ").val()) != ""){
			
		$.getJSON("/scripts/metodo_busca_cnpj.asp?cnpj="+$("#frmCNPJ").val(), function(data,textStatus){  
			// o getScript dá um eval no script, então é só ler!  
			//Se o resultado for igual a 1  
			if (data.Resultado == '1') {
				// troca o valor dos elementos  
				$("#frmRazao").val(unescape(data.Razao));		
				$("#frmFantasia").val(unescape(data.Fantasia));
				$("#frmPriProdut").val(unescape(data.Produto));
				$("#frmCEP").val(unescape(data.CEP));
				$("#frmEndereco").val(unescape(data.Endereco));
				$("#frmNumero").val(unescape(data.Numero));
				$("#frmComplemento").val(unescape(data.Complemento));
				$("#frmBairro").val(unescape(data.Bairro));
				$("#frmCidade").val(unescape(data.Cidade));
				$("select[name=frmEstado] option[sigla="+data.UF+"]").attr("selected","selected");
				$("select[name=frmPais] option[sigla="+data.Pais+"]").attr("selected","selected");
				$("#frmSite").val(unescape(data.Site));
				$("#frmCPF").val(unescape(data.CPF));
				$("#frmNome").val(unescape(data.NomeF));
				$("#frmNmCracha").val(unescape(data.NomeCredencialF));
				$("#frmDtNasc").val(unescape(data.DTNasc));
				$("select[name=frmCargo] option[sigla="+data.Cargo+"]").attr("selected","selected");
				$("select[name=frmDepto] option[sigla="+data.Departamento+"]").attr("selected","selected");
				$("#frmDDI").val(unescape(data.DDI1));
				$("#frmDDD").val(unescape(data.DDD1));
				$("#frmTelefone").val(unescape(data.Fone1));
				$("#frmDDI2").val(unescape(data.DDI2));
				$("#frmDDD2").val(unescape(data.DDD2));
				$("#frmTelefone2").val(unescape(data.Fone2));
				$("#frmEmail").val(unescape(data.Email));			
				alert(unescape(data.ResultadoTXT));
			}else{  
				alert("Nao encontrado");  
			}  
		});  
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
				alert(unescape(resultadoCEP["resultado_txt"]))
			}else{  
				$("#frmEndereco").val(''); 
				$("#frmNumero").val('');
				$("#frmComplemento").val('');
				$("#frmBairro").val('');  
				$("#frmCidade").val('');  
				$("select[name=frmEstado] option[sigla='-']").attr("selected","selected");
				alert("Endereço não encontrado");  
			}  
		});  
	}  
} 


function Enviar() {
	var erros = 0;
	// Limpar os Desativados
	$('select:disabled').each(function(i) { mudar_aviso(this.id, 'ok', false); } );
	$('input:disabled').each(function(i) { mudar_aviso(this.id, 'ok', false); } );
	// Verificar os Ativos
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case 'target':
				mudar_aviso(this.id, 'ok', false);
				break;
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
				valor_cnpj = $('#frmCNPJ').val();
					cnpj_incorreto = false;
					if (valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99') {
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
				break;
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
}