function verificaNumero(e) {
	if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {
	return false;
	}
}

// Buscar CNPJ
function getCadastroCNPJ() {  
	// Se o campo CNPJ não estiver vazio  

	var valor_cnpj = $("#frmCNPJ").val();

	cnpj_incorreto = false;
	if (valor_cnpj == '' || valor_cnpj == '__.___.___/____-__' || valor_cnpj == '00.000.000/0000-00' || valor_cnpj == '11.111.111/1111-11' || valor_cnpj == '22.222.222/2222-22' || valor_cnpj == '33.333.333/3333-33' || valor_cnpj == '44.444.444/4444-44' || valor_cnpj == '55.555.555/5555-55' || valor_cnpj == '66.666.666/6666-66' || valor_cnpj == '77.777.777/7777-77' || valor_cnpj == '88.888.888/8888-88' || valor_cnpj == '99.999.999/9999-99' || valor_cnpj == '__.___.___/____-__') {
		cnpj_incorreto = true;
		alert("CNPJ inválido");
	}
	if (cnpj_incorreto || ValidaCNPJ(valor_cnpj) == false) {
		cnpj_incorreto = true;
	}
	if (cnpj_incorreto) {
		mudar_aviso('frmCNPJ', 'x', false);
	} else {
		mudar_aviso('frmCNPJ', 'ok', false);


		
		$.getJSON("/scripts/metodo_busca_cnpj.asp?cnpj=" + $("#frmCNPJ").val(), function(data,textStatus){  
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
			} else {
				alert("Cadastro não localizado");
			}
		});  
	}
} 

function getCadastroCPF() {  
	// Se o campo CPF não estiver vazio  
	if($.trim($("#frmCPF").val()) != ""){
				
		$.getJSON("/scripts/metodo_busca_cpf.asp?cpf="+$("#frmCPF").val(), function(data,textStatus){  
			// Se o resultado for igual a 1  
			if (data.Resultado == '1') {
				// troca o valor dos elementos  
				$("#frmNome").val(unescape(data.NomeF));
				$("#frmNmCracha").val(unescape(data.NomeCredencialF));
				$("#frmDtNasc").val(unescape(data.DTNasc));
				
				$("select[name=frmCargo] option[sigla="+data.Cargo+"]").attr("selected","selected");
				$("select[name=frmDepto] option[sigla="+data.Departamento+"]").attr("selected","selected");
				
				$("#frmDDI").val(unescape(data.DDI1));
				$("#frmDDD").val(unescape(data.DDD1));
				$("#frmTelefone").val(unescape(data.Fone1));
				$("select[name=frmTipo] option[sigla="+data.ID_Tipo_Telefone1+"]").attr("selected","selected");
				$("#frmRamal").val(unescape(data.Ramal1));
				TipoTelefone(data.ID_Tipo_Telefone1);
				/*
				if (data.SMS1 == 'True') {
					$('#frmSMS').attr("checked", "checked");
				}
				*/
				$("#frmDDI2").val(unescape(data.DDI2));
				$("#frmDDD2").val(unescape(data.DDD2));
				$("#frmTelefone2").val(unescape(data.Fone2));
				$("select[name=frmTipo2] option[sigla="+data.ID_Tipo_Telefone2+"]").attr("selected","selected");
				$("#frmRamal2").val(unescape(data.Ramal2));
				TipoTelefone2(data.ID_Tipo_Telefone2);
				/*
				if (data.SMS2 == 'True') {
					$('#frmSMS2').attr("checked", "checked");
				}
				*/
				$("#frmEmail").val(unescape(data.Email));
				$("#frmEmailConf").val('');
				
				alert(unescape(data.ResultadoTXT));
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