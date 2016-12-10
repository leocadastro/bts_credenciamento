function mudar_aviso(qual, aviso, efeito) {
	if (aviso == 'ok') {
//		$('#' + qual).addClass("borda_text_field");
		$('#' + qual).removeClass("formulario_alerta");
		if ( efeito != false ) {
			$('#' + qual).hide().fadeIn("fast");
		} else { 
			//$('#' + qual).show(); 
		}
	} else if (aviso == 'x') {
//		$('#' + qual).removeClass("borda_text_field");
		$('#' + qual).addClass("formulario_alerta");
		if ( efeito != false ) {
			$('#' + qual).fadeOut("fast").fadeIn("fast").fadeOut("fast").fadeIn("fast");
		} else { 
			//$('#' + qual).show(); 
		}
	}
}
function verificar (qual, efeito, qtde) {
	/*
	if (qtde != undefined) {
		if ($('#' + qual).val().length < qtde) {
			mudar_aviso(qual, 'x', efeito);
			return 1;
		} else { 
			mudar_aviso(qual, 'ok', efeito);
			return 0;
		}
	*/
	if ( $('#' + qual).val() == '' || $('#' + qual).val() == '-') {
		mudar_aviso(qual, 'x', efeito);
		return 1;
	} else { 
		mudar_aviso(qual, 'ok', efeito);
		return 0;
	}
}
qdte_duvidas = 0;
function exec_duvida_form(texto, posicao, tamanho) {
	var larg_pag = $('#conteudo').innerWidth();
	var larg_relacionada = (larg_pag-870)/2;
	if (posicao.left + 300 < larg_pag) {
		$('#duvida').css('top', posicao.top - 20);
		$('#duvida').css('left', posicao.left + tamanho);
		$('#duvida').css('z-index', 99);
		$('#duvida').width(200);
		$('#tabela_duvida').width(200);
		$('#texto').html('');
		$('#texto').addClass('fs10px t_arial');
		$('#duvida').fadeOut().fadeIn();
		$('#texto').html(texto);
	} else {
		$('#duvida_invertida').css('top', posicao.top - 20);
		$('#duvida_invertida').css('left', posicao.left - 145);
		$('#duvida').css('z-index', 99);
		$('#texto_invertido').html(texto);
		$('#duvida_invertida').show();
	}
}
function valida_email_novo(qual) {
	var txt = '';
//	alert(qual);
	if (jQuery && navigator.appName != 'Microsoft Internet Explorer') {
		var form = $('#'+qual).val();
	} else {
		var form = $('#'+qual).val();
	}
//	alert(form);
//	alert('valida_email_novo - form: ' + form);
/*	if (qual == 'email') { form = document.getElementById(qual); }
	if (qual == 'amigo_email') { form = document.logar.amigo_email; }
*/
	//form = document.forms['logar'].elements[qual];
	//form = document.getElementById(qual);
	erro = 0;
	if (Rules_Vazio(form)) {
		txt = 'Por favor, o campo não pode ser vazio.';
		erro ++;
	} else if (Rules_Esp1(form) || Rules_Esp2(form) || Rules_Esp3(form)) {
		txt = 'Por favor, não utilize caracteres especiais.';
		erro ++;
	} else if (form.length < 6) {
		txt = 'E-mail muito curto.';
		erro ++;
	} else if (Rules_Email(form) == false) {
		var str = form; // email string
		var reg1 = /(@.*@)|(\.\.)|(@\.)|(\.@)|(^\.)/; // not valid
		var reg2 = /^.+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?)$/; // valid
		if (!reg1.test(str) && reg2.test(str)) { 
			// valido
		} else {
			// invalido
			txt = 'E-mail inválido, por favor digite corretamente.';
		}
		erro ++;
	} else {
		txt = '';
	}
	return txt;
	
}
/* Criacao de Rules*/
/* Vefifica c e numero				*/ function Rules_Numero(c) { return (((c >=-99999999*9999999) && (c <=99999999*9999999)) || (c.indexOf(",")>=0)) }
/* Vefifica { } ( ) < > [ ] | \ /  	*/ function Rules_Esp1(c) { return ((c.indexOf("{")>=0) || (c.indexOf("}")>=0) || (c.indexOf("(")>=0) || (c.indexOf(")")>=0) || (c.indexOf("<")>=0) || (c.indexOf(">")>=0) || (c.indexOf("[")>=0) || (c.indexOf("]")>=0) || (c.indexOf("|")>=0) || (c.indexOf("/")>=0)) }
/* Vefifica & * $ % ? ! ^ ~ ` ' "  	*/ function Rules_Esp2(c) { return ((c.indexOf("&")>=0) || (c.indexOf("*")>=0) || (c.indexOf("$")>=0) || (c.indexOf("%")>=0) || (c.indexOf("?")>=0) || (c.indexOf("!")>=0) || (c.indexOf("^")>=0) || (c.indexOf("~")>=0) || (c.indexOf("`")>=0) || (c.indexOf("\"")>=0) || (c.indexOf("`")>=0) || (c.indexOf("'")>=0)) }
/* Vefifica , ; : = #  				*/ function Rules_Esp3(c) { return ((c.indexOf(",")>=0) || (c.indexOf(";")>=0) || (c.indexOf(":")>=0) || (c.indexOf("=")>=0) || (c.indexOf("#")>=0)) }
/* Vefifica @ .  					*/ function Rules_Email(c) { return ((c.indexOf("@")>=0) && (c.indexOf(".")>=0)); }
/* Verifica se o valor e Nulo       */ function Rules_Vazio(c) { return ((c == null) || (c.length == 0)); }
/* Verifica se o valor e Nulo       */ function Rules_Pequeno(c) { return ((c.length < 6)); }
function ValidaCNPJ(cnpj) {

  var i = 0;
  var l = 0;
  var strNum = "";
  var strMul = "6543298765432";
  var character = "";
  var iValido = 1;
  var iSoma = 0;
  var strNum_base = "";
  var iLenNum_base = 0;
  var iLenMul = 0;
  var iSoma = 0;
  var strNum_base = 0;
  var iLenNum_base = 0;

  if (cnpj == "")
        return (false);

  l = cnpj.length;
  for (i = 0; i < l; i++) {
        caracter = cnpj.substring(i,i+1)
        if ((caracter >= '0') && (caracter <= '9'))
           strNum = strNum + caracter;
  };

  if(strNum.length != 14)
        return (false);

  strNum_base = strNum.substring(0,12);
  iLenNum_base = strNum_base.length - 1;
  iLenMul = strMul.length - 1;
  for(i = 0;i < 12; i++)
        iSoma = iSoma +
                        parseInt(strNum_base.substring((iLenNum_base-i),(iLenNum_base-i)+1),10) *
                        parseInt(strMul.substring((iLenMul-i),(iLenMul-i)+1),10);

  iSoma = 11 - (iSoma - Math.floor(iSoma/11) * 11);
  if(iSoma == 11 || iSoma == 10)
        iSoma = 0;

  strNum_base = strNum_base + iSoma;
  iSoma = 0;
  iLenNum_base = strNum_base.length - 1
  for(i = 0; i < 13; i++)
        iSoma = iSoma +
                        parseInt(strNum_base.substring((iLenNum_base-i),(iLenNum_base-i)+1),10) *
                        parseInt(strMul.substring((iLenMul-i),(iLenMul-i)+1),10)

  iSoma = 11 - (iSoma - Math.floor(iSoma/11) * 11);
  if(iSoma == 11 || iSoma == 10)
        iSoma = 0;
  strNum_base = strNum_base + iSoma;
  if(strNum != strNum_base)
        return (false);

  return (true);

}
function validacpf_2012(valor, aviso){
	if (idioma_atual != 1) {
		retorno = true;	
	} else if (valor.length > 0) {
		retorno = true;
		cpf = valor.replace('.','');
		cpf = cpf.replace('.','');
		cpf = cpf.replace('-','');
		var i;
		s = cpf;
		var c = s.substr(0,9);
		var dv = s.substr(9,2);
		var d1 = 0;
		for (i = 0; i < 9; i++)
		{
			d1 += c.charAt(i)*(10-i);
		}
		if (d1 == 0){
			retorno = false;
		}
		d1 = 11 - (d1 % 11);
		if (d1 > 9) d1 = 0;
		if (dv.charAt(0) != d1)
		{
			retorno = false;
		}
		d1 *= 2;
		for (i = 0; i < 9; i++)
		{
			d1 += c.charAt(i)*(11-i);
		}
		d1 = 11 - (d1 % 11);
		if (d1 > 9) d1 = 0;
		if (dv.charAt(1) != d1)
		{
			retorno = false;
		}
		if (retorno == false) { 
			if (aviso != 'sem_aviso') {
				//mostrar_aviso('CPF - Inválido') 
			}
			mudar_aviso('cpf_busca', 'x', false);
			return false;
		} else if (retorno == true) {
			mudar_aviso('cpf_busca', 'ok', false);
			return true;
		}
	} else {
		if (aviso != 'sem_aviso') {
			//mostrar_aviso('CPF - Não pode ser vazio', false); 
		}
		mudar_aviso('cpf_busca', 'x', false);				
		return false;
	}
}
function validarCPF(cpf) {
 
 	if (idioma_atual != 1) {
		retorno = true;	
	} else if (cpf.length > 0) {
		cpf = cpf.replace('.','');
		cpf = cpf.replace('-','');
		cpf = cpf.replace(/[^\d]+/g,'');
	 
		if(cpf == '') return false;
	 
		// Elimina CPFs invalidos conhecidos
		if (cpf.length != 11 ||
			cpf == "00000000000" ||
			cpf == "11111111111" ||
			cpf == "22222222222" ||
			cpf == "33333333333" ||
			cpf == "44444444444" ||
			cpf == "55555555555" ||
			cpf == "66666666666" ||
			cpf == "77777777777" ||
			cpf == "88888888888" ||
			cpf == "99999999999")
			return false;
		 
		// Valida 1o digito
		add = 0;
		for (i=0; i < 9; i ++)
			add += parseInt(cpf.charAt(i)) * (10 - i);
		rev = 11 - (add % 11);
		if (rev == 10 || rev == 11)
			rev = 0;
		if (rev != parseInt(cpf.charAt(9)))
			return false;
		 
		// Valida 2o digito
		add = 0;
		for (i = 0; i < 10; i ++)
			add += parseInt(cpf.charAt(i)) * (11 - i);
		rev = 11 - (add % 11);
		if (rev == 10 || rev == 11)
			rev = 0;
		if (rev != parseInt(cpf.charAt(10)))
			return false;
         
	    return true;
	} else {
		return false;
	}
    
}