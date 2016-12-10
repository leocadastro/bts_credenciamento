function mudar_aviso(qual, aviso, efeito) {
	if (aviso == 'ok') {
//		$('#' + qual).addClass("borda_text_field");
		$('#' + qual).hide();
		$('#' + qual).removeClass("form_alerta");
		if ( efeito != false ) {
			$('#' + qual).fadeIn("fast");
		} else { 
			$('#' + qual).show(); 
		}
	} else if (aviso == 'x') {
		$('#' + qual).removeClass("borda_text_field");
		$('#' + qual).addClass("form_alerta");
		if ( efeito != false ) {
			$('#' + qual).fadeOut("fast").fadeIn("fast").fadeOut("fast").fadeIn("fast");
		} else { 
			$('#' + qual).show(); 
		}
	}
}
function Trim(str){
	return str.replace(/^\s+|\s+$/g,"");
}
function verificar (qual, efeito) {
	qual = Trim(qual);
	if ( $('#' + qual).val() == '' || $('#' + qual).val() == '-' ) {
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
/* Criacao de Rules*/
/* Vefifica c e numero				*/ function Rules_Numero(c) { return (((c >=-99999999*9999999) && (c <=99999999*9999999)) || (c.indexOf(",")>=0)) }
/* Vefifica { } ( ) < > [ ] | \ /  	*/ function Rules_Esp1(c) { return ((c.indexOf("{")>=0) || (c.indexOf("}")>=0) || (c.indexOf("(")>=0) || (c.indexOf(")")>=0) || (c.indexOf("<")>=0) || (c.indexOf(">")>=0) || (c.indexOf("[")>=0) || (c.indexOf("]")>=0) || (c.indexOf("|")>=0) || (c.indexOf("/")>=0)) }
/* Vefifica & * $ % ? ! ^ ~ ` ' "  	*/ function Rules_Esp2(c) { return ((c.indexOf("&")>=0) || (c.indexOf("*")>=0) || (c.indexOf("$")>=0) || (c.indexOf("%")>=0) || (c.indexOf("?")>=0) || (c.indexOf("!")>=0) || (c.indexOf("^")>=0) || (c.indexOf("~")>=0) || (c.indexOf("`")>=0) || (c.indexOf("\"")>=0) || (c.indexOf("`")>=0) || (c.indexOf("'")>=0)) }
/* Vefifica , ; : = #  				*/ function Rules_Esp3(c) { return ((c.indexOf(",")>=0) || (c.indexOf(";")>=0) || (c.indexOf(":")>=0) || (c.indexOf("=")>=0) || (c.indexOf("#")>=0)) }
/* Vefifica @ .  					*/ function Rules_Email(c) { return ((c.indexOf("@")>=0) && (c.indexOf(".")>=0)); }
/* Verifica se o valor e Nulo       */ function Rules_Vazio(c) { return ((c == null) || (c.length == 0)); }
/* Verifica se o valor e Nulo       */ function Rules_Pequeno(c) { return ((c.length < 6)); }
