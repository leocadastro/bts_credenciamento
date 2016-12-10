$(document).ready(function(){
	// Ao terminar de carregar o documento executar:
	cor_muito_clara(cor_fundo,'txt_1');
	$('#aviso').hide();
	
	$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
	$('#txt_1').hide();
	setTimeout(function() {
		$('#txt_1').fadeIn();
	},1000);
	
});