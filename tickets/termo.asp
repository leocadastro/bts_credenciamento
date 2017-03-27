<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">
<link href="/css/checkbox.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/jquery.screwdefaultbuttons.js"></script>
<script language="javascript" src="/js/validar_forms.js"></script>
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<script language="javascript" src="/js/tipos.js"></script>
<script language="javascript" src="/js/jquery.translate.js"></script>
<!-- Script desta página -->
<script language="javascript" src="default.js" charset="utf-8"></script>
<!-- Script desta página FIM -->

<script language="javascript">
$(function() {

  var t = {
    text1: {
	  pt: "Olá",
      en: "Hello",
	  es: "Hola"
    },
	text2: {
	  pt: "Menu",
      en: "Menu",
	  es: "Menú"
    },
	text3: {
	  pt: "Meus Pedidos",
      en: "My purchase orders",
	  es: "Mis pedidos"
    },
	text4: {
	  pt: "Continuar pedido",
      en: "Continue with the purchase order",
	  es: "Continuar pedido"
    },
	text5: {
	  pt: "Suporte: visitante.abf@informa.com",
      en: "Support: visitante.abf@informa.com",
	  es: "Asistencia: visitante.abf@informa.com"
    },
	text6: {
	  pt: "Suporte:",
      en: "Support:",
	  es: "Asistencia:"
    },
	text7: {
	  pt: "Pedido n°:",
      en: "Purchase Order number:",
	  es: "Pedido número:"
    },
	text23: {
	  pt: "Novo Pedido",
      en: "New purchase order",
	  es: "Nuevo pedido"
    },
	text8: {
	  pt: "Termos & Condições",
      en: "Terms & conditions",
	  es: "Términos & condiciones"
    },
	text9: {
	  pt: "<strong>TERMOS E CONDIÇÕES</strong><br><br> Antes de finalizar sua compra conheça os Termos e Condições de venda de Ingresso para o evento ABF Franchising Expo 2017: <br><br> <strong>1 - DO OBJETO</strong><br> 1.1. A compra estará sujeita à disponibilidade de ingressos e à aprovação da operadora de seu cartão de crédito. Não será permitida a entrada de menores de 16 anos desacompanhados. <br><br> <strong>2 - DA AQUISIÇÃO DO INGRESSO</strong><br> 2.1. Confira atentamente os dados de seu pedido antes da confirmação de sua compra. Não será permitido cancelamento ou devolução de ingressos. O seu ingresso é um produto único, ou seja, após sua compra ele não estará mais disponível para venda. <br><br> <strong>3 - DA DESISTÊNCIA DA COMPRA E/OU CANCELAMENTO</strong><br> 3.1. Em caso de arrependimento do cliente, o reembolso do valor do ingresso será efetuado, desde que esse direito seja requisitado em até 7 (sete) dias da data da compra, até o limite de 48 (quarenta e oito) horas antes do evento. <br><br> 3.2. O estorno do valor do ingresso somente será efetuado mediante o envio de um documento escrito de próprio punho pelo cliente solicitando o reembolso do valor do ingresso para o endereço informado abaixo*. O estorno na fatura do cartão seguirá as normas de cada operadora/banco emissor, podendo ser creditado na fatura seguinte ou na subsequente, de acordo com a data de fechamento da fatura. <br><br> * Endereço para envio do documento escrito, em caso de pedido de estorno: <br> BTS INFORMA FEIRAS EVENTOS E EDITORA LTDA<br> Rua Bela Cintra, 967 - 11º andar - Conj. 112-A – Cerqueira Cesar – São Paulo - SP<br> CEP: 01415-003 <br><br> A/C Departamento de marketing – ABF Expo <br/><br/> 3.3. Em caso de cancelamento do evento por parte da Organização, o valor do ingresso será devolvido. <br><br> <strong>4 – DA RETIRADA DO INGRESSO</strong><br> 4.1. Para sua segurança, todas as compras, via internet, somente poderão ser entregues ao proprietário do número do CPF utilizado na compra, mediante a apresentação dos seguintes documentos: comprovante de pagamento da compra e original de um dos seguintes documentos de identificação com foto e dentro do prazo de validade: Cédula de Identidade (RG) que contenha o número do Cadastro da Pessoa Física (CPF), Carteira de Órgão ou Conselho de Classe, Carteira Nacional de Habilitação (CNH). 4.2. O cliente deverá retirar o seu ingresso nos guichês de atendimento exclusivos para quem comprou o ingresso online, os quais estarão localizados na entrada do evento durante o período de realização deste. Local: Expo Center Norte - Pavilhões Azul e Branco - São Paulo Rua José Bernardo Pinto, 333 - Vila Guilherme, sem a cobrança de taxa de conveniência. <br><br> 4.3. Para a retirada de ingressos por terceiros, este deverá apresentar instrumento de procuração do titular do pedido com firma reconhecida em cartório e poderes específicos para a retirada dos ingressos, a qual ficará retida, devendo, ainda, o outorgado/representante comparecer portando seus documentos originais de identificação. 4.4. Ocorrendo à impossibilidade de retirada do ingresso, por ausência de pessoa habilitada para o recebimento, não ocorrerá qualquer devolução de valores, sendo que o comprador está ciente dos Termos e Condições. <br><br><strong>5 - DAS CONDIÇÕES GERAIS DE USO</strong><br> 5.1. Terceira idade: Pessoas com mais de 60 anos podem adquirir seu ingresso com desconto de 50% de acordo com o Estatuto do Idoso (Lei nº10.741/2003, cap V, art 23). Um terceiro poderá adquirir os ingressos mediante a apresentação do documento original do idoso. <br><br>5.2. A meia-entrada somente será vendida na bilheteria do evento mediante apresentação do documento comprobatório pelo próprio cliente. <br><br>5.3. Cada ingresso é válido para os 04 dias da feira, podendo ser adquiridos antecipadamente por meio do site do evento ao custo de R$ 60,00 até dia 20 de junho de 2017. Caso deixe para comprar durante a realização do evento, de 21 a 24 de junho de 2017, na bilheteria local ou pelo site custará R$ 70,00. <br><br>5.4. Visando a proteção dos direitos dos clientes e de terceiros, caso ocorra à tentativa ou a efetiva utilização indevida dos serviços de conveniência de compra online de ingressos, a proprietária do evento poderá indicar os dados do cliente às autoridades públicas, aos serviços de proteção ao crédito, dentre outros, para início dos procedimentos legais e administrativos cabíveis. Excluindo-se os casos de atos ilícitos mencionados acima, a proprietária do evento se compromete a não divulgar, ceder, vender ou transferir a terceiros os dados pessoais fornecidos pelo cliente. <br><br>5.5. O ato de credenciamento e compra de ingressos para a feira ABF Franchising Expo não dá direito à participação nas Apresentações, Congressos e Seminários da Franchising Week, que ocorrerão de 19 a 23 de junho, nas salas Cantareira - 2º andar do Expo Center Norte. Este evento paralelo exige credenciamento específico e seu conteúdo está direcionado ao público de expositores e franqueadores. <br><br>5.6. Caso qualquer disposição do presente Termos e Condições seja considerada nula ou sem efeito, não resultará na nulidade total dos Termos e Condições, permanecendo em vigor nas demais disposições, permanecendo os direitos e obrigações ora estipulados. <br><br>5.7. Fica eleito o foro da Comarca de São Paulo, Estado de São Paulo, para dirimir quaisquer dúvidas oriundas dos Termos e Condições, excluindo-se qualquer outro foro, por mais privilegiado que seja.",
      en: "<strong>TERMS AND CONDITIONS </strong> <br> <br> Before finalizing your purchase, know the Terms and Conditions of sale of the ticket for the event ABF Franchising Expo 2017: <br> <br> <strong> 1 - DO OBJECT </strong> <br> 1.1. The purchase will be subject to the availability of tickets and the approval of the operator of your credit card. Children under 16 years old will not be allowed to enter unaccompanied. <br> <br> <strong> 2 - THE ACQUISITION OF THE TICKET </strong> <br> 2.1. Please carefully check your order details before confirming your purchase. No cancellation or refund will be allowed. Your ticket is a unique product, meaning after your purchase it will no longer be available for sale. <br><br> <strong> 3 - DENIAL OF PURCHASE AND/OR CANCELLATION </strong> <br> 3.1. In case of client's regret, the refund of the ticket amount will be made, provided that this right is requested within 7 (seven) days of the date of purchase, up to the limit of 48 (forty eight) hours before the event. <br><br> 3.2. The refund of the value of the ticket will only be made by sending a written document of own hand by the client requesting the refund of the value of the ticket to the address informed below *. The reversal on the invoice of the card will follow the rules of each operator/issuing bank and can be credited to the next or subsequent invoice, according to the closing date of the invoice. <br><br> * Address to send the written document, in case of a request for reversal: <br> BTS INFORMA FAIRAS EVENTOS E EDITORA LTDA <br> Rua Bela Cintra, 967 - 11º andar - Conj. 112-A - Cerqueira Cesar - São Paulo - SP <br> CEP: 01415-003 <br> <br> A / C Marketing Department - ABF Expo <br/><br/> 3.3. In case of cancellation of the event by the Organization, the ticket amount will be refunded. <br><br> <strong> 4 - WITHDRAWAL OF THE TICKET </strong><br> 4.1. For your safety, all purchases, via the Internet, can only be delivered to the owner of the CPF number used in the purchase, by presenting the following documents: proof of payment of the purchase and original of one of the following identification documents with photo and inside Of the period of validity: Identity Card (RG) that contains the Individual Registration Number (CPF), Organ Portfolio or Class Council, National Driver's License (CNH). 4.2. The customer should withdraw his ticket from the exclusive ticket booths for those who bought the ticket online, which will be located at the entrance of the event during the period of its completion. Location: Expo Center Norte - Pavilhões Azul e Branco - São Paulo Rua José Bernardo Pinto, 333 - Vila Guilherme, without the charge of convenience fee. <br><br> 4.3. For the withdrawal of tickets by third parties, this must present a proxy instrument of the holder of the application with a notarized signature and specific powers for the withdrawal of the tickets, which will be retained, and the grantor/representative must also be present with the original documents Of identification. 4.4. In the event of the impossibility of withdrawing the ticket, due to the absence of a person authorized to receive, no refund of value will occur, and the buyer is aware of the Terms and Conditions. <br><br><strong>5 - GENERAL CONDITIONS OF USE </strong> <br><br> 5.1. Senior Citizens: People over 60 can purchase their ticket at a discount of 50% in accordance with the Statute of the Elderly (Law nº 10.741 / 2003, cap V, art 23). A third party may purchase the tickets by presenting the original document of the senior citizen. <br><br> 5.2. The half-ticket will only be sold at the box office of the event upon presentation of the supporting document by the customer. <br><br> 5.3. Each ticket is valid for the 04 days of the fair, and can be purchased in advance through the website of the event at a cost of R$ 60.00 until June 20, 2017. If you leave to buy during the event, from 21 to June 24, 2017, at the local box office or through the website will cost R$ 70.00. <br><br>5.4. In order to protect the rights of customers and third parties, in case of attempted or undue use of convenience services for online purchase of tickets, the owner of the event may indicate the client's data to public authorities, credit protection services , Among others, to initiate legal and administrative procedures. Excluding the cases of illicit acts mentioned above, the owner of the event undertakes not to disclose, assign, sell or transfer to third parties the personal data provided by the client. Page 5 The accreditation and purchase of tickets for the ABF Franchising Expo will not entitle you to participate in the Franchising Week Presentations, Congresses and Seminars, which will take place from 19 to 23 June, in the Cantareira rooms - 2nd floor of Expo Center Norte. This parallel event requires specific credentialing and its content is targeted to the public of exhibitors and franchisors. 5.6. Should any provision of these Terms and Conditions be considered null or void, it will not result in the total nullity of the Terms and Conditions, remaining in force in the other provisions, remaining the rights and obligations stipulated herein. <br><br> 5.7. It is elected the forum of the Region of São Paulo, State of São Paulo, to resolve any doubts arising from the Terms and Conditions, excluding any other forum, however privileged it may be.",
	  es: "<strong> Términos y Condiciones </strong> <br> <br> Antes de finalizar su compra conocer Términos y Condiciones La venta de entradas para el evento de 2017 ABF Franchising Expo: <br> <strong> 1 - DO OBJETO </strong> Filmografía 1.1. La adquisición está sujeta a la disponibilidad de entradas y aprobación por parte del operador de su tarjeta de crédito. No se le permitirá entrar en menores de 16 años no acompañados. <br><br> <strong> 2 - ADMISIÓN DE ADQUISICIÓN </strong> Filmografía 2.1. Comprobar cuidadosamente los datos de su pedido antes de confirmar su compra. No se le permitirá cancelar o billetes de vuelta. Su entrada es de un solo producto, es decir, después de su compra ya no estará disponible para la venta. <br><br> <strong> 3 - ADQUISICIÓN DE RETIRADA Y/O CANCELACIÓN </strong> <br> 3.1. En caso de arrepentimiento del comprador, se efectuará el ingreso de la cantidad de reembolso, siempre y cuando así lo solicite dentro de los siete (7) días a partir de la fecha de compra, hasta el límite de 48 (cuarenta y ocho) horas antes del evento. <br><br> 3.2. La reversión del valor de entradas sólo se hará mediante el envío de un informe escrito por el propio mango documento que solicita la devolución del valor del boleto a la dirección que aparece a continuación *. La inversión en la factura de la tarjeta va a seguir las reglas de cada operador/banco emisor, puede ser acreditado a la cuenta siguiente o subsiguiente, de acuerdo a la fecha de cierre de la factura. <br><br> * Dirección para el envío del documento escrito, en caso de solicitud de reversión <br> BTS INFORMA FERIAS Y EVENTOS Publishing Ltd <br> Rua Bela Cintra, 967 - piso 11 - Conj. 112-A - Cerqueira César - Sao Paulo - SP <br> CEP: 01415-003 <br> departamento de A / C de Marketing - ABF Expo <br/><br/> 3.3. En caso de cancelación de eventos por la Organización, se devolverá el valor del boleto. <br><br> <strong> 4 - LA RETIRADA DE ENTRADA </strong><br> 4.1. Para su seguridad, todas las compras a través de Internet, sólo pueden ser entregados al número de seguridad social del propietario utilizada en la compra, mediante la presentación de los siguientes documentos: comprobante de pago de la compra original y uno de los siguientes documentos de identificación con foto y dentro de fecha de caducidad: Tarjeta de identidad (RG) que contiene el número del Registro de las personas Físicas (CPF), la cartera de órgano o consejo de clase, carné de conducir nacional (CNH). 4.2. El cliente deberá recoger las entradas en las taquillas de servicio exclusivo para aquellos que han comprado el billete en línea, que se ubicará en el caso de entrada durante el período de aplicación de la presente. Lugar: Expo Center Norte - Pabellón Azul y Blanco - Sao Paulo Rua José Bernardo Pinto, 333 - Vila Guilherme sin cobrar cuota de conveniencia. <br><br> 4.3. Para la retirada de billetes por parte de terceros, lo que deberá presentar el poder de la solicitud del titular del abogado con una firma reconocida y facultades específicas para retirar las entradas, que será retenido y también debe concedido / asistir representativa teniendo sus documentos originales ID. 4.4. Dándose la imposibilidad de retirada de la admisión, por ninguna persona que tenga derecho a recibir, no habrá valores de retorno, y el comprador es consciente de los términos y condiciones. <br><br><strong>5 - USO DE LAS CONDICIONES GENERALES </strong> 5.1. De la tercera edad: Las personas mayores de 60 pueden comprar el billete con un descuento del 50% de acuerdo con el Estatuto de edad avanzada (Ley nº10.741 / 2003, Capítulo V, artículo 23). Una tercera puede comprar boletos presentando el original del documento antiguo. <br><br> 5.2. La entrada de la mitad se venderá sólo en la taquilla del evento en la presentación de pruebas documentales por parte del cliente. <br><br> 5.3. Cada billete es válido durante 04 días de la feria y se pueden adquirir con antelación a través de la página web del evento con un costo de R$ 60,00 hasta el 20 de junio de 2017. Si deja de comprar durante el evento, del 21 al 24 de de junio de, 2017, en la taquilla local o en el sitio tendrá un costo de R$70,00. <br><br> 5.4. Con el fin de proteger los derechos de los clientes y terceros en el caso de la tentativa o la utilización indebida de los ticket de servicio conveniencia de compras en línea, el propietario del evento puede indicar los datos del cliente a las autoridades públicas, servicios de protección de crédito , entre otros, para el inicio de los procedimientos legales y administrativas aplicables. Excluyendo los casos de actos ilegales mencionados anteriormente, el propietario del evento se compromete a no divulgar, ceder, vender o transferir a terceros los datos personales proporcionados por el cliente. <br><br> 5.5. El acto de billetes de acreditación y de compra de la feria ABF Franchising Expo no da derecho a la participación en presentaciones, conferencias y seminarios para franquicias Semana, que se llevará a cabo del 19 al 23 de junio en salas de Cantareira - 2 ° piso de la Expo Center Norte. Este evento paralelo requiere autorización específica y su contenido se dirige a la audiencia de expositores y franquiciadores. <br><br> 5.6. Si se considera cualquier disposición de estos Términos y Condiciones nulo no dará lugar a la nulidad total de los Términos y Condiciones permanecerán en vigor, en otras disposiciones derechos y obligaciones estipulados en el presente documento restantes. <br><br> 5.7. Es el foro elegido de la región de Sao Paulo, Estado de Sao Paulo, para resolver cualquier asunto que surja de los Términos y Condiciones, con exclusión de cualquier otra jurisdicción, sin embargo privilegiado."
    },
	text10: {
	  pt: "Recusar",
      en: "Refuse",
	  es: "Rechazar"
    },
	text11: {
	  pt: "Aprovar",
      en: "Accept",
	  es: "Aceptar"
    }
	
  };
   var cookieLang = readCookie("lang");
  if(cookieLang == null)
	var _t = $('body').translate({lang: "pt", t: t});
   else
   var _t = $('body').translate({lang: cookieLang, t: t});
   
  var str = _t.g("translate");

  
  $(".lang_selector").click(function(ev) {
    var lang = $(this).attr("data-value");
    _t.lang(lang);
	createCookie("lang",lang,100);
    ev.preventDefault();
  });
});
  </script>

<%

'response.end
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
    response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
End If

ID_Edicao               = Session("cliente_edicao")
Idioma                  = Session("cliente_idioma")
ID_TP_Credenciamento    = Session("cliente_tipo")
TP_Formulario           = Session("cliente_formulario")
IRC                     = Session("cliente_cadastro")
ID_Empresa              = Session("cliente_empresa")
ID_Visitante            = Session("cliente_visitante")
Nome_Visitante          = Session("cliente_nome")
CPF_Visitante           = Session("cliente_cpf")

	' Verifica Idioma a ser apresentado
	Select Case (Idioma)
		Case "1"
			SgIdioma = "PTB"
		Case "2"
			SgIdioma = "ESP"
		Case "3"
			SgIdioma = "ENG"
		Case Else
			SgIdioma = "PTB"
	End Select

	Pagina_ID 	= 2

	SQL_Textos	=	" Select " &_
					"	ID_Texto, " &_
					"	ID_Tipo, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					" From Paginas_Textos " &_
					" Where  " &_
					"	ID_Idioma = " & idioma & " " &_
					"	AND ID_Pagina = " & Pagina_ID & " " &_
					" Order By Ordem "
	Set RS_Textos = Server.CreateObject("ADODB.Recordset")
	RS_Textos.Open SQL_Textos, Conexao

	If not RS_Textos.BOF or not RS_Textos.EOF Then
		total_registros = 0
		While not RS_Textos.EOF
			total_registros = total_registros + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.MoveFirst
		ReDim textos_array(total_registros-1)
		n = 0
		While not RS_Textos.EOF
			id = RS_Textos("id_texto")
			ident = RS_Textos("identificacao")
			texto = RS_Textos("texto")
			url_img = RS_Textos("url_imagem")
			textos_array(n) = Array(id, ident, texto, url_img)
			n = n + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.Close
	End If


'	For i = Lbound(textos_array) to Ubound(textos_array)
'		response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'	Next
'===========================================================
%>
<% If Limpar_texto(Request("teste")) = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
'	response.write(SQL_Img_Faixa & "<br>")
	Set RS_Img_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Img_Faixa.CursorType = 0
	RS_Img_Faixa.LockType = 1
	RS_Img_Faixa.Open SQL_Img_Faixa, Conexao
		img_faixa = RS_Img_Faixa("img_faixa")
	RS_Img_Faixa.Close

	' Faixa TOPO
	SQL_Faixa	= 	"Select " &_
					"	Cor, " &_
					"	Logo_Negativo, " &_
					"	Faixa_Fundo " &_
					"From Edicoes_configuracao " &_
					"Where  " &_
					"	ID_Edicao = " & ID_Edicao
	Set RS_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Faixa.CursorType = 0
	RS_Faixa.LockType = 1
	RS_Faixa.Open SQL_Faixa, Conexao

		faixa_cor	= RS_Faixa("cor")
		faixa_logo	= RS_Faixa("logo_negativo")
		faixa_fundo	= RS_Faixa("Faixa_Fundo")
	RS_Faixa.Close

	' Select de Eventos
	SQL_Evento	=	"SELECT " &_
					"	Nome_" & SgIdioma & " AS Evento, " &_
					"	Ano " &_
					"FROM Eventos as E " &_
					"INNER JOIN" &_
					"	Eventos_Edicoes as EE " &_
					"ON EE.ID_Evento = E.ID_Evento " &_
					"WHERE " &_
					"	E.Ativo = 1 " &_
					"	AND EE.ID_Edicao = " & ID_Edicao

	Set RS_Evento = Server.CreateObject("ADODB.Recordset")
	RS_Evento.CursorType = 0
	RS_Evento.LockType = 1
	RS_Evento.Open SQL_Evento, Conexao

	Evento = RS_Evento("Evento") & " " & RS_Evento("Ano")
	Rs_Evento.Close





	'Após o comprovante, caso seja realizado, enviar o email de comprovante
	Session("cliente_enviar_email") = 0

%>
<script language="javascript">
var idioma_atual 	= '<%=Session("cliente_idioma")%>';
var select       	= '<%=textos_array(36)(2)%>';
var cor_fundo 	 	= '<%=faixa_cor%>';
var tp_formulario 	= '';

$(document).ready(function(){
	var erro = '<%=erro%>';
	switch (erro) {
		case '1':
			$('#txt_topo').html('Campos Vazios');
			$('#aviso_topo').show();
			break;
		case '2':
			$('#txt_topo').html('Dados incorretos');
			$('#aviso_topo').show();
			break;
	}
});
</script>
</head>

<body>
<div style="width: 100%; position: absolute; height:750px; float:left; z-index:100; background:#CCC; display:none" id="loading" class="transparent">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
<!--#include virtual="/includes/cabecalho.asp"-->
<div style="width: 100%; position: absolute; left:0px; float:left; z-index:10; height: 115px;" id="faixa_selecionada">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center" height="45">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/faixa_fundo_esq.gif); height:45px; width:100%; margin-top:50px;"></div>
    <!-- Faixa Lateral -->
    </td>
    <td width="870" align="center">
    <!-- Faixa -->
        <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="50">&nbsp;</td>
          </tr>
          <tr>
            <td>
				<table width="870" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="189" height="45" background="/img/geral/faixa_fundo_sesq.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45"></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" style="background-repeat:repeat-x; position:relative;" class="atencao_13px cor_branco">
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"></div>
                        <div style="position:absolute; top:-45px; right:0px;" align="right"><img id="img_logo_selecionado" src="<%=faixa_logo%>" hspace="10"></div>
                    </td>
                  </tr>
                </table>
            </td>
          </tr>
      </table>
    <!-- Faixa -->
    </td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Lateral -->
    	<div style="background:url(<%=faixa_fundo%>); height:45px; width:100%; margin-top:50px;" id="faixa_dir"></div>
    <!-- Faixa Lateral	 -->
    </td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="center" valign="top">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left; display:;" id="conteudo">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="130" colspan="3">&nbsp;</td>
  </tr>
</table>
    <!-- Form Container -->
    <div id="contForm">
    <!-- Form -->

            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div><br/>
            <!-- End Alert error -->

            <table width="850" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b>Olá</b> <%=Nome_Visitante%></td>
                    <td width="500" height="30" bgcolor="#414042" align="right"><img src="/img/botoes/voltar.gif" width="47" height="15" hspace="5" class="cursor" onClick="link('status.asp');"></td>
                    <td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="sair();"></td>
                </tr>
            </table>
            <br/>

            	<%
				SQL_Consulta_Pedidos = 	"Select " &_
										"	P.* " &_
										"From " &_
										"	Pedidos As P " &_

										"Where " &_
										"	P.ID_Edicao = '" & Session("cliente_edicao") & "' " &_
										"	And P.ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "' " &_
										"	And P.ID_Visitante = '" & Session("cliente_visitante")  & "' " &_
										"	And P.Status_Pedido = 1"
				'Response.Write(SQL_Consulta_Pedidos)

				Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3

				If Not RS_Consulta_Pedidos.Eof Then

					Tickets 		= True
					Numero_Pedido 	= RS_Consulta_Pedidos("Numero_Pedido")
					'Session("pedido") = RS_Lista_Pedidos("Numero_Pedido")
					ID_Pedido 		= RS_Consulta_Pedidos("ID_Pedido")
					Idioma_Pedido	= RS_Consulta_Pedidos("ID_Idioma")
					Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)


				Else

					Tickets = False

				End If
				%>

				<!--#Include virtual="/tickets/menu_lateral.asp"-->

        <form action="/tickets/pedido.asp?acao=novo" method="post" id="termos_condicoes" name="termos_condicoes" >
            <fieldset style="float: right; width: 580px;">

	            <input type="hidden" value="0" name="aceito" id="aceito">

            	<legend class="trn" data-trn-key="text8">Termos & Condições</legend>
				<div id="parcAssis" class="div_parceria" style="width: 580px;">

                	<div style="width: 100%; height: 250px; overflow: auto; line-height: 150%" class="trn" data-trn-key="text9">
<strong>TERMOS E CONDIÇÕES</strong><br><br>
Antes de finalizar sua compra conheça os Termos e Condições de venda de Ingresso para o evento ABF Franchising Expo 2017:
<br><br>
<strong>1 - DO OBJETO</strong><br>
1.1. A compra estará sujeita à disponibilidade de ingressos e à aprovação da operadora de seu cartão de crédito. Não será permitida a entrada de menores de 16 anos desacompanhados.
<br><br>
<strong>2 - DA AQUISIÇÃO DO INGRESSO</strong><br>
2.1. Confira atentamente os dados de seu pedido antes da confirmação de sua compra. Não será permitido cancelamento ou devolução de ingressos. O seu ingresso é um produto único, ou seja, após sua compra ele não estará mais disponível para venda.
<br><br>
<strong>3 - DA DESISTÊNCIA DA COMPRA E/OU CANCELAMENTO</strong><br>
3.1. Em caso de arrependimento do cliente, o reembolso do valor do ingresso será efetuado, desde que esse direito seja requisitado em até 7 (sete) dias da data da compra, até o limite de 48 (quarenta e oito) horas antes do evento.
<br><br>
3.2. O estorno do valor do ingresso somente será efetuado mediante o envio de um documento escrito de próprio punho pelo cliente solicitando o reembolso do valor do ingresso para o endereço informado abaixo*. O estorno na fatura do cartão seguirá as normas de cada operadora/banco emissor, podendo ser creditado na fatura seguinte ou na subsequente, de acordo com a data de fechamento da fatura.
<br><br>
* Endereço para envio do documento escrito, em caso de pedido de estorno: <br>
BTS INFORMA FEIRAS EVENTOS E EDITORA LTDA<br>
Rua Bela Cintra, 967 - 11º andar - Conj. 112-A – Cerqueira Cesar – São Paulo - SP<br>
CEP: 01415-003
<br><br>
A/C Departamento de marketing – ABF Expo
<br/><br/>
3.3. Em caso de cancelamento do evento por parte da Organização, o valor do ingresso será devolvido.
<br><br>
<strong>4 – DA RETIRADA DO INGRESSO</strong><br>
4.1. Para sua segurança, todas as compras, via internet, somente poderão ser entregues ao proprietário do número do CPF utilizado na compra, mediante a apresentação dos seguintes documentos: comprovante de pagamento da compra e original de um dos seguintes documentos de identificação com foto e dentro do prazo de validade: Cédula de Identidade (RG) que contenha o número do Cadastro da Pessoa Física (CPF), Carteira de Órgão ou Conselho de Classe, Carteira Nacional de Habilitação (CNH).
4.2. O cliente deverá retirar o seu ingresso nos guichês de atendimento exclusivos para quem comprou o ingresso online, os quais estarão localizados na entrada do evento durante o período de realização deste. Local: Expo Center Norte - Pavilhões Azul e Branco - São Paulo
Rua José Bernardo Pinto, 333 - Vila Guilherme, sem a cobrança de taxa de conveniência.
<br><br>
4.3. Para a retirada de ingressos por terceiros, este deverá apresentar instrumento de procuração do titular do pedido com firma reconhecida em cartório e poderes específicos para a retirada dos ingressos, a qual ficará retida, devendo, ainda, o outorgado/representante comparecer portando seus documentos originais de identificação.
4.4. Ocorrendo à impossibilidade de retirada do ingresso, por ausência de pessoa habilitada para o recebimento, não ocorrerá qualquer devolução de valores, sendo que o comprador está ciente dos Termos e Condições.
<strong>5 - DAS CONDIÇÕES GERAIS DE USO</strong><br>
5.1. Terceira idade: Pessoas com mais de 60 anos podem adquirir seu ingresso com desconto de 50% de acordo com o Estatuto do Idoso (Lei nº10.741/2003, cap V, art 23). Um terceiro poderá adquirir os ingressos mediante a apresentação do documento original do idoso.
<br><br>5.2. A meia-entrada somente será vendida na bilheteria do evento mediante apresentação do documento comprobatório pelo próprio cliente.
<br><br>5.3. Cada ingresso é válido para os 04 dias da feira, podendo ser adquiridos antecipadamente por meio do site do evento ao custo de R$ 60,00 até dia 20 de junho de 2017. Caso deixe para comprar durante a realização do evento, de 21 a 24 de junho de 2017, na bilheteria local ou pelo site custará R$ 70,00.
<br><br>5.4. Visando a proteção dos direitos dos clientes e de terceiros, caso ocorra à tentativa ou a efetiva utilização indevida dos serviços de conveniência de compra online de ingressos, a proprietária do evento poderá indicar os dados do cliente às autoridades públicas, aos serviços de proteção ao crédito, dentre outros, para início dos procedimentos legais e administrativos cabíveis. Excluindo-se os casos de atos ilícitos mencionados acima, a proprietária do evento se compromete a não divulgar, ceder, vender ou transferir a terceiros os dados pessoais fornecidos pelo cliente.
<br><br>5.5. O ato de credenciamento e compra de ingressos para a feira ABF Franchising Expo não dá direito à participação nas Apresentações, Congressos e Seminários da Franchising Week, que ocorrerão de 19 a 23 de junho, nas salas Cantareira - 2º andar do Expo Center Norte. Este evento paralelo exige credenciamento específico e seu conteúdo está direcionado ao público de expositores e franqueadores.
<br><br>5.6. Caso qualquer disposição do presente Termos e Condições seja considerada nula ou sem efeito, não resultará na nulidade total dos Termos e Condições, permanecendo em vigor nas demais disposições, permanecendo os direitos e obrigações ora estipulados.
<br><br>5.7. Fica eleito o foro da Comarca de São Paulo, Estado de São Paulo, para dirimir quaisquer dúvidas oriundas dos Termos e Condições, excluindo-se qualquer outro foro, por mais privilegiado que seja.
                    </div>

					<br><br>

					<a href="#_" onclick="aprovacao(0);"><div id="loader_2" class="bt_recusar_termo trn" style="float: left; display: block" data-trn-key="text10">Recusar</div></a>
					<a href="#_" onclick="aprovacao(1);"><div id="loader_2" class="bt_aceitar_termo trn" style="float: right; display: block" data-trn-key="text11">Aprovar</div></a>

                    <br>
                </div>
            </fieldset>
            <br/>

            <!-- Alert error -->
            <div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px;">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
        </form>
        <!-- Form End -->
	</div>

    <script language="javascript">
    	function aprovacao(termo){

			document.getElementById("aceito").value = termo;
			document.getElementById("termos_condicoes").submit();

			}
    </script>
    <!-- End Form Container -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute;float:left; display:none; z-index:100" id="loading">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
</body>
</html>
<%
Conexao.Close
%>
