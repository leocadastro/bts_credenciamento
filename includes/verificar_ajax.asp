<!-- Detect support for cookies and Ajax and display message if not -->        
<div id="supportError" class="atencao_13px cor_vermelho">
	<noscript>
        <br>
        <center><img src="/img/alert/important.gif" width="32" height="32"></center><br>
        JavaScript parece estar desativado ou n&atilde;o suportado pelo seu navegador.<br>
        Esta web aplica&ccedil;&atilde;o requer JavaScript para funcionar corretamente.<br>
        Favor habilitar o JavaScript nas configura&ccedil;&otilde;es do navegador,<br>
        ou atualizar para um navegador com suporte a JavaScript e tente novamente.<hr size=1>
    </noscript>
    
    <script type="text/javascript">
    <!--
    
    if (!browserSupportsCookies())
    {
        var msg = '<br><center><img src="/img/alert/important.gif" width="32" height="32"></center><br>';
		msg += 'Os cookies parecem estar desativado ou n&atilde;o suportado pelo seu navegador.<br> ';
        msg += 'Esta web aplica&ccedil;&atilde;o requer Cookies para funcionar corretamente.<br> ';
        msg += 'Ative os cookies em seu navegador ';
        msg += 'ou atualizar para um navegador com suporte a cookies e tente novamente.<hr size=1>'
        
        document.write(msg);
    }else{
		//document.write("<br>Seu navegador suporta cookies");
	}
    
    if (!browserSupportsAjax())
    {
        var msg = '<br><center><img src="/img/alert/important.gif" width="32" height="32"></center><br>';
        msg += 'Seu navegador parece n&atilde;o suportar a tecnologia Ajax.<br>';
        msg += 'Esta web aplica&ccedil;&atilde;o requer Ajax para funcionar corretamente.<br>';
        msg += 'Por favor, atualize para um navegador com suporte a Ajax e tente novamente.<hr size=1>';
        
        document.write(msg);
    }else{
		//document.write("<br>Seu navegador suporta Ajax");
	}
        
    if (!ActiveXEnabledOrUnnecessary())
    {
        var msg = '<hr size=1>ActiveX parece ser desativado no seu browser ';
        msg += 'This web application requires Ajax technology to function properly. ';
        msg += 'Esta web aplicação requer tecnologia Ajax para funcionar corretamente. ';
        msg += 'Por favor ativar o ActiveX nas suas configurações de segurança do navegador ';
        msg += 'ou atualizar para um navegador com suporte a Ajax e tente novamente.';
        
        //document.write(msg);
    }else{
		//document.write("<br>Seu navegador suporta Activex");
	}
    -->
    </script>

</div>