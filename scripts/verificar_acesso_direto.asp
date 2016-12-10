<%
' Funcao para Verificar Parametros
function Soma(Out)
	
	' Setando Array
	Dim AcessoExterno

	' Separando Valores
	AcessoExterno = Split(Out,",")

	' Setando valores nas Variaveis
	Cliente_Edicao 		= AcessoExterno(0)
	Cliente_Idioma 		= AcessoExterno(1)
	Cliente_Tipo 		= AcessoExterno(2)
	Cliente_Formulario 	= AcessoExterno(3)

	' Verificando se a Edicao Existe
	'===========================================================
	' Listagem de Feiras por DATA
	SQL_Feiras	= 	"Select " &_
					"	Distinct " &_
					"	Ee.ID_Edicao, " &_
					"	Ecv.Cor, " &_
					"	Ecv.Logo_Box, " &_
					"	Ecv.Logo_Negativo, " &_
					"	Ecv.Faixa_Fundo, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano, " &_
					"	Ee.Data_Inicio_Feira " &_
					"From Edicoes_Configuracao as Ecv  " &_
					"Inner Join Eventos_Edicoes as Ee ON Ee.ID_Edicao = Ecv.ID_Edicao  " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento  " &_
					"Inner Join Edicoes_Tipo as Et ON Et.ID_Edicao = Ecv.ID_Edicao " &_
					"Where " &_
					"	Ecv.Ativo = 1 " &_
					"	AND Ee.Ativo = 1 " &_
					"	AND E.Ativo = 1 " &_
					"	AND Et.Ativo = 1 " &_
					"Order by Ee.Data_Inicio_Feira, Evento "
					"	AND getDate() >= Et.Inicio " &_
					"	AND getDate() <= Et.Fim " &_	
	response.write("<b>SQL_Feiras</b><br>" & SQL_Feiras & "<hr>")
	
	Set RS_Feiras = Server.CreateObject("ADODB.Recordset")
	RS_Feiras.CursorType = 0
	RS_Feiras.LockType = 1
	RS_Feiras.Open SQL_Feiras, Conexao, 1
'===========================================================


	' Verificando se o Idioma esta Disponivel


	' Verificando o Tipo do Formulário


	' Verificando Cliente Formulário




	Session("cliente_edicao") 		= AcessoExterno(0)
	Session("cliente_idioma") 		= AcessoExterno(1)
	Session("cliente_tipo") 		= AcessoExterno(2)
	Session("cliente_formulario") 	= AcessoExterno(3)

End Function
%>