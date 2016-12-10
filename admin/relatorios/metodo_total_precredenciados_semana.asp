<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%

'===================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'===================================================

QuantidadeRegistros = " TOP 100 "

'===================================================
' 00 - Verificando todos os Registros da Feira
'===================================================
SQL_Dados = "SELECT " & QuantidadeRegistros & " " &_
			"	E.Nome_PTB as Evento " &_
			"	,RC.ID_Relacionamento_Cadastro  as RelacionamentoCadastro " &_
			"	,TC.Nome as TipoCredenciamento " &_
			"	,RC.Data_Cadastro as DataCadastro " &_
			"FROM  " &_
			"	Relacionamento_Cadastro as RC " &_
			"	INNER JOIN Eventos_Edicoes as EE " &_
			"		ON RC.ID_Edicao = EE.ID_Edicao " &_
			"	INNER JOIN Eventos as E " &_
			"		ON EE.ID_Evento = E.ID_Evento " &_
			"	INNER JOIN Tipo_Credenciamento as TC " &_
			"		ON RC.ID_Tipo_Credenciamento = TC.ID_Tipo_Credenciamento " &_
			"WHERE  " &_
			"	EE.ID_Edicao = 24 " &_
			"	AND EE.Ano = 2013 " &_
			"ORDER BY " &_
			"	RC.Data_Cadastro ASC"
response.Write("<hr><strong>SQL_Dados:</strong><hr>" & SQL_Dados & "<hr>")
Set RS_Dados = Server.CreateObject("ADODB.Recordset")
	RS_Dados.CursorType = 0
	RS_Dados.LockType = 3
RS_Dados.Open SQL_Dados, Conexao 

' Verificando se existem dados
If not RS_Dados.BOF or not RS_Dados.EOF Then
	
	' Contador Total de Registros
	ContarRegistro = 0
	
	' Looping dos Resultados
	While not RS_Dados.EOF 
		
		ContarRegistro = ContarRegistro + 1
		
			response.Write("<hr>" & RS_Dados("TipoCredenciamento") & " - " & RS_Dados("DataCadastro") & "<hr>")
		
		RS_Dados.MoveNext
		Wend
	RS_Dados.Close		
	
	response.Write(ContarRegistro)
			
Else
	response.Write("Nenhum regostro encontrado!")
End If
%>
