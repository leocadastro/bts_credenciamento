<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
DIM StartTime
Dim EndTime
StartTime = Timer()

Server.ScriptTimeout = 999999999

id	= Limpar_Texto(Request("id"))

If IsNumeric(id) = false Then response.Redirect("default.asp?msg=erro_nao_encontrado")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

	SQL_Evento_Autorizado = "Select " &_
							"	EE.Ano " &_
							"	,E.Nome_PTB as Feira " &_
							"From Administradores_Edicoes AE " &_
							"Inner Join Eventos_Edicoes as EE ON Ae.ID_Edicao = EE.ID_Edicao " &_
							"Inner Join Eventos as E ON E.ID_Evento = EE.ID_Evento " &_
							"Where  " &_
							"	Ae.ID_Admin = " & session("admin_id_usuario") & " " &_
							"	AND Ae.ID_Edicao = " & ID
	Set RS_Evento_Autorizado = Server.CreateObject("ADODB.Recordset")
	RS_Evento_Autorizado.Open SQL_Evento_Autorizado, Conexao
	
	If RS_Evento_Autorizado.BOF or RS_Evento_Autorizado.EOF Then
		response.Redirect("default.asp?msg=erro_nao_autorizado")
	Else
		Feira = Replace(RS_Evento_Autorizado("Ano") & "-" & RS_Evento_Autorizado("Feira"), " ", "_")
		Feira = Replace(Feira, "&", "")
		RS_Evento_Autorizado.Close
	End If

%>
<!--#include virtual="/admin/inc/acentuacao.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administração Cred. 2012</title>
</head>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
});
</script>

<body>
<!--#include virtual="/admin/inc/menu_top.asp"-->
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/admin/images/img_tabela_branca_top.jpg" width="955" height="15" /></td>
  </tr>
</table>
<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="left" bgcolor="#FFFFFF" class="conteudo_site">
    
    <table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100" height="50">&nbsp;</td>
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Evento: <%=nome_ptb%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso Processando, por favor não interrompa.</span></div>

      <div style="padding-left:150px; padding-right:150px; width:600px;">

                <%
            
                SQL_Cadastros = "SET TRANSACTION ISOLATION LEVEL REPEATABLE READ; " &_
								"BEGIN TRANSACTION; " &_
								"	Select " &_
                                "	RC.ID_Relacionamento_Cadastro " &_
                                "	,E.CNPJ " &_
                                "	,E.Razao_Social as Empresa " &_
                                "	,E.Nome_Fantasia as NomeFantasia " &_
                                "	,V.Nome_Completo " &_
                                "	,V.Nome_Credencial as Credencial " &_
                                "	,V.CPF " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select " &_
                                "			RT.DDI + '' " &_
                                "		From Relacionamento_Telefones as RT " &_
                                "		Inner Join Tipo_Telefone as TT ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as DDI " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select " &_
                                "			RT.DDD + '' " &_
                                "		From Relacionamento_Telefones as RT " &_
                                "		Inner Join Tipo_Telefone as TT ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as DDD " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select " &_
                                "			RT.Numero + '' " &_
                                "		From Relacionamento_Telefones as RT " &_
                                "		Inner Join Tipo_Telefone as TT ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as Tel " &_
                                "	,V.Email " &_
                                "	,C.Cargo_PTB AS Cargo " &_
                                "	,Sc.Subcargo_PTB AS SubCargo " &_
                                "	,D.Depto_PTB AS Depto " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select  " &_
                                "			Ra.Ramo_PTB  + '; ' " &_
                                "		From Relacionamento_Ramo as Rr " &_
                                "		Inner Join RamodeAtividade as Ra ON Ra.ID_Ramo = Rr.ID_Ramo " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as Ramos " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select  " &_
                                "			Rr.Ramo_Outros + '; ' " &_
                                "		From Relacionamento_Ramo as Rr " &_
                                "		Inner Join RamodeAtividade as Ra ON Ra.ID_Ramo = Rr.ID_Ramo " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as Ramo_Outros " &_
                                "	,Stuff " &_
                                "	( " &_
                                "		( " &_
                                "		Select  " &_
                                "			Ae.Atividade_PTB + ' ', " &_
                                "			Ra.Atividade_Outros + '; ' " &_
                                "		From Relacionamento_Atividade as Ra " &_
                                "		Inner Join AtividadeEconomica as Ae ON Ae.ID_Atividade = Ra.ID_Atividade " &_
                                "		Where ID_Empresa = E.ID_Empresa " &_
                                "		For XML PATH ('') " &_
                                "		), 1, 0, '' " &_
                                "	) as Atividade " &_
                                "	,V.Cargo_Outros AS Cargo_Outros " &_
                                "	,V.Depto_Outros AS Depto_Outros " &_
                                "	,Re.Endereco AS Endereco " &_
                                "	,Re.Numero AS Numero " &_
                                "	,Re.Complemento AS Complemento " &_
                                "	,Re.Cidade " &_
                                "	,Re.CEP AS CEP " &_
								"	,P.Pais " &_
								"	,U.Sigla as UF " &_
                                "From Relacionamento_Cadastro as RC  " &_
                                "Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento  " &_
                                "Inner Join Formularios as F ON F.ID_Formulario = TC.ID_Formulario  " &_
                                "Inner Join Eventos_Edicoes as EE ON EE.ID_Edicao = RC.ID_Edicao  " &_
                                "Inner Join Eventos as EV ON EV.ID_Evento = EE.ID_Evento  " &_
                                "Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante  " &_
                                "Inner Join Idiomas as I ON I.ID_Idioma = RC.ID_Idioma " &_
                                "Left  Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa  " &_
                                "Left  Join Cargo as C ON C.ID_Cargo = V.ID_Cargo " &_
                                "Left  Join SubCargo as SC ON SC.ID_SubCargo = V.ID_SubCargo " &_
                                "Left  Join Depto as D ON D.ID_Depto = V.ID_Depto " &_
                                "Left  Join Relacionamento_Enderecos as RE  " &_
                                "	ON ((RE.ID_Empresa = RC.ID_Empresa) OR (RE.ID_Visitante = RC.ID_Visitante)) " &_
								"Left Join Pais as P ON P.ID_Pais = Re.ID_Pais " &_
								"Left Join UF as U ON U.ID_UF = Re.ID_UF " &_
                                "Where  " &_
                                "	RC.ID_Edicao = " & ID & " " &_
                                "	AND RC.Exportado = 0 " &_
                                "Order by RC.Data_Cadastro; " &_
								"COMMIT TRANSACTION; " 
            
            response.write(SQL_Cadastros & "<hr>")
            Response.Flush
			
                Set RS_Cadastros = Server.CreateObject("ADODB.RecordSet")
				RS_Cadastros.CursorType = 0
				RS_Cadastros.LockType = 1
                RS_Cadastros.Open SQL_Cadastros, Conexao
                
                Function trocar(qual)
                    If Len(qual) > 0 Then
                        limpar = qual 
                        limpar = Replace(limpar, "&", "&amp;")
                        limpar = Replace(limpar, "<", "&lt;")
                        limpar = Replace(limpar, ">", "&gt;")
                        limpar = Replace(limpar, """", "&quot;")
                        trocar = limpar
                    Else
                        trocar = qual
                    End If
                End Function
                
                Function limpar_formatacao(qual)
                    If Len(qual) > 0 Then
                        limpar = qual 
                        limpar = Replace(limpar, ".", "")
                        limpar = Replace(limpar, "-", "")
                        limpar = Replace(limpar, "/", "")
                        limpar_formatacao = limpar
                    Else
                        limpar_formatacao = qual
                    End If	
                End Function
                
                If not RS_Cadastros.BOF or not RS_Cadastros.EOF Then
            
                    d = Day(Now)
                    m = Month(Now)
                    a = Year(Now)
                    h = Hour(Now)
                    n = Minute(Now)
                    s = Second(Now)
                    If Len(d) < 2 Then d = "0" & d
                    If Len(m) < 2 Then m = "0" & m
                    If Len(h) < 2 Then h = "0" & h
                    If Len(n) < 2 Then n = "0" & n
                    If Len(s) < 2 Then s = "0" & s
                    data = d & "-" & m & "-" & a & "_" & h & "-" & n & "-" & s
            
                    arquivo = "pre_credenciados"
                    
            '		If Session("admin_xml_evento") = "9" Then
            '			feira = "FT2010"
            '		ElseIf Session("admin_xml_evento") = "17" Then
            '			feira = "FHR2010"
            '		ElseIf Session("admin_xml_evento") = "13" Then
            '			feira = "MA2010"
            '		End If
            
                    
                    extensao = ".xml"
            
                    Filename = arquivo & "_" & feira & "_" & data & extensao ' file to read
            
                    Const ForReading = 1, ForWriting = 2, ForAppending = 3
                    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
            
                    ' Create a filesystem object
                    Dim FSO
                    set FSO = server.createObject("Scripting.FileSystemObject")
            
                    ' Map the logical path to the physical system path
                    Dim Filepath
                    Filepath = Server.MapPath("arquivos/" & Filename)
            
                    Set oFiletxt = FSO.CreateTextFile(Filepath, True)
                    sPath = FSO.GetAbsolutePathName(Filepath)
                    sFilename = FSO.GetFileName(sPath)
            
                    oFiletxt.WriteLine("<?xml version='1.0' encoding='iso-8859-1'?>")
                    oFiletxt.WriteLine("<credenciamento>")
            
                    total = 0
                    If not RS_Cadastros.BOF or not RS_Cadastros.EOF Then
                        While not RS_Cadastros.EOF			
                            total = total + 1
                            response.write("N: " & total & " - ID: " & RS_Cadastros("ID_Relacionamento_Cadastro") & " - CPF: " & Ucase(RS_Cadastros("cpf")) & " - Nome: " & Ucase(RS_Cadastros("nome_completo")) & "<br>")
                            response.write("<script>self.scrollBy(0,400)</script>")
                            Response.Flush

                            oFiletxt.WriteLine("<cadastro>")
                            oFiletxt.WriteLine("<cod_identificacao>" 	& trocar(Ucase(RS_Cadastros("ID_Relacionamento_Cadastro"))) & "</cod_identificacao>")
                            oFiletxt.WriteLine("<cnpj>" 				& limpar_formatacao(trocar(Ucase(RS_Cadastros("cnpj")))) & "</cnpj>")
                            oFiletxt.WriteLine("<empresa>" 				& trocar(Ucase(RS_Cadastros("empresa"))) & "</empresa>")
                            oFiletxt.WriteLine("<nome_fantasia>"		& trocar(Ucase(RS_Cadastros("NomeFantasia"))) & "</nome_fantasia>")
                            oFiletxt.WriteLine("<nome>" 				& trocar(Ucase(RS_Cadastros("Nome_Completo"))) & "</nome>")
                            oFiletxt.WriteLine("<credencial>" 			& trocar(Ucase(RS_Cadastros("credencial"))) & "</credencial>")
                            oFiletxt.WriteLine("<cpf>" 					& limpar_formatacao(trocar(Ucase(RS_Cadastros("cpf")))) & "</cpf>")
                            oFiletxt.WriteLine("<ddi>" 					& trocar(Ucase(RS_Cadastros("ddi"))) & "</ddi>")
                            oFiletxt.WriteLine("<ddd>" 					& trocar(Ucase(RS_Cadastros("ddd"))) & "</ddd>")
                            oFiletxt.WriteLine("<fone>" 				& limpar_formatacao(trocar(Ucase(RS_Cadastros("tel")))) & "</fone>")
                            oFiletxt.WriteLine("<email>" 				& trocar(Ucase(RS_Cadastros("email"))) & "</email>")
                            oFiletxt.WriteLine("<cargo>"				& trocar(Ucase(RS_Cadastros("cargo"))) & "</cargo>")
                            oFiletxt.WriteLine("<cargo_outros>" 		& trocar(Ucase(RS_Cadastros("cargo_outros"))) & "</cargo_outros>")
                            oFiletxt.WriteLine("<subcargo>" 			& trocar(Ucase(RS_Cadastros("SubCargo"))) & "</subcargo>")
                            oFiletxt.WriteLine("<departamento>" 		& trocar(Ucase(RS_Cadastros("depto"))) & "</departamento>")
                            oFiletxt.WriteLine("<departamento_outros>" 	& trocar(Ucase(RS_Cadastros("depto_outros"))) & "</departamento_outros>")
                            oFiletxt.WriteLine("<ramo_atividade>" 		& trocar(Ucase(RS_Cadastros("ramos"))) & "</ramo_atividade>")
                            oFiletxt.WriteLine("<ramo_atividade_outros>" & trocar(Ucase(RS_Cadastros("ramo_outros"))) & "</ramo_atividade_outros>")
                            oFiletxt.WriteLine("<endereco>" 			& trocar(Ucase(RS_Cadastros("endereco"))) & "</endereco>")
                            oFiletxt.WriteLine("<nro>" 					& trocar(Ucase(RS_Cadastros("numero"))) & "</nro>")
                            oFiletxt.WriteLine("<complemento>" 			& trocar(Ucase(RS_Cadastros("complemento"))) & "</complemento>")
                            oFiletxt.WriteLine("<cep>" 					& limpar_formatacao(trocar(Ucase(RS_Cadastros("cep")))) & "</cep>")
                            oFiletxt.WriteLine("<cidade>" 				& trocar(Ucase(RS_Cadastros("cidade"))) & "</cidade>")
                            oFiletxt.WriteLine("<uf>" 					& trocar(Ucase(RS_Cadastros("uf"))) & "</uf>")
                            oFiletxt.WriteLine("<pais>" 				& trocar(Ucase(RS_Cadastros("pais"))) & "</pais>")
							
                            SQL_RespostasPesquisa = 	"Select " &_
                                                        "	P.Pergunta_PTB as Pergunta, " &_
                                                        "	Stuff " &_
                                                        "	( " &_
                                                        "		( " &_
                                                        "		Select  " &_
                                                        "			Opcao_PTB + '; ' " &_
                                                        "		From Perguntas_Opcoes as PO " &_
                                                        "		Where  " &_
                                                        "			PO.ID_Perguntas = RP.ID_Perguntas " &_
                                                        "			AND PO.ID_Opcoes = RP.ID_Opcoes " &_
                                                        "		For XML PATH ('') " &_
                                                        "		), 1, 0, '' " &_
                                                        "	) as Respostas " &_
                                                        "From Relacionamento_Perguntas as RP " &_
                                                        "Inner Join Perguntas as P  " &_
                                                        "	ON P.ID_Perguntas = RP.ID_Perguntas " &_
                                                        "Where RP.ID_Relacionamento_Cadastro = " & RS_Cadastros("ID_Relacionamento_Cadastro")
                                                        
            '	response.write("<hr>" & SQL_RespostasPesquisa & "<hr>")
                                                        
                            Set RS_RespostasPesquisa = Server.CreateObject("ADODB.RecordSet")
                            RS_RespostasPesquisa.CursorType = 0
                            RS_RespostasPesquisa.LockType = 1
                            RS_RespostasPesquisa.Open SQL_RespostasPesquisa, Conexao
                            
                            If not RS_RespostasPesquisa.BOF or not RS_RespostasPesquisa.EOF Then
                                oFiletxt.WriteLine("<pesquisa>")
                                While not RS_RespostasPesquisa.EOF
                                    pergunta = RS_RespostasPesquisa("Pergunta")
                                    resposta = RS_RespostasPesquisa("Respostas")
                                    If Len(pergunta) > 0 Then
                                        oFiletxt.WriteLine("<pergunta questao='" & pergunta & "' resposta='" & trocar(resposta) & "'/>")
                                    End If
                                    RS_RespostasPesquisa.MoveNext
                                Wend
                                oFiletxt.WriteLine("</pesquisa>")
                                RS_RespostasPesquisa.Close
                            End If				
                            oFiletxt.WriteLine("</cadastro>")
            
                            SQL_Exportado = "Update Relacionamento_Cadastro " &_
                                            "Set " &_
                                            "	exportado = 1 " &_
                                            "	,Exportacao_DATA = getDate() " &_
                                            "	,Exportado_por_ID_Admin = " & Session("admin_id_usuario") & " " &_
                                            "Where ID_Relacionamento_Cadastro = " & RS_Cadastros("ID_Relacionamento_Cadastro")
                            Set RS_Exportado = Server.CreateObject("ADODB.RecordSet")
                            RS_Exportado.Open SQL_Exportado, Conexao
            
                            RS_Cadastros.MoveNext()
		                    response.write("<script>self.scrollBy(0,400)</script>")
                            response.Flush()
                        Wend
                        RS_Cadastros.Close
                    End If
            
                    oFiletxt.WriteLine("</credenciamento>")
                    oFiletxt.Close
            
                    SQL_Arquivos = 	"Insert Into Arquivos_XML " &_
                                    "(arquivo, total, Id_Edicao) " &_
                                    "values " &_
                                    "('" & filename & "'," & total & "," & id & ")"
                   	Set RS_Arquivos = Server.CreateObject("ADODB.RecordSet")
                   	RS_Arquivos.Open SQL_Arquivos, Conexao
            
            
                    %>
                    <hr>
                    Arquivo <B><%=Filename%></B> criado com sucesso<br><br>
                    Total de Cadastros Listados : <b><%=total%></b><br>
                    <a href="arquivos/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo</a><br>
                    * Botão direito > Salvar Como
                    <%
                    response.write("<script>self.scrollBy(0,1000)</script>")
                    Response.Flush
                    %>
                <% Else %>
                    Não existem novos cadastros.
                <% End IF %>
	<hr>
        <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
        <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco "><a href="evento.asp?id=<%=id%>" style="color: #fff">Listar Arquivos</a></div>
        <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>

    </div>
	<%
	EndTime = Timer()
	response.write("<br>Tempo de processamento: " & (EndTime - StartTime) & " segundos<br>")
	%>
	</td>
  </tr>
</table>
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/admin/images/img_tabela_branca_inferior.jpg" width="955" height="15" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
<% Conexao.Close %>