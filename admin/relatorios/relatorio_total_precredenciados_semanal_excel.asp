<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
Function RemoverAcentuacao(texto)
	limpar = Ucase(texto)
	If Len(limpar) <= 0 Then
	Else
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "À", "A")
		limpar = Replace(limpar, "Â", "A")
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "Ã", "A")
		limpar = Replace(limpar, "É", "E")
		limpar = Replace(limpar, "È", "E")
		limpar = Replace(limpar, "Ê", "E")
		limpar = Replace(limpar, "Í", "I")
		limpar = Replace(limpar, "Ì", "I")
		limpar = Replace(limpar, "Î", "I")
		limpar = Replace(limpar, "Ó", "O")
		limpar = Replace(limpar, "Ò", "O")
		limpar = Replace(limpar, "Ô", "O")
		limpar = Replace(limpar, "Õ", "O")
		limpar = Replace(limpar, "Ú", "U")
		limpar = Replace(limpar, "Ù", "U")
		limpar = Replace(limpar, "Û", "U")
		limpar = Replace(limpar, "Ç", "Ç")
		limpar = Replace(limpar, "&", "E")
		limpar = Replace(limpar, " ", "E")
	End If
	RemoverAcentuacao = limpar
End Function

Function getWeek(mes,ano)
	If mes < 10 Then
		mes = "0" & mes
	End If
	dt = fncLastDay(mes,ano) & "-" & mes & "-" & ano
	ret = Datepart("ww",dt)
	getWeek = ret
End Function

Function fncLastDay(intMonth,intYear)
	Dim intDay
	Select Case intMonth
		Case 1, 3, 5, 7, 8, 10, 12
			intDay = 31
		Case 4, 6, 9, 11
			intDay = 30
		Case 2
			If intYear mod 4 = 0 Then
				If intYear mod 100 = 0 AND intYear mod 400 <> 0 Then
					intDay = 28
				Else
					intDay = 29
				End If
			Else
				intDay = 28
			End If
	End Select
	fncLastDay = intDay
End Function

Function fcDescMes(s)
	Select Case s
		Case "1"
			s = "Jan"
		Case "2"
			s = "Fev"
		Case "3"
			s = "Mar"
		Case "4"
			s = "Abr"
		Case "5"
			s = "Mai"
		Case "6"
			s = "Jun"
		Case "7"
			s = "Jul"
		Case "8"
			s = "Ago"
		Case "9"
			s = "Set"
		Case "10"
			s = "Out"
		Case "11"
			s = "Nov"
		Case "12"
			s = "Dez"
	End Select
	fcDescMes=s
End Function

Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

selEdicao = Limpar_Texto(Request("selEdicao"))
arquivo = "pre_credenciados_semanal"
extensao = ".XLS"
SQL_File=	"SELECT (Nome_PTB + '_' + Convert(varchar(4),Ano)) as arquivo " &_
			"FROM Eventos_Edicoes EE " &_
			"INNER JOIN Eventos EV " &_
			"	ON EE.ID_Evento=EV.ID_Evento " &_
			"WHERE ID_Edicao=" & selEdicao
Set RS_File = Server.CreateObject("ADODB.Recordset")
RS_File.Open SQL_File, Conexao
If Not RS_File.Eof Then
	feira=RS_File("arquivo")
End If

Filename = arquivo & "_" & RemoverAcentuacao(feira)  & extensao

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath("/admin/exportar_xml/arquivos_2012/" & Filename)

Set oFiletxt = FSO.CreateTextFile(Filepath, True)
sPath = FSO.GetAbsolutePathName(Filepath)
sFilename = FSO.GetFileName(sPath)
oFiletxt.WriteLine("<table>")
oFiletxt.WriteLine("<tr>")
oFiletxt.WriteLine("<td colspan='2'>Período</td>")
oFiletxt.WriteLine("<td colspan='3'>Empresa</td>")
oFiletxt.WriteLine("<td colspan='3'>Entidade</td>")
oFiletxt.WriteLine("<td colspan='3'>Pessoa Física</td>")
oFiletxt.WriteLine("<td>&nbsp;</td>")
oFiletxt.WriteLine("</tr>")
oFiletxt.WriteLine("<tr>")
oFiletxt.WriteLine("<td>Mês</td>")
oFiletxt.WriteLine("<td>Semana</td>")
oFiletxt.WriteLine("<td>Português</td>")
oFiletxt.WriteLine("<td>Inglês</td>")
oFiletxt.WriteLine("<td>Espanhol</td>")
oFiletxt.WriteLine("<td>Português</td>")
oFiletxt.WriteLine("<td>Inglês</td>")
oFiletxt.WriteLine("<td>Espanhol</td>")
oFiletxt.WriteLine("<td>Português</td>")
oFiletxt.WriteLine("<td>Inglês</td>")
oFiletxt.WriteLine("<td>Espanhol</td>")
oFiletxt.WriteLine("<td>Totais</td>")
oFiletxt.WriteLine("</tr>")

SQL_Dados =	"SELECT	 DISTINCT(datepart(ww,RC.Data_Cadastro)) As Semana " &_
			"		,Convert(int,MONTH(RC.Data_Cadastro)) As Mes " &_
			"		,YEAR(RC.Data_Cadastro) As Ano " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=1  " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Empresa_PT " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=2  " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Empresa_ESP " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=3  " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Empresa_EN " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=4  " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Entidade_PT " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=5 " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Entidade_ES " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=6 " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Entidade_EN " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=10 " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Fisica_PT " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=11 " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Fisica_ES " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento=12 " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Fisica_EN " &_
			"		,( " &_
			"			SELECT COUNT(EMP.ID_Relacionamento_Cadastro)  " &_
			"			FROM Relacionamento_Cadastro EMP  " &_
			"			WHERE EMP.ID_Edicao=" & selEdicao & " " &_
			"			AND EMP.ID_Tipo_Credenciamento IN (1,2,3,4,5,6,10,11,12) " &_
			"			AND datepart(ww,RC.Data_Cadastro)=datepart(ww,EMP.Data_Cadastro)  " &_
			"		) AS Total " &_
			"FROM Relacionamento_Cadastro RC " &_
			"Inner Join Tipo_Credenciamento as TC " &_
			"	ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
			"Where RC.ID_Edicao =  " & selEdicao & " " &_
			"And RC.ID_Tipo_Credenciamento IN (1,2,3,4,5,6,10,11,12)"

                
Set RS_Dados = Server.CreateObject("ADODB.Recordset")
RS_Dados.Open SQL_Dados, Conexao
bFstMes = true
If Not RS_Dados.Eof Then
	MesAnt=RS_Dados("Mes")
	corAnt = "#EAEAEA"
	While Not RS_Dados.Eof
		Semana			=RS_Dados("Semana")
		Mes				=Cint(RS_Dados("Mes"))
		Ano				=RS_Dados("Ano")
		If MesAnt<>Mes Then
			i = 1
			bPrimeiraMes=false
		Else
			If bFstMes Then
				If MesAnt > 1 Then
					nUltimaSemana = getWeek((MesAnt-1),ano)
					i = (Semana - nUltimaSemana)
				else
					i=Semana
				End if
				bFstMes=False
			End If
		End If											
		Empresa_PT		=RS_Dados("Empresa_PT")
		Empresa_ESP		=RS_Dados("Empresa_ESP")
		Empresa_EN		=RS_Dados("Empresa_EN")
		Entidade_PT		=RS_Dados("Entidade_PT")
		Entidade_ES		=RS_Dados("Entidade_ES")
		Entidade_EN		=RS_Dados("Entidade_EN")
		Fisica_PT		=RS_Dados("Fisica_PT")
		Fisica_ES		=RS_Dados("Fisica_ES")
		Fisica_EN		=RS_Dados("Fisica_EN")
		DescMes 		=fcDescMes(Cint(RS_Dados("Mes")))
		Total			=RS_Dados("Total")
		oFiletxt.WriteLine("<tr>")
		oFiletxt.WriteLine("<td>" & DescMes & "</td>")
		oFiletxt.WriteLine("<td>" & i & "&ordf; Semana</td>")
		oFiletxt.WriteLine("<td>" & Empresa_PT & "</td>")
		oFiletxt.WriteLine("<td>" & Empresa_ESP & "</td>")
		oFiletxt.WriteLine("<td>" & Empresa_EN & "</td>")
		oFiletxt.WriteLine("<td>" & Entidade_PT & "</td>")
		oFiletxt.WriteLine("<td>" & Entidade_ES & "</td>")
		oFiletxt.WriteLine("<td>" & Entidade_EN & "</td>")
		oFiletxt.WriteLine("<td>" & Fisica_PT & "</td>")
		oFiletxt.WriteLine("<td>" & Fisica_ES & "</td>")
		oFiletxt.WriteLine("<td>" & Fisica_EN & "</td>")
		oFiletxt.WriteLine("<td>" & Total & "</td>")
		oFiletxt.WriteLine("</tr>")
		i=i+1
		MesAnt=Mes
		corAnt=cor
		RS_Dados.Movenext
	Wend
End If
oFiletxt.WriteLine("</table>")
Conexao.Close
%>
<html>
<head>
    <meta content="text/html; charset=iso-8859-1" http-equiv="Content-type">
    <title>Administração Cred. 2012</title>
    <link type="text/css" rel="stylesheet" href="/admin/css/bts.css">
    <link type="text/css" rel="stylesheet" href="/admin/css/admin.css">
    <link media="screen" type="text/css" rel="stylesheet" href="/css/calendar.css">
    <script src="/js/jquery-1.3.2.min.js" language="javascript"></script>
    <script src="/admin/js/validar_forms.js" language="javascript"></script>
    <script src="/js/colorpicker/colorpicker.js" language="javascript"></script>
    <script src="/js/Calendario/calendar.js" language="javascript"></script>
    <script src="/js/jquery.maskedinput-1.2.2.min.js" language="javascript"></script>
</head>
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
            <td align="center" bgcolor="#FFFFFF">
                <table width="900" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="100" height="50">&nbsp;</td>
                        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center">
                    <span style="color: #B01D22">Listagem de Pr&eacute;-Credenciados - Semanal</span></td>
                        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
                            <a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td></td>
        </tr>
        <tr>
        	<td style="background-color:#FFF;">&nbsp;</td>
        </tr>
        <tr>
            <td style="background-color:#FFF; text-align:center;">
            	<table style="width:600px; height:90px; margin:0 auto;" class="bordaRedonda">
                	<tr>
                    	<td style="text-align:center;">
                            Arquivo <B><%=Filename%></B> criado com sucesso<br><br>
                            <span style="background-color:#FF0;">
                                <a href="/admin/exportar_xml/arquivos_2012/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo.</a>
                            </span>&nbsp;* Botão direito > Salvar Como
						</td>
					</tr>
				</table>
			</td>
        </tr>
        <tr>
            <td style="background-color:#FFF;">&nbsp;</td>
        </tr>
    </table>

</body>
</html>
