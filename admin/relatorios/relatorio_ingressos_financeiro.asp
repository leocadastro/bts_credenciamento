<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar =   	"Select " &_
					"	Pe.ID_Pedido " &_
					"	,Pe.ID_Edicao " &_
					"	,Pe.ID_Idioma " &_
					"	,Pe.Numero_Pedido " &_
					"	,Pe.ID_Rel_Cadastro " &_
					"	,Pe.ID_Visitante As ID_Comprador " &_
					"	,Vc.Email As Email_Comprador " &_
					"	,Vc.Nome_Completo As Nome_Comprador " &_
					"	,Vc.CPF As CPF_Comprador " &_
					"	,Vc.Passaporte As PASS_Comprador " &_
					"	,Vc.Senha " &_
					"	,Pe.Valor_Pedido " &_
					"	,Pe.Data_Pedido " &_
					"	,Pe.Status_Pedido As ID_Status " &_
					"	,St.Status_PTB As Status_Pedido " &_
					"From " &_
					"	Pedidos_Carrinho As Pc " &_
					"	Inner Join Pedidos As Pe " &_
					"		On Pc.ID_Pedido = Pe.ID_Pedido " &_
					"	Inner Join Pedidos_Status As St " &_
					"		On St.ID_Pedido_Status = Pe.Status_Pedido " &_
					"	Inner Join Visitantes As Vc " &_
					"		On Vc.ID_Visitante = Pe.ID_Visitante " &_
					"Where " &_
					"	Pe.Valor_Pedido > 1 " &_
					"	AND Status_Pedido = 3 " &_
					"ORDER BY " &_
					"	Numero_Pedido "
					'Response.Write(SQL_Listar)
	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.CursorLocation = 3
	RS_Listar.Open SQL_Listar, Conexao

%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração Cred. 2012</title>
    <script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
</head>

<body>
<table style="font-family:Arial, Helvetica, sans-serif; font-size:12px;" cellpadding="3" cellspacing="3" width="100%">
  <tr>
    <td class="borda_dir linha_16px" width="10"><strong>Qtde.</strong></td>
    <td class="borda_dir linha_16px" width="130"><strong>N. Pedido</strong></td>
    <td class="borda_dir linha_16px" align="center" width="70"><strong>Data</strong></td>
    <td class="borda_dir linha_16px" align="center" width="80"><strong>CPF</strong></td>
    <td class="borda_dir linha_16px" align="center" width="90"><strong>Passaporte</strong></td>
    <td class="borda_dir linha_16px" width="*" nowrap><strong>Nome do Comprador</strong></td>
    <td class="borda_dir linha_16px" width="*"><strong>E-mail</strong></td>
    <td class="borda_dir linha_16px" align="center" width="70"><strong>Valor</strong></td>
    <td class="borda_dir linha_16px" align="center" width="20"><strong>Série</strong></td>
    <td class="borda_dir linha_16px" align="center" width="40"><strong>Tipo</strong></td>
    <td class="borda_dir linha_16px" align="center" width="100"><strong>N. Autorização</strong></td>
    <td class="borda_dir linha_16px" align="center" width="50"><strong>Bandeira</strong></td>
    <td class="borda_dir linha_16px" align="center" width="70"><strong>Data</strong></td>
  </tr>
  <%
	If Not RS_Listar.EOF Then

		While Not RS_Listar.EOF
			Contador = Contador + 1	
			
			Numero_Pedido 	= RS_Listar("Numero_Pedido")
			CPF_Comprador	= RS_Listar("CPF_Comprador")
			Nome_Comprador	= RS_Listar("Nome_Comprador")
			Email_Comprador	= RS_Listar("Email_Comprador")
			Valor_Pedido	= FormatNumber(RS_Listar("Valor_Pedido"))
			
			'If Numero_Pedido = RS_Listar("Numero_Pedido") Then
				
				
			
%>
  <tr>
    <td class="borda_dir linha_16px"><%=Contador%></td>
    <td class="borda_dir linha_16px"><%=RS_Listar("Numero_Pedido")%></td>
    <td class="borda_dir linha_16px"><%
									If Day(RS_Listar("Data_Pedido")) < 10 Then
										Dia = "0" & Day(RS_Listar("Data_Pedido"))
									Else
										Dia = Day(RS_Listar("Data_Pedido"))
									End If
									
									If Month(RS_Listar("Data_Pedido")) < 10 Then
										Mes = "0" & Month(RS_Listar("Data_Pedido"))
									Else
										Mes = Month(RS_Listar("Data_Pedido"))
									End If
									Response.Write(Dia &"/"& Mes &"/"& Year(RS_Listar("Data_Pedido")))
									%>	</td>
    <td class="borda_dir linha_16px" align="center"><%=RS_Listar("CPF_Comprador")%></td>
	<td class="borda_dir linha_16px" align="center"><%=RS_Listar("PASS_Comprador")%></td>
    <td class="borda_dir linha_16px" nowrap><%=RS_Listar("Nome_Comprador")%></td>
	<td class="borda_dir linha_16px"><%=RS_Listar("Email_Comprador")%></td>
    <td class="borda_dir linha_16px" align="right">R$ 60,00</td>
    <td class="borda_dir linha_16px" align="center">E</td>
    <td class="borda_dir linha_16px" align="center">Crédito</td>

  </tr>
<%
		RS_Listar.MoveNext
		Wend
	End If
%>
</table>
</body>
</html>