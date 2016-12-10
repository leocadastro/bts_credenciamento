<% 
sub rw(s,b)
	response.Write("<textarea>"&s&"</textarea>")
	if b then
		response.End()
	end if
end sub


'==================================================
' Checar se está logado
	If Session("admin_logado") <> true Then 
		Session("admin_msg") = "novo_login" 
		Session("admin_url") = Lcase(Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING"))
		response.Redirect("/admin/")
	End if
'==================================================

'==================================================
' Checar se tem permissão para acessar esta página
	pagina_atual = Lcase(Request.ServerVariables("URL"))
	'response.write(pagina_atual)
	'******************************************************
	' Inserir aqui as PAGINAS e os perfis que podem acessar
	'******************************************************
	Dim paginas_permissao(91)
								  ' endereço , id_perfil permitido
	paginas_permissao(0)  = Array("/admin/menu.asp", "1,2,3,4")

	paginas_permissao(1)  = Array("/admin/eventos/default.asp", "1,2")
	paginas_permissao(2)  = Array("/admin/eventos/editar.asp", "1,2")

	paginas_permissao(3)  = Array("/admin/edicoes/default.asp", "1,2")
	paginas_permissao(4)  = Array("/admin/edicoes/editar.asp", "1,2")

	paginas_permissao(5)  = Array("/admin/edicoes_visual/default.asp", "1,2")
	paginas_permissao(6)  = Array("/admin/edicoes_visual/editar.asp", "1,2")

	paginas_permissao(7)  = Array("/admin/tipo_credenciamento/default.asp", "1,2")
	paginas_permissao(8)  = Array("/admin/tipo_credenciamento/editar.asp", "1,2")

	paginas_permissao(9)  = Array("/admin/rel_tipos_edicoes/default.asp", "1,2")
	paginas_permissao(10) = Array("/admin/rel_tipos_edicoes/editar.asp", "1,2")

	paginas_permissao(11) = Array("/admin/ramos_atividades/default.asp", "1,2")
	paginas_permissao(12) = Array("/admin/ramos_atividades/editar.asp", "1,2")

	paginas_permissao(13) = Array("/admin/paginas_web/default.asp", "1,2")
	paginas_permissao(14) = Array("/admin/paginas_web/editar.asp", "1,2")

	paginas_permissao(15) = Array("/admin/paginas_textos/default.asp", "1,2")
	paginas_permissao(16) = Array("/admin/paginas_textos/editar.asp", "1,2")

	paginas_permissao(17) = Array("/admin/paginas_tipo/default.asp", "1,2")
	paginas_permissao(18) = Array("/admin/paginas_tipo/editar.asp", "1,2")

	paginas_permissao(19) = Array("/admin/cargo/default.asp", "1,2")
	paginas_permissao(20) = Array("/admin/cargo/editar.asp", "1,2")

	paginas_permissao(21) = Array("/admin/sub_cargo/default.asp", "1,2")
	paginas_permissao(22) = Array("/admin/sub_cargo/editar.asp", "1,2")

	paginas_permissao(23) = Array("/admin/funcionarios/default.asp", "1,2")
	paginas_permissao(24) = Array("/admin/funcionarios/editar.asp", "1,2")

	paginas_permissao(25) = Array("/admin/ramos/default.asp", "1,2")
	paginas_permissao(26) = Array("/admin/ramos/editar.asp", "1,2")

	paginas_permissao(27) = Array("/admin/atividade/default.asp", "1,2")
	paginas_permissao(28) = Array("/admin/atividade/editar.asp", "1,2")

	paginas_permissao(29) = Array("/admin/area_interesse/default.asp", "1,2")
	paginas_permissao(30) = Array("/admin/area_interesse/editar.asp", "1,2")

	paginas_permissao(31) = Array("/admin/area_atuacao/default.asp", "1,2")
	paginas_permissao(32) = Array("/admin/area_atuacao/editar.asp", "1,2")

	paginas_permissao(33) = Array("/admin/relatorios/default.asp", "1,2,3")

	paginas_permissao(34) = Array("/admin/rel_edicoes_ramos/default.asp", "1,2")
	paginas_permissao(35) = Array("/admin/rel_edicoes_ramos/editar.asp", "1,2")

	paginas_permissao(36) = Array("/admin/relatorios/relatorio_precredenciados.asp", "1,2")

	paginas_permissao(37) = Array("/admin/rel_edicoes_atividade/default.asp", "1,2")
	paginas_permissao(38) = Array("/admin/rel_edicoes_atividade/editar.asp", "1,2")
	paginas_permissao(39) = Array("/admin/rel_edicoes_interesse/default.asp", "1,2")
	paginas_permissao(40) = Array("/admin/rel_edicoes_interesse/editar.asp", "1,2")
	paginas_permissao(41) = Array("/admin/rel_edicoes_atuacao/default.asp", "1,2")
	paginas_permissao(42) = Array("/admin/rel_edicoes_atuacao/editar.asp", "1,2")
	paginas_permissao(43) = Array("/admin/rel_edicoes_interessefeira/default.asp", "1,2")
	paginas_permissao(44) = Array("/admin/rel_edicoes_interessefeira/editar.asp", "1,2")

	paginas_permissao(45) = Array("/admin/interesse_feira/default.asp", "1,2")
	paginas_permissao(46) = Array("/admin/interesse_feira/editar.asp", "1,2")

	paginas_permissao(47) = Array("/admin/rel_edicoes_cargo/default.asp", "1,2")
	paginas_permissao(48) = Array("/admin/rel_edicoes_cargo/editar.asp", "1,2")
	paginas_permissao(49) = Array("/admin/rel_edicoes_subcargo/default.asp", "1,2")
	paginas_permissao(50) = Array("/admin/rel_edicoes_subcargo/editar.asp", "1,2")

	paginas_permissao(51) = Array("/admin/administradores/atualizar.asp", "1,2,3,4")
	paginas_permissao(52) = Array("/admin/administradores/default.asp", "1,2")
	paginas_permissao(53) = Array("/admin/administradores/editar.asp", "1,2")

	paginas_permissao(54) = Array("/admin/relatorios/relatorio_total_precredenciados.asp", "1,2,3")

	paginas_permissao(55) = Array("/admin/depto/default.asp", "1,2")
	paginas_permissao(56) = Array("/admin/depto/editar.asp", "1,2")

	paginas_permissao(57) = Array("/admin/relatorios/relatorio_precredenciados_detalhes.asp", "1,2")

	paginas_permissao(58) = Array("/admin/administradores/relacionar_edicoes/default.asp", "1,2")

	paginas_permissao(59) = Array("/admin/exportar_xml/default.asp", "1,2,4")
	paginas_permissao(60) = Array("/admin/exportar_xml/evento.asp", "1,2,4")
	paginas_permissao(61) = Array("/admin/exportar_xml/exportar.asp", "1,2,4")
	
	paginas_permissao(62) = Array("/admin/produtos/default.asp", "1,2")
	paginas_permissao(63) = Array("/admin/produtos/editar.asp", "1,2")
	
	paginas_permissao(64) = Array("/admin/outros/default.asp", "1,2")
	paginas_permissao(65) = Array("/admin/outros/listar.asp", "1,2")
	paginas_permissao(66) = Array("/admin/outros/editar.asp", "1,2")
	
	paginas_permissao(67) = Array("/admin/relatorios/exportar_precredenciados.asp", "1,2")
	paginas_permissao(68) = Array("/admin/relatorios/exportar.asp", "1,2")
	paginas_permissao(69) = Array("/admin/relatorios/exportar_precredenciados_correio.asp", "1,2")
	paginas_permissao(70) = Array("/admin/relatorios/exportar_correio.asp", "1,2")
	

	paginas_permissao(71) = Array("/admin/perguntas/default.asp", "1,2,4")
	paginas_permissao(72) = Array("/admin/perguntas/editar.asp", "1,2,4")
	paginas_permissao(73) = Array("/admin/perguntas/metodos.asp", "1,2,4")

	paginas_permissao(74) = Array("/admin/rel_edicoes_perguntas/default.asp", "1,2,4")
	paginas_permissao(75) = Array("/admin/rel_edicoes_perguntas/editar.asp", "1,2,4")
	paginas_permissao(76) = Array("/admin/rel_edicoes_perguntas/metodos.asp", "1,2,4")
	paginas_permissao(77) = Array("/admin/perguntas/editar_opcoes.asp", "1,2,4")	
	
	paginas_permissao(78) = Array("/admin/outros/subcargos.asp", "1,2,4")	
	paginas_permissao(79) = Array("/admin/outros/subcargos_editar.asp", "1,4")	
	
	paginas_permissao(80) = Array("/admin/links_feira/default.asp", "1,2,4")
	
	paginas_permissao(81) = Array("/admin/faq/default.asp", "1,2,4")	
	
	paginas_permissao(82) = Array("/admin/ramos_v2/default.asp", "1,2,4")
	paginas_permissao(83) = Array("/admin/ramos_v2/editar.asp", "1,2,4")
	
	paginas_permissao(84) = Array("/admin/rel_edicoes_ramos_v2/default.asp", "1,2,4")
	paginas_permissao(85) = Array("/admin/rel_edicoes_ramos_v2/editar.asp", "1,2,4")
	
	paginas_permissao(86) = Array("/admin/relatorios/relatorio_ingressos.asp", "1,2,3")
	
	paginas_permissao(87) = Array("/admin/relatorios/relatorio_total_precredenciados_semanal.asp", "1,2,3")
	paginas_permissao(88) = Array("/admin/relatorios/relatorio_total_precredenciados_semanal_excel.asp", "1,2,3")
	
	paginas_permissao(89) = Array("/admin/exportar_xml/exportar_internacional_ABF.asp", "1,2,3,4,5,6,7,8,9,10")
	
	paginas_permissao(90) = Array("/admin/relatorios/relatorio_alunos.asp", "1,2,3,4,5,6,7,8,9,10")
	paginas_permissao(91) = Array("/admin/relatorios/relatorio_alunos_gerar.asp", "1,2,3,4,5,6,7,8,9,10")

	
	'******************************************************
	
	pagina_encontrada = false
	exibir = false
	For p = LBound(paginas_permissao) to Ubound(paginas_permissao)
		'===========================================
		If Cstr(paginas_permissao(p)(0)) = Cstr(pagina_atual) Then
			pagina_encontrada = true
			' Pagina atual encontrada
			paginas_permissao_item = Split(paginas_permissao(p)(1), ",")
			For p1 = LBound(paginas_permissao_item) to Ubound(paginas_permissao_item)
				' Se seu perfil tem permissão à esta página
				If Cstr(Session("admin_id_perfil")) = Cstr(paginas_permissao_item(p1)) Then
					exibir = true
				End If
			Next
		End If
		'===========================================
	Next
	
	If exibir = false OR pagina_encontrada = false Then
		Session("admin_msg") = "pag_proibida"
		response.redirect("/admin/menu.asp")
	End If
'==================================================
%>