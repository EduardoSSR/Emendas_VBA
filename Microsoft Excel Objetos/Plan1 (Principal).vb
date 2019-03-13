	Option Explicit
	Public cod_emenda      As String
	Public cod_ibge        As String
	Public check_editar     As Boolean
	Public check_ed_autor          As Boolean
	Private Sub Abrir_Formulário_Click()
	Dim I As Integer
	Dim tabela As String
	I = 1
	For I = 1 To 13
	   tabela = Retorna_Tabela(I)
	   Call Listar(tabela, I)
	   'MsgBox "Listas Atualizadas."
	Next
	'Me.Range("E3", "E2000").NumberFormat = "$ #,##0.00"
	'Call Listar("Autores", 1)
	Call Listar_Emendas
	'MsgBox "Listas Atualizadas"
	'Cadastramento.Show

	Me.B_autor.Visible = True
	Me.B_emendas.Visible = True
	Me.B_lei.Visible = True

	End Sub

	Private Sub B_alerta_Click()
	Dim l As Integer

	Call Listar("Legislação", 8)

	l = 2
	Do Until Sheets("Legislação").Cells(l, 1) = ""
		Sheets("Dados_Alertas").Cells(l, 1) = Sheets("Legislação").Cells(l, 1).Value
	l = l + 1
	Loop

	Alarmes.Show
	End Sub

	Private Sub B_autor_Click()

	Me.B_autor.Visible = False
	Me.B_emendas.Visible = False
	Me.B_lei.Visible = False
	novo_autor.check_btn = False
	novo_autor.Show

	End Sub

	Private Sub B_editar_Click()

	Dim W   As Worksheet
	Dim result          As VbMsgBoxResult

	Dim num_emenda      As String
	Dim valor_emenda    As String
	Dim ano             As String
	Dim cod_acao        As String
	Dim cod_ibge        As String
	Dim cod_autor       As String
	Dim cod_fonte       As String
	Dim cod_gnd         As String
	Dim cod_modalidade  As String
	Dim cod_programa    As String
	Dim cod_status      As String
	Dim cod_instrumento As String
	Dim localizador     As String
	Dim cnpj            As String
	Dim beneficiario    As String
	Dim objeto          As String
	Dim proposta        As String
	Dim convenio        As String
	Dim limite          As String
	Dim empenhado       As String
	Dim repasse         As String
	Dim contrapartida   As String
	Dim nota_empenho    As String
	Dim data_inicio     As String
	Dim data_fim        As String
	Dim impedimento     As Integer
	Dim obs             As String
	Dim pendencia       As String

	Dim lei             As String
	Dim descricao       As String
	Dim data_beneficiario   As String
	Dim data_proposta       As String
	Dim data_apresentacao   As String
	Dim data_limite         As String

	Dim autor               As String
	Dim partido             As String
	Dim cargo               As String

	Dim fonte               As String
	Dim gnd                 As String
	Dim instrumento         As String
	Dim uf                  As String
	Dim municipio           As String
	Dim programa            As String
	Dim status              As String
	Dim valor_global        As Double
	Dim modalidade          As String
	Dim acao                As String

	Dim endereco        As Range
	Dim local_ano       As Range
	Dim local_autor     As Range
	Dim local_fonte     As Range
	Dim local_gnd       As Range
	Dim local_inst      As Range
	Dim local_ibge      As Range
	Dim local_prog      As Range
	Dim local_status    As Range
	Dim local_mod       As Range
	Dim local_acao      As Range
	Dim I               As Integer
	Dim tabela          As String
		
	Set W = Sheets("Emendas")
		
	Select Case ComboBox1.Text

	Case "Emendas"
		
		ComboBox2.Visible = False
		
		cod_emenda = Me.Cells(ActiveCell.Row, 1).Value
		
		'Cadastramento.emenda = Me.Cells(ActiveCell.Row, 1).Value
		
		If ActiveCell.Row >= 3 And cod_emenda <> "" Then
			num_emenda = Me.Cells(ActiveCell.Row, 3).Value
			valor_emenda = Me.Cells(ActiveCell.Row, 5).Value
			valor_emenda = Format(valor_emenda, "R$ #,##0.00")
			result = MsgBox("Deseja editar a emenda de NÚMERO " & num_emenda & " com VALOR igual a " & valor_emenda & " ?", vbYesNo)
			If result = vbYes Then
				
				Sheets("PARÂMETROS").Range("B5") = "FALSE"
				I = 1
				For I = 1 To 13
					tabela = Retorna_Tabela(I)
				Call Listar(tabela, I)
				Next
				
				Call Listar_Emendas
				Set endereco = W.Range("A1", "A2000").Find(cod_emenda)
				ano = W.Cells(endereco.Row, 2).Value
				cod_acao = W.Cells(endereco.Row, 9).Value
				cod_ibge = W.Cells(endereco.Row, 3).Value
				cod_autor = W.Cells(endereco.Row, 4).Value
				cod_fonte = W.Cells(endereco.Row, 5).Value
				cod_gnd = W.Cells(endereco.Row, 6).Value
				cod_modalidade = W.Cells(endereco.Row, 7).Value
				cod_programa = W.Cells(endereco.Row, 8).Value
				cod_status = W.Cells(endereco.Row, 10).Value
				cod_instrumento = W.Cells(endereco.Row, 11).Value
				localizador = W.Cells(endereco.Row, 14).Value
				cnpj = W.Cells(endereco.Row, 15).Value
				beneficiario = W.Cells(endereco.Row, 16).Value
				objeto = W.Cells(endereco.Row, 17).Value
				proposta = W.Cells(endereco.Row, 18).Value
				convenio = W.Cells(endereco.Row, 19).Value
				limite = W.Cells(endereco.Row, 20).Value
					limite = Format(limite, "R$ #,##0.00")
				empenhado = W.Cells(endereco.Row, 21).Value
					empenhado = Format(empenhado, "R$ #,##0.00")
				repasse = W.Cells(endereco.Row, 23).Value
				contrapartida = W.Cells(endereco.Row, 24).Value
				If repasse <> "" And contrapartida <> "" Then
				valor_global = (CDbl(contrapartida) + CDbl(repasse))
				Else
					valor_global = 0
				End If
					repasse = Format(repasse, "R$ #,##0.00")
					contrapartida = Format(contrapartida, "R$ #,##0.00")
				nota_empenho = W.Cells(endereco.Row, 22).Value
				data_inicio = W.Cells(endereco.Row, 25).Value
				data_fim = W.Cells(endereco.Row, 26).Value
				impedimento = W.Cells(endereco.Row, 27).Value
				obs = W.Cells(endereco.Row, 28).Value
				pendencia = W.Cells(endereco.Row, 29).Value
				
				Set local_ano = Sheets("Legislação").Range("A1", "A100").Find(ano)
				
				lei = Sheets("Legislação").Cells(local_ano.Row, 2).Value
				descricao = Sheets("Legislação").Cells(local_ano.Row, 3).Value
				data_beneficiario = Sheets("Legislação").Cells(local_ano.Row, 4).Value
				data_proposta = Sheets("Legislação").Cells(local_ano.Row, 5).Value
				data_apresentacao = Sheets("Legislação").Cells(local_ano.Row, 6).Value
				data_limite = Sheets("Legislação").Cells(local_ano.Row, 7).Value
				
				Set local_autor = Sheets("Autores").Range("A1", "A500").Find(cod_autor)
				
				autor = Sheets("Autores").Cells(local_autor.Row, 2).Value
				partido = Sheets("Autores").Cells(local_autor.Row, 4).Value
				cargo = Sheets("Autores").Cells(local_autor.Row, 3).Value
				
				Set local_fonte = Sheets("Fonte").Range("A1", "A100").Find(cod_fonte)
				fonte = Sheets("Fonte").Cells(local_fonte.Row, 2).Value
				
				Set local_gnd = Sheets("GND").Range("A1", "A100").Find(cod_gnd)
				gnd = Sheets("GND").Cells(local_gnd.Row, 2).Value
				
				
				Set local_inst = Sheets("Instrumento").Range("A1", "A100").Find(cod_instrumento)
				instrumento = Sheets("Instrumento").Cells(local_inst.Row, 2).Value
				
				Set local_ibge = Sheets("Município").Range("A1", "A7000").Find(cod_ibge)
				uf = Sheets("Município").Cells(local_ibge.Row, 3).Value
				municipio = Sheets("Município").Cells(local_ibge.Row, 2).Value
				
				Set local_prog = Sheets("Programa").Range("A1", "A100").Find(cod_programa)
				programa = Sheets("Programa").Cells(local_prog.Row, 2).Value
				
				Set local_status = Sheets("Status").Range("A1", "A100").Find(cod_status)
				status = Sheets("Status").Cells(local_status.Row, 2).Value
				
				Set local_mod = Sheets("Modalidade").Range("A1", "A100").Find(cod_modalidade)
				modalidade = Sheets("Modalidade").Cells(local_mod.Row, 2).Value
				
				Set local_acao = Sheets("Ação_Orçamentaria").Range("A1", "A100").Find(cod_acao)
				acao = Sheets("Ação_Orçamentaria").Cells(local_acao.Row, 2).Value
				
				Cadastramento.T_ano_emenda.Text = ano
				Cadastramento.T_valor.Text = valor_emenda
				Cadastramento.T_num_emenda.Text = num_emenda
				Cadastramento.T_acao.Text = acao
				Cadastramento.T_autor.Text = autor
				Cadastramento.T_beneficiario.Text = beneficiario
				Cadastramento.T_cargo.Text = cargo
				Cadastramento.T_cnpj_beneficiario.Text = cnpj
				Cadastramento.T_data_analise.Text = data_proposta
				Cadastramento.T_data_apresentacao.Text = data_apresentacao
				Cadastramento.T_data_beneficiario.Text = data_beneficiario
				Cadastramento.T_data_fim.Text = data_fim
				Cadastramento.T_data_inicio.Text = data_inicio
				Cadastramento.T_data_limite.Text = data_limite
				Cadastramento.T_descricacao.Text = descricao
				Cadastramento.T_fonte.Text = fonte
				Cadastramento.T_gnd.Text = gnd
				Cadastramento.T_ibge.Text = cod_ibge
				Cadastramento.T_instrumento.Text = instrumento
				Cadastramento.T_Lei.Text = lei
				Cadastramento.T_limite_empenho.Text = limite
				Cadastramento.T_localizxador.Text = localizador
				Cadastramento.T_municipio.Text = municipio
				Cadastramento.T_modalidade = modalidade
				Cadastramento.T_nota_empenho.Text = nota_empenho
				Cadastramento.T_num_siconv.Text = convenio
				Cadastramento.T_objeto.Text = objeto
				Cadastramento.T_observacao.Text = obs
				Cadastramento.T_partido.Text = partido
				Cadastramento.T_pendencia.Text = pendencia
				Cadastramento.T_programa.Text = programa
				Cadastramento.T_proposta.Text = proposta
				Cadastramento.T_status.Text = status
				Cadastramento.T_uf.Text = uf
				Cadastramento.T_valor_contrapartida.Text = contrapartida
				Cadastramento.T_valor_convenio.Text = Format(valor_global, "R$ #,##0.00")
				Cadastramento.T_valor_empenhado.Text = empenhado
				Cadastramento.T_valor_repasse.Text = repasse
				
				If impedimento = 0 Then
					Cadastramento.OptionButtonN = True
				ElseIf impedimento = 1 Then
					Cadastramento.OptionButtonS = True
				Else
					Cadastramento.OptionButtonS = False
					Cadastramento.OptionButtonN = False
				End If
				
				Cadastramento.f_ano = "'" & ano & "'"
				If beneficiario <> "" Then
					Cadastramento.f_beneficiario = "'" & beneficiario & "'"
				Else
					Cadastramento.f_beneficiario = "NULL"
				End If
				If cnpj <> "" Then
					Cadastramento.f_cnpj = "'" & cnpj & "'"
				Else
					Cadastramento.f_cnpj = "NULL"
				End If
				Cadastramento.f_cod_acao = "'" & cod_acao & "'"
				Cadastramento.f_cod_autor = "'" & cod_autor & "'"
				Cadastramento.f_cod_emenda = "'" & cod_emenda & "'"
				Cadastramento.f_cod_fonte = "'" & cod_fonte & "'"
				Cadastramento.f_cod_gnd = "'" & cod_gnd & "'"
				'Cadastramento.f_cod_instrumento = "'" & cod_instrumento & "'"
				If cod_instrumento <> "" Then
					Cadastramento.f_cod_instrumento = "'" & cod_instrumento & "'"
				Else
					Cadastramento.f_cod_instrumento = "NULL"
				End If
				Cadastramento.f_cod_modalidade = "'" & cod_modalidade & "'"
				Cadastramento.f_cod_programa = "'" & cod_programa & "'"
				If cod_status <> "" Then
					Cadastramento.f_cod_status = "'" & cod_status & "'"
				Else
					Cadastramento.f_cod_status = "NULL"
				End If
				If convenio <> "" Then
					Cadastramento.f_convenio = "'" & convenio & "'"
				Else
					Cadastramento.f_convenio = "NULL"
				End If
				If data_fim <> "" Then
					Cadastramento.f_data_fim = "'" & data_fim & "'"
				Else
					Cadastramento.f_data_fim = "NULL"
				End If
				If data_inicio <> "" Then
					Cadastramento.f_data_inicio = "'" & data_inicio & "'"
				Else
					Cadastramento.f_data_inicio = "NULL"
				End If
				If empenhado <> "" Then
					empenhado = Format(empenhado, " 0.00")
					empenhado = "'" & Replace(empenhado, ",", ".") & "'"
					Cadastramento.f_empenhado = empenhado
				Else
					Cadastramento.f_empenhado = "NULL"
				End If
				If cod_ibge <> "" Then
					Cadastramento.f_ibge = "'" & cod_ibge & "'"
				Else
					Cadastramento.f_ibge = "NULL"
				End If
				If limite <> "" Then
					limite = Format(limite, " 0.00")
					limite = "'" & Replace(limite, ",", ".") & "'"
					Cadastramento.f_limite = limite
				Else
					Cadastramento.f_limite = "NULL"
				End If
				If localizador <> "" Then
					Cadastramento.f_localizador = "'" & localizador & "'"
				Else
					Cadastramento.f_localizador = "NULL"
				End If
				If nota_empenho <> "" Then
					Cadastramento.f_nota_empenho = "'" & nota_empenho & "'"
				Else
					Cadastramento.f_nota_empenho = "NULL"
				End If
				Cadastramento.f_num_emenda = "'" & num_emenda & "'"
				If objeto <> "" Then
					Cadastramento.f_objeto = "'" & objeto & "'"
				Else
					Cadastramento.f_objeto = "NULL"
				End If
				If obs <> "" Then
					Cadastramento.f_obs = "'" & obs & "'"
				Else
					Cadastramento.f_obs = "NULL"
				End If
				If pendencia <> "" Then
					Cadastramento.f_pendencia = "'" & pendencia & "'"
				Else
					Cadastramento.f_pendencia = "NULL"
				End If
				If proposta <> "" Then
					Cadastramento.f_proposta = "'" & proposta & "'"
				Else
					Cadastramento.f_proposta = "NULL"
				End If
				If contrapartida <> "" Then
					contrapartida = Format(contrapartida, " 0.00")
					contrapartida = "'" & Replace(contrapartida, ",", ".") & "'"
					Cadastramento.f_valor_contrapartida = contrapartida
				Else
					Cadastramento.f_valor_contrapartida = "NULL"
				End If
				If valor_emenda <> "" Then
					valor_emenda = Format(valor_emenda, " 0.00")
					valor_emenda = "'" & Replace(valor_emenda, ",", ".") & "'"
					Cadastramento.f_valor_emenda = valor_emenda
				Else
					Cadastramento.f_valor_emenda = "NULL"
				End If
				If repasse <> "" Then
					repasse = Format(repasse, " 0.00")
					repasse = "'" & Replace(repasse, ",", ".") & "'"
					Cadastramento.f_valor_repasse = repasse
				Else
					Cadastramento.f_valor_repasse = "NULL"
				End If
				Cadastramento.Show
			Else
				Exit Sub
			End If
		Else
			MsgBox "Selecione uma linha que contenha dados. "
			Exit Sub
		End If
	Case "Autor"
		Dim set_autor As Range
		If ComboBox2.Text <> "" Then
			'Editar_autor.T_autor.Text = ComboBox2.Text
			Set set_autor = Sheets("Autores").Range("b1", "b500").Find(Plan1.ComboBox2.Text)
				Editar_autor.cod_autor = Sheets("Autores").Cells(set_autor.Row, 1).Value
				Editar_autor.T_autor.Text = ComboBox2.Text
				Editar_autor.T_cargo.Text = Sheets("Autores").Cells(set_autor.Row, 3).Value
				Editar_autor.T_partido.Text = Sheets("Autores").Cells(set_autor.Row, 4).Value
			
			Editar_autor.check_btn = False
			Editar_autor.Show
		Else
			MsgBox ("Selecione um autor. ")
			Exit Sub
		End If
	Case "Legislação"
		Dim set_lei As Range
		If ComboBox2.Text <> "" Then
			Editar_lei.T_ano.Text = ComboBox2.Text
			Set set_lei = Sheets("Legislação").Range("A1", "A500").Find(Plan1.ComboBox2.Text)
				Editar_lei.T_ano.Text = Sheets("Legislação").Cells(set_lei.Row, 1).Value
				Editar_lei.T_Lei.Text = Sheets("Legislação").Cells(set_lei.Row, 2).Value
				Editar_lei.T_descricacao = Sheets("Legislação").Cells(set_lei.Row, 3).Value
				Editar_lei.T_data_beneficiario.Text = Sheets("Legislação").Cells(set_lei.Row, 4).Value
				Editar_lei.T_data_analise.Text = Sheets("Legislação").Cells(set_lei.Row, 5).Value
				Editar_lei.T_data_apresentacao.Text = Sheets("Legislação").Cells(set_lei.Row, 6).Value
				Editar_lei.T_data_limite.Text = Sheets("Legislação").Cells(set_lei.Row, 7).Value
				
			Editar_lei.check_btn = False
			Editar_lei.Show
		
		Else
			MsgBox ("Selecione uma Lei.")
			Exit Sub
		End If
	End Select
	End Sub

	Private Sub B_emendas_Click()

	Me.B_autor.Visible = False
	Me.B_emendas.Visible = False
	Me.B_lei.Visible = False

	Sheets("PARÂMETROS").Range("B5") = "Ok"
	'Cadastramento.who_open = True
	Cadastramento.Show

	End Sub

	Private Sub B_lei_Click()
	Me.B_autor.Visible = False
	Me.B_emendas.Visible = False
	Me.B_lei.Visible = False
	Nova_lei.check_btn = False
	Nova_lei.Show
	End Sub

	Private Sub ComboBox1_Change()
		Dim Linha As Integer
		If ComboBox1.Text = "Emendas" Then
			ComboBox2.Visible = False
		ElseIf ComboBox1.Text = "Autor" Then
			Linha = 2
			ComboBox2.Clear
			ComboBox2.Width = 200
			Call Listar("Autores", 1)
			Do Until Sheets("Autores").Cells(Linha, 2) = ""
				ComboBox2.AddItem Sheets("Autores").Cells(Linha, 2).Value
				Linha = Linha + 1
			Loop
			ComboBox2.Visible = True
		ElseIf ComboBox1.Text = "Legislação" Then
			Linha = 2
			ComboBox2.Clear
			ComboBox2.Width = 50
			Do Until Sheets("Legislação").Cells(Linha, 1) = ""
				ComboBox2.AddItem Sheets("Legislação").Cells(Linha, 1)
				Linha = Linha + 1
			Loop
			ComboBox2.Visible = True
		End If
	End Sub

	Private Sub ComboBox2_Change()
		
	End Sub

	Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
		cod_emenda = Me.Cells(ActiveCell.Row, 1).Value
		If ActiveCell.Row <= 3 And cod_emenda <> "" Then
			Cadastramento.Show
		End If
	End Sub

	Private Sub Worksheet_SelectionChange(ByVal Target As Range)

	End Sub
