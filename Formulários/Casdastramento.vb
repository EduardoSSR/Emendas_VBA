Option Explicit

Public dpFrom1          As DateTimePicker
Public dpFrom2          As DateTimePicker
Dim check_numEmenda     As Boolean
Public check_autor         As Boolean
Public check_ano           As Boolean
Dim check_valor         As Boolean
Dim check_gnd           As Boolean
Dim check_fonte         As Boolean
Dim check_programa      As Boolean
Dim check_modalidade    As Boolean
Dim check_acao          As Boolean

Dim check_cnpj          As Boolean
Dim check_status        As Boolean
Dim check_instrumento   As Boolean

Dim CHECK_1             As Boolean
Dim CHECK_2             As Boolean

Dim name_uf             As String
Dim name_municipio      As String

Public who_open         As Boolean

Public f_cod_emenda          As String
Public f_ano                 As String
Public f_ibge                As String
Public f_cod_autor           As String
Public f_cod_fonte           As String
Public f_cod_gnd             As String
Public f_cod_modalidade      As String
Public f_cod_programa        As String
Public f_cod_acao            As String
Public f_cod_status          As String
Public f_cod_instrumento     As String
Public f_num_emenda          As String
Public f_valor_emenda        As String
Public f_localizador         As String
Public f_cnpj                As String
Public f_beneficiario        As String
Public f_objeto              As String
Public f_proposta            As String
Public f_convenio            As String
Public f_limite              As String
Public f_empenhado           As String
Public f_nota_empenho        As String
Public f_valor_repasse       As String
Public f_valor_contrapartida As String
Public f_data_inicio         As String
Public f_data_fim            As String
Public f_impedimento         As String
Public f_obs                 As String
Public f_pendencia           As String

Public emenda                As String

Dim f_cod_projeto As String
'Dim check_valor_empenhado As Boolean


Private Sub CommandButton1_Click()

'Dim cod_projeto         As String
Dim SQL                 As String
Dim col                 As Integer
Dim cn                  As New ADODB.Connection
Dim rs                  As New ADODB.Recordset
Dim ln                  As Long
Dim ctl As MSForms.Control
Dim cod_check           As Integer
Dim Linha               As Integer
Dim qtd_proj            As Integer
Dim la                  As Integer
Dim I                   As Integer

If who_open = True Then
'MsgBox "Insert"

'-----Verificar campos obrigatórios----------------------------
    cod_check = 0
    If check_numEmenda = False Then
    cod_check = 1
    ElseIf check_autor = False Then
    cod_check = 2
    ElseIf check_ano = False Then
    cod_check = 3
    ElseIf check_valor = False Then
    cod_check = 4
    ElseIf check_gnd = False Then
    cod_check = 5
    ElseIf check_fonte = False Then
    cod_check = 6
    ElseIf check_programa = False Then
    cod_check = 7
    ElseIf check_modalidade = False Then
    cod_check = 8
    ElseIf check_acao = False Then
    cod_check = 9
    ElseIf check_status = False Then
    cod_check = 10
    ElseIf check_instrumento = False Then
    cod_check = 11
    Else
    cod_check = 0
    End If
    
    If Verifica_Obrigatorios(cod_check) <> "OK" Then
        MsgBox Verifica_Obrigatorios(cod_check)
        Exit Sub
    ElseIf check_cnpj = False Then
        If Me.T_cnpj_beneficiario <> "" Then
        MsgBox "Digite um CNPJ válido, ou deixe-o em branco."
        Exit Sub
        End If
    Else
        'MsgBox "Chegou em ok."
        f_data_inicio = Me.T_data_inicio.Text
        f_data_fim = Me.T_data_fim.Text
        
        If CHECK_1 Then f_data_inicio = "'" & Format(Me.dpFrom1.Value(1), "yyyy-mm-dd") & "'"
        
        If CHECK_2 Then f_data_fim = "'" & Format(Me.dpFrom2.Value(1), "yyyy-mm-dd") & "'"
        
        If f_data_inicio = "" Then f_data_inicio = "NULL"
        If f_data_fim = "" Then f_data_fim = "NULL"
        
        If Me.OptionButtonN.Value Then
            f_impedimento = "False"
        ElseIf Me.OptionButtonS.Value Then
            f_impedimento = "True"
        Else
            f_impedimento = "NULL"
        End If
        
        cn.Open ConexaoDB
    
        SQL = "INSERT INTO `emenda_db`.`emenda` (`ano`, `cod_ibge`, `cod_autor`, `cod_fonte`, `cod_gnd`, `cod_modalidade`, `cod_programa_governo`"
        SQL = SQL & ", `cod_acao_orcamentaria`, `cod_status`, `cod_instrumento`, `num_emenda`, `valor_emenda`, `localizador`"
        SQL = SQL & ", `cnpj_beneficiario`, `beneficiario`,`objeto`,`proposta_siconv`, `convenio_siconv`, `lim_empenho`, `empenhado`"
        SQL = SQL & ", `nota_empenho`, `valor_repasse`, `valor_contrapartida`, `dt_ini_conv`,`dt_fim_conv`, `impedimento`"
        SQL = SQL & ", `obs`, `pendencia`) VALUES ("
        SQL = SQL & f_ano & ","
        SQL = SQL & f_ibge & ","
        SQL = SQL & f_cod_autor & ","
        SQL = SQL & f_cod_fonte & ","
        SQL = SQL & f_cod_gnd & ","
        SQL = SQL & f_cod_modalidade & ","
        SQL = SQL & f_cod_programa & ","
        SQL = SQL & f_cod_acao & ","
        SQL = SQL & f_cod_status & ","
        SQL = SQL & f_cod_instrumento & ","
        SQL = SQL & f_num_emenda & ","
        SQL = SQL & f_valor_emenda & ","
        SQL = SQL & f_localizador & ","
        SQL = SQL & f_cnpj & ","
        SQL = SQL & f_beneficiario & ","
        SQL = SQL & f_objeto & ","
        SQL = SQL & f_proposta & ","
        SQL = SQL & f_convenio & ","
        SQL = SQL & f_limite & ","
        SQL = SQL & f_empenhado & ","
        SQL = SQL & f_nota_empenho & ","
        SQL = SQL & f_valor_repasse & ","
        SQL = SQL & f_valor_contrapartida & ","
        SQL = SQL & f_data_inicio & ","
        SQL = SQL & f_data_fim & ","
        SQL = SQL & f_impedimento & ","
        SQL = SQL & f_obs & ","
        SQL = SQL & f_pendencia & ");"
    
        rs.Open SQL, cn
    
    
        cn.Close
        
        Set cn = New ADODB.Connection
        
        cn.Open ConexaoDB
        
        Set rs = New ADODB.Recordset
        
        SQL = "select max(cod_emenda) from emenda;"
        
        rs.Open SQL, cn
        
        Sheets("PARÂMETROS").Cells(10, 2).CopyFromRecordset rs
        
        rs.Close
        
        cn.Close
        
        
        f_cod_emenda = "'" & Sheets("PARÂMETROS").Cells(10, 2) & "'"
            
        qtd_proj = 0
        For la = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(la) = True Then
                qtd_proj = qtd_proj + 1
            End If
        Next la
        If qtd_proj > 0 Then
            For I = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(I) = True Then
                    Linha = 2
                    Do Until Sheets("Projeto").Cells(Linha, 2) = ""
                        If Sheets("Projeto").Cells(Linha, 2) = ListBox1.list(I) Then
                            f_cod_projeto = "'" & Sheets("Projeto").Cells(Linha, 1) & "'"
                        End If
                    Linha = Linha + 1
                    Loop
                    cn.Open ConexaoDB
                    SQL = "INSERT INTO `emenda_projeto`(`cod_emenda`, `cod_projeto`) VALUES ("
                    SQL = SQL & f_cod_emenda & ","
                    SQL = SQL & f_cod_projeto & ");"
                    rs.Open SQL, cn
                    cn.Close
                    'MsgBox ListBox1.Selected(I)
                End If
            Next I
        End If
        MsgBox "Nova emenda adicionada. "
        Me.Hide
        'Dim I As Integer
        Dim tabela As String
        I = 1
        For I = 1 To 13
            tabela = Retorna_Tabela(I)
            Call Listar(tabela, I)
        Next
    
        Call Listar_Emendas
        Exit Sub
    End If

Else
'MsgBox "Update"
    cod_check = 0
    If check_numEmenda = False Then
    cod_check = 1
    ElseIf check_autor = False Then
    cod_check = 2
    ElseIf check_ano = False Then
    cod_check = 3
    ElseIf check_valor = False Then
    cod_check = 4
    ElseIf check_gnd = False Then
    cod_check = 5
    ElseIf check_fonte = False Then
    cod_check = 6
    ElseIf check_programa = False Then
    cod_check = 7
    ElseIf check_modalidade = False Then
    cod_check = 8
    ElseIf check_acao = False Then
    cod_check = 9
    ElseIf check_status = False Then
    cod_check = 10
    ElseIf check_instrumento = False Then
    cod_check = 11
    Else
    cod_check = 0
    End If
    
    If Verifica_Obrigatorios(cod_check) <> "OK" Then
        MsgBox Verifica_Obrigatorios(cod_check)
        Exit Sub
    ElseIf check_cnpj = False Then
        If Me.T_cnpj_beneficiario <> "" Then
        MsgBox "Digite um CNPJ válido, ou deixe-o em branco."
        Exit Sub
        End If
    Else
        'MsgBox "Chegou em ok."
        f_data_inicio = Me.T_data_inicio.Text
        f_data_fim = Me.T_data_fim.Text
        
        If CHECK_1 Then f_data_inicio = "'" & Format(Me.dpFrom1.Value(1), "yyyy-mm-dd") & "'"
        
        If CHECK_2 Then f_data_fim = "'" & Format(Me.dpFrom2.Value(1), "yyyy-mm-dd") & "'"
        
        If f_data_inicio = "" Then f_data_inicio = "NULL"
        If f_data_fim = "" Then f_data_fim = "NULL"
        
        If Me.OptionButtonN.Value Then
            f_impedimento = "False"
        ElseIf Me.OptionButtonS.Value Then
            f_impedimento = "True"
        Else
            f_impedimento = "NULL"
        End If
     
        cn.Open ConexaoDB


        SQL = "UPDATE emenda SET `ano` ="
        SQL = SQL & f_ano & ","
        SQL = SQL & "`cod_ibge` ="
        SQL = SQL & f_ibge & ","
        SQL = SQL & "`cod_autor` ="
        SQL = SQL & f_cod_autor & ","
        SQL = SQL & "`cod_fonte` ="
        SQL = SQL & f_cod_fonte & ","
        SQL = SQL & "`cod_gnd` ="
        SQL = SQL & f_cod_gnd & ","
        SQL = SQL & "`cod_modalidade` ="
        SQL = SQL & f_cod_modalidade & ","
        SQL = SQL & "`cod_acao_orcamentaria` ="
        SQL = SQL & f_cod_acao & ","
        SQL = SQL & "`cod_programa_governo` ="
        SQL = SQL & f_cod_programa & ","
        SQL = SQL & "`cod_status` ="
        SQL = SQL & f_cod_status & ","
        SQL = SQL & "`cod_instrumento` ="
        SQL = SQL & f_cod_instrumento & ","
        SQL = SQL & "`num_emenda` ="
        SQL = SQL & f_num_emenda & ","
        SQL = SQL & "`valor_emenda` ="
        SQL = SQL & f_valor_emenda & ","
        SQL = SQL & "`localizador` ="
        SQL = SQL & f_localizador & ","
        SQL = SQL & "`cnpj_beneficiario` ="
        SQL = SQL & f_cnpj & ","
        SQL = SQL & "`beneficiario` ="
        SQL = SQL & f_beneficiario & ","
        SQL = SQL & "`objeto` ="
        SQL = SQL & f_objeto & ","
        SQL = SQL & "`proposta_siconv` ="
        SQL = SQL & f_proposta & ","
        SQL = SQL & "`convenio_siconv` ="
        SQL = SQL & f_convenio & ","
        SQL = SQL & "`lim_empenho` ="
        SQL = SQL & f_limite & ","
        SQL = SQL & "`empenhado` ="
        SQL = SQL & f_empenhado & ","
        SQL = SQL & "`nota_empenho` ="
        SQL = SQL & f_nota_empenho & ","
        SQL = SQL & "`valor_repasse` ="
        SQL = SQL & f_valor_repasse & ","
        SQL = SQL & "`valor_contrapartida` ="
        SQL = SQL & f_valor_contrapartida & ","
        SQL = SQL & "`dt_ini_conv` ="
        SQL = SQL & f_data_inicio & ","
        SQL = SQL & "`dt_fim_conv` ="
        SQL = SQL & f_data_fim & ","
        SQL = SQL & "`impedimento` ="
        SQL = SQL & f_impedimento & ","
        SQL = SQL & "`obs` ="
        SQL = SQL & f_obs & ","
        SQL = SQL & "`pendencia` ="
        SQL = SQL & f_pendencia & " WHERE `cod_emenda`="
        SQL = SQL & f_cod_emenda & ";"
        
    
        rs.Open SQL, cn
        
        
        cn.Close
        End If
        MsgBox "Emenda Atualizada. "
        Me.Hide
        'Dim I As Integer
        'Dim tabela As String
        I = 1
        For I = 1 To 13
            tabela = Retorna_Tabela(I)
            Call Listar(tabela, I)
        Next
    
        Call Listar_Emendas
        Exit Sub
    'End If
        
End If

'-----Limpar Formulário----------------------------------------
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.Value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl
    
    
End Sub

Private Sub CommandButton2_Click()
Dim set_autor As Range
If check_autor Then
    
        Set set_autor = Sheets("Autores").Range("b1", "b500").Find(Me.T_autor.Text)
            Editar_autor.cod_autor = Sheets("Autores").Cells(set_autor.Row, 1).Value
            Editar_autor.T_autor.Text = Me.T_autor.Text
            Editar_autor.T_cargo.Text = Sheets("Autores").Cells(set_autor.Row, 3).Value
            Editar_autor.T_partido.Text = Sheets("Autores").Cells(set_autor.Row, 4).Value
        
    Editar_autor.check_btn = True
    Editar_autor.Show
Else
    novo_autor.check_btn = True
    novo_autor.Show
End If

End Sub

Private Sub CommandButton3_Click()
If ListBox1.Visible Then
'Insert.Visible = False
Dim I As Integer
CommandButton3.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\down.ico")
'MsgBox ListBox1.ListCount
For I = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(I) = True Then
        'Insert.Enabled = True
        'Me.T_projeto.AddItem ListBox1.list(I)
        If Me.Text_projeto.Text = "" Then
            Me.Text_projeto.Text = ListBox1.list(I)
        Else
            Me.Text_projeto.Text = Me.Text_projeto.Text & ", " & ListBox1.list(I)
        End If
    End If
Next I
ListBox1.Visible = False
Else
CommandButton3.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\up.ico")
'Insert.Visible = True
'Insert.Enabled = True
Text_projeto.Text = ""
ListBox1.Visible = True

End If
End Sub

Private Sub CommandButton4_Click()
If check_ano Then
    Editar_lei.check_btn = True
    Editar_lei.Show
Else
    Nova_lei.Show
    Nova_lei.check_btn = True
End If

End Sub

Private Sub Insert_Click()

Dim I As Integer
MsgBox ListBox1.ListCount
For I = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(I) = True Then
        'Insert.Enabled = True
        'Me.T_projeto.AddItem ListBox1.list(I)
        If Me.Text_projeto.Text = "" Then
            Me.Text_projeto.Text = ListBox1.list(I)
        Else
            Me.Text_projeto.Text = Me.Text_projeto.Text & ", " & ListBox1.list(I)
        End If
    End If
Next I

ListBox1.Visible = False
'T_projeto.Visible = True
Insert.Visible = False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub T_acao_Change()
Dim acao As String
Dim t As Integer
t = 0
f_cod_acao = "NULL"
If Len(T_acao) > 0 Then
    If IsNumeric(Right(T_acao.Value, 1)) Then
        T_acao = Left(T_acao, Len(T_acao.Text) - 1)
    End If
End If

acao = T_acao.Text

If acao = "" Then
    check_acao = False
    T_acao.BackColor = RGB(251, 172, 159)
End If
Dim Linha
Linha = 2
Do Until Sheets("Ação_Orçamentaria").Cells(Linha, 2) = ""
    If Sheets("Ação_Orçamentaria").Cells(Linha, 2) = acao Then
     t = t + 1
     f_cod_acao = "'" & Sheets("Ação_Orçamentaria").Cells(Linha, 1) & "'"
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_acao = True
    T_acao.BackColor = RGB(151, 247, 162)
Else
    check_acao = False
    T_acao.BackColor = RGB(251, 172, 159)
End If

End Sub

Private Sub T_ano_emenda_AfterUpdate()
Dim t As Integer
Dim Linha As Integer
Dim ano As String

t = 0
ano = T_ano_emenda.Text

If ano = "" Then
    check_ano = False
    T_ano_emenda.BackColor = RGB(251, 172, 159)
    CommandButton4.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\plus.ico")

End If
Linha = 2
Do Until Sheets("Legislação").Cells(Linha, 1) = ""

    If Sheets("Legislação").Cells(Linha, 1).Text = ano Then
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_ano = True
    T_ano_emenda.BackColor = RGB(151, 247, 162)
    f_ano = "'" & ano & "'"
    CommandButton4.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\pencil3.ico")
   
Else
    check_ano = False
    T_ano_emenda.BackColor = RGB(251, 172, 159)
    CommandButton4.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\plus.ico")
    Nova_lei.T_ano.Text = ano
End If

Dim lei As String
Dim descricao As String
Dim data_analise        As String
Dim data_apresentacao   As String
Dim data_beneficiario   As String
Dim data_limite         As String

Dim I As Integer
ano = T_ano_emenda.Text
Linha = 1

Do Until Sheets("Legislação").Cells(Linha, 1) = ""
Linha = Linha + 1
Loop

For I = 2 To Linha
    If Sheets("Legislação").Cells(I, 1).Text = ano Then
        lei = Sheets("Legislação").Cells(I, 2).Value
        descricao = Sheets("Legislação").Cells(I, 3).Value
        data_analise = Sheets("Legislação").Cells(I, 5).Value
        data_beneficiario = Sheets("Legislação").Cells(I, 4).Value
        data_apresentacao = Sheets("Legislação").Cells(I, 6).Value
        data_limite = Sheets("Legislação").Cells(I, 7).Value
    End If
Next I

Me.T_Lei.Text = lei
Me.T_descricacao.Text = descricao
Me.T_data_analise.Text = data_analise
Me.T_data_apresentacao.Text = data_apresentacao
Me.T_data_beneficiario.Text = data_beneficiario
Me.T_data_limite.Text = data_limite
Editar_lei.T_ano.Text = ano
Editar_lei.T_data_analise = data_analise
Editar_lei.T_data_apresentacao = data_apresentacao
Editar_lei.T_data_beneficiario = data_beneficiario
Editar_lei.T_data_limite = data_limite
Editar_lei.T_descricacao = descricao
Editar_lei.T_Lei = lei

End Sub

Private Sub T_ano_emenda_Change()
If Len(T_ano_emenda.Text) > 0 Then
            If Not IsNumeric(Right(T_ano_emenda.Value, 1)) Then
            T_ano_emenda = Left(T_ano_emenda, Len(T_ano_emenda.Text) - 1)
        End If
End If
End Sub

Private Sub T_ano_emenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_ano_emenda.MaxLength = 4
End Sub

Private Sub T_autor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_autor.MaxLength = 45
End Sub

Private Sub T_autor_AfterUpdate()
Dim autor As String
Dim t As Integer
t = 0
autor = Trim(T_autor.Text)

If autor = "" Then
    check_autor = False
    T_autor.BackColor = RGB(251, 172, 159)
    CommandButton2.Enabled = True
    CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\plus.ico")
End If
Dim Linha
Linha = 2
Do Until Sheets("Autores").Cells(Linha, 2) = ""

    If Sheets("Autores").Cells(Linha, 2) = autor Then
     t = t + 1
     f_cod_autor = "'" & Sheets("Autores").Cells(Linha, 1) & "'"
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_autor = True
    T_autor.BackColor = RGB(151, 247, 162)
    CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\pencil3.ico")
    Editar_autor.T_autor.Text = autor
    'Editar_autor.T_autor.Enabled = False
    CommandButton2.Enabled = True
Else
    check_autor = False
    T_autor.BackColor = RGB(251, 172, 159)
    CommandButton2.Enabled = True
    CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\plus.ico")
    novo_autor.T_autor.Text = autor
End If

End Sub

Private Sub T_autor_Change()
Dim nome_autor As String
Dim cargo As String
Dim partido As String
Dim Linha As Integer
Dim I As Integer
nome_autor = T_autor.Text
Linha = 1
Do Until Sheets("Autores").Cells(Linha, 1) = ""
Linha = Linha + 1
Loop
For I = 2 To Linha
    If Sheets("Autores").Cells(I, 2).Value = nome_autor Then
        cargo = Sheets("Autores").Cells(I, 3).Value
        partido = Sheets("Autores").Cells(I, 4).Value
    End If
Next I

Me.T_cargo = cargo
Me.T_partido = partido
End Sub

Private Sub T_beneficiario_Change()
If T_beneficiario.Text <> "" Then f_beneficiario = "'" & T_beneficiario.Text & "'"
End Sub

Private Sub T_cnpj_beneficiario_AfterUpdate()
Dim cnpj As String
cnpj = SeparaNumeros(T_cnpj_beneficiario.Text, True)
f_cnpj = "NULL"
If cnpj <> "" Then
    f_cnpj = "'" & cnpj & "'"
    T_cnpj_beneficiario.Text = Format(cnpj, "00"".""000"".""000""/""0000-00")
End If
End Sub

Private Sub T_cnpj_beneficiario_Change()
Dim cnpj As String

cnpj = SeparaNumeros(T_cnpj_beneficiario.Text, True)
check_cnpj = CheckCNPJ(cnpj)
If check_cnpj Then
    T_cnpj_beneficiario.BackColor = RGB(151, 247, 162)
Else
    T_cnpj_beneficiario.BackColor = RGB(251, 172, 159)
End If

If cnpj = "" Then
    T_cnpj_beneficiario.BackColor = RGB(255, 255, 255)
End If

End Sub

Private Sub T_cnpj_beneficiario_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_cnpj_beneficiario.MaxLength = 18 '07.454.325/0001-41
Select Case KeyAscii
Case 8 'Aceita o BACK SPACE
Case 13: SendKeys "{TAB}" 'Emula o TAB
Case 48 To 57
If T_cnpj_beneficiario.SelStart = 2 Then T_cnpj_beneficiario.SelText = "."
If T_cnpj_beneficiario.SelStart = 6 Then T_cnpj_beneficiario.SelText = "."
If T_cnpj_beneficiario.SelStart = 10 Then T_cnpj_beneficiario.SelText = "/"
If T_cnpj_beneficiario.SelStart = 15 Then T_cnpj_beneficiario.SelText = "-"
Case Else: KeyAscii = 0 'Ignora os outros caracteres
End Select
End Sub


Private Sub T_data_fim_DropButtonClick()
 CHECK_2 = True
 Set dpFrom2 = New DateTimePicker
    With dpFrom2
        .Add T_data_fim
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFA, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub


Private Sub T_data_inicio_DropButtonClick()
CHECK_1 = True

Set dpFrom1 = New DateTimePicker
    
    With dpFrom1
        .Add T_data_inicio
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFF, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub


Private Sub T_fonte_Change()
Dim fonte As String
Dim t As Integer
t = 0

If Len(T_fonte) > 0 Then
    If IsNumeric(Right(T_fonte.Value, 1)) Then
        T_fonte = Left(T_fonte, Len(T_fonte.Text) - 1)
    End If
End If

fonte = T_fonte.Text

If fonte = "" Then
    check_fonte = False
    T_fonte.BackColor = RGB(251, 172, 159)
End If
Dim Linha
Linha = 2
Do Until Sheets("Fonte").Cells(Linha, 2) = ""

    If Sheets("Fonte").Cells(Linha, 2) = fonte Then
     f_cod_fonte = "'" & Sheets("Fonte").Cells(Linha, 1) & "'"
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_fonte = True
    T_fonte.BackColor = RGB(151, 247, 162)
Else
    check_fonte = False
    T_fonte.BackColor = RGB(251, 172, 159)
End If

End Sub

Private Sub T_gnd_Change()
Dim gnd As String
Dim t As Integer
t = 0

If Len(T_gnd) > 0 Then
    If IsNumeric(Right(T_gnd.Value, 1)) Then
        T_gnd = Left(T_gnd, Len(T_gnd.Text) - 1)
    End If
End If

gnd = T_gnd.Text

If gnd = "" Then
    check_gnd = False
    T_gnd.BackColor = RGB(251, 172, 159)
End If
Dim Linha
Linha = 2
Do Until Sheets("GND").Cells(Linha, 2) = ""

    If Sheets("GND").Cells(Linha, 2) = gnd Then
        f_cod_gnd = "'" & Sheets("GND").Cells(Linha, 1) & "'"
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_gnd = True
    T_gnd.BackColor = RGB(151, 247, 162)
Else
    check_gnd = False
    T_gnd.BackColor = RGB(251, 172, 159)
End If

End Sub

Private Sub T_instrumento_Change()
Dim instrumento As String
Dim Linha As Integer
Dim t As Integer
Linha = 2
instrumento = T_instrumento.Text
t = 0

If Len(T_instrumento) > 0 Then
    If IsNumeric(Right(T_instrumento.Value, 1)) Then
        T_instrumento = Left(T_instrumento, Len(T_instrumento.Text) - 1)
    End If
End If

If instrumento <> "" Then
    Do Until Sheets("Instrumento").Cells(Linha, 2) = ""
        If Sheets("Instrumento").Cells(Linha, 2) = instrumento Then
          f_cod_instrumento = "'" & Sheets("Instrumento").Cells(Linha, 1).Value & "'"
          t = t + 1
        End If
    Linha = Linha + 1
    Loop
End If

If t > 0 Then
    check_instrumento = True
    Me.T_instrumento.BackColor = RGB(151, 247, 162)
Else
    Me.T_instrumento.BackColor = RGB(251, 172, 159)
    check_instrumento = False
    f_cod_instrumento = "NULL"
End If

If instrumento = "" Then
    T_instrumento.BackColor = RGB(255, 255, 255)
    f_cod_instrumento = "NULL"
    check_instrumento = True
End If

End Sub


Private Sub T_limite_empenho_AfterUpdate()
If T_limite_empenho.Text <> "" Then
    T_limite_empenho.Text = Format(T_limite_empenho.Text, "R$ #,##0.00")
    'check_valor_empenhado = True
    f_limite = Format(T_limite_empenho.Text, " 0.00")
    f_limite = "'" & Replace(f_limite, ",", ".") & "'"
    T_limite_empenho.BackColor = RGB(151, 247, 162)
Else
    T_limite_empenho.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub T_limite_empenho_Change()
If Len(T_limite_empenho.Text) > 0 Then
      If Not IsNumeric(Right(T_limite_empenho.Value, 1)) Then
        If Not (Right(T_limite_empenho.Text, 1)) = "," Then
           T_limite_empenho = Left(T_limite_empenho, Len(T_limite_empenho.Text) - 1)
        End If
      End If
End If
End Sub


Private Sub T_localizxador_Change()
If Me.T_localizxador.Text <> "" Then f_localizador = "'" & Me.T_localizxador.Text & "'"
End Sub

Private Sub T_modalidade_Change()

Dim modalidade As String
Dim t As Integer
t = 0

modalidade = T_modalidade.Text

If modalidade = "" Then
    check_modalidade = False
    T_modalidade.BackColor = RGB(251, 172, 159)
End If
Dim Linha
Linha = 2
Do Until Sheets("Modalidade").Cells(Linha, 2) = ""

    If Sheets("Modalidade").Cells(Linha, 2) = modalidade Then
     f_cod_modalidade = "'" & Sheets("Modalidade").Cells(Linha, 1) & "'"
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_modalidade = True
    T_modalidade.BackColor = RGB(151, 247, 162)
Else
    check_modalidade = False
    T_modalidade.BackColor = RGB(251, 172, 159)
End If

End Sub



Private Sub T_nota_empenho_Change()
If Me.T_nota_empenho <> "" Then f_nota_empenho = "'" & Me.T_nota_empenho.Text & "'"
End Sub

Private Sub T_num_emenda_AfterUpdate()

If Len(T_num_emenda.Text) < 12 Or T_num_emenda = "" Then
    check_numEmenda = False
    T_num_emenda.BackColor = RGB(251, 172, 159)
ElseIf Len(T_num_emenda.Text) = 12 Then
    check_numEmenda = True
    f_num_emenda = "'" & T_num_emenda.Text & "'"
    T_num_emenda.BackColor = RGB(151, 247, 162)
End If

End Sub

Private Sub T_num_emenda_Change()
If Len(T_num_emenda.Text) > 0 Then
        If Not IsNumeric(Right(T_num_emenda.Value, 1)) Then
            T_num_emenda = Left(T_num_emenda, Len(T_num_emenda.Text) - 1)
        End If
End If
End Sub

Private Sub T_num_emenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_num_emenda.MaxLength = 12
End Sub

Private Sub T_num_siconv_Change()
If Me.T_num_siconv <> "" Then f_convenio = "'" & Me.T_num_siconv.Text & "'"
End Sub

Private Sub T_objeto_Change()
If T_objeto.Text <> "" Then f_objeto = "'" & Me.T_objeto.Text & "'"
End Sub

Private Sub T_observacao_Change()
If T_observacao.Text <> "" Then f_obs = "'" & T_observacao.Text & "'"
End Sub

Private Sub T_pendencia_Change()
If T_pendencia.Text <> "" Then f_pendencia = "'" & T_pendencia.Text & "'"
End Sub

Private Sub T_programa_Change()
Dim programa As String
Dim t As Integer
t = 0

If Len(T_programa) > 0 Then
    If IsNumeric(Right(T_programa.Value, 1)) Then
        T_programa = Left(T_programa, Len(T_programa.Text) - 1)
    End If
End If

programa = T_programa.Text

If programa = "" Then
    check_programa = False
    T_programa.BackColor = RGB(251, 172, 159)
End If
Dim Linha
Linha = 2
Do Until Sheets("Programa").Cells(Linha, 2) = ""

    If Sheets("Programa").Cells(Linha, 2) = programa Then
     f_cod_programa = "'" & Sheets("Programa").Cells(Linha, 1) & "'"
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_programa = True
    T_programa.BackColor = RGB(151, 247, 162)
Else
    check_programa = False
    T_programa.BackColor = RGB(251, 172, 159)
End If
End Sub


Private Sub T_projeto_Change()

End Sub

Private Sub T_projeto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
T_projeto.Visible = False
Insert.Visible = True
Insert.Enabled = True
ListBox1.Visible = True
End Sub

Private Sub T_projeto_DropButtonClick()
If ListBox1.Visible = True Then
    ListBox1.Visible = False
    Insert.Visible = False
Else
    T_projeto.Clear
    Insert.Visible = True
    Insert.Enabled = True
    ListBox1.Visible = True
    'T_projeto.Visible = False
End If

End Sub

Private Sub T_proposta_Change()
If T_proposta.Text <> "" Then f_proposta = "'" & Me.T_proposta.Text & "'"
End Sub

Private Sub T_status_Change()
Dim status As String
Dim Linha As Integer
Dim t As Integer
Linha = 2
status = T_status.Text
t = 0

If Len(T_status) > 0 Then
    If IsNumeric(Right(T_status.Value, 1)) Then
        T_status = Left(T_status, Len(T_status.Text) - 1)
    End If
End If

If status <> "" Then
    Do Until Sheets("Status").Cells(Linha, 2) = ""
        If Sheets("Status").Cells(Linha, 2) = status Then
          f_cod_status = "'" & Sheets("Status").Cells(Linha, 1).Value & "'"
          t = t + 1
        End If
    Linha = Linha + 1
    Loop
End If

If t > 0 Then
    check_status = True
    Me.T_status.BackColor = RGB(151, 247, 162)
Else
    Me.T_status.BackColor = RGB(251, 172, 159)
    check_status = False
    f_cod_status = "NULL"
End If

If status = "" Then
    T_status.BackColor = RGB(255, 255, 255)
    f_cod_status = "NULL"
    check_status = True
End If

End Sub

Private Sub T_uf_Change()
T_municipio.Clear
Dim uf As String
Dim Linha As Integer

If Len(T_uf) > 0 Then
    If IsNumeric(Right(T_uf.Value, 1)) Then
        T_uf = Left(T_uf, Len(T_uf.Text) - 1)
    End If
End If

uf = T_uf.Text
'MsgBox T_uf.ListIndex

Linha = 2

Do Until Sheets("Município").Cells(Linha, 1) = ""

If Sheets("Município").Cells(Linha, 3) = uf Then
    T_municipio.AddItem Sheets("Município").Cells(Linha, 2)
    T_ibge.AddItem Sheets("Município").Cells(Linha, 1)
Linha = Linha + 1
Else
Linha = Linha + 1
End If
Loop

'linha = 0
'Do Until Me.T_municipio.ListIndex(linha) = ""
'    if
'    linha = linha + 1
'Loop
End Sub
Private Sub T_municipio_Change()
T_ibge.ListIndex = (T_municipio.ListIndex)
f_ibge = "'" & T_ibge.Text & "'"
End Sub

Private Sub T_uf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_uf.MaxLength = 2
End Sub

Private Sub T_valor_AfterUpdate()
If T_valor.Text <> "" Then
    T_valor.Text = Format(T_valor.Text, "R$ #,##0.00")
    
    f_valor_emenda = Format(T_valor.Text, " 0.00")
    f_valor_emenda = "'" & Replace(f_valor_emenda, ",", ".") & "'"
    check_valor = True
    'f_valor_emenda = ""&  &""
    T_valor.BackColor = RGB(151, 247, 162)
Else
    T_valor.BackColor = RGB(251, 172, 159)
    check_valor = False
End If
If T_valor.Value = 0 Then
    T_valor.BackColor = RGB(251, 172, 159)
    check_valor = False
End If

End Sub

Private Sub T_valor_Change()
If Len(T_valor.Text) > 0 Then

      If Not IsNumeric(Right(T_valor.Value, 1)) Then
        If Not (Right(T_valor.Text, 1)) = "," Then
           T_valor = Left(T_valor, Len(T_valor.Text) - 1)
        End If
      End If
       
End If
End Sub

Private Sub T_valor_empenhado_AfterUpdate()
If T_valor_empenhado.Text <> "" Then
    T_valor_empenhado.Text = Format(T_valor_empenhado.Text, "R$ #,##0.00")
    'check_valor_empenhado = True
    f_empenhado = Format(T_valor_empenhado.Text, " 0.00")
    f_empenhado = "'" & Replace(f_empenhado, ",", ".") & "'"
    T_valor_empenhado.BackColor = RGB(151, 247, 162)
Else
    T_valor_empenhado.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub T_valor_empenhado_Change()
If Len(T_valor_empenhado.Text) > 0 Then
      If Not IsNumeric(Right(T_valor_empenhado.Value, 1)) Then
        If Not (Right(T_valor_empenhado.Text, 1)) = "," Then
           T_valor_empenhado = Left(T_valor_empenhado, Len(T_valor_empenhado.Text) - 1)
        End If
      End If
End If
End Sub
Private Sub T_valor_repasse_AfterUpdate()
If T_valor_repasse.Text <> "" Then
    T_valor_repasse.Text = Format(T_valor_repasse.Text, "R$ #,##0.00")
    f_valor_repasse = Format(T_valor_repasse.Text, " 0.00")
    f_valor_repasse = "'" & Replace(f_valor_repasse, ",", ".") & "'"
    'check_valor_repasse = True
    T_valor_repasse.BackColor = RGB(151, 247, 162)
Else
    T_valor_repasse.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub T_valor_repasse_Change()
Dim contrapartida As Double
Dim repasse       As Double

If Len(T_valor_repasse.Text) > 0 Then
    If Not IsNumeric(Right(T_valor_repasse.Value, 1)) Then
        If Not (Right(T_valor_repasse.Text, 1)) = "," Then
            T_valor_repasse = Left(T_valor_repasse, Len(T_valor_repasse.Text) - 1)
        End If
    End If
    If Len(T_valor_repasse.Text) > 0 Then
        repasse = T_valor_repasse.Value
    Else
        repasse = 0
    End If
Else
    repasse = 0
End If
If T_valor_contrapartida.Text = "" Then
    contrapartida = 0
Else
    contrapartida = T_valor_contrapartida.Value
End If
      T_valor_convenio.Text = Format((contrapartida + repasse), "R$ #,###0.00")
End Sub
Private Sub T_valor_contrapartida_AfterUpdate()
If T_valor_contrapartida.Text <> "" Then
    T_valor_contrapartida.Text = Format(T_valor_contrapartida.Value, "R$ #,###0.00")
    f_valor_contrapartida = Format(T_valor_contrapartida.Text, " 0.00")
    f_valor_contrapartida = "'" & Replace(f_valor_contrapartida, ",", ".") & "'"
    'check_valor_contrapartida = True
    T_valor_contrapartida.BackColor = RGB(151, 247, 162)
Else
    T_valor_contrapartida.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub T_valor_contrapartida_Change()
Dim contrapartida As Double
Dim repasse       As Double

If Len(T_valor_contrapartida.Text) > 0 Then

    If Not IsNumeric(Right(T_valor_contrapartida.Value, 1)) = True Then
        If Not (Right(T_valor_contrapartida.Text, 1)) = "," Then
            T_valor_contrapartida = Left(T_valor_contrapartida, Len(T_valor_contrapartida.Text) - 1)
        End If
    End If
    If Len(T_valor_contrapartida.Text) > 0 Then
        contrapartida = T_valor_contrapartida.Value
    Else
        contrapartida = 0
    End If
Else
    contrapartida = 0
End If
If T_valor_repasse.Text = "" Then
      repasse = 0
 Else
      repasse = T_valor_repasse.Value
End If

T_valor_convenio.Text = Format((contrapartida + repasse), "R$ #,###0.00")
End Sub
Private Sub T_valor_convenio_AfterUpdate()
If T_valor_convenio.Text <> "" Then
    T_valor_convenio.Text = Format(T_valor_convenio.Value, "R$ #,###0.00")
    'check_valor_convenio = True
    T_valor_convenio.BackColor = RGB(151, 247, 162)
Else
    'T_valor_convenio.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Text_projeto_Change()

End Sub

Private Sub UserForm_Initialize()

Dim Linha As Integer
Linha = 2
Do Until Sheets("Legislação").Cells(Linha, 1) = ""
    Cadastramento.T_ano_emenda.AddItem Sheets("Legislação").Cells(Linha, 1)
    Linha = Linha + 1
Loop
'Cadastramento.T_ano_emenda.RowSource = ThisWorkbook.Sheets("Legislação").Range("ANO").Value

Linha = 2
Do Until Sheets("Autores").Cells(Linha, 2) = ""
    Cadastramento.T_autor.AddItem Sheets("Autores").Cells(Linha, 2)
    Linha = Linha + 1
Loop
   ' Me.T_ano_emenda.RowSource = Sheets("Legislação").Range("A2", "A")
Linha = 2
Do Until Sheets("GND").Cells(Linha, 2) = ""
    Cadastramento.T_gnd.AddItem Sheets("GND").Cells(Linha, 2)
    Linha = Linha + 1
Loop
    
Linha = 2
Do Until Sheets("Fonte").Cells(Linha, 2) = ""
    Cadastramento.T_fonte.AddItem Sheets("Fonte").Cells(Linha, 2)
    Linha = Linha + 1
Loop
    
Linha = 2
Do Until Sheets("Programa").Cells(Linha, 2) = ""
    Cadastramento.T_programa.AddItem Sheets("Programa").Cells(Linha, 2)
    Linha = Linha + 1
Loop

Linha = 2
Do Until Sheets("Ação_Orçamentaria").Cells(Linha, 2) = ""
    Cadastramento.T_acao.AddItem Sheets("Ação_Orçamentaria").Cells(Linha, 2)
    Linha = Linha + 1
Loop

Linha = 2
Do Until Sheets("Modalidade").Cells(Linha, 2) = ""
    Cadastramento.T_modalidade.AddItem Sheets("Modalidade").Cells(Linha, 2)
    Linha = Linha + 1
Loop

Linha = 2
Do Until Sheets("Status").Cells(Linha, 2) = ""
    Cadastramento.T_status.AddItem Sheets("Status").Cells(Linha, 2)
    Linha = Linha + 1
Loop

Linha = 2
Do Until Sheets("Instrumento").Cells(Linha, 2) = ""
    Cadastramento.T_instrumento.AddItem Sheets("Instrumento").Cells(Linha, 2)
    Linha = Linha + 1
Loop


Linha = 2
Do Until Sheets("Projeto").Cells(Linha, 2) = ""
    Cadastramento.ListBox1.AddItem Sheets("Projeto").Cells(Linha, 2)
    Linha = Linha + 1
Loop
Linha = 2
    Sheets("Município").Unprotect ("1234")
    Worksheets("Município").Range("C1").AutoFilter
    Worksheets("Município").AutoFilter.Sort.SortFields.Clear
    Worksheets("Município").AutoFilter.Sort.SortFields.Add Key:= _
        Range("C1:C5571"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With Sheets("Município").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Município").Protect ("1234")
Do Until Sheets("Município").Cells(Linha, 2) = ""
    If Not VerificaDuplicidade(Me.T_uf, Sheets("Município").Cells(Linha, 3).Value) Then
        Cadastramento.T_uf.AddItem Sheets("Município").Cells(Linha, 3)
    End If
    Linha = Linha + 1
Loop

CommandButton3.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\down.ico")

If Sheets("PARÂMETROS").Range("B5") = "Ok" Then
   ' MsgBox "Quem abriu foi o adicionar"
    who_open = True
    Me.CommandButton1.Caption = "Salvar"
    
    CHECK_1 = False
    CHECK_2 = False
    f_cod_emenda = "NULL"
    f_ano = "NULL"
    f_ibge = "NULL"
    f_cod_autor = "NULL"
    f_cod_fonte = "NULL"
    f_cod_gnd = "NULL"
    f_cod_modalidade = "NULL"
    f_cod_programa = "NULL"
    f_cod_acao = "NULL"
    f_cod_status = "NULL"
    f_cod_instrumento = "NULL"
    f_num_emenda = "NULL"
    f_valor_emenda = "NULL"
    f_localizador = "NULL"
    f_cnpj = "NULL"
    f_beneficiario = "NULL"
    f_objeto = "NULL"
    f_proposta = "NULL"
    f_convenio = "NULL"
    f_limite = "NULL"
    f_empenhado = "NULL"
    f_nota_empenho = "NULL"
    f_valor_repasse = "NULL"
    f_valor_contrapartida = "NULL"
    f_data_inicio = "NULL"
    f_data_fim = "NULL"
    f_impedimento = "NULL"
    f_obs = "NULL"
    f_pendencia = "NULL"
        
    check_numEmenda = False
    check_autor = False
    check_ano = False
    check_valor = False
    check_gnd = False
    check_fonte = False
    check_programa = False
    check_modalidade = False
    check_acao = False
    check_cnpj = True
    check_status = True
    check_instrumento = True

Else
    'MsgBox "Quem abriu foi o editar"
    who_open = False
    CommandButton1.Caption = "Atualizar"
        
    check_numEmenda = True
    T_num_emenda.BackColor = RGB(151, 247, 162)
    check_autor = True
    T_autor.BackColor = RGB(151, 247, 162)
    check_ano = True
    T_ano_emenda.BackColor = RGB(151, 247, 162)
    check_valor = True
    T_valor.BackColor = RGB(151, 247, 162)
    check_gnd = True
    T_gnd.BackColor = RGB(151, 247, 162)
    check_fonte = True
    T_fonte.BackColor = RGB(151, 247, 162)
    check_programa = True
    T_programa.BackColor = RGB(151, 247, 162)
    check_modalidade = True
    T_modalidade.BackColor = RGB(151, 247, 162)
    check_acao = True
    T_acao.BackColor = RGB(151, 247, 162)
    check_cnpj = True
    If Me.T_cnpj_beneficiario.Text <> "" Then
        Me.T_cnpj_beneficiario.BackColor = RGB(151, 247, 162)
    End If
    check_status = True
    check_instrumento = True
    
    CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\pencil3.ico")
    Linha = 2
    Dim cod_projeto As Integer
    Do Until Sheets("Projeto_Emenda").Cells(Linha, 1) = ""
        If Sheets("Projeto_Emenda").Cells(Linha, 1).Value = Plan1.cod_emenda Then
            cod_projeto = Sheets("Projeto_Emenda").Cells(Linha, 2).Value
                If Me.Text_projeto.Text = "" Then
                    Me.Text_projeto.Text = Sheets("Projeto").Cells(cod_projeto + 1, 2).Value
                Else
                    Me.Text_projeto.Text = Text_projeto.Text & ", " & Sheets("Projeto").Cells(cod_projeto + 1, 2).Value
                End If
                'ListBox1.Selected(cod_projeto) = True
                ListBox1.Selected(cod_projeto - 1) = True
            'MsgBox cod_projeto
        End If
    Linha = Linha + 1
    Loop
    
    Dim local_ibge      As Range
    Set local_ibge = Sheets("Município").Range("A1", "A7000").Find(Plan1.cod_ibge)
        name_uf = Sheets("Município").Cells(local_ibge.Row, 3).Value
        name_municipio = Sheets("Município").Cells(local_ibge.Row, 2).Value
        T_uf.Text = name_uf

End If
    

End Sub

