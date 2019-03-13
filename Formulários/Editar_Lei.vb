Public dpFrom1 As DateTimePicker
Public dpFrom2 As DateTimePicker
Public dpFrom3 As DateTimePicker
Public dpFrom4 As DateTimePicker

Dim CHECK_1 As Boolean
Dim CHECK_2 As Boolean
Dim CHECK_3 As Boolean
Dim CHECK_4 As Boolean

Public check_btn As Boolean

Private Sub CommandButton1_Click()

Dim SQL                 As String
Dim col                 As Integer
Dim cn                  As New ADODB.Connection
Dim rs                  As New ADODB.Recordset
Dim ln                  As Long
Dim ano                 As String
Dim lei                 As Variant
Dim descricao           As Variant
Dim data_analise        As String
Dim data_apresentacao   As String
Dim data_beneficiario   As String
Dim data_limite         As String
Dim jocker              As String

ano = Me.T_ano
If ano = "" Then
    MsgBox "Digite um nome para o autor"
    Exit Sub
End If


lei = Me.T_Lei
descricao = Me.T_descricacao

data_analise = Me.T_data_analise.Text
data_apresentação = Me.T_data_apresentacao.Text
data_beneficiario = Me.T_data_beneficiario.Text
data_limite = Me.T_data_limite.Text

If CHECK_3 Then data_analise = Me.dpFrom3.Value(1)
If CHECK_2 Then data_apresentacao = Me.dpFrom2.Value(1)
If CHECK_1 Then data_beneficiario = Me.dpFrom1.Value(1)
If CHECK_4 Then data_limite = Me.dpFrom4.Value(1)


If data_analise = "" Then
    data_analise = "NULL"
Else
    data_analise = "'" & Format(data_analise, "yyyy-mm-dd") & "'"
End If
If data_apresentacao = "" Then
    data_apresentacao = "NULL"
Else
    data_apresentacao = "'" & Format(data_apresentacao, "yyyy-mm-dd") & "'"
End If
If data_beneficiario = "" Then
    data_beneficiario = "NULL"
Else
    data_beneficiario = "'" & Format(data_beneficiario, "yyyy-mm-dd") & "'"
End If
If data_limite = "" Then
    data_limite = "NULL"
Else
    data_limite = "'" & Format(data_limite, "yyyy-mm-dd") & "'"
End If


cn.Open ConexaoDB


'`dt_limite_celebracao`='2019-02-13' WHERE `ano`='2019';

SQL = "UPDATE legislacao SET `legislacao` = '"
SQL = SQL & lei & "',"
SQL = SQL & "`descricao` ='"
SQL = SQL & descricao & "', "
SQL = SQL & "`dt_indicacao_beneficiario`="
SQL = SQL & data_beneficiario & ", "
SQL = SQL & "`dt_cadastramento_proposta`="
SQL = SQL & data_apresentacao & ", "
SQL = SQL & "`dt_analise_proposta`="
SQL = SQL & data_analise & ", "
SQL = SQL & " `dt_celebracao_convenio`= "
SQL = SQL & data_limite & " WHERE `ano`='"
SQL = SQL & ano & "';"


rs.Open SQL, cn


cn.Close

MsgBox "Legislação Atualizada. "
Me.Hide
If check_btn Then
    Call Listar_Emendas
    Cadastramento.Hide
    Cadastramento.T_ano_emenda.Text = ano
    Cadastramento.T_Lei.Text = lei
    Cadastramento.T_descricacao.Text = descricao
    If CHECK_3 Then Cadastramento.T_data_analise.Text = Me.dpFrom3.Value(1)
    If CHECK_2 Then Cadastramento.T_data_apresentacao.Text = Me.dpFrom2.Value(1)
    If CHECK_1 Then Cadastramento.T_data_beneficiario.Text = Me.dpFrom1.Value(1)
    If CHECK_4 Then Cadastramento.T_data_limite.Text = Me.dpFrom4.Value(1)
    
    'Cadastramento.T_autor.AddItem autor
    Cadastramento.check_ano = True
        Cadastramento.T_ano_emenda.BackColor = RGB(151, 247, 162)
    'Cadastramento.CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\block.ico")
    'Cadastramento.CommandButton2.Enabled = False
'Call Inserir
End If

End Sub

Private Sub T_ano_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_ano.MaxLength = 4
End Sub


Private Sub T_data_analise_DropButtonClick()
CHECK_3 = True

Set dpFrom3 = New DateTimePicker
    With dpFrom3
        .Add T_data_analise
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFF, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub

Private Sub T_data_apresentacao_DropButtonClick()
CHECK_2 = True
Set dpFrom2 = New DateTimePicker
    With dpFrom2
        .Add T_data_apresentacao
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFA, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub


Private Sub T_data_beneficiario_DropButtonClick()
CHECK_1 = True
Set dpFrom1 = New DateTimePicker
    With dpFrom1
        .Add T_data_beneficiario
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFF, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub

Private Sub T_data_limite_DropButtonClick()
 CHECK_4 = True

 Set dpFrom4 = New DateTimePicker
    With dpFrom4
        .Add T_data_limite
        .Create Me, "dd/mmm/yyyy", _
            BackColor:=&H125FFFA, _
            TitleBack:=&H808000, _
            Trailing:=&H99FFFF, _
            TitleFore:=&HFFFFFF
    End With
End Sub

Private Sub UserForm_Initialize()
    
    CHECK_1 = False
    CHECK_2 = False
    CHECK_3 = False
    CHECK_4 = False
    If check_btn Then
        T_ano.Text = Cadastramento.T_ano_emenda.Text
        Me.T_Lei.Text = Cadastramento.T_Lei.Text
        Me.T_descricacao.Text = Cadastramento.T_descricacao.Text
        Me.T_data_analise.Text = Cadastramento.T_data_analise.Text
        Me.T_data_apresentacao.Text = Cadastramento.T_data_apresentacao.Text
        Me.T_data_beneficiario.Text = Cadastramento.T_data_beneficiario.Text
        Me.T_data_limite.Text = Cadastramento.T_data_limite.Text
    End If
    
    T_ano.Enabled = False
End Sub
