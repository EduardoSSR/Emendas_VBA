VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Nova_lei 
   Caption         =   "Adicionar Nova Lei"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295.001
   OleObjectBlob   =   "Nova_lei.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Nova_lei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DATA1 As DateTimePicker
Public DATA2 As DateTimePicker
Public DATA3 As DateTimePicker
Public DATA4 As DateTimePicker

Dim CHECK_1 As Boolean
Dim CHECK_2 As Boolean
Dim CHECK_3 As Boolean
Dim CHECK_4 As Boolean

Public check_btn As Boolean

Dim check_ano As Boolean



Private Sub CommandButton1_Click()

If Not check_ano Then
    MsgBox "Digite um ano válido"
    Exit Sub
End If

Dim SQL                 As String
Dim col                 As Integer
Dim cn                  As New ADODB.Connection
Dim rs                  As New ADODB.Recordset
Dim ln                  As Long
Dim ano                 As String
Dim lei                 As Variant
Dim descricao           As Variant
Dim data_analise        As Variant
Dim data_apresentacao   As Variant
Dim data_beneficiario   As Variant
Dim data_limite         As Variant
Dim jocker              As Variant

ano = Me.T_ano.Text

lei = Me.T_Lei
descricao = Me.T_descricacao

data_analise = Me.T_data_analise.Text
data_apresentação = Me.T_data_apresentacao.Text
data_beneficiario = Me.T_data_beneficiario.Text
data_limite = Me.T_data_limite.Text


If CHECK_3 Then data_analise = Me.DATA3.Value(1)
If CHECK_2 Then data_apresentacao = Me.DATA3.Value(1)
If CHECK_1 Then data_beneficiario = Me.DATA1.Value(1)
If CHECK_4 Then data_limite = Me.DATA4.Value(1)


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

SQL = "INSERT INTO `emenda_db`.`legislacao` (`ano`, `legislacao`, `descricao`, `dt_indicacao_beneficiario`"
SQL = SQL & ", `dt_cadastramento_proposta`, `dt_analise_proposta`, `dt_celebracao_convenio`) VALUES ('"
SQL = SQL & ano & "','"
SQL = SQL & lei & "','"
SQL = SQL & descricao & "',"
SQL = SQL & data_beneficiario & ","
SQL = SQL & data_analise & ","
SQL = SQL & data_apresentacao & ","
SQL = SQL & data_limite & ");"

rs.Open SQL, cn


cn.Close

MsgBox "Nova Lei Adicionada "
Me.Hide

If check_btn Then
    Call Listar_Emendas
    Cadastramento.Hide
    Cadastramento.T_ano_emenda.Text = ano
    Cadastramento.check_ano = True
    Cadastramento.T_Lei.Text = lei
    Cadastramento.T_descricacao.Text = descricao
    If CHECK_3 Then Cadastramento.T_data_analise.Text = Me.DATA3.Value(1)
    If CHECK_2 Then Cadastramento.T_data_apresentacao.Text = Me.DATA2.Value(1)
    If CHECK_1 Then Cadastramento.T_data_beneficiario.Text = Me.DATA1.Value(1)
    If CHECK_4 Then Cadastramento.T_data_limite.Text = Me.DATA4.Value(1)
    
    Cadastramento.T_ano_emenda.BackColor = RGB(151, 247, 162)
       
    'Call Inserir
End If

End Sub

Private Sub T_ano_AfterUpdate()
Dim t As Integer
Dim ano As String
t = 0

ano = T_ano.Text

Linha = 2
Do Until Sheets("Legislação").Cells(Linha, 1) = ""
    If Sheets("Legislação").Cells(Linha, 1).Text = ano Then
     t = t + 1
    End If
Linha = Linha + 1
Loop
If t > 0 Then
    check_ano = False
    T_ano.BackColor = RGB(251, 172, 159)
    Me.Alerta_ano.Caption = "Já existe legislação com esse ano."
    Me.Alerta_ano.Visible = True
Else
    check_ano = True
    T_ano.BackColor = RGB(151, 247, 162)
    Me.T_data_limite = "31/12/" & Me.T_ano.Text
    Me.Alerta_ano.Visible = False
End If
If T_ano.Text = "" Then
    check_ano = False
    T_ano.BackColor = RGB(251, 172, 159)
    Me.Alerta_ano.Visible = True
    Me.Alerta_ano.Caption = "Ano é um campo obrigatório."
End If
End Sub

Private Sub T_ano_Change()
If Len(T_ano.Text) > 0 Then
            If Not IsNumeric(Right(T_ano.Value, 1)) Then
            T_ano = Left(T_ano, Len(T_ano.Text) - 1)
        End If
End If
End Sub

Private Sub T_ano_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_ano.MaxLength = 4
End Sub

Private Sub T_data_analise_DropButtonClick()
 CHECK_3 = True
 Set DATA3 = New DateTimePicker
 With DATA3
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
Set DATA2 = New DateTimePicker
With DATA2
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
Set DATA1 = New DateTimePicker
 With DATA1
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
Set DATA4 = New DateTimePicker
    With DATA4
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
    
    If T_ano.Text <> "" Then
    check_ano = True
    End If
    
End Sub

