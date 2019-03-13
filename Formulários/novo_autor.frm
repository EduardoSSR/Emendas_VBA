VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} novo_autor 
   Caption         =   "Adicionar Autor"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "novo_autor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "novo_autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public check_btn    As Boolean
Private Sub CommandButton1_Click()

Dim SQL             As String
Dim col             As Integer
Dim cn              As New ADODB.Connection
Dim rs              As New ADODB.Recordset
Dim ln              As Long
Dim autor           As String
Dim cargo           As String
Dim partido         As String
Dim l           As String

l = 2
autor = Me.T_autor

If autor = "" Then
    MsgBox "Digite um nome para o autor"
    Exit Sub
End If
Do Until Sheets("Autores").Cells(l, 1) = ""
    If Sheets("Autores").Cells(l, 2) = autor Then
    MsgBox "O autor: " & autor & ", já está cadastrado. "
    Exit Sub
    End If
l = l + 1
Loop

cargo = Me.T_cargo
If cargo = "" Then
    MsgBox "O cargo é obrigatório. Escolha um na lista"
    Exit Sub
End If

partido = Me.T_partido
If partido = "" Then
    MsgBox "O partido é obrigatório. Escolha um na lista"
    Exit Sub
End If

cn.Open ConexaoDB

SQL = "Insert into autor"
SQL = SQL & " (autor, cargo, partido) "
SQL = SQL & " values "
SQL = SQL & " ('" & autor & "','"
SQL = SQL & cargo & "',"
SQL = SQL & "'" & partido & "')"

rs.Open SQL, cn

cn.Close

Call Listar("Autores", 1)

If autor = "" Then
    MsgBox "Digite um nome para o autor"
    Exit Sub
End If
Do Until Sheets("Autores").Cells(l, 1) = ""
    If Sheets("Autores").Cells(l, 2) = autor Then
    MsgBox "Autor: " & autor & ", cadastrado com sucesso. "
    Cadastramento.f_cod_autor = "'" & Sheets("Autores").Cells(l, 1) & "'"
    End If
l = l + 1
Loop

Me.Hide

If check_btn Then


    Call Listar_Emendas
    
    Cadastramento.Hide
    Cadastramento.T_autor.Text = autor
    Cadastramento.T_cargo.Text = cargo
    Cadastramento.T_partido.Text = partido
    Cadastramento.T_autor.AddItem autor
    
    Cadastramento.check_autor = True
    
    Cadastramento.T_autor.BackColor = RGB(151, 247, 162)
    Cadastramento.CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\pencil3.ico")
    Cadastramento.CommandButton2.Enabled = True
        
    Cadastramento.Show
'Call Inserir
End If

End Sub


Private Sub T_autor_Change()

End Sub

Private Sub UserForm_Initialize()
Dim Linha As Integer
Linha = 2
Do Until Sheets("Dados_autor").Cells(Linha, 1) = ""
Me.T_cargo.AddItem Sheets("Dados_autor").Cells(Linha, 1)
Linha = Linha + 1
Loop
Linha = 2
Do Until Sheets("Dados_autor").Cells(Linha, 2) = ""
Me.T_partido.AddItem Sheets("Dados_autor").Cells(Linha, 2)
Linha = Linha + 1
Loop


End Sub
