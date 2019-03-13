VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Editar_autor 
   Caption         =   "Editar Autor"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Editar_autor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Editar_autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public check_btn As Boolean
Public cod_autor       As String
Private Sub CommandButton1_Click()
Dim SQL             As String
Dim col             As Integer
Dim cn              As New ADODB.Connection
Dim rs              As New ADODB.Recordset
Dim ln              As Long
Dim autor           As String
Dim cargo           As String
Dim partido         As String
Dim local_autor     As Range

'MsgBox cod_autor

autor = Me.T_autor
If autor = "" Then
    MsgBox "Digite um nome para o autor"
    Exit Sub
End If




cargo = Me.T_cargo
If cargo = "" Then
    MsgBox "Selecione um cargo."
    Exit Sub
End If
partido = Me.T_partido
If partido = "" Then
    MsgBox "Selecione um partido."
    Exit Sub
End If
cn.Open ConexaoDB


SQL = "UPDATE autor SET `autor` = '"
SQL = SQL & autor & "',"
SQL = SQL & "`cargo` = '"
SQL = SQL & cargo & "',"
SQL = SQL & "`partido` ='"
SQL = SQL & partido & "' WHERE `cod_autor`='"
SQL = SQL & cod_autor & "';"


rs.Open SQL, cn


cn.Close

MsgBox "Autor Atualizado. "


Call Listar_Emendas
Me.Hide
If check_btn Then
    'Call Listar_Emendas
    Cadastramento.Hide
    Cadastramento.T_autor.Text = autor
    Cadastramento.T_cargo.Text = cargo
    Cadastramento.T_partido.Text = partido
    'Cadastramento.T_autor.AddItem autor
    Cadastramento.check_autor = True
        Cadastramento.T_autor.BackColor = RGB(151, 247, 162)
        'Cadastramento.CommandButton2.Picture = LoadPicture("R:\DEIDI\COSIS\ATUAL\ARQUIVOS DE SUPORTE\EMENDA PARLAMENTAR\Itens_Cadastro\block.ico")
        'Cadastramento.CommandButton2.Enabled = False
    Cadastramento.Show
    'Call Inserir
'Else
'    Plan1.ComboBox1.Text = "Emendas"
'    Plan1.check_ed_autor = True
Else
    Plan1.ComboBox1.Text = "Emendas"
    Plan1.ComboBox2.Visible = False
    
End If
End Sub

Private Sub T_autor_AfterUpdate()
Dim t               As Integer
Dim Linha           As Integer

t = 0
Linha = 2
Do Until Sheets("Autores").Cells(Linha, 1) = ""
    If Sheets("Autores").Cells(Linha, 2).Value = T_autor.Text Then
        If Sheets("Autores").Cells(Linha, 1).Value <> cod_autor Then
            t = t + 1
        End If
    End If
    If t > 0 Then
        MsgBox "Já existe autor com esse nome."
        Exit Sub
    End If
    Linha = Linha + 1
Loop
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
