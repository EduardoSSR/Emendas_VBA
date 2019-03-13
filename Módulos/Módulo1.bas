Attribute VB_Name = "Módulo1"
Global Conexao As ADODB.Connection
Dim W   As Worksheet
Dim SQL As String

Sub Conectar()

Set Conexao = New ADODB.Connection

Conexao.ConnectionString = _
"driver={mysql odbc 5.1 driver};" & _
"server=172.25.116.18;database=emenda_db;uid=eduardo;pwd=eduardo;"
'"server=172.25.116.18;database=emenda_db;uid=root;pwd=root;"
'"server=127.0.0.1;database=emenda_db;uid=root;pwd=root;"
Conexao.CursorLocation = adUseClient

Conexao.Open
If Conexao.State = 1 Then
    MsgBox (" Conectado ")
     
Else
    MsgBox "Não foi possível estabelecer comunicação com o Servidor. Verifique seu Host e sua chave/Senha.", vbCritical, "Impossível criar Banco de dados."
End If
End Sub

Sub Desconectar()

Conexao.Close
MsgBox (" Desconectado ")
Set Conexao = Nothing

End Sub

Function ConexaoDB()

'Dim arq     As String

'arq = ActiveWorkbook.Path & "\bdAlunos.accdb"

ConexaoDB = _
"driver={mysql odbc 5.1 driver};" & _
"server=172.25.116.18;database=emenda_db;uid=eduardo;pwd=eduardo;"
'"server=127.0.0.1;database=emenda_db;uid=root;pwd=root;"

End Function
Sub Inserir()

Cadastramento.Show

End Sub

Function RetornaSQL(vQuery As Integer)

Select Case vQuery
Case 1
    SQL = "select * from autor"
Case 2
    SQL = "select * from acao_orcamentaria"
Case 3
    SQL = "select * from emenda"
Case 4
    SQL = "select * from emenda_projeto"
Case 5
    SQL = "select * from fonte"
Case 6
    SQL = "select * from gnd"
Case 7
    SQL = "select * from instrumento"
Case 8
    SQL = "select * from legislacao"
Case 9
    SQL = "select * from modalidade"
Case 10
    SQL = "select * from municipio"
Case 11
    SQL = "select * from programa_governo"
Case 12
    SQL = "select * from projeto"
Case 13
    SQL = "select * from status"
Case 14
    SQL = "SELECT emenda.cod_emenda, emenda.ano, emenda.num_emenda, emenda.beneficiario, emenda.valor_emenda, autor.autor,autor.cargo,autor.partido, status.descricao AS status " _
            & "FROM emenda " _
            & "INNER JOIN autor " _
            & "ON emenda.cod_autor = autor.cod_autor " _
            & "LEFT JOIN status " _
            & "ON emenda.cod_status = status.cod_status"
End Select


RetornaSQL = SQL


End Function
Sub list()
Call Listar("Autores", 1)
End Sub

Sub Listar(planilha As String, cod_select As Integer)

Application.EnableEvents = False
Application.ScreenUpdating = False

'Atribuição de variáveis
'---------------------------------------------

Dim SQL         As String
Dim cn          As New ADODB.Connection
Dim rs          As New ADODB.Recordset
Dim I           As Long
Dim FD          As ADODB.Field
Dim col         As Integer
Dim arq         As String
Dim vCol        As Range
Dim vRNG        As Range
Dim sel         As String
Set W = Sheets(planilha)

'W.Select
'W.Range("A1").Select

'Desproteger a planilha
'---------------------------------------------
W.Unprotect ("1234")

W.AutoFilterMode = False
'Apagar as células utilizadas anteriormente
'---------------------------------------------
With W.Range("A1")
    .CurrentRegion.ClearContents
    .CurrentRegion.ClearFormats
End With

'Iniciar a inserção dos dados
'---------------------------------------------
col = 1

'Criar a conexão com o Banco de Dados
'---------------------------------------------
Set cn = New ADODB.Connection

'Abrir a conexão de dados
'----------------------------------------------
cn.Open ConexaoDB

'Criar um recordset
'----------------------------------------------
Set rs = New ADODB.Recordset

SQL = RetornaSQL(cod_select)

'Realiza a consulta
'-----------------------------------------------
rs.Open SQL, cn

'Verifica se há dados no Recordset
'-----------------------------------------------
If rs.EOF = False Then  'EOF = End of File

    'W.Range("A1").Select
    
    'Adicionar o nome das colunas
    '-------------------------------------------
    
    For Each FD In rs.fields
    
        With W.Cells(1, col)
            .Value = FD.name
            .Font.Bold = True
            .Interior.Color = RGB(196, 215, 155)
        End With
                
        col = col + 1
        
    Next FD
    
    'Inserir dados do Recordset na planilha
    '------------------------------------------
    
    W.Cells(3, 1).CopyFromRecordset rs
    
    'MsgBox W.Range(a2, d12).Text
    Application.StatusBar = "Consulta concluída..."
    
    
Else
 
    Application.StatusBar = "Não há dados para serem trazidos..."
    
End If

'Fechar Recordset
'----------------------------------------------
rs.Close

'Fechar conexão com banco de dados
'----------------------------------------------
cn.Close

'Cria registro das alteradas
'----------------------------------------------
'With W.Range("N2")
'    .Value = "Alterada"
'    .Interior.Color = RGB(37, 219, 119)
'End With

W.UsedRange.EntireColumn.AutoFit

'Classifa planilha com última classificação
'realizada
'-----------------------------------------------
Set vCol = W.Range(Sheets("PARÂMETROS").Range("B3").Value)

'W.Range("A3").CurrentRegion.Select
Set vRNG = W.Range("A3")

W.Sort.SortFields.Clear
W.Sort.SortFields.Add Key:=vCol, _
                      SortOn:=xlSortOnValues, _
                      Order:=xlAscending, _
                      DataOption:=xlSortNormal
With W.Sort
    .SetRange vRNG
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

W.Cells(2, 1).EntireRow.Delete
'W.Range("A1", "D1").Select
If W.Range("A1").Value = True Then
W.Range("A1", "N1").AutoFilter
End If
If planilha = "Legislação" Then
'MsgBox "Arrumei as datas..."
W.Range("D2", "G1000").NumberFormat = "m/d/yyyy"
End If

W.Protect _
    Password:="1234", _
    AllowFiltering:=True
'W.Range("A2").Select


Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada"

End Sub

Sub Listar_Emendas()

Application.EnableEvents = False
Application.ScreenUpdating = False

'Atribuição de variáveis
'---------------------------------------------

Dim SQL         As String
Dim cn          As New ADODB.Connection
Dim rs          As New ADODB.Recordset
Dim I           As Long
Dim FD          As ADODB.Field
Dim col         As Integer
Dim arq         As String
Dim vCol        As Range
Dim vRNG        As Range
Dim sel         As String

Set W = Sheets("Principal")

W.Select
W.Range("A1").Select

'Desproteger a planilha
'---------------------------------------------
W.Unprotect ("1234")

W.AutoFilterMode = False
'Apagar as células utilizadas anteriormente
'---------------------------------------------
With W.Range("A2")
    .CurrentRegion.ClearContents
    .CurrentRegion.ClearFormats
End With

'Iniciar a inserção dos dados
'---------------------------------------------
col = 1

'Criar a conexão com o Banco de Dados
'---------------------------------------------
Set cn = New ADODB.Connection

'Abrir a conexão de dados
'----------------------------------------------
cn.Open ConexaoDB

'Criar um recordset
'----------------------------------------------
Set rs = New ADODB.Recordset

SQL = RetornaSQL(14)

'Realiza a consulta
'-----------------------------------------------
rs.Open SQL, cn

'Verifica se há dados no Recordset
'-----------------------------------------------
If rs.EOF = False Then  'EOF = End of File

    W.Range("A2").Select
    
    'Adicionar o nome das colunas
    '-------------------------------------------
    
    For Each FD In rs.fields
     
        With W.Cells(2, col)
            .Value = FD.name
            .Font.Bold = True
            .Interior.Color = RGB(196, 215, 155)
        End With
                
        col = col + 1
        
    Next FD
    
    'Inserir dados do Recordset na planilha
    '------------------------------------------
    
    W.Cells(4, 1).CopyFromRecordset rs
    
    'MsgBox W.Range(a2, d12).Text
    Application.StatusBar = "Consulta concluída..."
    
    
Else
 
    Application.StatusBar = "Não há dados para serem trazidos..."
    
    
End If

'Fechar Recordset
'----------------------------------------------
rs.Close

'Fechar conexão com banco de dados
'----------------------------------------------
cn.Close

'Cria registro das alteradas
'----------------------------------------------
'With W.Range("N2")
'    .Value = "Alterada"
'    .Interior.Color = RGB(37, 219, 119)
'End With

W.UsedRange.EntireColumn.AutoFit

'Classifa planilha com última classificação
'realizada
'-----------------------------------------------
'Set vCol = W.Range(Sheets("PARÂMETROS").Range("B3").Value)
'
'W.Range("A4").CurrentRegion.Select
'Set vRNG = Selection
'
'W.Sort.SortFields.Clear
'W.Sort.SortFields.Add Key:=vCol, _
'                      SortOn:=xlSortOnValues, _
'                      Order:=xlAscending, _
'                      DataOption:=xlSortNormal
'With W.Sort
'    .SetRange vRNG
'    .Header = xlNo
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
'End With
'
W.Cells(3, 1).EntireRow.Delete

W.Columns("A:A").EntireColumn.Hidden = True
'W.Range("A2", "AC2").Select
'If W.Range("A2").Value = True Then
'Selection.AutoFilter
'End If

W.Range("e3", "e2000").NumberFormat = "$ #,##0.00"

W.Protect _
    Password:="1234", _
    AllowFiltering:=True
W.Range("A3").Select


Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada"

End Sub

Sub Desproteger()
Attribute Desproteger.VB_ProcData.VB_Invoke_Func = "u\n14"

Set W = Sheets(ActiveSheet.name)
W.Unprotect ("1234")
End Sub

Function Retorna_Tabela(Numero As Integer)
Dim SQL As String
Select Case Numero
Case 1
    SQL = "Autores"
Case 2
    SQL = "Ação_Orçamentaria"
Case 3
    SQL = "Emendas"
Case 4
    SQL = "Projeto_Emenda"
Case 5
    SQL = "Fonte"
Case 6
    SQL = "GND"
Case 7
    SQL = "Instrumento"
Case 8
    SQL = "Legislação"
Case 9
    SQL = "Modalidade"
Case 10
    SQL = "Município"
Case 11
    SQL = "Programa"
Case 12
    SQL = "Projeto"
Case 13
    SQL = "Status"

End Select


Retorna_Tabela = SQL

End Function

Public Function VerificaDuplicidade(ByRef ctrlList As MSForms.Control, _
                                    ByVal strValor As String) As Boolean

    Dim intList As Integer
    
    VerificaDuplicidade = False
    
    With ctrlList
        
        For intList = 0 To .ListCount - 1
        
            If VBA.LCase(VBA.Trim(.list(intList, 0))) = _
                VBA.LCase(VBA.Trim(strValor)) Then
            
                VerificaDuplicidade = True
                Exit Function
                
            End If
        
        Next intList
    
    End With
    
End Function

Public Function CheckCNPJ(ByVal sCNPJ As String) As Boolean
'Objetivo: esta function valida o CNPJ
'Parametro: sCnpj número do sCnpj
'Retorno: True se for válido - False se não for válido

   Dim VAR1, VAR2, VAR3, VAR4, VAR5
         
   If Len(sCNPJ) = 8 And Val(sCNPJ) > 0 Then
      VAR1 = 0
      VAR2 = 0
      VAR4 = 0
      For VAR3 = 1 To 7
         VAR1 = Val(Mid(sCNPJ, VAR3, 1))
         If (VAR1 Mod 2) <> 0 Then
            VAR1 = VAR1 * 2
         End If
         If VAR1 > 9 Then
            VAR2 = VAR2 + Int(VAR1 / 10) + (VAR1 Mod 10)
         Else
            VAR2 = VAR2 + VAR1
         End If
      Next VAR3
      VAR4 = IIf((VAR2 Mod 10) <> 0, 10 - (VAR2 Mod 10), 0)
      If VAR4 = Val(Mid(sCNPJ, 8, 1)) Then
         CheckCNPJ = True
      Else
         CheckCNPJ = False
      End If
   Else
      If Len(sCNPJ) = 14 And Val(sCNPJ) > 0 Then
         VAR1 = 0
         VAR3 = 0
         VAR4 = 0
         VAR5 = 0
         VAR2 = 5
         For VAR3 = 1 To 12
            VAR1 = VAR1 + (Val(Mid(sCNPJ, VAR3, 1)) * VAR2)
            VAR2 = IIf(VAR2 > 2, VAR2 - 1, 9)
         Next VAR3
         VAR1 = VAR1 Mod 11
         VAR4 = IIf(VAR1 > 1, 11 - VAR1, 0)
         VAR1 = 0
         VAR3 = 0
         VAR2 = 6
         For VAR3 = 1 To 13
            VAR1 = VAR1 + (Val(Mid(sCNPJ, VAR3, 1)) * VAR2)
            VAR2 = IIf(VAR2 > 2, VAR2 - 1, 9)
         Next VAR3
         VAR1 = VAR1 Mod 11
         VAR5 = IIf(VAR1 > 1, 11 - VAR1, 0)
         If (VAR4 = Val(Mid(sCNPJ, 13, 1)) And VAR5 = Val(Mid(sCNPJ, 14, 1))) Then
            CheckCNPJ = True
         Else
            CheckCNPJ = False
         End If
      Else
         CheckCNPJ = False
      End If
   End If
End Function
Public Function SeparaNumeros(rng As String, e_num As Boolean) As String
    Dim X As Long, xstr As String
    X = VBA.Len(rng)
 
    For I = 1 To X
 
    xstr = VBA.Mid(rng, I, 1)
 
    If ((VBA.IsNumeric(xstr) And e_num) Or (Not (VBA.IsNumeric(xstr)) And Not (e_num))) Then
 
    SeparaNumeros = SeparaNumeros + xstr
 
    End If
    Next

End Function

Public Function Verifica_Obrigatorios(check As Integer) As String

Dim resposta As String

Select Case check
Case 0
    resposta = "OK"
Case 1
    resposta = "Número da Emenda é inválido."
Case 2
    resposta = "Autor inválido. Escolha um na lista, ou adicione outro. "
Case 3
    resposta = " Ano da Legislação inválido. Escolha um na lista, ou adicione outro. "
Case 4
    resposta = " A emenda precisa de um 'Valor'. "
Case 5
    resposta = " O campo GND é obrigatório, selecione-o na lista. "
Case 6
    resposta = " O campo Fonte é obrigatório, selecione-o na lista."
Case 7
    resposta = " O campo Programa é obrigatório, selecione-o na lista."
Case 8
    resposta = " O campo Modalidade é obrigatório, selecione-o na lista."
Case 9
    resposta = " O campo Ação é obrigatório, selecione-o na lista."
Case 10
    resposta = " Preencha o campo 'Status' com um valor válido, ou deixe-o em branco."
Case 11
    resposta = " Preencha o campo 'Instrumento' com um valor válido, ou deixe-o em branco."

End Select

Verifica_Obrigatorios = resposta

End Function

Sub set_combo1()
Plan1.ComboBox1.Clear

Plan1.ComboBox1.AddItem "Emendas"
Plan1.ComboBox1.AddItem "Autor"
Plan1.ComboBox1.AddItem "Legislação"
Plan1.ComboBox1.Text = "Emendas"
'Plan1.ComboBox1.TopIndex (0)

End Sub
