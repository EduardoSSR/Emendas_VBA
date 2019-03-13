VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Alarmes 
   Caption         =   "Alertas"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12195
   OleObjectBlob   =   "Alarmes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Alarmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check_ano As Boolean
Dim VALID_cmail As Boolean
Dim VALID_mail As Boolean
Private Sub CommandButton1_Click()
Dim ano As Integer
Dim lin   As Integer
If check_ano Then
    ano = Me.T_ano
    lin = 2
    Do Until Sheets("Dados_Alertas").Cells(lin, 1) = ""
        If Sheets("Dados_Alertas").Cells(lin, 1) = ano Then
            If VALID_mail Then
                Sheets("Dados_Alertas").Cells(lin, 2) = Trim(Me.Tdestinatario)
            End If
            If VALID_cmail Then
                Sheets("Dados_Alertas").Cells(lin, 3) = Trim(Me.T_copiado)
            End If
            If Me.CheckBeneficiario.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 4) = "5"
            Else
                Sheets("Dados_Alertas").Cells(lin, 4) = ""
            End If
            If Me.CheckApresentacao.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 5) = "5"
            Else
                Sheets("Dados_Alertas").Cells(lin, 5) = ""
            End If
            If Me.CheckAnalise.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 6) = "5"
            Else
                Sheets("Dados_Alertas").Cells(lin, 6) = ""
            End If
            If Me.Check30.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 7) = "30"
            Else
                Sheets("Dados_Alertas").Cells(lin, 7) = ""
            End If
            If Me.Check15.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 8) = "15"
            Else
                Sheets("Dados_Alertas").Cells(lin, 8) = ""
            End If
            If Me.Check10.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 9) = "10"
            Else
                Sheets("Dados_Alertas").Cells(lin, 9) = ""
            End If
            If Me.Check5.Value = True Then
                Sheets("Dados_Alertas").Cells(lin, 10) = "5"
            Else
                Sheets("Dados_Alertas").Cells(lin, 10) = ""
            End If
        End If
    lin = lin + 1
    Loop
Else
MsgBox "Escolha uma legislação válida. "
Exit Sub
End If

Me.Hide
Sheets("Dados_Alertas").UsedRange.EntireColumn.AutoFit
End Sub

Private Sub T_copiado_AfterUpdate()
Dim copias() As String
Dim I        As Integer
Dim tam      As Integer
If T_copiado <> "" Then
    
    copias = Split(T_copiado.Text, ",")
    tam = UBound(copias) + 1
    'MsgBox tam
    For I = 0 To tam - 1
        VALID_cmail = ValidEMail(copias(I))
            If VALID_cmail = False Then
                T_copiado.BackColor = RGB(251, 172, 159)
                MsgBox "O " & I + 1 & "º email é inválido"
                Exit Sub
            Else
                T_copiado.BackColor = RGB(151, 247, 162)
            End If
    Next I
Else
    VALID_cmail = False
    T_copiado.BackColor = RGB(255, 247, 255)
End If
End Sub

Private Sub T_copiado_Change()

End Sub

Private Sub Tdestinatario_AfterUpdate()

If Tdestinatario <> "" Then
    VALID_mail = ValidEMail(Me.Tdestinatario)
        
        If VALID_mail = False Then
            Tdestinatario.BackColor = RGB(251, 172, 159)
        Else
            Tdestinatario.BackColor = RGB(151, 247, 162)
        End If
Else
    VALID_mail = False
    Tdestinatario.BackColor = RGB(255, 247, 255)
End If

End Sub

Private Sub Tdestinatario_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Dim Linha As Integer
Linha = 2
Do Until Sheets("Legislação").Cells(Linha, 1) = ""
    Me.T_ano.AddItem Sheets("Legislação").Cells(Linha, 1)
    Linha = Linha + 1
Loop
Linha = 2

End Sub

Private Sub T_ano_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
T_ano.MaxLength = 4
End Sub

Private Sub T_ano_Change()
    
    If Len(T_ano.Text) > 0 Then
            If Not IsNumeric(Right(T_ano.Value, 1)) Then
                T_ano = Left(T_ano, Len(T_ano.Text) - 1)
            End If
    End If
    
Dim t                   As Integer
Dim Linha               As Integer
Dim ano                 As String
Dim lei                 As String
Dim descricao           As String
Dim data_analise        As String
Dim data_apresentacao   As String
Dim data_beneficiario   As String
Dim data_limite         As String
Dim destinatario        As String
Dim copiado             As String
Dim I                   As Integer

t = 0
ano = T_ano.Text

If ano = "" Then
    check_ano = False
    T_ano.BackColor = RGB(251, 172, 159)
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
    T_ano.BackColor = RGB(151, 247, 162)
  
Else
    check_ano = False
    T_ano.BackColor = RGB(251, 172, 159)
End If

Linha = 1
Do Until Sheets("Legislação").Cells(Linha, 1) = ""
    Linha = Linha + 1
Loop

For I = 2 To Linha
    If Sheets("Legislação").Cells(I, 1).Text = ano Then
        lei = Sheets("Legislação").Cells(I, 2).Value
        descricao = Sheets("Legislação").Cells(I, 3).Value
        data_analise = Sheets("Legislação").Cells(I, 6).Value
        data_beneficiario = Sheets("Legislação").Cells(I, 4).Value
        data_apresentacao = Sheets("Legislação").Cells(I, 5).Value
        data_limite = Sheets("Legislação").Cells(I, 7).Value
        destinatario = Sheets("Dados_Alertas").Cells(I, 2)
        copiado = Sheets("Dados_Alertas").Cells(I, 3)
        If Sheets("Dados_Alertas").Cells(I, 4) Then
            Me.CheckBeneficiario.Value = True
        Else
            Me.CheckBeneficiario.Value = False
        End If
        If Sheets("Dados_Alertas").Cells(I, 5) Then
            Me.CheckApresentacao.Value = True
        Else
            Me.CheckApresentacao.Value = False
        End If
        If Sheets("Dados_Alertas").Cells(I, 6) Then
            Me.CheckAnalise.Value = True
        Else
            Me.CheckAnalise.Value = False
        End If
        If Sheets("Dados_Alertas").Cells(I, 7) Then
            Me.Check30.Value = True
        Else
            Me.Check30.Value = False
        End If
        If Sheets("Dados_Alertas").Cells(I, 8) Then
            Me.Check15.Value = True
        Else
            Me.Check15.Value = False
        End If
        If Sheets("Dados_Alertas").Cells(I, 9) Then
            Me.Check10.Value = True
        Else
            Me.Check10.Value = False
        End If
        
        If Sheets("Dados_Alertas").Cells(I, 10) Then
            Me.Check5.Value = True
        Else
            Me.Check5.Value = False
        End If
    End If
Next I

Me.T_Lei.Text = lei
Me.T_descricao.Text = descricao
Me.T_proposta.Text = data_analise
Me.T_apresentacao.Text = data_apresentacao
Me.T_data_beneficiario.Text = data_beneficiario
Me.T_limite.Text = data_limite
Me.T_copiado = copiado
Me.Tdestinatario = destinatario
    
End Sub
