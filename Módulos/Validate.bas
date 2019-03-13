Attribute VB_Name = "Validate"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Dim check As Boolean
Sub verifica_validade()
Dim W   As Worksheet
Set W = Sheets("Dados_Alertas")

Dim Linha As Integer, coluna As Integer, validade As Integer, lei As String, destinatario As String, _
    copiado As String, data As Date, qual_data As Integer, I As Integer, TESTE As String

Linha = 2
Do Until W.Cells(Linha, 1).Value = ""
    'For linha = 2 To 50
            If W.Cells(Linha, 2).Value <> "" Then
                
                lei = W.Cells(Linha, 1)
                destinatario = W.Cells(Linha, 2)
                
                If W.Cells(Linha, 3).Value <> "" Then
                    copiado = W.Cells(Linha, 3).Value
                Else
                    copiado = ""
                End If
                
                For coluna = 4 To 7
                    
                    If W.Cells(Linha, coluna) <> "" Then
                        dias = W.Cells(Linha, coluna).Value
                        I = 2
                        Do Until Sheets("Legislação").Cells(I, 1) = ""
                            If Sheets("Legislação").Cells(I, 1) = lei Then
                                If coluna > 7 Then
                                    data = Sheets("Legislação").Cells(I, 7)
                                Else
                                    data = Sheets("Legislação").Cells(I, coluna)
                                End If
                            End If
                        I = I + 1
                        Loop
                        validade = data - Date
                        
                    Else
                        dias = 0
                    End If
                    If dias > 0 Then
                        If validade <= dias And validade >= 0 Then
                            'MsgBox "Email ENVIADO"
                            Call SendEmail(lei, destinatario, copiado, validade, coluna)
                            Application.StatusBar = "Email enviado para " & destinatario
                        End If
                    End If
                Next coluna
            End If
Linha = Linha + 1
Loop
If check Then
 Call Timer
End If
End Sub

Sub Timer()
    'Application.OnTime Now + TimeValue("24:00:00"), "verifica_validade"
    Dim hora_atual
    check = True
    hora_atual = Format(DateTime.Now, "h:mm")
    If hora_atual = "11:00" Then
        check = False
        Call verifica_validade
        'MsgBox "São 9:38"
        Sleep (2000)
    End If
    Call TESTE
End Sub


Sub TESTE()
    'Call SendEmail("2020", "eduardo070294@gmail.com", "", 5, 1)
    Application.OnTime Now + TimeValue("00:01:00"), "Timer"
End Sub
