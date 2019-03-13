Attribute VB_Name = "ValidMail"
'Validar e-mails digitados
Public Function ValidEMail(sEMail As String) As Boolean
  Dim nCharacter As Integer
  Dim Count As Integer
  Dim sLetra As String
  'Verifica se o e-mail tem no M�NIMO 5 caracteres (a@b.c)
  If Len(sEMail) < 5 Then
    'O e-mail � inv�lido, pois tem menos de 5 caracteres
    ValidEMail = False
    MsgBox "O e-mail digitado tem menos de 5 caracteres."
    Exit Function
  End If
  'Verificar a existencia de arrobas (@) no e-mail
  For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "@" Then
      'OPA!!! Achou uma arroba!!!
      'Soma 1 ao contador
      Count = Count + 1
    End If
  Next
  'Verifica o n�mero de arrobas.
  'TEM que ter """UMA""" arroba
  If Count <> 1 Then
    'O e-mail � inv�lido, pois tem 0 ou mais de 1 arroba
    ValidEMail = False
    MsgBox "O n� de arrobas '@' do e-mail � inv�lido."
    Exit Function
  Else
    'O e-mail tem 1 arroba.
    'Verificar a posi��o da arroba
    If InStr(sEMail, "@") = 1 Then
      'O e-mail � inv�lido, pois come�a com uma @
      ValidEMail = False
      MsgBox "O e-mail foi iniciado com uma arroba '@'."
      Exit Function
    ElseIf InStr(sEMail, "@") = Len(sEMail) Then
      'O e-mail � inv�lido, pois termina com uma @
      ValidEMail = False
      MsgBox "O e-mail termina com uma arroba '@'."
      Exit Function
    End If
  End If
  nCharacter = 0
  Count = 0
  'Verificar a existencia de pontos (.) no e-mail
  For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "." Then
      'OPA!!! Achou um ponto!!!
      'Soma 1 ao contador
      Count = Count + 1
    End If
  Next
  'Verifica o n�mero de pontos.
  'TEM que ter PELO MENOS UM ponto.
  If Count < 1 Then
    'O e-mail � inv�lido, pois n�o tem pontos.
    ValidEMail = False
    MsgBox "O e-mail � inv�lido, pois n�o cont�m pontos '.'."
    Exit Function
  Else
    'O e-mail tem pelo menos 1 ponto.
    'Verificar a posi��o do ponto:
    If InStr(sEMail, ".") = 1 Then
      'O e-mail � inv�lido, pois come�a com um ponto
      ValidEMail = False
      MsgBox "O e-mail foi iniciado com um ponto '.'."
      Exit Function
    ElseIf InStr(sEMail, ".") = Len(sEMail) Then
      'O e-mail � inv�lido, pois termina com um ponto.
      ValidEMail = False
      MsgBox "O e-mail termina com um ponto '.'."
      Exit Function
    ElseIf InStr(InStr(sEMail, "@"), sEMail, ".") = 0 Then
      'O e-mail � inv�lido, pois termina com um ponto.
      ValidEMail = False
      MsgBox "O e-mail n�o tem nenhum ponto '.' ap�s a arroba '@'."
      Exit Function
    End If
  End If
  nCharacter = 0
  Count = 0
  'Verifica se o e-mail n�o tem pontos consecutivos (..) ap�s a arroba (@).
  If InStr(sEMail, "..") > InStr(sEMail, "@") Then
    'O e-mail � inv�lido, tem pontos consecutivos ap�s o @.
    ValidEMail = False
    MsgBox "O e-mail cont�m pontos consecutivos '..' ap�s o arroba '@'."
    Exit Function
  End If
  'Verifica se o e-mail tem caracteres inv�lidos
  For nCharacter = 1 To Len(sEMail)
    sLetra = Mid$(sEMail, nCharacter, 1)
    If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
      'O e-mail � inv�lido, pois tem caracteres inv�lidos
      ValidEMail = False
      MsgBox "Foi digitado um caracter inv�lido no e-mail."
      Exit Function
    End If
  Next
  nCharacter = 0
  'Bem, se a verifica��o chegou at� aqui � porque o e-mail � v�lido, ent�o...
  ValidEMail = True

End Function
'No evento que quiser usar

