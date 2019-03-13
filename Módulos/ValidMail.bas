Attribute VB_Name = "ValidMail"
'Validar e-mails digitados
Public Function ValidEMail(sEMail As String) As Boolean
  Dim nCharacter As Integer
  Dim Count As Integer
  Dim sLetra As String
  'Verifica se o e-mail tem no MÍNIMO 5 caracteres (a@b.c)
  If Len(sEMail) < 5 Then
    'O e-mail é inválido, pois tem menos de 5 caracteres
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
  'Verifica o número de arrobas.
  'TEM que ter """UMA""" arroba
  If Count <> 1 Then
    'O e-mail é inválido, pois tem 0 ou mais de 1 arroba
    ValidEMail = False
    MsgBox "O nº de arrobas '@' do e-mail é inválido."
    Exit Function
  Else
    'O e-mail tem 1 arroba.
    'Verificar a posição da arroba
    If InStr(sEMail, "@") = 1 Then
      'O e-mail é inválido, pois começa com uma @
      ValidEMail = False
      MsgBox "O e-mail foi iniciado com uma arroba '@'."
      Exit Function
    ElseIf InStr(sEMail, "@") = Len(sEMail) Then
      'O e-mail é inválido, pois termina com uma @
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
  'Verifica o número de pontos.
  'TEM que ter PELO MENOS UM ponto.
  If Count < 1 Then
    'O e-mail é inválido, pois não tem pontos.
    ValidEMail = False
    MsgBox "O e-mail é inválido, pois não contém pontos '.'."
    Exit Function
  Else
    'O e-mail tem pelo menos 1 ponto.
    'Verificar a posição do ponto:
    If InStr(sEMail, ".") = 1 Then
      'O e-mail é inválido, pois começa com um ponto
      ValidEMail = False
      MsgBox "O e-mail foi iniciado com um ponto '.'."
      Exit Function
    ElseIf InStr(sEMail, ".") = Len(sEMail) Then
      'O e-mail é inválido, pois termina com um ponto.
      ValidEMail = False
      MsgBox "O e-mail termina com um ponto '.'."
      Exit Function
    ElseIf InStr(InStr(sEMail, "@"), sEMail, ".") = 0 Then
      'O e-mail é inválido, pois termina com um ponto.
      ValidEMail = False
      MsgBox "O e-mail não tem nenhum ponto '.' após a arroba '@'."
      Exit Function
    End If
  End If
  nCharacter = 0
  Count = 0
  'Verifica se o e-mail não tem pontos consecutivos (..) após a arroba (@).
  If InStr(sEMail, "..") > InStr(sEMail, "@") Then
    'O e-mail é inválido, tem pontos consecutivos após o @.
    ValidEMail = False
    MsgBox "O e-mail contém pontos consecutivos '..' após o arroba '@'."
    Exit Function
  End If
  'Verifica se o e-mail tem caracteres inválidos
  For nCharacter = 1 To Len(sEMail)
    sLetra = Mid$(sEMail, nCharacter, 1)
    If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
      'O e-mail é inválido, pois tem caracteres inválidos
      ValidEMail = False
      MsgBox "Foi digitado um caracter inválido no e-mail."
      Exit Function
    End If
  Next
  nCharacter = 0
  'Bem, se a verificação chegou até aqui é porque o e-mail é válido, então...
  ValidEMail = True

End Function
'No evento que quiser usar

