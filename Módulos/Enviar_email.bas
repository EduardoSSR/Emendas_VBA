Attribute VB_Name = "Enviar_email"
Function SendEmail(lei As String, destinatario As String, copia_para As String, validade As Integer, qdata As Integer)

    'On Error GoTo Err

    Dim NewMail As Object
    Dim mailConfig As Object
    Dim fields As Variant
    Dim msConfigURL As String
    Dim dias_restantes As Integer
    Dim lin As Integer
    
            Set NewMail = CreateObject("CDO.Message")
            Set mailConfig = CreateObject("CDO.Configuration")
        
            ' load all default configurations
            mailConfig.Load -1
        
            Set fields = mailConfig.fields
        
        'Set All Email Properties
        
            With NewMail
                .Subject = "Validade Emendas Legislação de " & lei
                .From = "notification@mctic.gov.br"
                '.To = "eduardo070294@gmail.com"
                .to = destinatario
                .CC = copia_para
                .BCC = ""
                '.HTMLBody = " A emenda: " & emenda & " com número de processo: " & process & ", vencerá em " & validade & " dias."
                .HTMLBody = BuildHtmlBody(lei, qdata, validade)
                
            End With
        
            msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
            
            With fields
                'Enable SSL Authentication
                .Item(msConfigURL & "/smtpusessl") = True
        
                'Make SMTP authentication Enabled=true (1)
                .Item(msConfigURL & "/smtpauthenticate") = 1
        
                'Set the SMTP server and port Details
                'To get these details you can get on Settings Page of your Gmail Account
                .Item(msConfigURL & "/smtpserver") = "correio.mctic.gov.br"
                .Item(msConfigURL & "/smtpserverport") = 25
                .Item(msConfigURL & "/sendusing") = 2
        
                'Set your credentials of your Gmail Account
                .Item(msConfigURL & "/sendusername") = "eduardo.rodrigues@mctic.gov.br"
                .Item(msConfigURL & "/sendpassword") = "Edu-7913"
        
                'Update the configuration fields
                .Update
        
            End With
            NewMail.Configuration = mailConfig
            NewMail.Send
            
            If qdata = 7 Or qdata = 8 Or qdata = 9 Then
            lin = 2
            Do Until Sheets("Dados_Alertas").Cells(lin, 1) = ""
                If Sheets("Dados_Alertas").Cells(lin, 1) = lei Then
                    Sheets("Dados_Alertas").Cells(lin, qdata) = ""
                End If
            lin = lin + 1
            Loop
            End If
            'MsgBox ("Mail has been Sent")
        
            
'Exit_Err:
'
'    Set NewMail = Nothing
'    Set mailConfig = Nothing
'    End
'
'Err:
'    Select Case Err.Number
'
'    Case -2147220973  'Could be because of Internet Connection
'        MsgBox " Could be no Internet Connection !!  -- " & Err.Description
'
'    Case -2147220975  'Incorrect credentials User ID or password
'        MsgBox "Incorrect Credentials !!  -- " & Err.Description
'
'    Case Else   'Rest other errors
'        MsgBox "Error occured while sending the email !!  -- " & Err.Description
'    End Select
'
'    Resume Exit_Err
'
'
    
End Function
