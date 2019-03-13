Attribute VB_Name = "Criar_HTML"
Function BuildHtmlBody(lei As String, q_data As Integer, validade As Integer)
    Dim oSheet
    Set oSheet = ThisWorkbook.Sheets("Principal")
    Dim I, rows
    

    Dim html, name, address, age, department, data_limite, descricao
    
    Select Case q_data
        Case 4
            descricao = "Data de Indicação de Beneficiário "
        Case 5
            descricao = "Data de Cadastramento da Proposta "
        Case 6
            descricao = "Data de Análise da Proposta "
        Case 7
            descricao = "Data Limite para Celebração do Convênio"
        Case 8
            descricao = "Data Limite para Celebração do Convênio"
        Case 9
            descricao = "Data Limite para Celebração do Convênio"
        Case 10
            descricao = "Data Limite para Celebração do Convênio"
    End Select
        
    
    rows = Sheets("Legislação").UsedRange.rows.Count
    For I = 2 To rows
        If Sheets("Legislação").Cells(I, 1) = lei Then
            If q_data <= 7 Then
                data_limite = Sheets("Legislação").Cells(I, q_data)
            Else
                data_limite = Sheets("Legislação").Cells(I, 7)
            End If
        End If
    Next I
    
    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:'Segoe UI', Calibri, Arial, Helvetica; font-size: 14px; max-width: 768px;"">"
    html = html & "Olá, <br />" '<br />Este é um email de teste, utilizando o VBA. <br />"
    html = html & "A <b>" & descricao & " (" & data_limite & ")</b> das seguintes emendas está se aproximando (Faltam <b>" & validade & "</b> dias):<br /><br />"
    html = html & "<table style='border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"

    ' Build a html table based on rows data
    
    rows = oSheet.UsedRange.rows.Count
    For I = 2 To rows
        If oSheet.Cells(I, 2) = lei Then
        
        name = Trim(oSheet.Cells(I, 2))
        address = Trim(oSheet.Cells(I, 3))
        age = Trim(oSheet.Cells(I, 4))
        department = Trim(oSheet.Cells(I, 6))
        
        html = html & "<tr>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & name & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & address & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & age & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & department & "</td>"
        html = html & "</tr>"
        
        ElseIf oSheet.Cells(I, 2) = "ano" Then
        
        name = Trim(oSheet.Cells(I, 2))
        address = Trim(oSheet.Cells(I, 3))
        age = Trim(oSheet.Cells(I, 4))
        department = Trim(oSheet.Cells(I, 6))
        
        html = html & "<tr>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & name & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & address & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & age & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & department & "</td>"
        html = html & "</tr>"
        
        End If
    Next

    html = html & "</table></div></body></html>"
    BuildHtmlBody = html
End Function
