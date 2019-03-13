Attribute VB_Name = "M�dulo3"
Sub lsLigarTelaCheia()
Attribute lsLigarTelaCheia.VB_ProcData.VB_Invoke_Func = "M\n14"
    'Oculta todas as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de f�rmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = False
    
    'Alterar o nome do Excel
    Application.Caption = "Controle de manuten��o de ve�culos 3.0"
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar guias das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os t�tulos de linha e coluna
        .DisplayHeadings = False
        
        'Oculta valores zero na planilha
        .DisplayZeros = False
        
        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
End Sub

Sub lsDesligarTelaCheia()
Attribute lsDesligarTelaCheia.VB_ProcData.VB_Invoke_Func = "m\n14"
    'Reexibe os menus
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    'Reexibir a barra de f�rmulas
    Application.DisplayFormulaBar = True
    
    'Reexibir a barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = True
    
    'Reexibir o cabe�alho da Pasta de trabalho
    ActiveWindow.DisplayHeadings = True
    
    'Retornar o nome do Excel
    Application.Caption = ""
    
    With ActiveWindow
        'Reexibir barra horizontal
        .DisplayHorizontalScrollBar = True
        
        'Reexibir barra vertical
        .DisplayVerticalScrollBar = True
        
        'Reexibir guias das planilhas
        .DisplayWorkbookTabs = True
        
        'Reexibir os t�tulos de linha e coluna
        .DisplayHeadings = True
        
        'Reexibir valores zero na planilha
        .DisplayZeros = True
        
        'Reexibir as linhas de grade da planilha
        .DisplayGridlines = True
    End With
End Sub
