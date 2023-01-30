# Botão para deixar a planilha com a tela cheia


## Tela cheia
    Sub lsLigarTelaCheia()
        'Oculta todas as guias de menu
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        
        'Ocultar barra de fórmulas
        Application.DisplayFormulaBar = False
        
        'Ocultar barra de status, disposta ao final da planilha
        Application.DisplayStatusBar = False
        
        'Alterar o nome do Excel
        Application.Caption = "Sistema financeiro Optical"
        
        With ActiveWindow
            'Ocultar barra horizontal
            .DisplayHorizontalScrollBar = True
            
            'Ocultar barra vertical
            .DisplayVerticalScrollBar = True
            
            'Ocultar guias das planilhas
            .DisplayWorkbookTabs = False
            
            'Oculta os títulos de linha e coluna
            .DisplayHeadings = False
            
            'Oculta valores zero na planilha
            .DisplayZeros = True
            
            'Oculta as linhas de grade da planilha
            .DisplayGridlines = False
        End With
    End Sub

## Minimizar
    Sub lsDesligarTelaCheia()
        'Reexibe os menus
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        
        'Reexibir a barra de fórmulas
        Application.DisplayFormulaBar = True
        
        'Reexibir a barra de status, disposta ao final da planilha
        Application.DisplayStatusBar = True
        
        'Reexibir o cabeçalho da Pasta de trabalho
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
            
            'Reexibir os títulos de linha e coluna
            .DisplayHeadings = True
            
            'Reexibir valores zero na planilha
            .DisplayZeros = True
            
            'Reexibir as linhas de grade da planilha
            .DisplayGridlines = True
        End With
    End Sub