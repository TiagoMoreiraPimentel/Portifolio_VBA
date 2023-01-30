## Botão que fecha a planilha e salva, mas antes da as opção Yes, No e cencelar para o usuario escolher
# Crie um modulo e cole o codigo a seguir e depois chema a função no botão:

    Sub Fechar_Planilha_Salvar()
    Dim resposta As Integer
        Dim ANS As Integer
        resposta = vbYesNoCancel + vbQuestion + vbDefaultButton2
        ANS = MsgBox("Deseja salvar e sair desta planilha?", resposta, "Sistema Financeiro Optical")
        
        'função com botao Yes, No, Cancelar que se 'sim' salva e fecha, se 'no' a penas fecha e se 'cancelar' não faz nada.
        If ANS = vbYes Then
            ActiveWorkbook.Save
            Application.Quit
        ElseIf ANS = vbNo Then
            ThisWorkbook.Application.Quit
            ThisWorkbook.Close SaveChanges:=False
        Else
            Cancel = True
        End If
    
    End Sub