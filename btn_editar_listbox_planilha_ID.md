## Botão que edita os valores de uma planilha a través da seleção dos dados por uma listbox com o mouse

# Crie um botão e cole no modo click

    ' codigo para alterar o cadastro de acordo com o ID
    On Error GoTo Erro

    'desprotege a planilha
    Planilha7.Unprotect "123"

    Dim Resp As String

    'pergunta ao usuario se quer mesmo editar
    Resp = MsgBox("Confirmar Alteração?", vbYesNo, "CONFIRMAR")

    'se a resposta fpo não
    If Resp = VBA.vbNo Then
        MsgBox "Alteração Cancelada", vbInformation, "ALTERAR"
        Exit Sub
        
        'se for sim
        ElseIf Resp = VBA.vbYes Then
        
        Dim IDpesquisar As Double
        
        If CadastroClientes.TextBox36.Text = "" Then
            
            MsgBox "Não existem um ID para alterar!", vbInformation, "EDITAR"
            Exit Sub
            
        Else
            'compara se o id que está na coluna A:A, é igual ao ID que está escrito na textbox36 no forms
            IDpesquisar = Planilha7.Range("A:A").Find(CadastroClientes.TextBox36.Value, lookat:=1).Row
        
            'define que o ID comparado na variavel IDpesquisar, vai pular para a proxima coluna e atribuir o valor do textbox_nomecompleto do forms
            Planilha7.Cells(IDpesquisar, 2) = CadastroClientes.TextBox_nomecompleto.Value
            Planilha7.Cells(IDpesquisar, 3) = CadastroClientes.TextBox_os.Value
            Planilha7.Cells(IDpesquisar, 4) = CadastroClientes.TextBox_fixo.Value
            Planilha7.Cells(IDpesquisar, 5) = CadastroClientes.TextBox_celular.Value
            Planilha7.Cells(IDpesquisar, 6) = CadastroClientes.TextBox_cpf.Value
            Planilha7.Cells(IDpesquisar, 7) = CadastroClientes.TextBox_cep.Value
            Planilha7.Cells(IDpesquisar, 8) = CadastroClientes.TextBox_endereco.Value
            Planilha7.Cells(IDpesquisar, 9) = CadastroClientes.TextBox_complemento.Value
            Planilha7.Cells(IDpesquisar, 10) = CadastroClientes.TextBox_bairro.Value
            Planilha7.Cells(IDpesquisar, 11) = CadastroClientes.TextBox_localidade.Value
            Planilha7.Cells(IDpesquisar, 12) = CadastroClientes.TextBox_uf.Value
            Planilha7.Cells(IDpesquisar, 13) = CadastroClientes.TextBox_observacao.Value
            Planilha7.Cells(IDpesquisar, 14) = CadastroClientes.TextBox_data.Value
            Planilha7.Cells(IDpesquisar, 15) = CadastroClientes.TextBox12.Value
            Planilha7.Cells(IDpesquisar, 16) = CadastroClientes.TextBox13.Value
            Planilha7.Cells(IDpesquisar, 17) = CadastroClientes.TextBox14.Value
            Planilha7.Cells(IDpesquisar, 18) = CadastroClientes.TextBox15.Value
            Planilha7.Cells(IDpesquisar, 19) = CadastroClientes.TextBox16.Value
            Planilha7.Cells(IDpesquisar, 20) = CadastroClientes.TextBox17.Value
            Planilha7.Cells(IDpesquisar, 21) = CadastroClientes.TextBox35.Value
            Planilha7.Cells(IDpesquisar, 22) = CadastroClientes.TextBox19.Value
            Planilha7.Cells(IDpesquisar, 23) = CadastroClientes.TextBox20.Value
            Planilha7.Cells(IDpesquisar, 24) = CadastroClientes.TextBox21.Value
            Planilha7.Cells(IDpesquisar, 25) = CadastroClientes.TextBox22.Value
            Planilha7.Cells(IDpesquisar, 26) = CadastroClientes.TextBox23.Value
            Planilha7.Cells(IDpesquisar, 27) = CadastroClientes.TextBox24.Value
            Planilha7.Cells(IDpesquisar, 28) = CadastroClientes.TextBox25.Value
            Planilha7.Cells(IDpesquisar, 29) = CadastroClientes.TextBox26.Value
            Planilha7.Cells(IDpesquisar, 30) = CadastroClientes.TextBox27.Value
            Planilha7.Cells(IDpesquisar, 31) = CadastroClientes.TextBox34.Value
            Planilha7.Cells(IDpesquisar, 32) = CadastroClientes.TextBox31.Value
            Planilha7.Cells(IDpesquisar, 33) = CadastroClientes.TextBox32.Value
            Planilha7.Cells(IDpesquisar, 34) = CadastroClientes.TextBox33.Value

            'protege a planilha
            Planilha7.Protect "123"
        
            Unload Me
        
            Exit Sub
            
        End If
        
    End If

    Erro:
    MsgBox "Erro!", vbCritical, "EXCLUIR"