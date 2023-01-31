## Este codigo serve para excluir os dados de uma listbox e da planilha de dados, selecionando a linha com o mouser no formulario e clicando no botão de excluir.

# Crie um Botão e cole o codigo no modo Click

    ' codigo para apagar as linhas da listbox e da planilha de cadastro com acionamento do botão
    On Error GoTo Erro

    'definição das variaveis
    Dim ID As Double, Linha As Double, linhalist As Double
    Dim Plan As String
    Dim C As Variant
    Dim Resp As Integer

    'define na variavel linhalist a seleção da linha inteira da listbox que estiver selecionada com o mouser
    linhalist = ListBox1.ListIndex

    'verifica se é a primeira linha da listbox, se for fechar o if
    If linhalist = 0 Then
    Exit Sub
    End If

    'pergunta ao usuario se ele pretende excluir a linha ou não
    Resp = MsgBox("Confirmar Exclusão?", VBA.vbYesNo, "EXCLUIR")

    'caso responda não, então apresenta a mensagem e fecha
    If Resp = VBA.vbNo Then
        MsgBox "Exclusão Cancelada", vbInformation, "EXCLUIR"
        Exit Sub
    End If

    'define o caminho da planilha com os dados a serem excluidos na variavel Plan
    Plan = Planilha7.Name
    'define que o ID sera o valor na primeira coluna da linha inteira.
    ID = ListBox1.List(linhalist, 0)

    'define a comparação com a coluna do ID la na planilha.
    With Worksheets(Plan).Range("A:A")
        
        'define que sera comparado o valor inteiro da coluna da planilha com o da listbox
        Set C = .Find(ID, LookIn:=xlValues, Lookat:=xlWhole)
        
        If Not C Is Nothing Then
            Linha = C.Row
            'exclui a linha da planilha
            Worksheets(Plan).Rows(Linha).Delete
            'exclui a linha da listbox
            ListBox1.RemoveItem (linhalist)
            MsgBox "Excluido com sucesso!", vbInformation, "EXCLUIR"
            
            Else
            MsgBox "Não Localizado!", vbInformation, "EXCLUIR"
            
        End If
        
    End With

    Set C = Nothing

    Exit Sub
    Erro:
    MsgBox "Erro!", vbCritical, "EXCLUIR"