# Carregar um grafico como GIF em um forms VBA
## Crie um modulo normal para esta função  
    
    Sub Carregar_Grafico1()

    'Tratamento de erro'
    On Error GoTo Erro

    Planilha1.Activate 'altere para a planilha que contenha o grafico

    'Variaveis'
    Dim Plan As String
    Dim PastaNome As String

    'Define que a variavel plan vai receber o nome da planilha1'
    Plan = Planilha1.Name
    'seleciona o grafico1 da planilha armazenada em plan'
    CurrentChart = Sheets(Plan).ChartObjects(1).Activate
    Set CurrentChart = Sheets(Plan).ChartObjects(1).Chart

    'Prepara para salvar a imagem do grafico na mesma pasta da planilha'
    PastaNome = ThisWorkbook.Path & Application.PathSeparator & "grafico1.gif"
    CurrentChart.Export Filename:=PastaNome, filtername:="GIF"
    'Salva a imagem do grafico no fomrato gif na mesma pasta que está a planilha salva'
    Image1.Picture = LoadPicture(PastaNome)

    Planilha3.Activate 'altere para a planilha que pretende que fique aberta depois de carregar

    Exit Sub
    Erro:
    MsgBox "Erro!", vbCritical, "ERRO"

    End Sub