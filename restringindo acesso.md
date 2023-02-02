# Fazer com que a planilha só abra em 1 computador

## Para usufruir da verificação de máquina no Excel cole o seguinte código na pasta de trabalho Activet_workbook

    Dim CompName As String
    CompName = Environ$("ComputerName")

    'Aqui você irá colocar o nome da máquina autorizada
    If CompName <> "PC_Max" Then 

        'Mensagem de erro exibida se o nome não bater
        MsgBox "Este computador não tem direito de executar esta aplicação."
        ActiveWorkbook.Close SaveChanges:=False

    End If