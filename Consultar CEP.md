# Crie um modulo normal
## Dentro do modulo, cole o codigo a seguir.

    Function ConsultaCEP(valorCep As String, tipoCampo As String)

    Dim oXmlDoc As DOMDocument
    Dim oXmlNode As IXMLDOMNode
    Dim oXmlNodes As IXMLDOMNodeList

    Set oXmlDoc = New DOMDocument
    oXmlDoc.async = False

    oXmlDoc.Load (httpsviacep.com.brws + valorCep + xml)

    Set oXmlNodes = oXmlDoc.SelectNodes(xmlcep + tipoCampo)
        
    For Each oXmlNode In oXmlNodes
        ConsultaCEP = oXmlNode.Text
    Next

    End Function

# Em seguida dentro da textbox que receberá o CEP, cole este codigo na parte de KeyPress:
## Para pesquisar digite o CEP e ao digitar pressione a tecla 'ZERO' mais uma vez para realizar a busca

    Private Sub TextBox_cep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim valorEndereco, valorComplemento, valorBairro, valorLocalidade, valorUf As String

    If Len(TextBox_cep.Text) = 9 Then

            ' trata o campo endereço
            ' instancia a função consultaCEP passando os dois parametros
            valorEndereco = ConsultaCEP(TextBox_cep.Value, "logradouro")
            TextBox_endereco.Enabled = True
            TextBox_endereco.BackColor = RGB(255, 255, 255)
            TextBox_endereco.Text = valorEndereco
            ' trata o campo complemento
            ' instancia a função consultaCEP passando os dois parametros
            valorComplemento = ConsultaCEP(TextBox_cep.Value, "complemento")
            TextBox_complemento.Enabled = True
            TextBox_complemento.BackColor = RGB(255, 255, 255)
            TextBox_complemento.Text = valorComplemento
            ' trata o campo bairro
            ' instancia a função consultaCEP passando os dois parametros
            valorBairro = ConsultaCEP(TextBox_cep.Value, "bairro")
            TextBox_bairro.Enabled = True
            TextBox_bairro.BackColor = RGB(255, 255, 255)
            TextBox_bairro.Text = valorBairro
            ' trata o campo localidade
            ' instancia a função consultaCEP passando os dois parametros
            valorLocalidade = ConsultaCEP(TextBox_cep.Value, "localidade")
            TextBox_localidade.Enabled = True
            TextBox_localidade.BackColor = RGB(255, 255, 255)
            TextBox_localidade.Text = valorLocalidade
            ' trata o campo uf
            ' instancia a função consultaCEP passando os dois parametros
            valorUf = ConsultaCEP(TextBox_cep.Value, "uf")
            TextBox_uf.Enabled = True
            TextBox_uf.BackColor = RGB(255, 255, 255)
            TextBox_uf.Text = valorUf
            
    End If

    End Sub