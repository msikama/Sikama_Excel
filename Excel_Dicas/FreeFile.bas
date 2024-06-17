Dim MyIndex, FileNumber
For MyIndex = 1 To 5    ' Fazer o loop 5 vezes.
    FileNumber = FreeFile    ' Obtém o número de arquivo
        ' usuário.
    Open "TEST" & MyIndex For Output As #FileNumber    ' Cria o nome de arquivo.
    Write #FileNumber, "Este é um exemplo."    ' Gera texto.
    Close #FileNumber    ' Fecha o arquivo.
Next MyIndex

