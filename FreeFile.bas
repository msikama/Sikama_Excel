Dim MyIndex, FileNumber
For MyIndex = 1 To 5    ' Fazer o loop 5 vezes.
    FileNumber = FreeFile    ' Obt�m o n�mero de arquivo
        ' usu�rio.
    Open "TEST" & MyIndex For Output As #FileNumber    ' Cria o nome de arquivo.
    Write #FileNumber, "Este � um exemplo."    ' Gera texto.
    Close #FileNumber    ' Fecha o arquivo.
Next MyIndex

