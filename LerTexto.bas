Dim MyChar
Open "TESTFILE" For Input As #1    ' Abre o arquivo.
Do While Not EOF(1)    ' Faz o loop at� o fim do arquivo.
    MyChar = Input(1, #1)    ' Obt�m um caractere.
    Debug.Print MyChar    ' Imprima na janela Immediate.
Loop
Close #1    ' Fecha o arquivo.

