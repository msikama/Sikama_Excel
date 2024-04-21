Attribute VB_Name = "Sikama_OpenTXT_ADODB"

Private Sub OpenTXT_ADODB()

    str_file = "G:\#001# Pfizer\Marisa\service-requests-filtered (3).csv"
 
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Type = 2
    objStream.Open
    objStream.LoadFromFile = str_file
'   objStream.LineSeparator = 10
    
    
    While Not objStream.EOS
    
        TextLine = objStream.ReadText(-2)
        DoEvents

    Wend

End Sub
