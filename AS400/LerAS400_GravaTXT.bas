Attribute VB_Name = "sikama"

Public Sub Verifica_VRVxMAPS()

On Error GoTo msgerro:

   Dim cn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   
   Dim str_sql As String

    Sheets("VRVxMAPS").Select
    Sheets("VRVxMAPS").Rows("1:65000").Delete Shift:=xlUp

    dbl_header = 3
    dbl_linha = 3
    UserForm1.Show 1
    
    If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
       Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
    
        cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
        cn.Open

        Dim strPath As String

        str_sql = ""
        str_sql = str_sql & "Select "
        str_sql = str_sql & "     * "
        str_sql = str_sql & "from   Tabela "

        rs.Open str_sql, cn, adOpenForwardOnly
        Do While Not rs.EOF



          rs.MoveNext
        Loop


    End If

   MsgBox "Terminou"
   Exit Sub

msgerro: MsgBox Err.Number & " - " & Err.Description, vbCritical
 
End Sub


Private Sub Ler_Grava_arquivo_texto()

Dim MyRecord As Record, RecordNumber
strPath = ActiveWorkbook.Path & "\NF_1511.TXT"

Dim TextLine
Dim TextLineB

TextLineB = ""

Open strPath For Input As #1
Do While Not EOF(1)
    Line Input #1, TextLine
    
    If TextLine <> TextLineB Then
    
    
    End If
    
    TextLine = TextLineB
    
Loop
Close #1



End Sub

Public Sub ler_Text_De_Um_Diretorio()

    Dim dbl_dta As Date

    MyPath = ActiveWorkbook.Path & "\*.txt"
    MyName = Dir(MyPath, vbArchive)
    dNew = 3
    
    Rows("4:65000").Delete Shift:=xlUp
    
    Do While MyName <> ""
        
        Dim TextLine
        
        sDir = ActiveWorkbook.Path & "\" & MyName
        dlin = 0
        
        Open sDir For Input As #1
        Do While Not EOF(1)
                
            dlin = dlin + 1
            
            Line Input #1, TextLine
        
        
        Loop
        Close #1    ' Fecha o arquivo.
        
        MyName = Dir
    Loop
        
    Range("B4").Select

End Sub


