Attribute VB_Name = "Sikama_ProventosRecebidos"
Public Sub Valida_Rendimento_StatusInvest()

    Dim sTicker As String

    On Error Resume Next
       Sheets("Validando").Select
       If Err.Number <> 0 Then
          Sheets.Add After:=ActiveSheet
          ActiveSheet.Name = "Validando"
       Else
          Sheets("Validando").Cells.Delete Shift:=xlUp
       End If
    On Error GoTo 0:
    
    Sheets(1).Select
    DoEvents

    dCTRL = 0
    dNovo = 4
    
    Sheets("Validando").Columns("B:B").ColumnWidth = 10
    Sheets("Validando").Columns("C:C").ColumnWidth = 10
    Sheets("Validando").Columns("D:D").ColumnWidth = 15
    Sheets("Validando").Columns("E:E").ColumnWidth = 30
    
    Sheets("Validando").Columns("F:F").NumberFormat = "$ #,##0.00"
    Sheets("Validando").Columns("G:G").NumberFormat = "$ #,##0.00"
    Sheets("Validando").Columns("F:F").ColumnWidth = 20
    Sheets("Validando").Columns("G:G").ColumnWidth = 20
    Sheets("Validando").Columns("H:H").ColumnWidth = 15
    Sheets("Validando").Columns("I:I").ColumnWidth = 15
    
    Sheets("Validando").Range("B3").Value = "Ano"
    Sheets("Validando").Range("C3").Value = "Mes"
    Sheets("Validando").Range("D3").Value = "Ticker"
    Sheets("Validando").Range("E3").Value = "Tipo"
    Sheets("Validando").Range("F3").Value = "Rendimento"
    Sheets("Validando").Range("G3").Value = "Valor"
    Sheets("Validando").Range("H3").Value = "Chave_Atual"
    Sheets("Validando").Range("I3").Value = "Chave_Posterior"
    
    Sheets("Validando").Rows("3:3").VerticalAlignment = xlCenter
    
    dLin = 2
    
    For dLin = 1 To 10000
    
        If dCTRL > 10 Then
           Exit For
        ElseIf Len(Trim(Range("C" & dLin).Value)) = 0 Then
           dCTRL = dCTRL + 1
        Else
           dCTRL = 0
        End If
    
        If Len(Trim(Sheets(1).Range("C" & dLin).Value)) > 0 And _
           Trim(Sheets(1).Range("C" & dLin).Value) <> "ATIVO" Then
    
           sTab_Date = Split(Sheets(1).Range("K" & dLin).Value, "/")
           sTab_Ticker = Sheets(1).Range("C" & dLin).Value
        
           dMesAnt = CInt(sTab_Date(1)) + 1
           dAnoAnt = sTab_Date(2)
           If dMesAnt > 12 Then
               sMesAnt = "Janeiro"
               sAnoAnt = CStr(Int(dAnoAnt) + 1)
           Else
               sMesAnt = Ver_Mes(dMesAnt)
               sAnoAnt = sTab_Date(2)
           End If
        
           sMes = Ver_Mes(CInt(sTab_Date(1)))
           sAno = sTab_Date(2)
           sTicker = sTab_Ticker 'Trim(sTab_Ticker(0))
           dNovo = ver_Linha(sTicker, sMes)
            
           Sheets("Validando").Range("B" & dNovo).Value = sAno
           Sheets("Validando").Range("C" & dNovo).Value = sMes
           Sheets("Validando").Range("D" & dNovo).Value = sTicker
           Sheets("Validando").Range("E" & dNovo).Value = Sheets(1).Range("E" & dLin).Value
          
           If sTicker = Trim(sTab_Ticker) Then
               Sheets("Validando").Range("F" & dNovo).Value = Sheets(1).Range("G" & dLin).Value
           End If
           
           Sheets("Validando").Range("G" & dNovo).Value = Sheets("Validando").Range("G" & dNovo).Value + Sheets(1).Range("i" & dLin).Value
           Sheets("Validando").Range("H" & dNovo).Value = Ver_Ordem(sMes, sAno, Trim(sTicker))
           Sheets("Validando").Range("I" & dNovo).Value = Ver_Ordem(sMesAnt, sAnoAnt, Trim(sTicker))
    
        End If
    
    Next

    
End Sub
    
Public Sub Gravando_MySQL()

    Dim cn          As Object
    Dim strSql As String
    Dim strPath As String
    
    Sheets(2).Select
    DoEvents
    
    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    dLin = 4
    
    Do While Len(Trim(Sheets(2).Range("C" & dLin).Value)) > 0
    
       Sheets(2).Range("C" & dLin).Select
    
       If Len(Trim(Sheets("Validando").Range("H" & dLin).Value)) > 0 Then
    
          sSQL = ""
          sSQL = sSQL & " update Tab_Aplicacao "
          sSQL = sSQL & " set    cc_VAT_RN   = " & Str(Sheets("Validando").Range("G" & dLin).Value)
          sSQL = sSQL & "      , cc_VAT_Cota = " & Str(Sheets("Validando").Range("F" & dLin).Value)
          sSQL = sSQL & " where cc_Controle = '" & Sheets("Validando").Range("H" & dLin).Value & "' "
            
          conn.Execute (sSQL), dResult
      
          Sheets("Validando").Range("J" & dLin).Value = dResult
      
       End If
      
       If Len(Trim(Sheets("Validando").Range("I" & dLin).Value)) > 0 Then
    
          sSQL = ""
          sSQL = sSQL & " update Tab_Aplicacao "
          sSQL = sSQL & " set    cc_VAN_RN   = " & Str(Sheets("Validando").Range("G" & dLin).Value)
          sSQL = sSQL & "      , cc_VAN_Cota = " & Str(Sheets("Validando").Range("F" & dLin).Value)
          sSQL = sSQL & " where cc_Controle = '" & Sheets("Validando").Range("I" & dLin).Value & "' "
            
          conn.Execute (sSQL), dResult
      
          Sheets("Validando").Range("K" & dLin).Value = dResult
      
       End If
      
      
       dLin = dLin + 1
    Loop

End Sub

    
    
    

Public Sub Valida_Rendimento_B3()

    Dim sTicker As String

    On Error Resume Next
       Sheets("Validando").Select
       If Err.Number <> 0 Then
          Sheets.Add After:=ActiveSheet
          ActiveSheet.Name = "Validando"
       Else
          Sheets("Validando").Cells.Delete Shift:=xlUp
       End If
    On Error GoTo 0:
    
    Sheets("Proventos Recebidos").Select
    DoEvents

    dCTRL = 0
    dNovo = 4
    
    Sheets("Validando").Columns("B:B").ColumnWidth = 10
    Sheets("Validando").Columns("C:C").ColumnWidth = 10
    Sheets("Validando").Columns("D:D").ColumnWidth = 15
    Sheets("Validando").Columns("E:E").ColumnWidth = 30
    
    Sheets("Validando").Columns("F:F").NumberFormat = "$ #,##0.00"
    Sheets("Validando").Columns("G:G").NumberFormat = "$ #,##0.00"
    Sheets("Validando").Columns("F:F").ColumnWidth = 20
    Sheets("Validando").Columns("G:G").ColumnWidth = 20
    Sheets("Validando").Columns("H:H").ColumnWidth = 15
    Sheets("Validando").Columns("I:I").ColumnWidth = 15
    
    Sheets("Validando").Range("B3").Value = "Ano"
    Sheets("Validando").Range("C3").Value = "Mes"
    Sheets("Validando").Range("D3").Value = "Ticker"
    Sheets("Validando").Range("E3").Value = "Tipo"
    Sheets("Validando").Range("F3").Value = "Rendimento"
    Sheets("Validando").Range("G3").Value = "Valor"
    Sheets("Validando").Range("H3").Value = "Chave_Atual"
    Sheets("Validando").Range("I3").Value = "Chave_Posterior"
    
    Sheets("Validando").Rows("3:3").VerticalAlignment = xlCenter
    
    dLin = 2
    Do While Len(Trim(Sheets("Proventos Recebidos").Range("A" & dLin).Value)) > 0
    
       sTab_Date = Split(Sheets("Proventos Recebidos").Range("B" & dLin).Value, "/")
       sTab_Ticker = Split(Sheets("Proventos Recebidos").Range("A" & dLin).Value, "-")
    
       dMesAnt = CInt(sTab_Date(1)) + 1
       dAnoAnt = sTab_Date(2)
       If dMesAnt > 12 Then
           sMesAnt = "Janeiro"
           sAnoAnt = CStr(Int(dAnoAnt) + 1)
       Else
           sMesAnt = Ver_Mes(dMesAnt)
           sAnoAnt = sTab_Date(2)
       End If
    
       sMes = Ver_Mes(CInt(sTab_Date(1)))
       sAno = sTab_Date(2)
       sTicker = Trim(sTab_Ticker(0))
       dNovo = ver_Linha(sTicker, sMes)
        
       Sheets("Validando").Range("B" & dNovo).Value = sAno
       Sheets("Validando").Range("C" & dNovo).Value = sMes
       Sheets("Validando").Range("D" & dNovo).Value = sTicker
       Sheets("Validando").Range("E" & dNovo).Value = Sheets("Proventos Recebidos").Range("C" & dLin).Value
      
       If sTicker = Trim(sTab_Ticker(0)) Then
           Sheets("Validando").Range("F" & dNovo).Value = Sheets("Proventos Recebidos").Range("F" & dLin).Value
       End If
       
       Sheets("Validando").Range("G" & dNovo).Value = Sheets("Validando").Range("G" & dNovo).Value + Sheets("Proventos Recebidos").Range("G" & dLin).Value
       Sheets("Validando").Range("H" & dNovo).Value = Ver_Ordem(sMes, sAno, Trim(sTicker))
       Sheets("Validando").Range("I" & dNovo).Value = Ver_Ordem(sMesAnt, sAnoAnt, Trim(sTicker))
       dNovo = dNovo + 1
    
       dLin = dLin + 1
    Loop
    
End Sub
    
Private Function Update_MySQL(ByVal dLin As Double) As Boolean

    Dim cn          As Object
    
    Dim strSql As String
    Dim strPath As String

    Dim sControle As String: sControle = ""


    Dim dtaWrk As Date

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " update cc_Controle "
    sSQL = sSQL & " from Tab_Aplicacao "
    sSQL = sSQL & " where cc_Controle = '" & sMes & "' "
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_Ordem = rs!cc_Controle
    End If
    rs.Close

End Function
    
Private Function Ver_Mes(ByVal iMes As Integer)
   
    Dim Tab_Mes(1 To 12) As String

    Tab_Mes(1) = "Janeiro"
    Tab_Mes(2) = "Fevereiro"
    Tab_Mes(3) = "Março"
    Tab_Mes(4) = "Abril"
    Tab_Mes(5) = "Maio"
    Tab_Mes(6) = "Junho"
    Tab_Mes(7) = "Julho"
    Tab_Mes(8) = "Agosto"
    Tab_Mes(9) = "Setembro"
    Tab_Mes(10) = "Outubro"
    Tab_Mes(11) = "Novembro"
    Tab_Mes(12) = "Dezembro"

    Ver_Mes = Tab_Mes(iMes)

End Function

Private Function ver_Linha(ByRef sTicker As String, ByVal sMes As String) As Double

    dFind = 4
    
    If Right(sTicker, 2) = "12" Or Right(sTicker, 2) = "13" Then
       sTicker = Replace(sTicker, "12", "11")
       sTicker = Replace(sTicker, "13", "11")
    End If
    

    Do While Len(Trim(Sheets("Validando").Range("D" & dFind).Value)) > 0
    
       If InStr(1, Sheets("Validando").Range("D" & dFind).Value, sTicker) And _
          Sheets("Validando").Range("C" & dFind).Value = sMes Then
          Exit Do
       End If
       dFind = dFind + 1
    Loop
    
    ver_Linha = dFind

End Function

    
    
Private Function Ver_Ordem(ByVal sMes As String, ByVal sAno As String, ByVal sTiker As String) As String

    Dim cn          As Object
    Dim rs          As Object
    Dim strSql As String
    Dim strPath As String

    Dim dtaWrk As Date

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " Select cc_Controle "
    sSQL = sSQL & " from Tab_Aplicacao "
    sSQL = sSQL & " where cc_Mes = '" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Descricao like '%" + sTiker + "%' "
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_Ordem = rs!cc_Controle
    End If
    rs.Close

End Function
