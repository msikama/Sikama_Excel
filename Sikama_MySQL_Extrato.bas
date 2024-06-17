Attribute VB_Name = "Sikama_MySQL_Extrato"
Private sMes As String
Private sAno As String
Private dNew As Double
Private bFLG As Boolean
Private dULI As Double
Private sTAB_Ordens_Usada As String

Public Sub Valida_Extrato()
     
    bFLG = False
    dNew = 5

    sTAB_Ordens_Usada = ""

    On Error Resume Next
       Sheets("Validando").Select
       If Err.Number <> 0 Then
          Sheets.Add After:=ActiveSheet
          ActiveSheet.Name = "Validando"
       Else
          Sheets("Validando").Cells.Delete Shift:=xlUp
       End If
    On Error GoTo 0:
    
    Sheets("Lançamentos").Select
    DoEvents

    sTab_Date = Split(Sheets("Lançamentos").Range("A12").Value, Space(1))

    Dim dtaWrk As Date

    sMes = UCase(Left(Format(sTab_Date(0), "MMMM"), 1)) & Mid(Format(sTab_Date(0), "MMMM"), 2, 30)
    sAno = Year(sTab_Date(0))
    
    Sheets("Validando").Range("D2").Value = sMes
    Sheets("Validando").Range("E2").Value = sAno
    Sheets("Validando").Range("B4").Value = "data"
    Sheets("Validando").Range("C4").Value = "lançamento"
    Sheets("Validando").Range("D4").Value = "ag./origem"
    Sheets("Validando").Range("E4").Value = "valor (R$)"
    Sheets("Validando").Range("G4").Value = "Validando"
    
    Sheets("Validando").Range("B4").HorizontalAlignment = xlCenter
    Sheets("Validando").Range("C4").HorizontalAlignment = xlLeft
    Sheets("Validando").Range("D4").HorizontalAlignment = xlCenter
    Sheets("Validando").Range("E4").HorizontalAlignment = xlRight
    Sheets("Validando").Range("G4").HorizontalAlignment = xlCenter

    Sheets("Validando").Range("B4:E4").Interior.Color = RGB(0, 51, 0)
    Sheets("Validando").Range("B4:E4").Font.Color = RGB(255, 255, 255)
    Sheets("Validando").Range("G4:G4").Interior.Color = RGB(0, 0, 51)
    Sheets("Validando").Range("G4:G4").Font.Color = RGB(255, 255, 255)

    Sheets("Validando").Range("B4:E4").Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Validando").Range("B4:E4").Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("Validando").Range("B4:E4").Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Validando").Range("B4:E4").Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeLeft).Weight = xlThin
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeTop).Weight = xlThin
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeBottom).Weight = xlThin
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Validando").Range("B4:E4").Borders(xlEdgeRight).Weight = xlThin
    Sheets("Validando").Range("B4:E4").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Validando").Range("B4:E4").Borders(xlInsideVertical).Weight = xlThin

    Sheets("Validando").Columns("B:B").ColumnWidth = 20
    Sheets("Validando").Columns("C:C").ColumnWidth = 50
    Sheets("Validando").Columns("D:D").ColumnWidth = 15
    Sheets("Validando").Columns("E:E").ColumnWidth = 20
    Sheets("Validando").Columns("F:F").ColumnWidth = 2
    Sheets("Validando").Columns("G:G").ColumnWidth = 20
    
    dCTRL = 0

    For dIDX = 1 To 100000

       If dCTRL > 10 Or InStr(Sheets("Lançamentos").Range("A" & dIDX).Value, "lançamentos futuros") > 0 Then
           Exit For
       ElseIf Len(Trim(Sheets("Lançamentos").Range("A" & dIDX).Value)) = 0 Then
           dCTRL = dCTRL + 1
       Else
           dCTRL = 0
           Range("A" & dIDX).Select
           DoEvents
       End If
       
       If IsDate(Sheets("Lançamentos").Range("A" & dIDX).Value) = True Then
    
           If InStr(Sheets("Lançamentos").Range("A" & dIDX).Value, "/") > 0 And _
              InStr(UCase(Sheets("Lançamentos").Range("B" & dIDX).Value), "SALDO") > 0 And _
              bFLG = False Then
    
                bFLG = True
    
                Sheets("Validando").Range("B" & dNew).HorizontalAlignment = xlCenter
                Sheets("Validando").Range("C" & dNew).HorizontalAlignment = xlLeft
                Sheets("Validando").Range("D" & dNew).HorizontalAlignment = xlCenter
                Sheets("Validando").Range("E" & dNew).HorizontalAlignment = xlRight
                
                sDia_A = Day(Sheets("Lançamentos").Range("A" & dIDX).Value)
                sMes_A = Month(Sheets("Lançamentos").Range("A" & dIDX).Value)
                sAno_A = Year(Sheets("Lançamentos").Range("A" & dIDX).Value)
                
                Sheets("Validando").Range("B" & dNew).NumberFormat = "@"
                Sheets("Validando").Range("B" & dNew).Value = Right("00" & sDia_A, 2) & "/" & Right("00" & sMes_A, 2) & "/" & sAno_A
                Sheets("Validando").Range("C" & dNew).Value = Sheets("Lançamentos").Range("B" & dIDX).Value
                Sheets("Validando").Range("D" & dNew).Value = Sheets("Lançamentos").Range("C" & dIDX).Value
                
                If InStr(1, Sheets("Lançamentos").Range("E" & dIDX - 1).Value, "-") > 0 Then
                   Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = True
                   Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(51, 0, 0)
                   sValor = Replace(Sheets("Lançamentos").Range("E" & dIDX).Value, "-", "")
                   dValor = CDbl(sValor) * -1
                Else
                   Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = True
                   Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(0, 0, 51)
                   dValor = CDbl(Sheets("Lançamentos").Range("E" & dIDX).Value)
                End If
    
                Sheets("Validando").Range("E" & dNew).NumberFormat = "#,##0.00"
                Sheets("Validando").Range("E" & dNew).Value = dValor
                dNew = dNew + 1
    
           ElseIf InStr(Sheets("Lançamentos").Range("A" & dIDX).Value, "/") > 0 And _
                 InStr(UCase(Sheets("Lançamentos").Range("B" & dIDX).Value), "SALDO") > 0 Then
                 
                 dULI = dIDX
    
           Else
               
               If InStr(Sheets("Lançamentos").Range("A" & dIDX).Value, "/") > 0 And _
                  InStr(UCase(Sheets("Lançamentos").Range("B" & dIDX).Value), "SALDO") = 0 Then
        
                    Sheets("Validando").Range("B" & dNew).HorizontalAlignment = xlCenter
                    Sheets("Validando").Range("C" & dNew).HorizontalAlignment = xlLeft
                    Sheets("Validando").Range("D" & dNew).HorizontalAlignment = xlCenter
                    Sheets("Validando").Range("E" & dNew).HorizontalAlignment = xlRight
                    
                    sDia_A = Day(Sheets("Lançamentos").Range("A" & dIDX).Value)
                    sMes_A = Month(Sheets("Lançamentos").Range("A" & dIDX).Value)
                    sAno_A = Year(Sheets("Lançamentos").Range("A" & dIDX).Value)
                    
                    Sheets("Validando").Range("B" & dNew).NumberFormat = "@"
                    Sheets("Validando").Range("B" & dNew).Value = Right("00" & sDia_A, 2) & "/" & Right("00" & sMes_A, 2) & "/" & sAno_A
                    Sheets("Validando").Range("C" & dNew).Value = Sheets("Lançamentos").Range("B" & dIDX).Value
                    Sheets("Validando").Range("D" & dNew).Value = Sheets("Lançamentos").Range("C" & dIDX).Value
        
                    '*** DIVIDENDOS
        
                    If InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "OPERACOES") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "DIVIDENDOS") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "JSCP") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "ACOES") > 0 Then
                       
                       Sheets("Validando").Range("D" & dNew).Value = "Dividendos"
                   
                    '*** A VISTA
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "RSHOP") > 0 Or _
                           InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "RSCCS") > 0 Or _
                           InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "RSCSS") > 0 Then
                       
                       Sheets("Validando").Range("D" & dNew).Value = "A_Vista"
                    
                    '*** PROVENTOS
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "RENDIMENTO") > 0 Then
                       Sheets("Validando").Range("D" & dNew).Value = "Proventos-FIIS"
                    
                    '*** POUPANÇA
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "POUP AUT") > 0 Then
                       Sheets("Validando").Range("D" & dNew).Value = "Itaú-Juros"
                    
                    
                    '*** MENSAL
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "INT PAG TIT") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "ELETROPAULO") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "VIVO-SP") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "VIVO-SP") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "PREMIO VGBL") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "SEGURO CARTAO") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "PERS BLACK") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "PERS INFINIT") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "ITAU BLACK") > 0 Or _
                       InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "MOBILEPAG") > 0 Then
                    
                       Sheets("Validando").Range("D" & dNew).Value = "Mensal"
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "REMUNERACAO/SALARIO") > 0 Then
                       Sheets("Validando").Range("D" & dNew).Value = "Luandre"
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "COR  SUBSC") > 0 Then
                       Sheets("Validando").Range("D" & dNew).Value = "PicPay-Inv"
                    End If
        
                    
                    If InStr(1, Sheets("Lançamentos").Range("D" & dIDX).Value, "-") > 0 Then
                       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = False
                       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(51, 0, 0)
                       sValor = Replace(Sheets("Lançamentos").Range("D" & dIDX).Value, "-", "")
                       dValor = CDbl(sValor) * -1
                    Else
                       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = False
                       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(0, 0, 51)
                       dValor = CDbl(Sheets("Lançamentos").Range("D" & dIDX).Value)
                    End If
    
                    Sheets("Validando").Range("E" & dNew).NumberFormat = "#,##0.00"
                    Sheets("Validando").Range("E" & dNew).Value = dValor
                    
                    If InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "PIX TRANSF  MARCIO") > 0 Then
                       
                       If dValor >= 0 Then
                          Sheets("Validando").Range("D" & dNew).Value = "PIX-Pagamento"
                       Else
                          Sheets("Validando").Range("D" & dNew).Value = "PIX-PicPay"
                       End If
                    
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "PIX") > 0 Then
                    
                       If dValor >= 0 Then
                          Sheets("Validando").Range("D" & dNew).Value = "PIX-Pagamento"
                       Else
                          Sheets("Validando").Range("D" & dNew).Value = "PIX-Depósito"
                       End If
                       
                    ElseIf InStr(Sheets("Lançamentos").Range("B" & dIDX).Value, "TED") > 0 Then
                    
                       If dValor >= 0 Then
                          Sheets("Validando").Range("D" & dNew).Value = "Transferencia"
                       Else
                          Sheets("Validando").Range("D" & dNew).Value = "Depósito"
                       End If
                       
                    End If
                    
                    dNew = dNew + 1
        
               End If
           
           End If
    
       End If
    
    Next


    Sheets("Validando").Range("B" & dNew).HorizontalAlignment = xlCenter
    Sheets("Validando").Range("C" & dNew).HorizontalAlignment = xlLeft
    Sheets("Validando").Range("D" & dNew).HorizontalAlignment = xlCenter
    Sheets("Validando").Range("E" & dNew).HorizontalAlignment = xlRight
    Sheets("Validando").Range("B" & dNew).Value = Sheets("Lançamentos").Range("A" & dULI).Value
    Sheets("Validando").Range("C" & dNew).Value = "SALDO FINAL"
    Sheets("Validando").Range("D" & dNew).Value = Sheets("Lançamentos").Range("C" & dULI).Value
    
    If InStr(1, Sheets("Lançamentos").Range("E" & dIDX - 1).Value, "-") > 0 Then
       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = True
       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(51, 0, 0)
       sValor = Replace(Sheets("Lançamentos").Range("E" & dIDX - 1).Value, "-", "")
       dValor = CDbl(sValor) * -1
    Else
       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Bold = True
       Sheets("Validando").Range("B" & dNew & ":E" & dNew).Font.Color = RGB(0, 0, 51)
       
       If Len(Trim(Sheets("Lançamentos").Range("E" & dULI).Value)) > 0 Then
          dValor = CDbl(Sheets("Lançamentos").Range("E" & dULI).Value)
       End If
    End If

    Sheets("Validando").Range("E" & dNew).NumberFormat = "#,##0.00"
    Sheets("Validando").Range("E" & dNew).Value = dValor
            
Valida_Extrato:

    Sheets("Validando").Select
    DoEvents

   ' Call Sort_extrato(dNew)

    Dim sEfetivo As String

    Call Ver_Movimento_RE

    bEsp = False

    dLin = 5
    Do While Len(Trim(Sheets("Validando").Range("B" & dLin).Value)) > 0

       If InStr(UCase(Sheets("Validando").Range("C" & dLin).Value), "SALDO") = 0 Then
       
           sData = Sheets("Validando").Range("B" & dLin).Value
           sLanc = Sheets("Validando").Range("C" & dLin).Value
           sValor = Sheets("Validando").Range("E" & dLin).Value
    
           If InStr(sLanc, "COR OPERACOES B3") > 0 Or InStr(sLanc, "ITAUCOR") > 0 Then
              Sheets("Validando").Range("G" & dLin).Value = Ver_OperacaoB3(sLanc, sData, sValor)
              Sheets("Validando").Range("F" & dLin).Value = "S"
              Sheets("Validando").Range("F" & dLin).Interior.Color = RGB(0, 51, 0)
              Sheets("Validando").Range("F" & dLin).Font.Color = RGB(255, 255, 255)
           ElseIf InStr(sLanc, "PIX TRANSF  MARCIO") > 0 Then
           
              If sValor = ActiveSheet.Range("H2").Value Then
                 Sheets("Validando").Range("G" & dLin).Value = Sheets("Validando").Range("G9").Value
                 Sheets("Validando").Range("F" & dLin).Value = "S"
                 Sheets("Validando").Range("F" & dLin).Interior.Color = RGB(0, 51, 0)
                 Sheets("Validando").Range("F" & dLin).Font.Color = RGB(255, 255, 255)
              Else
                 Sheets("Validando").Range("G" & dLin).Value = Ver_Movimento(sData, sLanc, sValor, sEfetivo)
                 Sheets("Validando").Range("F" & dLin).Value = sEfetivo
                 If sEfetivo = "S" Then
                    Sheets("Validando").Range("F" & dLin).Interior.Color = RGB(0, 51, 0)
                    Sheets("Validando").Range("F" & dLin).Font.Color = RGB(255, 255, 255)
                 End If
              End If
              
           ElseIf sValor = Sheets("Validando").Range("H2").Value And Sheets("Validando").Range("D" & dLin).Value = "PIX-Pagamento" Then
              Sheets("Validando").Range("G" & dLin).Value = Sheets("Validando").Range("G2").Value
              
              If Day(sData) >= 20 Then
                 Sheets("Validando").Range("F" & dLin).Value = "S"
                 Sheets("Validando").Range("F" & dLin).Interior.Color = RGB(0, 51, 0)
                 Sheets("Validando").Range("F" & dLin).Font.Color = RGB(255, 255, 255)
              End If
                 
           Else
           
             If (sValor = -125 Or sValor = -29) And _
                (InStr(Sheets("Validando").Range("C" & dLin).Value, "43599") > 0 Or _
                 InStr(Sheets("Validando").Range("C" & dLin).Value, "45599") > 0) Then
                sValor = -154
                bEsp = True
             End If
           
              Sheets("Validando").Range("G" & dLin).Value = Ver_Movimento(sData, sLanc, sValor, sEfetivo)
              Sheets("Validando").Range("F" & dLin).Value = sEfetivo
              If sEfetivo = "S" Then
                 Sheets("Validando").Range("F" & dLin).Interior.Color = RGB(0, 51, 0)
                 Sheets("Validando").Range("F" & dLin).Font.Color = RGB(255, 255, 255)
              End If
              
              If bEsp = True Then
                 Sheets("Validando").Range("G" & dLin).Value = Sheets("Validando").Range("G" & dLin).Value & "-E"
                 bEsp = False
              End If
              
           End If

           If Len(Trim(Sheets("Validando").Range("G" & dLin).Value)) > 0 Then
           
              Sheets("Validando").Range("G" & dLin).Interior.Color = RGB(0, 0, 51)
              Sheets("Validando").Range("G" & dLin).Font.Color = RGB(255, 255, 255)
              Sheets("Validando").Columns("G:G").EntireColumn.AutoFit
              
              If sTAB_Ordens_Usada <> "" Then
                 sTAB_Ordens_Usada = sTAB_Ordens_Usada & ","
              End If
                
              sTAB_Ordens_Usada = sTAB_Ordens_Usada & "'" & Range("G" & dLin).Value & "'"
              
           End If

           Sheets("Validando").Range("G" & dLin).HorizontalAlignment = xlCenter

       End If

       dLin = dLin + 1
    Loop
    
End Sub

Private Function Ver_Movimento(ByVal sData As String, ByVal sLancamento As String, ByVal dValor As Double, ByRef sEfetivo As String) As String

    Dim cn          As Object
    Dim rs          As Object
    Dim strSql As String
    Dim strPath As String

    sEfetivo = ""

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
   
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " Select * "
    sSQL = sSQL & " from Tab_Movimento "
    sSQL = sSQL & " where cc_Mes ='" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Lancamento = '" & sLancamento & "' "
    sSQL = sSQL & " and   cc_Valor = " & Str(dValor) & " "
    sSQL = sSQL & " and   cc_Ordem not in(" & sTAB_Ordens_Usada & ")"
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_Movimento = rs!cc_Ordem
       sEfetivo = rs!cc_Efetivo
       Exit Function
    End If
    
    rs.Close
    
    sSQL = ""
    sSQL = sSQL & " Select * "
    sSQL = sSQL & " from Tab_Movimento "
    sSQL = sSQL & " where cc_Mes ='" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Valor = " & Str(dValor) & " "
    sSQL = sSQL & " and   cc_Ordem not in(" & sTAB_Ordens_Usada & ")"
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_Movimento = rs!cc_Ordem
       sEfetivo = rs!cc_Efetivo
    End If

End Function

Private Function Ver_OperacaoB3(ByVal sLancamento As String, ByVal sData As String, ByVal dValor As Double) As String

    Dim cn          As Object
    Dim rs          As Object
    Dim strSql      As String
    Dim strPath     As String

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " Select ifnull(sum(cc_Valor), 0) cc_Valor, MIN(cc_Ordem) cc_Ordem "
    sSQL = sSQL & " from Tab_Movimento "
    sSQL = sSQL & " where cc_Mes = '" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Modulo = 'DI' "
    sSQL = sSQL & " and   cc_Efetivo = 'S' "
    sSQL = sSQL & " Group by cc_Data "
    
    Set rs = conn.Execute(sSQL)

    Do While Not rs.EOF
    
       If dValor = CDbl(rs!cc_Valor) Then
          Ver_OperacaoB3 = rs!cc_Ordem
          Exit Function
       End If
       
       rs.MoveNext
    Loop
    rs.Close

    sSQL = ""
    sSQL = sSQL & " Select cc_Ordem "
    sSQL = sSQL & " from Tab_Movimento "
    sSQL = sSQL & " where cc_Mes = '" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Lancamento = '" & sLancamento & "' "
    sSQL = sSQL & " and   cc_Efetivo = 'S' "
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_OperacaoB3 = rs!cc_Ordem
    End If
    rs.Close


End Function


Private Sub Ver_Movimento_RE()

    Dim cn          As Object
    Dim rs          As Object
    Dim strSql As String
    Dim strPath As String

    Dim dtaWrk As Date

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " Select ifnull(sum(cc_Valor), 0) cc_Valor, MIN(cc_Ordem) cc_Ordem "
    sSQL = sSQL & " from Tab_Movimento "
    sSQL = sSQL & " where cc_Mes = '" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Modulo = 'RE' "
    sSQL = sSQL & " and   cc_Efetivo <> 'T' "
    sSQL = sSQL & " and   cc_Efetivo <> 'N' "
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Sheets("Validando").Range("H2").Value = rs!cc_Valor
       Sheets("Validando").Range("G2").Value = rs!cc_Ordem
    End If
    rs.Close

End Sub

Private Sub Sort_extrato(ByVal dLastLine As Double)

    Sheets("Validando").Select
    Range("B5").Select
    ActiveWorkbook.Worksheets("Validando").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Validando").Sort.SortFields.Add2 Key:=Range( _
        "B5:B243"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("Validando").Sort.SortFields.Add2 Key:=Range( _
        "C5:C" & dLastLine), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Validando").Sort
        .SetRange Range("B4:E" & dLastLine)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

