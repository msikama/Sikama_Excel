Attribute VB_Name = "Sikama_MySQL_Cartao"
Private sMes As String
Private sAno As String
Private sCar As String
Private sTAB_Ordens As String

Public Sub Valida_Cartao()

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

    dCTRL = 0
    dNovo = 4
    
    Sheets("Validando").Columns("B:B").ColumnWidth = 10
    Sheets("Validando").Columns("C:C").ColumnWidth = 10
    Sheets("Validando").Columns("D:D").ColumnWidth = 50
    Sheets("Validando").Columns("E:E").ColumnWidth = 10
    
    Sheets("Validando").Columns("F:F").NumberFormat = "$ #,##0.00"
    Sheets("Validando").Columns("F:F").ColumnWidth = 20
    Sheets("Validando").Columns("G:G").ColumnWidth = 20
    
    Sheets("Validando").Range("B3").Value = "Final"
    Sheets("Validando").Range("C3").Value = "Data"
    Sheets("Validando").Range("D3").Value = "lançamento"
    Sheets("Validando").Range("E3").Value = "Origem"
    Sheets("Validando").Range("F3").Value = "Valor"
    Sheets("Validando").Range("G3").Value = "Registro"
    
    Sheets("Validando").Rows("3:3").VerticalAlignment = xlCenter
    Sheets("Validando").Rows("3:3").VerticalAlignment = xlCenter
    
    For dLin = 1 To 300
    
        If dCTRL > 10 Then
           Exit For
        ElseIf Len(Trim(Range("A" & dLin).Value)) = 0 Then
           dCTRL = dCTRL + 1
        Else
           dCTRL = 0
        End If
    
        If Len(Trim(Range("A" & dLin).Value)) <> 0 Then
    
           If InStr(Range("A" & dLin).Value, "total nacional") > 0 Then
               Sheets("Validando").Range("B" & dNovo).Value = sCar
               Sheets("Validando").Range("D" & dNovo).Value = "Total nacional do cartão"
               Sheets("Validando").Range("F" & dNovo).Value = CDbl(Range("D" & dLin).Value)
               dNovo = dNovo + 1
           ElseIf InStr(Range("A" & dLin).Value, "- final") > 0 Then
    
              dPos = InStr(UCase(Range("A" & dLin).Value), "FINAL")
              sCar = Mid(Range("A" & dLin).Value, dPos + 6, 4)
    
           ElseIf IsDate(Range("A" & dLin).Value) = True And _
                  InStr(Range("B" & dLin).Value, "PAGAMENTO EFETUADO") = 0 Then
    
                  dPos = InStr(UCase(Range("B" & dLin).Value), "/")
                  If dPos > 0 Then
                     sPres = Trim(Mid(Range("B" & dLin).Value, dPos - 2, 6))
                     sHist = Trim(Replace(Range("B" & dLin).Value, sPres, ""))
                     sHist = sHist & " [" & sPres & "]"
                     sOrigem = "Parcelado"
                  Else
                     sHist = Range("B" & dLin).Value
                     sOrigem = "A_Vista"
                  End If
    
                  If CDbl(Range("D" & dLin).Value) < 0 Then
                     sOrigem = "CashBack"
                  End If
    
                  Sheets("Validando").Range("B" & dNovo).Value = sCar
                  Sheets("Validando").Range("C" & dNovo).Value = CDate(Range("A" & dLin).Value)
                  Sheets("Validando").Range("D" & dNovo).Value = sHist
                  Sheets("Validando").Range("E" & dNovo).Value = sOrigem
                  Sheets("Validando").Range("F" & dNovo).Value = CDbl(Range("D" & dLin).Value)
                  Sheets("Validando").Range("G" & dNovo).Value = ""
                  dNovo = dNovo + 1
            
            End If
    
        End If
        
    Next

    Sheets("Validando").Select
    DoEvents

    sTab_Date = Split(Sheets("Lançamentos").Range("B2").Value, Space(1))

    Dim dtaWrk As Date

    sMes = UCase(Left(Format(sTab_Date(0), "MMMM"), 1)) & Mid(Format(sTab_Date(0), "MMMM"), 2, 30)
    sAno = Year(sTab_Date(0))

    dLin = 4
    
    Do While Len(Trim(Sheets("Validando").Range("B" & dLin).Value)) > 0
   
        sData = Range("C" & dLin).Value
        sLanc = Range("D" & dLin).Value
        sValor = Range("F" & dLin).Value
        sCartao = Right("0000" & Sheets("Validando").Range("B" & dLin).Value, 4)
         
        Range("G" & dLin).Value = Ver_Movimento(sData, sValor, sCartao)
        If Len(Trim(Range("G" & dLin).Value)) > 0 Then
           Range("G" & dLin).Interior.Color = RGB(0, 0, 51)
           Range("G" & dLin).Font.Color = RGB(255, 255, 255)
           Columns("G:G").EntireColumn.AutoFit
           
           If sTAB_Ordens <> "" Then
              sTAB_Ordens = sTAB_Ordens & ","
           End If
           
           sTAB_Ordens = sTAB_Ordens & "'" & Range("F" & dLin).Value & "'"
           
        End If
                  
       dLin = dLin + 1
    Loop
    
   
End Sub


Private Function Ver_Movimento(ByVal sData As String, ByVal dValor As Double, ByVal sCartao As String) As String

    Dim cn          As Object
    Dim rs          As Object
    Dim strSql As String
    Dim strPath As String

    strPath = "F:\02 - Banco\ControleComSQLite\@MS_Controle_Conta_Corrente-DB\MSikama_VS.db"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & strPath & ";"
    
    sSQL = ""
    sSQL = sSQL & " Select * "
    sSQL = sSQL & " from Tab_Cartao "
    sSQL = sSQL & " where cc_Mes ='" & sMes & "' "
    sSQL = sSQL & " and   cc_Ano = " & sAno & " "
    sSQL = sSQL & " and   cc_Cartao = '" & sCartao & "' "
    sSQL = sSQL & " and   cc_Valor = " & Str(dValor * -1) & " "
    
    If Len(Trim(sTAB_Ordens)) > 0 Then
       sSQL = sSQL & " and cc_Ordem not in(" & sTAB_Ordens & ")"
    End If
    
    Set rs = conn.Execute(sSQL)

    If Not rs.EOF Then
       Ver_Movimento = rs!cc_Ordem
    End If
    
    
    

End Function

