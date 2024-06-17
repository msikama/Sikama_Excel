Attribute VB_Name = "Sikama_NFe"


Private Sub teste()

   Dim dbl_Linha        As Double
   Dim dbl_LastLine     As Double
   
   Dim dNota(1 To 3)    As Double
   Dim dSerie(1 To 3)   As Double
   Dim dMotivo(1 To 3)  As Double

   Dim dbl_D10          As Double
   Dim dbl_D20          As Double
   Dim dbl_D60          As Double

   Sheets("Div10").Rows("3:65000").Delete Shift:=xlUp
   Sheets("Div20").Rows("3:65000").Delete Shift:=xlUp
   Sheets("Div60").Rows("3:65000").Delete Shift:=xlUp

   dbl_LastLine = Selection.End(xlDown).Row

   Range("B2:L" & dbl_LastLine).Sort Key1:=Range("B3"), Order1:=xlAscending, _
                                     DataOption1:=xlSortTextAsNumbers, _
                                     Key2:=Range("D3"), Order2:=xlAscending, _
                                     DataOption2:=xlSortTextAsNumbers, _
                                     Key3:=Range("E3"), Order3:=xlAscending, _
                                     DataOption3:=xlSortTextAsNumbers, _
                                     Header:=xlGuess, _
                                     OrderCustom:=1, _
                                     MatchCase:=False, _
                                     Orientation:=xlTopToBottom

   dbl_D10 = 3
   dbl_D20 = 3
   dbl_D60 = 3

   xx = 3
   Do While Len(Trim(Sheets("GERAL").Range("C" & xx).Value)) > 0
    
       Sheets("GERAL").Range("B" & xx & ":L" & xx).Interior.ColorIndex = 0
       Sheets("GERAL").Range("B" & xx & ":L" & xx).Font.ColorIndex = 1
   
      ' LOGÍSTICAS
   
        yy = 13
        Do While Len(Trim(Sheets("Inconsistencias").Range("A" & yy).Value)) > 0
          
           If InStr(1, Sheets("Inconsistencias").Range("A" & yy).Value, "OUTRAS OBSERVAÇÕES") > 0 Then
              Exit Do
           End If
          
           str_dist = Sheets("Inconsistencias").Range("A" & yy).Value
          
           dNota(1) = CDbl(Sheets("Inconsistencias").Range("B" & yy).Value)
           dSerie(1) = IIf(str_dist = "AGV", 28, IIf(str_dist = "DHL", 23, ""))
           dMotivo(1) = CDbl(Sheets("Inconsistencias").Range("C" & yy).Value)
          
           dNota(2) = CDbl(Sheets("Inconsistencias").Range("E" & yy).Value)
           dSerie(2) = IIf(str_dist = "AGV", 28, IIf(str_dist = "DHL", 23, ""))
           dMotivo(2) = CDbl(Sheets("Inconsistencias").Range("F" & yy).Value)
          
           dNota(3) = CDbl(Sheets("Inconsistencias").Range("H" & yy).Value)
           dSerie(3) = IIf(str_dist = "AGV", 28, IIf(str_dist = "DHL", 23, ""))
           dMotivo(3) = CDbl(Sheets("Inconsistencias").Range("I" & yy).Value)
          
           If CDbl(Sheets("GERAL").Range("E" & xx).Value) = dNota(1) And _
              CDbl(Sheets("GERAL").Range("D" & xx).Value) = dSerie(1) Then
              Call criticas(dMotivo(1), xx, "Logistica")
              Exit Do
                    
           ElseIf CDbl(Sheets("GERAL").Range("E" & xx).Value) = dNota(2) And _
                  CDbl(Sheets("GERAL").Range("D" & xx).Value) = dSerie(2) Then
              Call criticas(dMotivo(2), xx, "Validador Pfizer")
              Exit Do
           
           ElseIf CDbl(Sheets("GERAL").Range("E" & xx).Value = dNota(3)) And _
                  CDbl(Sheets("GERAL").Range("D" & xx).Value = dSerie(3)) Then
              Call criticas(dMotivo(3), xx, "SEFAZ")
              Exit Do
           Else
              
                If CDbl(Sheets("GERAL").Range("B" & xx).Value) = 10 Then
                   If CDbl(Sheets("GERAL").Range("D" & xx).Value) <> 21 Then
                      Sheets("GERAL").Range("K" & xx).Value = "Enviado"
                      Sheets("GERAL").Range("L" & xx).Value = "Correto"
                   Else
                      Sheets("GERAL").Range("K" & xx).Value = ""
                      Sheets("GERAL").Range("L" & xx).Value = ""
                   End If
                Else
                      Sheets("GERAL").Range("K" & xx).Value = "Enviado"
                      Sheets("GERAL").Range("L" & xx).Value = "Correto"
                End If
              
              
           End If
          
      If yy = 18 Then
        QW = ""
      End If
          
       yy = yy + 1
     Loop
       
       
     Range("A" & xx).Select
     DoEvents
     xx = xx + 1
   Loop

End Sub


Private Function criticas(ByVal dMSG As Double, ByVal xx As Double, ByVal strK As String) As String


    Select Case strK
    
        Case "Logistica"
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Interior.ColorIndex = 36
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.ColorIndex = 1
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.Bold = True
    
        Case "Validador Pfizer"
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Interior.ColorIndex = 34
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.ColorIndex = 5
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.Bold = True
    
        Case "SEFAZ"
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Interior.ColorIndex = 35
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.ColorIndex = 10
             Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.Bold = True
    
    End Select

    Select Case dMSG
     
        Case 1   '- Inconsistência\Estoque Bloqueado
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "001 - Inconsistência\Estoque Bloqueado"
             Exit Function
        Case 2   '- Código Emitente x Municipio
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "002 - Código Emitente x Municipio"
             Exit Function
        Case 3   '- Endereço do Destinatário - Complemento
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "003 - Endereço do Destinatário - Complemento"
             Exit Function
        Case 4   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "004 - Logistica informará assim que possível"
             Exit Function
        Case 5   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "005 - Data de Fabricação do Lote Inválida"
             Exit Function
        
        Case 6   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "006 - Problemas no sistema da logística "
             Exit Function
        
        Case 7   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "007 - Sem saldo para atender a solicitação"
             Exit Function
        
        Case 8   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "008 - Item solicitado em duplicidade"
             Exit Function
        
        Case 9   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "009 - Erro de conversão (Tamanho do Campo)"
             Exit Function
        
        Case 10   '- Logistica informará assim que possível
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "010 - Item solicitado em duplicidade"
             Exit Function
        
        Case 210 '- IE do destinatário inválida"
             Sheets("GERAL").Range("K" & xx).Value = strK
             Sheets("GERAL").Range("L" & xx).Value = "210 - IE do destinatário inválida"
             Exit Function

    End Select

    Sheets("GERAL").Range("B" & xx & ":N" & xx).Interior.ColorIndex = 0
    Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.ColorIndex = 1
    Sheets("GERAL").Range("B" & xx & ":N" & xx).Font.Bold = False

    If CDbl(Sheets("GERAL").Range("B" & xx).Value) = 10 Then
       If CDbl(Sheets("GERAL").Range("D" & xx).Value) <> 21 Then
          Sheets("GERAL").Range("K" & xx).Value = "Enviado"
          Sheets("GERAL").Range("L" & xx).Value = "Correto"
       Else
          Sheets("GERAL").Range("K" & xx).Value = ""
          Sheets("GERAL").Range("L" & xx).Value = ""
       End If
    Else
          Sheets("GERAL").Range("K" & xx).Value = "Enviado"
          Sheets("GERAL").Range("L" & xx).Value = "Correto"
    End If

End Function









