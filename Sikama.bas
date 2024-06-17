Attribute VB_Name = "Sikama"
Public Sub Extrato_Itau_Limpa()

    On Error Resume Next
     
       Sheets("Limpo").Select
       If Err.Number > 0 Then
          Sheets.Add After:=ActiveSheet
          ActiveSheet.Name = "Limpo"
       End If
       
    On Error GoTo 0:
    
    Sheets("Limpo").Select
    Sheets("Limpo").Cells.Delete Shift:=xlUp
   
    
    Sheets("Limpo").Range("D5").Value = "data"
    Sheets("Limpo").Range("E5").Value = "lançamento"
    Sheets("Limpo").Range("F5").Value = "ag./origem"
    Sheets("Limpo").Range("G5").Value = "valor (R$)"
    Sheets("Limpo").Range("H5").Value = "saldos (R$)"
    
    dLin = 12
    dNew = 6
    
    Do While Len(Trim(Sheets("Lançamentos").Range("B" & dLin).Value))

       If InStr(Sheets("Lançamentos").Range("A" & dLin).Value, "lançamentos futuros") > 0 Then
          Exit Do
       End If

       If InStr(Sheets("Lançamentos").Range("B" & dLin).Value, "SALDO") = 0 Then
          Sheets("Limpo").Range("D" & dNew).Value = Sheets("Lançamentos").Range("A" & dLin).Value
          Sheets("Limpo").Range("E" & dNew).Value = Sheets("Lançamentos").Range("B" & dLin).Value
          Sheets("Limpo").Range("F" & dNew).Value = Sheets("Lançamentos").Range("C" & dLin).Value
          Sheets("Limpo").Range("G" & dNew).Value = Sheets("Lançamentos").Range("D" & dLin).Value
          dNew = dNew + 1
       End If

       Range("B" & dLin).Select
       DoEvents

       dLin = dLin + 1
    Loop
       

End Sub


Public Sub Extrato_Itau_Limpas()

   Dim sTab(1 To 12) As String

   sTab(1) = "Janeiro"
   sTab(2) = "Fevereiro"
   sTab(3) = "Março"
   sTab(4) = "Abril"
   sTab(5) = "Maio"
   sTab(6) = "Junho"
   sTab(7) = "Julho"
   sTab(8) = "Agosto"
   sTab(9) = "Setembro"
   sTab(10) = "Outubro"
   sTab(11) = "Novembro"
   sTab(12) = "Dezembro"
   

   If Range("B5").Value <> "59548-9" And Range("C5").Value <> "59548-9" Then
      MsgBox "Não é uma planilha de Extrato - Itau"
      Exit Sub
   End If

   If Range("B5").Value = "59548-9" Then
      Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
      Range("B5").Value = "Mês"
   End If

   dLin = 10
   Do While Len(Trim(Range("B5").Value)) > 0

        If Range("C" & dLin).Value <> "SALDO ANTERIOR" And _
           Range("C" & dLin).Value <> "SALDO DO DIA" And _
           Range("C" & dLin).Value <> "SALDO ANTERIOR" Then
        
           If Range("C" & dLin).Value = "lançamentos futuros" Then
              Exit Do
           End If
        
           Range("A" & dLin).Value = Weekday(Range("B" & dLin).Value)
        
        
        
        
        End If

      dLin = dLin + 1
   Loop

End Sub
