Attribute VB_Name = "Módulo1"

Public Sub Verifica_MSTXDET()

    On Error GoTo msgerro:

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    
    Dim str_con As String
    Dim str_sql As String
 
    Dim int_err As Integer
    
    Dim str_usr As String
    Dim str_pas As String
 
    Dim str_itm As String
    Dim dbl_codigo As Double
    
    Dim dbr_lin As Double

    Sheets("Ver").Select
    
    Sheets("Erros").Rows("3:15999").Delete Shift:=xlUp
    Sheets("Ver").Rows("3:15999").Delete Shift:=xlUp

    dbl_header = 3
    dbl_linha = 3
    UserForm1.Show 1
    
    If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
       Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
    
        cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
        cn.Open
    
        dbr_lin = 3
    
        str_sql = " SELECT " & _
                  "     tdoc1 || tdoc2 as TIPO , " & _
                  "     anodup || mesdup || diadup as WDATA , " & _
                  "     DIVI, CLIENT, NDUPL,  " & _
                  "     VLTTAL, VLICM, DESPE, DCVLP, DCICM," & _
                  "     DCPIT, DCCDP " & _
                  " FROM " & _
                  "     sikama.FATMST " & _
                  " WHERE " & _
                  "     tdoc1||tdoc2 IN('10', '12', '15', '20', '22', '25') and " & _
                  "     ( divi = '20' or divi = '08' or  divi = '60' or divi = '10' ) " & _
                  " Order by " & _
                  "     divi, ndupl "
               
               ' "     tdoc1||tdoc2 IN('10', '12', '15', '20', '22', '25', '50', '52', '55') and " & _

        rs.Open str_sql, cn, adOpenForwardOnly
        Do While Not rs.EOF
    
            Sheets("Ver").Range("a" & dbl_linha).Select
            DoEvents
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "FATMST"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 11
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            
            Sheets("Ver").Range("C" & dbl_linha).Value = rs!divi
            Sheets("Ver").Range("D" & dbl_linha).Value = rs!ndupl
            Sheets("Ver").Range("E" & dbl_linha).Value = rs!client
            Sheets("Ver").Range("F" & dbl_linha).Value = rs!TIPO
            Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rs!wdata, 7, 2) & "/" & Mid(rs!wdata, 5, 2) & "/" & Mid(rs!wdata, 1, 4))
            Sheets("Ver").Range("J" & dbl_linha).Value = rs!VLTTAL
            Sheets("Ver").Range("K" & dbl_linha).Value = rs!DCPIT
            Sheets("Ver").Range("L" & dbl_linha).Value = rs!DESPE
            Sheets("Ver").Range("M" & dbl_linha).Value = rs!DCICM
            Sheets("Ver").Range("N" & dbl_linha).Value = rs!DCVLP
            Sheets("Ver").Range("O" & dbl_linha).Value = rs!VLICM
            Sheets("Ver").Range("P" & dbl_linha).Value = rs!DCCDP
           
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            dbr_lin = dbl_linha
            
            str_sql = " SELECT " & _
                      "    tdoc1 || tdoc2 as TIPO , cprod1 || cprod2 || fill06 as str_item , DIVI, NDUPL, CLIENT, anodup || mesdup || diadup as WDATA , " & _
                      "    VLNORM, VALLIQ, VALICM, DESCIT, DESCPE, DESCES, DESCRE, DESCDP, ICMZF " & _
                      " FROM " & _
                      "     sikama.FATDET " & _
                      " WHERE " & _
                      "     tdoc1||tdoc2 = '" & rs!TIPO & "' and " & _
                      "     ndupl = '" & rs!ndupl & "' and " & _
                      "     anodup || mesdup || diadup = '" & rs!wdata & "'"
        
            rt.Open str_sql, cn, adOpenForwardOnly
        
            str_formula = ""
        
            Sheets("Ver").Range("i" & dbr_lin).Value = 0
        
            Do While Not rt.EOF
        
                dbl_linha = dbl_linha + 1
            
                Sheets("Ver").Range("B" & dbl_linha).Value = "FATDET"
                Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 18
                Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
                
                Sheets("Ver").Range("C" & dbl_linha).Value = rt!divi
                Sheets("Ver").Range("D" & dbl_linha).Value = rt!ndupl
                Sheets("Ver").Range("E" & dbl_linha).Value = rt!client
                Sheets("Ver").Range("F" & dbl_linha).Value = rt!TIPO
                Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rt!wdata, 7, 2) & "/" & Mid(rt!wdata, 5, 2) & "/" & Mid(rt!wdata, 1, 4))
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(rt!VLNORM * -1, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(rt!VALLIQ * -1, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round((rt!DESCIT * -1), 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round((rt!DESCPE * -1) + (rt!DESCES * -1), 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = 0 'Round(rt!ICMZF * -1, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(rt!DESCRE * -1, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(rt!VALICM * -1, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(rt!DESCDP * -1, 2)
            
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                
                Sheets("Ver").Range("i" & dbr_lin).Value = Sheets("Ver").Range("i" & dbr_lin).Value + rt!VALLIQ + rt!DESCIT + rt!DESCPE + rt!DESCES + rt!DESCRE + rt!DESCDP + rt!ICMZF
                
               rt.MoveNext
            Loop
            
            rt.Close
            
            Rows(dbr_lin & ":" & dbl_linha).Rows.Group
            dbl_linha = dbl_linha + 1
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 13
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 15
            Sheets("Ver").Range("C" & dbl_linha).Value = Sheets("Ver").Range("C" & dbr_lin).Value
            Sheets("Ver").Range("D" & dbl_linha).Value = Sheets("Ver").Range("D" & dbr_lin).Value
            Sheets("Ver").Range("E" & dbl_linha).Value = Sheets("Ver").Range("E" & dbr_lin).Value
            Sheets("Ver").Range("F" & dbl_linha).Value = Sheets("Ver").Range("F" & dbr_lin).Value
            Sheets("Ver").Range("G" & dbl_linha).Value = Sheets("Ver").Range("G" & dbr_lin).Value
            
            For dbr_idx = dbr_lin To dbl_linha - 1
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(Sheets("Ver").Range("I" & dbl_linha).Value + Sheets("Ver").Range("I" & dbr_idx).Value, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(Sheets("Ver").Range("J" & dbl_linha).Value + Sheets("Ver").Range("J" & dbr_idx).Value, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round(Sheets("Ver").Range("K" & dbl_linha).Value + Sheets("Ver").Range("K" & dbr_idx).Value, 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round(Sheets("Ver").Range("L" & dbl_linha).Value + Sheets("Ver").Range("L" & dbr_idx).Value, 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = Round(Sheets("Ver").Range("M" & dbl_linha).Value + Sheets("Ver").Range("M" & dbr_idx).Value, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(Sheets("Ver").Range("N" & dbl_linha).Value + Sheets("Ver").Range("N" & dbr_idx).Value, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(Sheets("Ver").Range("O" & dbl_linha).Value + Sheets("Ver").Range("O" & dbr_idx).Value, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(Sheets("Ver").Range("P" & dbl_linha).Value + Sheets("Ver").Range("P" & dbr_idx).Value, 2)
            Next
            
            If Sheets("Ver").Range("I" & dbl_linha).Value <> 0 Then
               Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 37
            End If
            
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            ActiveSheet.Outline.ShowLevels RowLevels:=1
            
            dbl_linha = dbl_linha + 1
            
          rs.MoveNext
        Loop
    
        rs.Close

        Rows("3:" & dbl_linha - 1).Rows.Group
    
   End If

   cn.Close

    dbl_new = 3

    For dbl_linha = dbl_linha To 1 Step -1

        Range("b" & dbl_linha).Select

        dbl_dif = Sheets("Ver").Range("I" & dbl_linha).Value + _
                  Sheets("Ver").Range("J" & dbl_linha).Value + _
                  Sheets("Ver").Range("K" & dbl_linha).Value + _
                  Sheets("Ver").Range("L" & dbl_linha).Value + _
                  Sheets("Ver").Range("M" & dbl_linha).Value + _
                  Sheets("Ver").Range("N" & dbl_linha).Value + _
                  Sheets("Ver").Range("O" & dbl_linha).Value + _
                  Sheets("Ver").Range("P" & dbl_linha).Value

        If Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL" Then
        
            If dbl_dif <> 0 Then

            Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 13
            Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
            Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_linha).Value
            Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_linha).Value
            Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_linha).Value
            Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_linha).Value
            Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_linha).Value
            Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_linha).Value
            Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_linha).Value
            Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_linha).Value
            Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_linha).Value
            Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_linha).Value
            Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_linha).Value
            Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_linha).Value
            Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_linha).Value

            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            Sheets("Erros").Range("C" & dbl_new & ":Q" & dbl_new).Interior.ColorIndex = 37

            dbl_new = dbl_new + 1

          For dbl_erro = dbl_linha - 1 To 1 Step -1

                Range("b" & dbl_new).Select

                If Sheets("Ver").Range("B" & dbl_erro).Value = "TOTAL" Then
                   dbl_linha = dbl_erro + 1
                   Exit For
                End If


                Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_erro).Value
                Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_erro).Value
                Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_erro).Value
                Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_erro).Value
                Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_erro).Value
                Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_erro).Value
                Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_erro).Value
                Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_erro).Value
                Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_erro).Value
                Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_erro).Value
                Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_erro).Value
                Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_erro).Value
                Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_erro).Value
                Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_erro).Value
                
                If Sheets("Erros").Range("B" & dbl_new).Value = "FATDET" Then
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 18
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_linha).Value
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                Else
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 11
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                End If
                
                dbl_new = dbl_new + 1
          Next

         End If

        End If

    Next

   MsgBox "Terminou"
   Exit Sub

msgerro:

     MsgBox Err.Number & " - " & Err.Description, vbCritical
 
End Sub


Public Sub Verifica_MSTXDET_Animal()

    On Error GoTo msgerro:

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    
    Dim str_con As String
    Dim str_sql As String
 
    Dim int_err As Integer
    
    Dim str_usr As String
    Dim str_pas As String
 
    Dim str_itm As String
    Dim dbl_codigo As Double
    
    Dim dbr_lin As Double

    Sheets("Ver").Select
    
    Sheets("Erros").Rows("3:15999").Delete Shift:=xlUp
    Sheets("Ver").Rows("3:15999").Delete Shift:=xlUp

    dbl_header = 3
    dbl_linha = 3
    UserForm1.Show 1
    
    If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
       Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
    
        cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
        cn.Open
    
        dbr_lin = 3
    
        str_sql = " SELECT " & _
                  "     tdoc1 || tdoc2 as TIPO , " & _
                  "     anodup || mesdup || diadup as WDATA , " & _
                  "     DIVI, CLIENT, NDUPL,  " & _
                  "     VLTTAL, VLICM, DESPE, DCVLP, DCICM," & _
                  "     DCPIT, DCCDP " & _
                  " FROM " & _
                  "     sikama.FATMST " & _
                  " WHERE " & _
                  "     tdoc1||tdoc2 IN('10', '12', '15', '20', '22', '25', '50', '52', '55') and " & _
                  "     divi = '20'  " & _
                  " Order by " & _
                  "     divi, ndupl "
               
        rs.Open str_sql, cn, adOpenForwardOnly
        Do While Not rs.EOF
    
            Sheets("Ver").Range("a" & dbl_linha).Select
            DoEvents
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "FATMST"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 11
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            
            Sheets("Ver").Range("C" & dbl_linha).Value = rs!divi
            Sheets("Ver").Range("D" & dbl_linha).Value = rs!ndupl
            Sheets("Ver").Range("E" & dbl_linha).Value = rs!client
            Sheets("Ver").Range("F" & dbl_linha).Value = rs!TIPO
            Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rs!wdata, 7, 2) & "/" & Mid(rs!wdata, 5, 2) & "/" & Mid(rs!wdata, 1, 4))
            Sheets("Ver").Range("J" & dbl_linha).Value = rs!VLTTAL
            Sheets("Ver").Range("K" & dbl_linha).Value = rs!DCPIT
            Sheets("Ver").Range("L" & dbl_linha).Value = rs!DESPE
            Sheets("Ver").Range("M" & dbl_linha).Value = rs!DCICM
            Sheets("Ver").Range("N" & dbl_linha).Value = rs!DCVLP
            Sheets("Ver").Range("O" & dbl_linha).Value = rs!VLICM
            Sheets("Ver").Range("P" & dbl_linha).Value = rs!DCCDP
           
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            dbr_lin = dbl_linha
            
            str_sql = " SELECT " & _
                      "    tdoc1 || tdoc2 as TIPO , cprod1 || cprod2 || fill06 as str_item , DIVI, NDUPL, CLIENT, anodup || mesdup || diadup as WDATA , " & _
                      "    VLNORM, VALLIQ, VALICM, DESCIT, DESCPE, DESCES, DESCRE, DESCDP, ICMZF " & _
                      " FROM " & _
                      "     sikama.FATDET " & _
                      " WHERE " & _
                      "     tdoc1||tdoc2 = '" & rs!TIPO & "' and " & _
                      "     ndupl = '" & rs!ndupl & "' and " & _
                      "     anodup || mesdup || diadup = '" & rs!wdata & "'"
        
            rt.Open str_sql, cn, adOpenForwardOnly
        
            str_formula = ""
        
            Sheets("Ver").Range("i" & dbr_lin).Value = 0
        
            Do While Not rt.EOF
        
                dbl_linha = dbl_linha + 1
            
                Sheets("Ver").Range("B" & dbl_linha).Value = "FATDET"
                Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 18
                Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
                
                Sheets("Ver").Range("C" & dbl_linha).Value = rt!divi
                Sheets("Ver").Range("D" & dbl_linha).Value = rt!ndupl
                Sheets("Ver").Range("E" & dbl_linha).Value = rt!client
                Sheets("Ver").Range("F" & dbl_linha).Value = rt!TIPO
                Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rt!wdata, 7, 2) & "/" & Mid(rt!wdata, 5, 2) & "/" & Mid(rt!wdata, 1, 4))
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(rt!VLNORM * -1, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(rt!VALLIQ * -1, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round((rt!DESCIT * -1), 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round((rt!DESCPE * -1) + (rt!DESCES * -1), 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = Round(rt!ICMZF * -1, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(rt!DESCRE * -1, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(rt!VALICM * -1, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(rt!DESCDP * -1, 2)
            
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                
                Sheets("Ver").Range("i" & dbr_lin).Value = Sheets("Ver").Range("i" & dbr_lin).Value + rt!VALLIQ + rt!DESCIT + rt!DESCPE + rt!DESCES + rt!DESCRE + rt!DESCDP + rt!ICMZF
                
               rt.MoveNext
            Loop
            
            rt.Close
            
            Rows(dbr_lin & ":" & dbl_linha).Rows.Group
            dbl_linha = dbl_linha + 1
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 13
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 15
            Sheets("Ver").Range("C" & dbl_linha).Value = Sheets("Ver").Range("C" & dbr_lin).Value
            Sheets("Ver").Range("D" & dbl_linha).Value = Sheets("Ver").Range("D" & dbr_lin).Value
            Sheets("Ver").Range("E" & dbl_linha).Value = Sheets("Ver").Range("E" & dbr_lin).Value
            Sheets("Ver").Range("F" & dbl_linha).Value = Sheets("Ver").Range("F" & dbr_lin).Value
            Sheets("Ver").Range("G" & dbl_linha).Value = Sheets("Ver").Range("G" & dbr_lin).Value
            
            For dbr_idx = dbr_lin To dbl_linha - 1
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(Sheets("Ver").Range("I" & dbl_linha).Value + Sheets("Ver").Range("I" & dbr_idx).Value, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(Sheets("Ver").Range("J" & dbl_linha).Value + Sheets("Ver").Range("J" & dbr_idx).Value, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round(Sheets("Ver").Range("K" & dbl_linha).Value + Sheets("Ver").Range("K" & dbr_idx).Value, 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round(Sheets("Ver").Range("L" & dbl_linha).Value + Sheets("Ver").Range("L" & dbr_idx).Value, 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = Round(Sheets("Ver").Range("M" & dbl_linha).Value + Sheets("Ver").Range("M" & dbr_idx).Value, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(Sheets("Ver").Range("N" & dbl_linha).Value + Sheets("Ver").Range("N" & dbr_idx).Value, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(Sheets("Ver").Range("O" & dbl_linha).Value + Sheets("Ver").Range("O" & dbr_idx).Value, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(Sheets("Ver").Range("P" & dbl_linha).Value + Sheets("Ver").Range("P" & dbr_idx).Value, 2)
            Next
            
            If Sheets("Ver").Range("I" & dbl_linha).Value <> 0 Then
               Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 37
            End If
            
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            ActiveSheet.Outline.ShowLevels RowLevels:=1
            
            dbl_linha = dbl_linha + 1
            
          rs.MoveNext
        Loop
    
        rs.Close

        Rows("3:" & dbl_linha - 1).Rows.Group
    
   End If

   cn.Close

    dbl_new = 3

    For dbl_linha = dbl_linha To 1 Step -1

        Range("b" & dbl_linha).Select

        dbl_dif = Sheets("Ver").Range("I" & dbl_linha).Value + _
                  Sheets("Ver").Range("J" & dbl_linha).Value + _
                  Sheets("Ver").Range("K" & dbl_linha).Value + _
                  Sheets("Ver").Range("L" & dbl_linha).Value + _
                  Sheets("Ver").Range("M" & dbl_linha).Value + _
                  Sheets("Ver").Range("N" & dbl_linha).Value + _
                  Sheets("Ver").Range("O" & dbl_linha).Value + _
                  Sheets("Ver").Range("P" & dbl_linha).Value

        If Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL" Then
        
            If dbl_dif <> 0 Then

            Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 13
            Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
            Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_linha).Value
            Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_linha).Value
            Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_linha).Value
            Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_linha).Value
            Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_linha).Value
            Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_linha).Value
            Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_linha).Value
            Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_linha).Value
            Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_linha).Value
            Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_linha).Value
            Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_linha).Value
            Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_linha).Value
            Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_linha).Value

            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            Sheets("Erros").Range("C" & dbl_new & ":Q" & dbl_new).Interior.ColorIndex = 37

            dbl_new = dbl_new + 1

          For dbl_erro = dbl_linha - 1 To 1 Step -1

                Range("b" & dbl_new).Select

                If Sheets("Ver").Range("B" & dbl_erro).Value = "TOTAL" Then
                   dbl_linha = dbl_erro + 1
                   Exit For
                End If


                Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_erro).Value
                Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_erro).Value
                Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_erro).Value
                Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_erro).Value
                Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_erro).Value
                Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_erro).Value
                Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_erro).Value
                Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_erro).Value
                Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_erro).Value
                Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_erro).Value
                Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_erro).Value
                Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_erro).Value
                Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_erro).Value
                Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_erro).Value
                
                If Sheets("Erros").Range("B" & dbl_new).Value = "FATDET" Then
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 18
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_linha).Value
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                Else
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 11
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                End If
                
                dbl_new = dbl_new + 1
          Next

         End If

        End If

    Next

   MsgBox "Terminou"
   Exit Sub

msgerro:

     MsgBox Err.Number & " - " & Err.Description, vbCritical
 
End Sub


Public Sub Verifica_Seq_Notas()

    On Error GoTo msgerro:

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    
    Dim str_con As String
    Dim str_sql As String
    Dim dbl_nota(10) As Double
    Dim int_err As Integer
    
    Dim dbl_ctrl As Double
    
    Dim str_usr As String
    Dim str_pas As String
 
    Dim str_itm As String
    Dim dbl_codigo As Double
    
    Dim dbr_lin As Double

    Sheets("SeqNF").Select
    
    Sheets("SeqNF").Rows("4:15999").Delete Shift:=xlUp

    dbl_header = 3
    dbl_linha = 3
    UserForm1.Show 1
    
    If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
       Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
    
        cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
        cn.Open
    
    
        str_sql = " SELECT " & _
                  "     *  " & _
                  " FROM " & _
                  "     pfzdata.TABAGV " & _
                  " ORDER BY " & _
                  "    AGNFS desc"
        
        dbr_lin = 3

        rs.Open str_sql, cn, adOpenForwardOnly
        Do While Not rs.EOF
        
            Cells(3, dbr_lin).Value = rs!AGSER
            Cells(3, dbr_lin + 1).Value = rs!AGNFS
           
           dbr_lin = dbr_lin + 2
           rs.MoveNext
        Loop
    
        rs.Close
    
        For dbl_lin = 3 To dbr_lin
    
            If Len(Trim(Cells(3, dbl_lin).Value)) = 0 Then
                Exit For
            End If
    
            str_sql = " SELECT " & _
                      "     NDUPL, SDUPL " & _
                      " FROM " & _
                      "     sikama.FATMST " & _
                      " WHERE sdupl ='" & Cells(3, dbl_lin) & "' and " & _
                      "       (TDOC1 <> '1' and TDOC1 <> '2' and TDOC1 <> '7')  " & _
                      " Order by " & _
                      "     ndupl desc "
                       
            dbl_ctrl = 4
                   
            rs.Open str_sql, cn, adOpenForwardOnly
            Do While Not rs.EOF
        
                Sheets("SeqNF").Range("B" & dbl_ctrl).Value = "FATMST"
        
                Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin).Select
        
                Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin + 1).Value = rs!ndupl
              
                If Cells(dbl_ctrl - 1, dbl_lin + 1).Value - 1 = CDbl(rs!ndupl) Then
                    Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin).Value = "OK"
                Else
                    Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin).Value = "Erro"
                    Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin).Interior.ColorIndex = 9
                    Sheets("SeqNF").Cells(dbl_ctrl, dbl_lin).Font.ColorIndex = 2
                 End If
              
                dbl_ctrl = dbl_ctrl + 1
              rs.MoveNext
            Loop
        
            dbr_lin = dbr_lin + 1
            rs.Close
            
        Next
        
   End If

   cn.Close

    Sheets("SeqNF").Range("a1").Select

   MsgBox "Terminou"
   Exit Sub

msgerro:

     MsgBox Err.Number & " - " & Err.Description, vbCritical
 
End Sub




Public Sub Verifica_MSTXDET_Animal_PorNota()

    On Error GoTo msgerro:

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    
    Dim str_con As String
    Dim str_sql As String
 
    Dim int_err As Integer
    
    Dim str_usr As String
    Dim str_pas As String
 
    Dim str_itm As String
    Dim dbl_codigo As Double
    
    Dim dbr_lin As Double

    Sheets("Ver").Select
    
    Sheets("Erros").Rows("3:15999").Delete Shift:=xlUp
    Sheets("Ver").Rows("3:15999").Delete Shift:=xlUp

    dbl_header = 3
    dbl_linha = 3
    UserForm1.Show 1
    
    If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
       Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
    
        cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
        cn.Open
    
        dbr_lin = 3
    
        str_sql = " SELECT " & _
                  "     tdoc1 || tdoc2 as TIPO , " & _
                  "     anodup || mesdup || diadup as WDATA , " & _
                  "     DIVI, CLIENT, NDUPL,  " & _
                  "     VLTTAL, VLICM, DESPE, DCVLP, DCICM," & _
                  "     DCPIT, DCCDP " & _
                  " FROM " & _
                  "     sikama.FATMST " & _
                  " WHERE " & _
                  "     tdoc1||tdoc2 IN('10', '12', '15', '20', '22', '25', '50', '52', '55') and " & _
                  "      ndupl in('090520', '230520')" & _
                  " Order by " & _
                  "     divi, ndupl "
               
        rs.Open str_sql, cn, adOpenForwardOnly
        Do While Not rs.EOF
    
            Sheets("Ver").Range("a" & dbl_linha).Select
            DoEvents
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "FATMST"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 11
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            
            Sheets("Ver").Range("C" & dbl_linha).Value = rs!divi
            Sheets("Ver").Range("D" & dbl_linha).Value = rs!ndupl
            Sheets("Ver").Range("E" & dbl_linha).Value = rs!client
            Sheets("Ver").Range("F" & dbl_linha).Value = rs!TIPO
            Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rs!wdata, 7, 2) & "/" & Mid(rs!wdata, 5, 2) & "/" & Mid(rs!wdata, 1, 4))
            Sheets("Ver").Range("J" & dbl_linha).Value = rs!VLTTAL
            Sheets("Ver").Range("K" & dbl_linha).Value = rs!DCPIT
            Sheets("Ver").Range("L" & dbl_linha).Value = rs!DESPE
            Sheets("Ver").Range("M" & dbl_linha).Value = rs!DCICM
            Sheets("Ver").Range("N" & dbl_linha).Value = rs!DCVLP
            Sheets("Ver").Range("O" & dbl_linha).Value = rs!VLICM
            Sheets("Ver").Range("P" & dbl_linha).Value = rs!DCCDP
           
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            dbr_lin = dbl_linha
            
            str_sql = " SELECT " & _
                      "    tdoc1 || tdoc2 as TIPO , cprod1 || cprod2 || fill06 as str_item , DIVI, NDUPL, CLIENT, anodup || mesdup || diadup as WDATA , " & _
                      "    VLNORM, VALLIQ, VALICM, DESCIT, DESCPE, DESCES, DESCRE, DESCDP, ICMZF " & _
                      " FROM " & _
                      "     sikama.FATDET " & _
                      " WHERE " & _
                      "     tdoc1||tdoc2 = '" & rs!TIPO & "' and " & _
                      "     ndupl = '" & rs!ndupl & "' and " & _
                      "     anodup || mesdup || diadup = '" & rs!wdata & "'"
        
            rt.Open str_sql, cn, adOpenForwardOnly
        
            str_formula = ""
        
            Sheets("Ver").Range("i" & dbr_lin).Value = 0
        
            Do While Not rt.EOF
        
                dbl_linha = dbl_linha + 1
            
                Sheets("Ver").Range("B" & dbl_linha).Value = "FATDET"
                Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 18
                Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
                
                Sheets("Ver").Range("C" & dbl_linha).Value = rt!divi
                Sheets("Ver").Range("D" & dbl_linha).Value = rt!ndupl
                Sheets("Ver").Range("E" & dbl_linha).Value = rt!client
                Sheets("Ver").Range("F" & dbl_linha).Value = rt!TIPO
                Sheets("Ver").Range("G" & dbl_linha).Value = CDate(Mid(rt!wdata, 7, 2) & "/" & Mid(rt!wdata, 5, 2) & "/" & Mid(rt!wdata, 1, 4))
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(rt!VLNORM * -1, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(rt!VALLIQ * -1, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round((rt!DESCIT * -1), 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round((rt!DESCPE * -1) + (rt!DESCES * -1), 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = Round(rt!ICMZF * -1, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(rt!DESCRE * -1, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(rt!VALICM * -1, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(rt!DESCDP * -1, 2)
            
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
                Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                
                Sheets("Ver").Range("i" & dbr_lin).Value = Sheets("Ver").Range("i" & dbr_lin).Value + rt!VALLIQ + rt!DESCIT + rt!DESCPE + rt!DESCES + rt!DESCRE + rt!DESCDP + rt!ICMZF
                
               rt.MoveNext
            Loop
            
            rt.Close
            
            Rows(dbr_lin & ":" & dbl_linha).Rows.Group
            dbl_linha = dbl_linha + 1
            
            Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL"
            Sheets("Ver").Range("B" & dbl_linha).Interior.ColorIndex = 13
            Sheets("Ver").Range("B" & dbl_linha).Font.ColorIndex = 2
            Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 15
            Sheets("Ver").Range("C" & dbl_linha).Value = Sheets("Ver").Range("C" & dbr_lin).Value
            Sheets("Ver").Range("D" & dbl_linha).Value = Sheets("Ver").Range("D" & dbr_lin).Value
            Sheets("Ver").Range("E" & dbl_linha).Value = Sheets("Ver").Range("E" & dbr_lin).Value
            Sheets("Ver").Range("F" & dbl_linha).Value = Sheets("Ver").Range("F" & dbr_lin).Value
            Sheets("Ver").Range("G" & dbl_linha).Value = Sheets("Ver").Range("G" & dbr_lin).Value
            
            For dbr_idx = dbr_lin To dbl_linha - 1
                Sheets("Ver").Range("I" & dbl_linha).Value = Round(Sheets("Ver").Range("I" & dbl_linha).Value + Sheets("Ver").Range("I" & dbr_idx).Value, 2)
                Sheets("Ver").Range("J" & dbl_linha).Value = Round(Sheets("Ver").Range("J" & dbl_linha).Value + Sheets("Ver").Range("J" & dbr_idx).Value, 2)
                Sheets("Ver").Range("K" & dbl_linha).Value = Round(Sheets("Ver").Range("K" & dbl_linha).Value + Sheets("Ver").Range("K" & dbr_idx).Value, 2)
                Sheets("Ver").Range("L" & dbl_linha).Value = Round(Sheets("Ver").Range("L" & dbl_linha).Value + Sheets("Ver").Range("L" & dbr_idx).Value, 2)
                Sheets("Ver").Range("M" & dbl_linha).Value = Round(Sheets("Ver").Range("M" & dbl_linha).Value + Sheets("Ver").Range("M" & dbr_idx).Value, 2)
                Sheets("Ver").Range("N" & dbl_linha).Value = Round(Sheets("Ver").Range("N" & dbl_linha).Value + Sheets("Ver").Range("N" & dbr_idx).Value, 2)
                Sheets("Ver").Range("O" & dbl_linha).Value = Round(Sheets("Ver").Range("O" & dbl_linha).Value + Sheets("Ver").Range("O" & dbr_idx).Value, 2)
                Sheets("Ver").Range("P" & dbl_linha).Value = Round(Sheets("Ver").Range("P" & dbl_linha).Value + Sheets("Ver").Range("P" & dbr_idx).Value, 2)
            Next
            
            If Sheets("Ver").Range("I" & dbl_linha).Value <> 0 Then
               Sheets("Ver").Range("C" & dbl_linha & ":Q" & dbl_linha).Interior.ColorIndex = 37
            End If
            
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Ver").Range("B" & dbl_linha & ":Q" & dbl_linha).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            
            ActiveSheet.Outline.ShowLevels RowLevels:=1
            
            dbl_linha = dbl_linha + 1
            
          rs.MoveNext
        Loop
    
        rs.Close

        Rows("3:" & dbl_linha - 1).Rows.Group
    
   End If

   cn.Close

    dbl_new = 3

    For dbl_linha = dbl_linha To 1 Step -1

        Range("b" & dbl_linha).Select

        dbl_dif = Sheets("Ver").Range("I" & dbl_linha).Value + _
                  Sheets("Ver").Range("J" & dbl_linha).Value + _
                  Sheets("Ver").Range("K" & dbl_linha).Value + _
                  Sheets("Ver").Range("L" & dbl_linha).Value + _
                  Sheets("Ver").Range("M" & dbl_linha).Value + _
                  Sheets("Ver").Range("N" & dbl_linha).Value + _
                  Sheets("Ver").Range("O" & dbl_linha).Value + _
                  Sheets("Ver").Range("P" & dbl_linha).Value

        If Sheets("Ver").Range("B" & dbl_linha).Value = "TOTAL" Then
        
            If dbl_dif <> 0 Then

            Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 13
            Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
            Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_linha).Value
            Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_linha).Value
            Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_linha).Value
            Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_linha).Value
            Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_linha).Value
            Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_linha).Value
            Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_linha).Value
            Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_linha).Value
            Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_linha).Value
            Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_linha).Value
            Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_linha).Value
            Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_linha).Value
            Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_linha).Value

            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
            Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
            Sheets("Erros").Range("C" & dbl_new & ":Q" & dbl_new).Interior.ColorIndex = 37

            dbl_new = dbl_new + 1

          For dbl_erro = dbl_linha - 1 To 1 Step -1

                Range("b" & dbl_new).Select

                If Sheets("Ver").Range("B" & dbl_erro).Value = "TOTAL" Then
                   dbl_linha = dbl_erro + 1
                   Exit For
                End If


                Sheets("Erros").Range("B" & dbl_new).Value = Sheets("Ver").Range("B" & dbl_erro).Value
                Sheets("Erros").Range("C" & dbl_new).Value = Sheets("Ver").Range("C" & dbl_erro).Value
                Sheets("Erros").Range("D" & dbl_new).Value = Sheets("Ver").Range("D" & dbl_erro).Value
                Sheets("Erros").Range("E" & dbl_new).Value = Sheets("Ver").Range("E" & dbl_erro).Value
                Sheets("Erros").Range("F" & dbl_new).Value = Sheets("Ver").Range("F" & dbl_erro).Value
                Sheets("Erros").Range("G" & dbl_new).Value = Sheets("Ver").Range("G" & dbl_erro).Value
                Sheets("Erros").Range("I" & dbl_new).Value = Sheets("Ver").Range("I" & dbl_erro).Value
                Sheets("Erros").Range("J" & dbl_new).Value = Sheets("Ver").Range("J" & dbl_erro).Value
                Sheets("Erros").Range("K" & dbl_new).Value = Sheets("Ver").Range("K" & dbl_erro).Value
                Sheets("Erros").Range("L" & dbl_new).Value = Sheets("Ver").Range("L" & dbl_erro).Value
                Sheets("Erros").Range("M" & dbl_new).Value = Sheets("Ver").Range("M" & dbl_erro).Value
                Sheets("Erros").Range("N" & dbl_new).Value = Sheets("Ver").Range("N" & dbl_erro).Value
                Sheets("Erros").Range("O" & dbl_new).Value = Sheets("Ver").Range("O" & dbl_erro).Value
                Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_erro).Value
                
                If Sheets("Erros").Range("B" & dbl_new).Value = "FATDET" Then
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 18
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("P" & dbl_new).Value = Sheets("Ver").Range("P" & dbl_linha).Value
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                Else
                    Sheets("Erros").Range("B" & dbl_new).Interior.ColorIndex = 11
                    Sheets("Erros").Range("B" & dbl_new).Font.ColorIndex = 2
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).LineStyle = xlContinuous
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).Weight = xlThin
                    Sheets("Erros").Range("B" & dbl_new & ":Q" & dbl_new).Borders(xlInsideVertical).ColorIndex = xlAutomatic
                End If
                
                dbl_new = dbl_new + 1
          Next

         End If

        End If

    Next

   MsgBox "Terminou"
   Exit Sub

msgerro:

     MsgBox Err.Number & " - " & Err.Description, vbCritical
 
End Sub

