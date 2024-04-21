Attribute VB_Name = "ODBC400"
Public str_user As String
Public str_pws  As String
Public cn       As New ADODB.Connection
Public rs       As New ADODB.Recordset

Private Function AS400(ByVal act As Boolean)

    If act = True Then

        AS400 = False
 
        If cn.State = 0 Then          '0=Desconectado e 1=Concectado

            If Len(Trim(str_user)) = 0 And _
               Len(Trim(str_pws)) = 0 Then
            
               UserForm1.Show 1
                
                If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
                   Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
                   str_user = Trim(UserForm1.TextBox1.Text)
                   str_pws = Trim(UserForm1.TextBox2.Text)
                End If
                
             End If
        
           cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBR1;Uid=" & str_user & ";Pwd=" & str_pws & ";ForceTranslation=1 "
           cn.Open
            
           AS400 = True

        End If

    Else
        cn.Close
    End If

End Function

Public Sub LERAS400()

  If AS400(True) = True Then

        Cells.Delete Shift:=xlUp

        str_sql = ""
        str_sql = str_sql & " select * from pfz0439.lfs02h  "
                                                                                                   
        rs.Open str_sql, cn
    
        For cnt = 0 To rs.Fields.Count - 1
        Range("A1").Offset(0, cnt) = rs.Fields(cnt).Name   'cnt é aposição do campo dentro do record set
        Next
        
        Range("A2").CopyFromRecordset rs
        Range("A3").Activate
        
        'Esta parte de baixo é para dar uma formatadazinha nela para não ficar tão feia
        
        Set tbl = ActiveCell.CurrentRegion
        tbl.AutoFormat

        rs.Close

    Call AS400(False)
  End If

End Sub





