Attribute VB_Name = "Módulo1"
Public str_user As String
Public str_pws  As String
Public cn       As New ADODB.Connection
Public rs       As New ADODB.Recordset

Public Function AS400(ByVal act As Boolean)

    If act = True Then

        On Error GoTo msgerro:
        UserForm1.Show 1
        
        If Len(Trim(str_user)) > 0 And _
           Len(Trim(str_pws)) > 0 Then
            
            If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
               Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
            
               cn.ConnectionString = "Driver={iSeries Access ODBC Driver};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & " "
               cn.Open
        
               str_user = Trim(UserForm1.TextBox1.Text)
               str_pws = Trim(UserForm1.TextBox2.Text)
        
            End If
        
        End If

    Else
        cn.Close
    End If

End Function


