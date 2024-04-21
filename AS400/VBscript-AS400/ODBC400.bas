Attribute VB_Name = "ODBC400"
Public str_user As String
Public str_pws  As String
Public cn       As New ADODB.Connection
Public rs       As New ADODB.Recordset

Private Function AS400(ByVal act As Boolean)

    AS400 = True

    If act = True Then

        On Error GoTo msgerro:
        
        If Len(Trim(str_user)) = 0 And _
           Len(Trim(str_pws)) = 0 Then
           UserForm1.Show 1
            
            If Len(Trim(UserForm1.TextBox1.Text)) = 0 or _
               Len(Trim(UserForm1.TextBox2.Text)) = 0 Then
               AS400 = False
               Exit Function
            End If
        
        End If
            
        If Len(Trim(UserForm1.TextBox1.Text)) > 0 And _
           Len(Trim(UserForm1.TextBox2.Text)) > 0 Then
        
           cn.ConnectionString = "Driver={Client Access ODBC Driver (32-bit)};System=PFZBRSEC;Uid=" & UserForm1.TextBox1.Text & ";Pwd=" & UserForm1.TextBox2.Text & ";ForceTranslation=1; timeout = 800"
           cn.Open
    
           str_user = Trim(UserForm1.TextBox1.Text)
           str_pws = Trim(UserForm1.TextBox2.Text)
    
        End If

    Else
        cn.Close
    End If

    Exit Function

msgerro:

    AS400 = False
    MsgBox Err.Number & " - " & Err.Description

End Function

Public Sub LERAS400()

  If AS400(True) = True Then

     







    Call AS400(False)
  End If


End Sub





