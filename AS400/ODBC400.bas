Attribute VB_Name = "ODBC400"
Public str_user As String
Public str_pws  As String
Public cn       As New ADODB.Connection
Public rs       As New ADODB.Recordset
Public dLin     As Double
Public dNew     As Double

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
        
           cn.ConnectionString = "Provider=IBMDA400;Data Source=GUABRPRD;Force Translate=1;User Id=" & UserForm1.TextBox1.Text & ";Password=" & UserForm1.TextBox2.Text & ""
           cn.Open
            
           AS400 = True

        End If

    Else
        cn.Close
    End If

End Function

Public Sub LERAS400()

  If AS400(True) = True Then
    
     dLin = 2
     dNew = 2
     
     Do While Len(Trim(Sheets("vandas per03").Range("A" & dLin).Value)) > 0
     
        Sheets("Compara").Range("C" & dLin).Value = Sheets("vandas per03").Range("A" & dLin).Value
     
     
     
     
     Loop

    Call AS400(False)
  End If


End Sub
