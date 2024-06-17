VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sString         As String
Private conn            As Object
Private rs              As Object
Private dResult         As Double

Private sSeq            As String

Public sResult          As String

Public sANO             As Double
Public sMes             As String
Public sData            As String
Public sLancamento      As String
Public sOrigem          As String
Public sValor           As Double
Public sfetivo          As String
Public sOrdem           As String

Public sCartao          As String
Public sReserva         As String

Public sCorretagem      As String
Public sOrdAporte       As String
Public sControle        As String

Public sCaixinha        As String

Public Sub Connection(ByVal sConecta As Boolean)

    If sConecta = True Then
        Dim sPath   As String: sPath = ActiveWorkbook.Path & "\MSikama.db"
        Set conn = CreateObject("ADODB.Connection")
        conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & sPath
    Else
        conn.Close
    End If

End Sub

Public Function Sequencia_Table(ByVal sTabela As String, ByRef sMes As String) As String

     Dim tab_Mes(1 To 12) As String

     Dim dSEQ    As Double: dSEQ = 0
     Dim sPre    As String
     Dim sPer    As String
     
     tab_Mes(1) = "Janeiro"
     tab_Mes(2) = "Fevereiro"
     tab_Mes(3) = "Março"
     tab_Mes(4) = "Abril"
     tab_Mes(5) = "Maio"
     tab_Mes(6) = "Junho"
     tab_Mes(7) = "Julho"
     tab_Mes(8) = "Agosto"
     tab_Mes(9) = "Setembro"
     tab_Mes(10) = "Outubro"
     tab_Mes(11) = "Novembro"
     tab_Mes(12) = "Dezembro"

     For m = 1 To 12
       If tab_Mes(m) = sMes Then
          Exit For
       End If
     Next m

     Set rs = New Recordset

     If IsDate(sMes) = True Then
         sANO2 = Right(Year(sMes), 2)
         sPre = Replace(UCase(sTabela), "TAB_", "")
         sPer = Right("00" & Month(sMes), 2)
         sMes = tab_Mes(Month(sMes))
     Else
 
         sANO2 = Right(sR00_Ano, 2)
 
         If sTabela = "Tab_DC_MC" Or sTabela = "Tab_DC_VS" Then
            sPre = Replace(Replace(UCase(sTabela), "TAB_DC", "D"), "_", "")
            sPer = Right("00" & m, 2)
         Else
            sPre = Replace(UCase(sTabela), "TAB_", "")
            sPer = Right("00" & m, 2)
         End If
     End If

     sKey = sANO2 & sPre & sPer & "."

     sSQL = ""
     sSQL = sSQL & " Select IFNULL(max(cc_Ordem), 0) as sUltimo "
     sSQL = sSQL & " from   " & sTabela & " "
     sSQL = sSQL & " where  cc_Ordem like '" & sKey & "%'"
     sSQL = sSQL & " and    cc_Ordem not like '%.A%' "
     sSQL = sSQL & " and    cc_Ordem not like '%.B%' "
     sSQL = sSQL & " and    cc_Ordem not like '%.C%' "
     sSQL = sSQL & " and    cc_Ordem not like '%.T%' "
    
     rs.Open sSQL, conn
     If Not rs.EOF Then
        dSEQ = CDbl(Replace(rs!sUltimo, sKey, ""))
     End If
     rs.Close

     If dSEQ <> 0 Then
        dSEQ = dSEQ + 1
        Sequencia_Table = sKey & Right("000" & dSEQ, 3)
     Else
        Sequencia_Table = sKey & "001"
     End If

End Function

Public Function Update_Table(ByVal sTabela As String) As Boolean

     On Error GoTo err_Update_Table:

     Update_Table = False

     If Len(Trim(sOrdem)) = 0 Then
        sOrdem = "@@@@@@@@"
     End If

        If InStr(sLancamento, "Cartão Cash Back""") > 0 Then
           sLancamento = Replace(sLancamento, """""", "")
          ' Stop
        End If
        If InStr(sLancamento, "Tarifas") = 0 And sfetivo <> "T" And sfetivo <> "N" Then
           sControle = sCorretagem
        End If

    
     sSQL = ""
     sSQL = sSQL & " Update " & sTabela & " "
     sSQL = sSQL & "    SET cc_Ano        = " & sANO & " "
     sSQL = sSQL & "      , cc_Mes        = '" & sMes & "' "
     sSQL = sSQL & "      , cc_Data       = '" & sData & "' "
     sSQL = sSQL & "      , cc_Lancamento = '" & sLancamento & "' "
     sSQL = sSQL & "      , cc_Origem     = '" & sOrigem & "' "
     sSQL = sSQL & "      , cc_Valor      =  " & Str(Round(sValor, 2)) & " "
     sSQL = sSQL & "      , cc_Efetivo    = '" & UCase(sfetivo) & "' "

     If sTabela = "Tab_DV" Then
        sSQL = sSQL & "      , cc_Reserva    = '" & Trim(sReserva) & "' "
     ElseIf sTabela = "Tab_DC_MC" Or sTabela = "Tab_DC_VS" Then
        sSQL = sSQL & "      , cc_Cartao    = '" & UCase(sCartao) & "' "
        sSQL = sSQL & "      , cc_Reserva    = '" & UCase(sReserva) & "' "
     ElseIf sTabela = "Tab_DI" Then
        sSQL = sSQL & "      , cc_Corretagem    = '" & Trim(sCorretagem) & "' "
     ElseIf sTabela = "Tab_PP" Then
        sSQL = sSQL & "      , cc_Caixinha    = '" & Trim(sCaixinha) & "' "
     End If

     sSQL = sSQL & " Where cc_Ordem       = '" & sOrdem & "' "

     conn.Execute sSQL, dResult
    
     If dResult = 0 Then
    
         If sOrdem = "@@@@@@@@" Then
            sOrdem = Sequencia_Table(sTabela, sMes)
         End If

         sSQL = ""
         sSQL = sSQL & " Insert into " & sTabela & "    "
         sSQL = sSQL & "     (cc_Ano                "
         sSQL = sSQL & "     , cc_Mes               "
         sSQL = sSQL & "     , cc_Data              "
         sSQL = sSQL & "     , cc_Lancamento        "
         sSQL = sSQL & "     , cc_Origem            "
         sSQL = sSQL & "     , cc_Valor             "
         sSQL = sSQL & "     , cc_Efetivo           "

         If sTabela = "Tab_DV" Then
            sSQL = sSQL & "      , cc_Reserva       "
         ElseIf sTabela = "Tab_DC_MC" Or sTabela = "Tab_DC_VS" Then
            sSQL = sSQL & "  , cc_Cartao            "
            sSQL = sSQL & "  , cc_Reserva           "
         ElseIf sTabela = "Tab_DI" Then
            sSQL = sSQL & "  , cc_Corretagem        "
         ElseIf sTabela = "Tab_PP" Then
            sSQL = sSQL & "  , cc_Caixinha          "
         End If
    
         sSQL = sSQL & "     , cc_Ordem)"
         sSQL = sSQL & " values(" & sANO & " "
         sSQL = sSQL & "     , '" & sMes & "' "
         sSQL = sSQL & "     , '" & sData & "' "
         sSQL = sSQL & "     , '" & sLancamento & "' "
         sSQL = sSQL & "     , '" & sOrigem & "' "
         sSQL = sSQL & "     , " & Str(Round(sValor, 2)) & " "
         sSQL = sSQL & "     , '" & UCase(sfetivo) & "' "
    
         If sTabela = "Tab_DV" Then
            sSQL = sSQL & "      , '" & UCase(sReserva) & "' "
         ElseIf sTabela = "Tab_DC_MC" Or sTabela = "Tab_DC_VS" Then
            sSQL = sSQL & "      , '" & sCartao & "' "
            sSQL = sSQL & "      , '" & UCase(sReserva) & "' "
         ElseIf sTabela = "Tab_DI" Then
            sSQL = sSQL & "      , '" & Trim(sCorretagem) & "' "
         ElseIf sTabela = "Tab_PP" Then
            sSQL = sSQL & "      , '" & Trim(sCaixinha) & "' "
         End If
   
         sSQL = sSQL & "     , '" & sOrdem & "')"

         conn.Execute sSQL, dResult
    
         If dResult > 0 Then
            sResult = "Incluído com Sucesso"
            Update_Table = True
         Else
           sResult = "Não Incluído"
           Update_Table = False
           Stop
         End If
    
    Else
      sResult = "Alterado com Sucesso"
      Update_Table = True
    End If
  
    Exit Function

err_Update_Table:

    sResult = Err.Number & "-" & Err.Description
    Update_Table = False

End Function

Public Function Delete_Table(ByVal sTabela As String) As Boolean

     sResp = MsgBox("Deseja excluir esta linha?", vbExclamation + vbYesNo, "")
     If sResp = vbNo Then
        Exit Function
     End If

     If Len(Trim(sOrdem)) = 0 Then
        MsgBox "Linha sem Número de Ordem.", vbInformation + vbOKOnly, ""
        Exit Function
     End If

     On Error GoTo Delete_Table_Error:

     Delete_Table = False

        If InStr(sLancamento, "Cartão Cash Back""") > 0 Then
           sLancamento = Replace(sLancamento, """""", "")
        End If
    
        sSQL = ""
        sSQL = sSQL & " Delete from " & sTabela & " "
        sSQL = sSQL & " Where cc_Ordem = '" & sOrdem & "' "

         conn.Execute sSQL, dResult
          
         If dResult > 0 Then
            sResult = "Excluído com Sucesso"
            Delete_Table = True
         Else
            Stop
         End If
        
         Exit Function

Delete_Table_Error:

    Stop
    Resume


End Function

Public Function Atualiza_Chave_Table(ByVal sTabela As String, ByVal sOrdem As String, ByVal sAntigaOrdem As String) As Boolean

     On Error GoTo err_Update_CR:

     Update_CR = False
    
         sSQL = ""
         sSQL = sSQL & " Update " & sTabela & " "
         sSQL = sSQL & "    SET cc_Ordem        = '" & sOrdem & "' "
         sSQL = sSQL & " Where  cc_Ordem       = '" & sAntigaOrdem & "' "
    
         conn.Execute sSQL, dResult
          
         If dResult > 0 Then
            sResult = "Alterado com Sucesso"
            Update_CR = True
         Else
            Stop
         End If
        
         Exit Function

err_Update_CR:

    sResult = Err.Number & "-" & Err.Description
    Update_CR = False

End Function


Public Function Carrega_RS(ByVal sSQL As String) As Recordset

    Set rs = New Recordset

    rs.Open sSQL, conn

    Set Carrega_RS = rs

End Function



