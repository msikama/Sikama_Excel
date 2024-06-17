Attribute VB_Name = "Sikama_Plan_0000_R00"
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Public MyRibbon         As IRibbonUI
Public cSetup           As SQLite_Setup

Public bR00             As Boolean
Public bR01             As Boolean
Public bR02             As Boolean
Public bR03             As Boolean
Public bR04             As Boolean
Public bR05             As Boolean
Public bR06             As Boolean
Public bR07             As Boolean
Public bR08             As Boolean
Public bR09             As Boolean
Public bR10             As Boolean
Public bR11             As Boolean
Public bR12             As Boolean
Public bR13             As Boolean
Public bA00             As Boolean
Public bA01             As Boolean
Public bA02             As Boolean
Public bA03             As Boolean
Public bB01             As Boolean

Public bR80             As Boolean
Public bR82             As Boolean

Public sR00_Mes         As String
Public sR00_Ano         As Double

Public tab_Mes(1 To 13) As String
Public tab_R01(1 To 3)  As String
Public tab_R02(1 To 6)  As String
Public tab_R03(1 To 6)  As String
Public tab_R04(1 To 9)  As String
Public tab_R05(1 To 12)  As String
Public tab_R06(1 To 9)  As String
Public tab_R08(1 To 9)  As String
Public tab_A00(1 To 9)  As String
Public tab_A01(1 To 9)  As String
Public tab_A02(1 To 14)  As String
Public tab_A03(1 To 9)  As String

Sub onLoad(Ribbon As IRibbonUI)
    
    Set MyRibbon = Ribbon
    Set cSetup = New SQLite_Setup

    Plan_8000.Range("J02").Value = ObjPtr(MyRibbon)

    Atualizacao_Telas (False)

    bFlag = True

    bR00 = True
    bR01 = False
    bR02 = False
    bR03 = False
    bR04 = False
    bR05 = False
    bR06 = False
    bR07 = False
    bR08 = False
    bR09 = False
    bR10 = False
    bR82 = False
    bA00 = False
    bA01 = False
    bA02 = False
    bA03 = False
    bB01 = False

    cSetup.Carga_Inicial

    Plan_0001.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR01_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0002.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR02_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0003.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR03_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0004.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR04_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0005.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR05_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0006.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR06_Mes, Operator:=xlOr, Criteria2:="="
    Plan_0008.Range("C3").CurrentRegion.AutoFilter Field:=1, Criteria1:="=" & sR08_Mes, Operator:=xlOr, Criteria2:="="

   'Ribbon - CR
    tab_R01(1) = "0,00"
    tab_R01(2) = "0,00"
    tab_R01(3) = "0,00"
    
   'Ribbon - DF
    tab_R02(1) = "0,00"
    tab_R02(2) = "0,00"
    tab_R02(3) = "0,00"
    tab_R02(4) = "0,00"
    tab_R02(5) = "0,00"
    tab_R02(6) = "0,00"
 
   'Ribbon - DV
    tab_R03(1) = "0,00"
    tab_R03(2) = "0,00"
    tab_R03(3) = "0,00"
    tab_R03(4) = "0,00"
    tab_R03(5) = "0,00"
    tab_R03(6) = "0,00"
 
   'Ribbon - DI
    tab_R04(1) = "0,00"
    tab_R04(2) = "0,00"
    tab_R04(3) = "0,00"

   'Ribbon - DCB
    tab_R05(1) = "0,00"
    tab_R05(2) = "0,00"
    tab_R05(3) = "0,00"
    tab_R05(4) = "0,00"
    tab_R05(5) = "0,00"
    tab_R05(6) = "0,00"
    tab_R05(7) = "0,00"
    tab_R05(8) = "0,00"
    tab_R05(9) = "0,00"
    tab_R05(10) = "0,00"
    tab_R05(11) = "0,00"
    tab_R05(12) = "0,00"

   'Ribbon - DVS
    tab_R06(1) = "0,00"
    tab_R06(2) = "0,00"
    tab_R06(3) = "0,00"
    tab_R06(4) = "0,00"
    tab_R06(5) = "0,00"
    tab_R06(6) = "0,00"
    tab_R06(7) = "0,00"
    tab_R06(8) = "0,00"
    tab_R06(9) = "0,00"

   'Ribbon - Investimento


   'Ribbon - PicPay

    sR08_Caixa = "Principal"

    If sR00_Ano > 2021 Then
        dMes = Month(Now)
        If dMes = 1 Or dMes = 3 Or dMes = 5 Or dMes = 7 Or dMes = 9 Or dMes = 11 Then
           sR08_Caixa = "Caixinha [Impar]"
        Else
           sR08_Caixa = "Caixinha [Par]"
        End If
        cSetup.Update_Setup "R08", "Combo_R08_Caixinha", sR08_Caixa, "Caixinha Ativa"
    End If

    tab_R08(1) = "0,00"
    tab_R08(2) = "0,00"
    tab_R08(3) = "0,00"
    tab_R08(4) = "0,00"
    tab_R08(5) = "0,00"
    tab_R08(6) = "0,00"
    tab_R08(7) = "0,00"
    tab_R08(8) = "0,00"
    tab_R08(9) = "0,00"

   'Ribbon - Personalitte



   'Ribbon - Aplicação Movimento

    tab_A00(1) = "0000"
    tab_A00(2) = "0,00"
    tab_A00(3) = "0,00"
    tab_A00(4) = "0,00"
    tab_A00(5) = "0,00"
    tab_A00(6) = "0,00"
    tab_A00(7) = "0,00"
    tab_A00(8) = "0,00"
    tab_A00(9) = "0,00"

   'Ribbon - Aplicação Movimento

    tab_A01(1) = "0000"
    tab_A01(2) = "0,00"
    tab_A01(3) = "0,00"
    tab_A01(4) = "0,00"
    tab_A01(5) = "0,00"
    tab_A01(6) = "0,00"

   'Ribbon - Aplicação Movimento

    tab_A02(1) = "0,00"
    tab_A02(2) = "0,00"
    tab_A02(3) = "0,00"
    tab_A02(4) = "0,00"
    tab_A02(5) = "0,00"
    tab_A02(6) = "0,00"
    tab_A02(7) = "0,00"
    tab_A02(8) = "0,00"
    tab_A02(9) = "0,00"
    tab_A02(10) = "0,00"
    tab_A02(11) = "0,00"
    tab_A02(12) = "0,00"
    tab_A02(13) = "0,00"
    tab_A02(14) = "0,00"

    Plan_0000.Select
    Plan_0000.Cells(2, 2).Select

    Atualizacao_Telas (True)
    DoEvents

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    DoEvents

End Sub

'Callback for R00 getVisible
Sub Tab_R00_onTabSelected(control As IRibbonControl, ByRef returnedVal)
    returnedVal = bR00
End Sub

'Callback for button_SV onAction
Sub button_SV_OnAction(control As IRibbonControl)
    
  sResp = MsgBox("Salvar a Planilha", vbQuestion + vbYesNo)

  If sResp = vbYes Then

     Application.DisplayAlerts = False
     ActiveWorkbook.Save

     sResp = MsgBox("Fechar a Planilha", vbQuestion + vbYesNo)
     If sResp = vbYes Then
        Application.Quit
        ActiveWindow.Close
     End If

     Application.DisplayAlerts = True


  End If
    

End Sub

'Callback for button_CR onAction
Sub button_CR_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR01 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R01"

    MyRibbon.ActivateTab "R01"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0001.Select

End Sub

'Callback for button_DF onAction
Sub button_DF_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR02 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R02"

    MyRibbon.ActivateTab "R02"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0002.Select

End Sub

'Callback for button_DV onAction
Sub button_DV_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR03 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R03"

    MyRibbon.ActivateTab "R03"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0003.Select

End Sub

'Callback for button_DI onAction
Sub button_DI_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR04 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R04"

    MyRibbon.ActivateTab "R04"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0004.Select

End Sub

'Callback for button_MC onAction
Sub button_MC_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR05 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R05"

    MyRibbon.ActivateTab "R05"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0005.Select

End Sub

'Callback for button_VC onAction
Sub button_VS_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR06 = True

    MyRibbon.ActivateTab "R06"
    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R06"

    Call Close_Forms

    Atualizacao_Telas (True)
    Plan_0006.Select


End Sub

'Callback for button_IV onAction
Sub button_IV_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR07 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R07"

    Plan_0007.Select
    MyRibbon.ActivateTab "R07"

    Call Close_Forms

    Atualizacao_Telas (True)

End Sub

'Callback for button_IV onAction
Sub button_AP_OnAction(control As IRibbonControl)

    bR00 = False
    bA00 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "A00"
    MyRibbon.ActivateTab "A00"

    Call Close_Forms

    DoEvents

    Plan_0011.Select

End Sub

'Callback for button_PR onAction
Sub button_PR_OnAction(control As IRibbonControl)

    bR00 = False
    bB01 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "B01"
    MyRibbon.ActivateTab "B01"

    Call Close_Forms

    DoEvents

    Plan_0014.Select

End Sub

'Callback for button_R00_Setup onAction
Sub button_R00_Setup(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = False
    bR82 = True

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R82"

    Plan_8000.Select
    MyRibbon.ActivateTab "R82"

    Plan_8000.Select

    Call Close_Forms

    Atualizacao_Telas (True)

End Sub

'Callback for Combo_R00_Mes getText
Sub ComboBox_R00_Mes_getText(control As IRibbonControl, ByRef returnedVal)

    returnedVal = sR00_Mes

End Sub

'Callback for Combo_R00_Mes onChange
Sub ComboBox_R00_Mes_OnChange(control As IRibbonControl, text As String)

    sR00_Mes = text
    cSetup.Update_Mes (sR00_Mes)
    
    Plan_0000.Atualiza_Painel_Anual
    Plan_0000.Atualiza_Painel_Mensal

    Plan_0000.Range("G19").Value = Plan_0000.Range("O11").Value

End Sub

'Callback for Combo_R00_Ano getText
Sub ComboBox_R00_Ano_getText(control As IRibbonControl, ByRef returnedVal)

    returnedVal = sR00_Ano

End Sub

'Callback for Combo_R00_Ano onChange
Sub ComboBox_R00_Ano_OnChange(control As IRibbonControl, text As String)

    sR00_Ano = text

    dResp = MsgBox("Deseja Carregar o ano: " & sR00_Ano & "?", vbQuestion + vbYesNo, "")

    If dResp = vbYes Then
       sR00_Ano = text
       cSetup.Update_Setup "R00", "Combo_R00_Ano", sR00_Ano, "Ano"
       frm_Carrega_Dados_Ano.Show 1
    Else
        cSetup.Retrive_Setup "R00", "Combo_R00_Ano"
        sR00_Ano = CDbl(cSetup.sVal_01)
    End If

    MyRibbon.InvalidateControl "Combo_R00_Mes"

    Call Plan_0000.Atualiza_Painel_Mensal

End Sub

'Callback for button_R01_Volta onAction
Sub button_R00_Volta_OnAction(control As IRibbonControl)

    Atualizacao_Telas (False)

    bR00 = True
    bR01 = False
    bR02 = False
    bR03 = False
    bR04 = False
    bR05 = False
    bR06 = False
    bR07 = False
    bR08 = False
    bR09 = False
    bR10 = False
    bA00 = False
    bA01 = False
    bB01 = False

    MyRibbon.InvalidateControl "R00"
    MyRibbon.InvalidateControl "R01"
    MyRibbon.InvalidateControl "R02"
    MyRibbon.InvalidateControl "R03"
    MyRibbon.InvalidateControl "R04"
    MyRibbon.InvalidateControl "R05"
    MyRibbon.InvalidateControl "R06"
    MyRibbon.InvalidateControl "R07"
    MyRibbon.InvalidateControl "R08"
    MyRibbon.InvalidateControl "R09"
    MyRibbon.InvalidateControl "R11"
    MyRibbon.InvalidateControl "A00"
    MyRibbon.InvalidateControl "A01"
    MyRibbon.InvalidateControl "B01"

    Call Close_Forms

    Plan_0000.Select
    MyRibbon.ActivateTab "R00"

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    DoEvents

    Atualizacao_Telas (True)

End Sub

'Callback for R00_BX_Saldo onAction
Sub DialogBoxLauncher_R00_BX_Diff_OnAction(control As IRibbonControl)

    If bShow = False Then
       frm_Verifica_Juros.Show vbModeless + vbPopup
    Else
       MsgBox "Formulário já ativo", vbOKOnly, ""
    End If
 
End Sub

'Callback for button_R00_Atualiza onAction
Sub button_R00_Atualiza(control As IRibbonControl)
    Plan_0000.Atualiza_Painel_Anual
    Plan_0000.Atualiza_Painel_Mensal
End Sub


'=-----------------------------------------------------------------------------------------=
'=-----------------------------------------------------------------------------------------=

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
        Dim objRibbon As Object
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
End Function
