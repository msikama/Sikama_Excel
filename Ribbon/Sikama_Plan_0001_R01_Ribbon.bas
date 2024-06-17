Attribute VB_Name = "Sikama_Plan_0001_R01"
Public sR01_Mes As String

Sub Tab_R01_onTabSelected(control As IRibbonControl, ByRef returnedVal)
    returnedVal = bR01
End Sub

'Callback for R01_Grupo_Volta getLabel
Sub Label_R01_Grupo_Volta_getText(control As IRibbonControl, ByRef returnedVal)
   returnedVal = sR00_Ano
End Sub

'=---------------------------------------------=
'= Mês
'=---------------------------------------------=

'Callback for Combo_R01_Mes getText
Sub ComboBox_R01_Mes_getText(control As IRibbonControl, ByRef returnedVal)

    returnedVal = sR01_Mes
    Call Plan_0001.Ver_Filtros

End Sub

'Callback for Combo_R01_Mes onChange
Sub ComboBox_R01_Mes_OnChange(control As IRibbonControl, text As String)

    sR01_Mes = text

    If sR01_Mes = "[Ano Inteiro]" Then
       cSetup.Update_Setup "R01", "Combo_R01_Mes", sR01_Mes, "[Ano Inteiro]"
    Else
       cSetup.Update_Setup "R01", "Combo_R01_Mes", sR01_Mes, Ver_MesD(sR01_Mes)
    End If

    Call Plan_0001.Ver_Filtros

End Sub

'=---------------------------------------------=
'= Saldo
'=---------------------------------------------=

'Callback for R01_Label_01 getLabel
Sub Label_R01_02_getText(control As IRibbonControl, ByRef returnedVal)

    Dim sString As String: sString = Format(CDbl(tab_R01(1)), "###,##0.00")

    dZ01 = Len(sString)

    If sString = "0,00" Then
       dZ01 = 1
    End If

    dZ02 = 22 - dZ01

    sValor = "R$" & Space(dZ02) & sString

    returnedVal = sValor

End Sub

'Callback for R01_Label_04 getLabel
Sub Label_R01_04_getText(control As IRibbonControl, ByRef returnedVal)

    Dim sString As String: sString = Format(CDbl(tab_R01(2)), "###,##0.00")

    dZ01 = Len(sString)

    If sString = "0,00" Then
       dZ01 = 1
    End If

    dZ02 = 22 - dZ01

    sValor = "R$" & Space(dZ02) & sString

    returnedVal = sValor

End Sub

'Callback for R01_Label_06 getLabel
Sub Label_R01_06_getText(control As IRibbonControl, ByRef returnedVal)

    Dim sString As String: sString = Format(CDbl(tab_R01(3)), "###,##0.00")

    dZ01 = Len(sString)

    If sString = "0,00" Then
       dZ01 = 1
    End If

    dZ02 = 22 - dZ01

    sValor = "R$" & Space(dZ02) & sString

    ss = Len(sValor)

    returnedVal = sValor

End Sub

'=---------------------------------------------=
'=---------------------------------------------=

'Callback for button_R01_SQLite_Salva onAction
Sub button_R01_SQLite_Salva_OnAction(control As IRibbonControl)

     Call Plan_0001.Salva_CR

End Sub

'Callback for button_R01_Atualiza onAction
Sub button_R01_Atualiza(control As IRibbonControl)

    Call Plan_0001.Carga_CR
    Call Plan_0001.Ver_Filtros

End Sub

