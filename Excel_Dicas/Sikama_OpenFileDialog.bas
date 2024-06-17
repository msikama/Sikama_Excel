Attribute VB_Name = "Sikama_OpenFileDialog"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, sid As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long

'
''=*************************************************************************************************=
''=-------------------------------------------------------------------------------------------------=
''=-------------------------------------------------------------------------------------------------=
''=*************************************************************************************************=
'
Public Sub SelecionaFile()

    str_file = OpenFileDialog()
    If Len(Trim(str_file)) = 0 Then
       Exit Sub
    End If

    MsgBox "O arquivo selecionado foi:" & vbCrLf & vbCrLf & str_file

End Sub

'
''=*************************************************************************************************=
''=-------------------------------------------------------------------------------------------------=
''=-------------------------------------------------------------------------------------------------=
''=*************************************************************************************************=
'
Public Function OpenFileDialog() As String

  Dim Filter As String, Title As String
  Dim FilterIndex As Integer
  Dim FileName As Variant

  Filter = "Arquivos ServiceCenter (*.csv),*.csv,"

  FilterIndex = 3

  Title = "Selecione um arquivo"

  ChDir ("C:\")

  With Application

    FileName = .GetOpenFilename(Filter, FilterIndex, Title)

  End With

  If FileName = False Then

    MsgBox "Nenhum arquivo foi selecionado."
    Exit Function

  End If

  ' Retorna o caminho do arquivo

  OpenFileDialog = FileName

End Function


Private Function GetFolder() As String

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:

    GetFolder = sItem
    Set fldr = Nothing

End Function
