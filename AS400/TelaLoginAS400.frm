VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AS400"
   ClientHeight    =   1980
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   4710
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "TelaLoginAS400.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
  
   Me.Hide

End Sub

Private Sub CommandButton2_Click()

    TextBox1.Text = ""
    TextBox2.Text = ""

    Me.Hide

End Sub


