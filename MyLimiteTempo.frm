VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Password"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   360
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Entre com a sua senha em 15 segundos"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Text1 = "secret" Then
        Timer1.Enabled = False
        MsgBox "Bem-vindo ao sistema!"
        End
    Else
        MsgBox "Desculpe! Não sei quem é você"
    End If

End Sub

Private Sub Timer1_Timer()
    MsgBox "Desculpe! Seu tempo acabou."
End Sub
