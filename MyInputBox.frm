VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InputBox"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
  ' Variáveos do tipo Variant
    Dim Prompt, fullname
    Dim vTitle As String
    
    vTitle = "Sistema de entradas"
    
    Prompt = "Por gentileza, entre com o seu nome."
    fullname = InputBox(Prompt, vTitle) 'InputBox$(Prompt)
    Label1.Caption = fullname
    
End Sub

Private Sub Command2_Click()
    End
End Sub
