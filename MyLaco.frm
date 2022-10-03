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
   Begin VB.CommandButton Command5 
      Caption         =   "Laço - Do While no Fim = 0"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Laço - Do While no Início < 5"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Laço - For Next Decre."
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Laço - For Next + Step"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Laço - For Next"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    For i = 1 To 10
        FontSize = 10 + i
        Print "Line"; i
    Next i

End Sub

Private Sub Command2_Click()
    
    For i = 1 To 10 Step 2
        FontSize = 10 + i
        Print "Line"; i
    Next i

End Sub

Private Sub Command3_Click()
    
    For i = 10 To 1 Step -1
        FontSize = 10 + i
        Print "Line"; i
    Next i

End Sub

Private Sub Command4_Click()

    i = 0
    
    Do While i < 5
        FontSize = 10 + i
        Print "Line"; i
        i = i + 1
    Loop
    
End Sub

Private Sub Command5_Click()

    i = 0
    
    Do
        FontSize = 10 + i
        Print "Line"; i
        i = i + 1
    Loop While i = 0

End Sub
