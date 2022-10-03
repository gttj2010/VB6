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
      Caption         =   "Print Form Inteiro"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   3000
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    With Printer
         Printer.Print ""
        .FontName = "Arial"
        .FontSize = 18
        .FontBold = True
         Printer.Print "Mariners"
        .FontBold = False
         Printer.Print Text1
        .EndDoc
    End With
    
End Sub

Private Sub Command2_Click()
    Form1.PrintForm
End Sub
