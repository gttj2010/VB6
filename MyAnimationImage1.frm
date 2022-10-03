VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Approaching CD"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   585
      Left            =   1800
      Picture         =   "MyAnimationImage1.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    
    With Image1
        .Height = .Height + 200
        .Width = .Width + 200
    End With
    
End Sub
