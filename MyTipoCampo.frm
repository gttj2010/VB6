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
      Caption         =   "Sair"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tipo de variável"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   Dim vValor As Variant
   Dim vNum, vText, vAlfa As Boolean
   
   vValor = Text1
   
   For i = 1 To Len(vValor)
      If vNum And vText Then vAlfa = True: Exit For
      
      If IsNumeric(Mid(vValor, i, 1)) Then
         vNum = True
      Else
         vText = True
      End If
   Next i
      
   If vNum And vText Then
       Label1.Caption = "Campo alfanumérico"
   ElseIf vNum Then
       Label1.Caption = "Campo numérico"
   Else
       Label1.Caption = "Campo texto"
   End If
   
End Sub

Private Sub Command2_Click()
    End
End Sub
