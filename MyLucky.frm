VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "End"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Spin"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image img1 
      Height          =   4095
      Left            =   2880
      Picture         =   "MyLucky.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label lbl4 
      Caption         =   "Lucky Seven"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   3375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click(Index As Integer)
    
    img1.Visible = False  ' Oculta a imagem
    lbl1.Caption = Int(Rnd * 10)  ' Escolhe números
    lbl2(0).Caption = Int(Rnd * 10)
    lbl3(1).Caption = Int(Rnd * 10)
    
    If lbl1.Caption = 7 Or lbl2(0).Caption = 7 Or lbl3(1).Caption = 7 Then
       img1.Visible = True
       Beep
    End If
    
End Sub

Private Sub cmd2_Click(Index As Integer)
    
    End
    
End Sub

Private Sub Form_Load()
    Randomize
End Sub
