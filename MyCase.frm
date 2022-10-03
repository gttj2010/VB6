VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Choose a country"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Bem vindo ao programana Internacional"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()

    With List1
        .AddItem "England"
        .AddItem "Germany"
        .AddItem "Spain"
        .AddItem "Italy"
    End With
    
End Sub

Private Sub List1_Click()
    Label3.Caption = List1.Text
    
    Select Case List1.ListIndex
        Case 0
            Label4.Caption = "Hello, programmer"
        Case 1
            Label4.Caption = "Hallo, programmierer"
        Case 2
            Label4.Caption = "Hola, programador"
        Case 3
            Label4.Caption = "Ciao, programmatori"
    End Select
    
End Sub
