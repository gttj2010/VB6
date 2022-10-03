VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCloseItem 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "&Clock"
      Begin VB.Menu mnuTimeItem 
         Caption         =   "&Time"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDateItem 
         Caption         =   "&Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTextColorItem 
         Caption         =   "TextCo&lor..."
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCloseItem_Click()
    
    Image1.Picture = LoadPicture("")
    mnuCloseItem.Enabled = False
    
End Sub

Private Sub mnuDateItem_Click()
    
    Label1.Caption = Date
    
End Sub

Private Sub mnuExitItem_Click()
    
    End
    
End Sub

Private Sub mnuOpenItem_Click()
    
    With CommonDialog1
        .Filter = "Metafiles (*.JPG)|*.JPG"
        .ShowOpen
         Image1.Picture = LoadPicture(.FileName)
         mnuCloseItem.Enabled = True
    End With
    
End Sub

Private Sub mnuTextColorItem_Click()

    With CommonDialog1
        .Flags = &H1&
        .ShowColor
         Label1.ForeColor = .Color
    End With
    
End Sub

Private Sub mnuTimeItem_Click()
Label1.Caption = Time
End Sub
