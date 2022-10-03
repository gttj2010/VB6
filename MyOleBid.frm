VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OLE OLE3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Class           =   "Paint.Picture"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   4440
      OleObjectBlob   =   "MyOleBid.frx":0000
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OLE OLE2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Class           =   "Excel.SheetBinaryMacroEnabled.12"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   2400
      OleObjectBlob   =   "MyOleBid.frx":1C018
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Class           =   "Word.Document.12"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   360
      OleObjectBlob   =   "MyOleBid.frx":1FA30
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lbl5 
      Caption         =   "Site drawings"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lbl4 
      Caption         =   "Bid Calculator"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lbl3 
      Caption         =   "Estimate scraptchpad"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lbl2 
      Caption         =   "A Construction estimate front end featuring Word, Excel e Paint"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label lbl1 
      Caption         =   "Bid Estimator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5415
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
