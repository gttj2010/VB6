VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   360
      Pattern         =   "*.jpg*"
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4335
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   360
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    selectedfile = File1.Path & "\" & File1.FileName
    Image1.Picture = LoadPicture(selectedfile)
End Sub
