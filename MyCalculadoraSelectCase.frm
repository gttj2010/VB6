VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operações"
      Height          =   2775
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton optExponenciacao 
         Caption         =   "Exponenciação"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optRestoDivisao 
         Caption         =   "Resto da Divisão"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optDivisao 
         Caption         =   "Divisão"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optMultiplicacao 
         Caption         =   "Multiplicação"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optSubtracao 
         Caption         =   "Subtração"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optAdicao 
         Caption         =   "Adição"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblResultado 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Número 2"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Número 1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
       
    Select Case True
        Case txt1 = "" And txt2 = ""
            GoTo pular
        
        Case (Not IsNumeric(txt1) And txt1 <> "") Or (Not IsNumeric(txt2) And txt2 <> "")
            MsgBox "Atenção!, Não se pode calcular um texto. Favor verificar.", vbInformation
            GoTo pular
            
        Case optAdicao
            vValor = Val(txt1) + Val(txt2)
            
        Case optSubtracao
            vValor = Val(txt1) - Val(txt2)
            
        Case optMultiplicacao
            vValor = Val(txt1) * Val(txt2)
            
        Case optExponenciacao
            vValor = Val(txt1) ^ Val(txt2)
        
        Case optRestoDivisao
            vValor = Val(txt1) Mod Val(txt2)
        
        Case Val(txt2) = 0 Or txt2 = ""
            MsgBox "Atenção!, Não existe nenhum valor dividio por 0. Favor verificar.", vbInformation
            GoTo pular
        
        Case Else
            vValor = Val(txt1) / Val(txt2)
    End Select
    
    lblResultado.Caption = vValor
    
pular:

End Sub

Private Sub Command2_Click()
    txt1 = ""
    txt2 = ""
End Sub

Private Sub Form_Load()
    optAdicao = True
End Sub
