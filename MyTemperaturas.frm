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
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Prompt = "Entre com a temperatura Fahrenheit"
    
    Do
        FTEMP = InputBox(Prompt, "Fahrenheit para Celsius")
        
        If FTEMP <> "" Then
           Celsius = Int((FTEMP + 40) * 5 / 9 - 40) ' Aplicou-se o int aqui para se evitar números muito longos e que saiam da visão
           MsgBox (Celsius), , "Temperatura em Celsius"
        End If
    Loop While FTEMP <> ""
    
End Sub
