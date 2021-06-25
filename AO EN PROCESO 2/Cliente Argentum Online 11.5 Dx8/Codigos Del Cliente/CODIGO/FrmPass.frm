VERSION 5.00
Begin VB.Form FrmPass 
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cambiar Contraseña"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Nueva contraseña:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Repetir nueva contraseña:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
