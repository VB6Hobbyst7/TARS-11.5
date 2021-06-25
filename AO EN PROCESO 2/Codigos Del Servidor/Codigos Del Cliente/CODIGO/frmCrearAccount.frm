VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Cuenta"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox RePass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Mail 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   750
      TabIndex        =   8
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1005
      TabIndex        =   7
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1335
      TabIndex        =   6
      Top             =   1320
      Width           =   360
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Pass <> RePass Then
    MsgBox "Las passwords que tipeo no coinciden", , "Coco rules"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If

If Nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If

Call SendData("NACCNT" & Nombre & "," & Pass & "," & Mail)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
