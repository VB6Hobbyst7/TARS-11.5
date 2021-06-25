VERSION 5.00
Begin VB.Form frmCuent 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   8130
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "BorrarPJ"
      Height          =   615
      Left            =   4580
      TabIndex        =   37
      Top             =   7440
      Width           =   2150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deslogear"
      Height          =   615
      Left            =   6810
      TabIndex        =   36
      Top             =   7440
      Width           =   2150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   7440
      Width           =   2150
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Crear PJ"
      Height          =   615
      Left            =   2350
      TabIndex        =   33
      Top             =   7440
      Width           =   2150
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   7720
      MouseIcon       =   "frmCuent.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":1994
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   32
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   500
      MouseIcon       =   "frmCuent.frx":1C2F
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":28F9
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   31
      Top             =   5280
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   7720
      MouseIcon       =   "frmCuent.frx":2B94
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":385E
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   30
      Top             =   5280
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   500
      MouseIcon       =   "frmCuent.frx":3AF9
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":47C3
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   5760
      MouseIcon       =   "frmCuent.frx":4A5E
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":5728
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   2520
      MouseIcon       =   "frmCuent.frx":59C3
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":668D
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   5760
      MouseIcon       =   "frmCuent.frx":6928
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":75F2
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   2520
      MouseIcon       =   "frmCuent.frx":788D
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":8557
      ScaleHeight     =   1155
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   29
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   28
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   27
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   26
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   24
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   23
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   22
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   21
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   19
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   5250
      TabIndex        =   18
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5250
      TabIndex        =   16
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   3
      Left            =   5250
      TabIndex        =   14
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   13
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5250
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   2010
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   0
      Left            =   2010
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   5250
      TabIndex        =   8
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   5250
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   2010
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub

Private Sub Command4_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub Form_Load()
Dim i As Integer
Label3.Caption = UserName
End Sub
Private Sub Command2_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub


Private Sub nombre_dblClick(Index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub
Private Sub nombre_Click(Index As Integer)
PJClickeado = frmCuent.Nombre(Index).Caption
End Sub
Private Sub PJ_Click(Index As Integer)
PJClickeado = frmCuent.Nombre(Index).Caption
End Sub

Private Sub PJ_dblClick(Index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub


