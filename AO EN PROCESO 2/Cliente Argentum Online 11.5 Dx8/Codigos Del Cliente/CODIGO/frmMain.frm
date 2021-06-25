VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8640
   ClientLeft      =   360
   ClientTop       =   315
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   3960
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   9
      Left            =   5370
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   42
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   8
      Left            =   4800
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   7
      Left            =   4200
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   40
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   3600
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   39
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   2970
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1830
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1230
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   660
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   8400
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   90
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   8400
      Width           =   480
   End
   Begin VB.ListBox ChatContacts 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   1425
      Left            =   6660
      TabIndex        =   25
      Top             =   75
      Width           =   1560
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   3000
      Top             =   2040
   End
   Begin VB.TextBox SendCMSTXT 
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   45
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1575
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   2520
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Trabajo 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   2040
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   2040
      Top             =   2040
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7560
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8985
      Left            =   8415
      Picture         =   "frmMain.frx":6C214
      ScaleHeight     =   599
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   2
      Top             =   0
      Width           =   3585
      Begin VB.CommandButton DespInv 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   525
         MouseIcon       =   "frmMain.frx":83977
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   495
         MouseIcon       =   "frmMain.frx":83AC9
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4380
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2640
         Left            =   465
         ScaleHeight     =   174
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   168
         TabIndex        =   8
         Top             =   1830
         Width           =   2550
      End
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2790
         Left            =   450
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1860
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Oro"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   2670
         TabIndex        =   51
         Top             =   5700
         Width           =   495
      End
      Begin VB.Label lblClanes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clanes"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2205
         TabIndex        =   50
         Top             =   7890
         Width           =   1125
      End
      Begin VB.Label Image1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Estadisticas"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2145
         TabIndex        =   49
         Top             =   7440
         Width           =   1275
      End
      Begin VB.Label Image1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2175
         TabIndex        =   48
         Top             =   6960
         Width           =   1155
      End
      Begin VB.Label cmdInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Top             =   4800
         Width           =   675
      End
      Begin VB.Label CmdLanzar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LANZAR"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   210
         TabIndex        =   46
         Top             =   4770
         Width           =   1725
      End
      Begin VB.Label Menu 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   150
         Left            =   1050
         TabIndex        =   33
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label ItemName 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Nada"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   300
         TabIndex        =   30
         Top             =   5310
         Width           =   3015
      End
      Begin VB.Label Canjes 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canjes"
         ForeColor       =   &H0000C0C0&
         Height          =   300
         Left            =   720
         TabIndex        =   24
         Top             =   3480
         Width           =   2130
      End
      Begin VB.Label AguBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   7515
         Width           =   1455
      End
      Begin VB.Label HamBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   7170
         Width           =   1455
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   315
         TabIndex        =   21
         Top             =   6825
         Width           =   1290
      End
      Begin VB.Label ManaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   315
         TabIndex        =   20
         Top             =   6510
         Width           =   1290
      End
      Begin VB.Label StaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   315
         TabIndex        =   19
         Top             =   6165
         Width           =   1290
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   315
         Top             =   7575
         Width           =   1290
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   315
         Top             =   7245
         Width           =   1290
      End
      Begin VB.Shape Hpshp 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   315
         Top             =   6900
         Width           =   1290
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   315
         Top             =   6585
         Width           =   1290
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H0000FFFF&
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   315
         Top             =   6240
         Width           =   1290
      End
      Begin VB.Label lblPorcLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0%"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1845
         TabIndex        =   17
         Top             =   705
         Width           =   450
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   0
         Left            =   2985
         MouseIcon       =   "frmMain.frx":83C1B
         MousePointer    =   99  'Custom
         Top             =   2100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   1
         Left            =   2985
         MouseIcon       =   "frmMain.frx":83D6D
         MousePointer    =   99  'Custom
         Top             =   2520
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "   "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1065
         TabIndex        =   14
         Top             =   435
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label exp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia:"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   285
         TabIndex        =   13
         Top             =   675
         Width           =   1065
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   2
         Left            =   2040
         Top             =   6540
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   1
         Left            =   2040
         Top             =   6255
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   0
         Left            =   2040
         Top             =   5955
         Width           =   360
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2595
         TabIndex        =   12
         Top             =   5970
         Width           =   105
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   420
         TabIndex        =   11
         Top             =   60
         Width           =   2625
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HECHIZOS"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   480
         Left            =   1860
         MouseIcon       =   "frmMain.frx":83EBF
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1410
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIO"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   450
         Left            =   420
         MouseIcon       =   "frmMain.frx":84011
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1380
         Width           =   1350
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   5
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   735
         TabIndex        =   4
         Top             =   375
         Width           =   105
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   285
         TabIndex        =   3
         Top             =   435
         Width           =   465
      End
   End
   Begin VB.Timer Attack 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   2040
   End
   Begin VB.TextBox SendTxt 
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   45
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1575
      Visible         =   0   'False
      Width           =   7815
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1410
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   105
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   2487
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":84163
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   60
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   543
      TabIndex        =   18
      Top             =   1920
      Width           =   8175
      Begin VB.Label NombreMapa 
         BackStyle       =   0  'Transparent
         Caption         =   "Bargomaud"
         BeginProperty Font 
            Name            =   "Morpheus"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   1170
         Left            =   6585
         TabIndex        =   43
         Top             =   5850
         Width           =   3570
      End
   End
   Begin VB.Label lblMinimizar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MINIMIZAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6030
      TabIndex        =   45
      Top             =   8460
      Width           =   1065
   End
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   7335
      TabIndex        =   44
      Top             =   8460
      Width           =   645
   End
   Begin VB.Label H 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   7875
      TabIndex        =   31
      Top             =   1575
      Width           =   345
   End
   Begin VB.Label Casco 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   11250
      TabIndex        =   29
      Top             =   8550
      Width           =   615
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   10530
      TabIndex        =   28
      Top             =   8550
      Width           =   615
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   9750
      TabIndex        =   27
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Armadura 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   9000
      TabIndex        =   26
      Top             =   8520
      Width           =   615
   End
   Begin VB.Menu cmdA 
      Caption         =   "Amigos"
      Visible         =   0   'False
      Begin VB.Menu cmdAddC 
         Caption         =   "Agregar contacto"
      End
      Begin VB.Menu cmdElimC 
         Caption         =   "Eliminar contacto"
      End
      Begin VB.Menu cmdChat 
         Caption         =   "Iniciar chat"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu cmdH 
      Caption         =   "Modo de Habla"
      Visible         =   0   'False
      Begin VB.Menu cmdNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu cmdG 
         Caption         =   "Gritar"
      End
      Begin VB.Menu cmdClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu cmdDenuncia 
         Caption         =   "Denuncia"
      End
      Begin VB.Menu cmdParty 
         Caption         =   "Party"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Enum H

  cNormal = 0

  Gritar = 1

  Clan = 2

  cDenuncia = 3

  cParty = 4

End Enum

Dim TxtHabla(0 To 4) As String

Dim ModoHabla As Integer

Private Sub Canjes_Click()

SendData ("CCANJE")

End Sub

Private Sub cmdClan_Click()

ModoHabla = 2

End Sub

Private Sub cmdG_Click()

ModoHabla = 1

End Sub

Private Sub cmdDenuncia_Click()

ModoHabla = 3

'Call AddtoRichTextBox(RecTxt, "Modo de Habla: /RMSG" & vbCrLf, 255, 255, 255, True, False, False)

End Sub

Private Sub cmdNormal_Click()

ModoHabla = 0

End Sub

Private Sub cmdParty_Click()

ModoHabla = 4

End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub FPS_Timer()

If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub



Private Sub lblMapaName_Click()

End Sub

Private Sub h_Click()
PopUpMenu cmdH
End Sub

Private Sub Label2_Click()
PopUpMenu cmdH
frmOpciones.Show
End Sub

Private Sub InvEqu_Click()

End Sub

Private Sub lblClanes_Click()
frmGuildAdm.Show
End Sub

Private Sub lblEstadisticas_Click()
frmEstadisticas.Show
End Sub

Private Sub lblOpciones_Click()
frmOpciones.Show
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lblSalir_Click()
    If MsgBox("¿Estas seguro que quieres salir?", vbYesNo) = vbYes Then
        Call SendData("/SALIR")
    End If
End Sub

Private Sub Menu_Click()
'InvEqu.Picture = LoadPicture(App.path & "\Recursos\Graficos\Menu.jpg")

ItemName.Visible = False

picInv.Visible = False

hlst.Visible = False

cmdInfo.Visible = False

CmdLanzar.Visible = False

cmdMoverHechi(0).Visible = False

cmdMoverHechi(1).Visible = False

End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicAU_Click()
    AddtoRichTextBox frmMain.RecTxt, "Hay actualizaciones pendientes. Cierra el juego y ejecuta el autoupdate. (el mismo debe descargarse del sitio oficial http://ao.alkon.com.ar, y deberás conectarte al puerto 7667 con la IP tradicional del juego)", 255, 255, 255, False, False, False
End Sub

Private Sub PicMH_Click()
    AddtoRichTextBox frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
End Sub

Private Sub PicSeg_Click()
    AddtoRichTextBox frmMain.RecTxt, "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
End Sub

Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_DblClick()
Call Form_DblClick
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
   If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     TIMERS                         '
''''''''''''''''''''''''''''''''''''''

Private Sub Trabajo_Timer()
    'NoPuedeUsar = False
End Sub

Private Sub Attack_Timer()
    'UserCanAttack = 1
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("LH" & hlst.ListIndex + 1)
        Call SendData("UK" & Magia)
        UsaMacro = True
        'UserCanAttack = 0
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub Form_Click()

    If Cartel Then Cartel = False
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
        
            Select Case KeyCode
                Case vbKeyM:
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyC:
                    Call SendData("TAB")
                    IScombate = Not IScombate
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyS:
                    Call SendData("/SEG")
                Case vbKeyZ:
                    Call SendData("/SEGCLAN")
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
            End Select
        End If
        
                If KeyCode >= 112 And KeyCode <= 121 Then
            UsarMacro KeyCode - 112
            Exit Sub
        End If
        
        Select Case KeyCode
            Case vbKeyReturn:
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyDelete:
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
            Case vbKeyF4:
            Call SendData("/SALIR")
            Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                        'If IScombate Then
                        ''[ANIM ATAK]
                        'charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
                        'charlist(UserCharIndex).Arma.WeaponAttack = GrhData(charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).grhindex).NumFrames + 1
                        'End If
                End If
            Case vbKeyF5:
                Call frmOpciones.Show(vbModeless, frmMain)
            Case vbKeyF7:
                Call SendData("/MEDITAR")
            Case vbKeyF12:
                          
        End Select
        
End Sub

Private Sub Form_Load()
    
    ModoHabla = 1


    
   Me.left = 0
   Me.top = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()

    Call Audio.PlayWave(SND_CLICK)

    'InvEqu.Picture = LoadPicture(App.path & "\Recursos\Graficos\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ItemName.Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    'InvEqu.Picture = LoadPicture(App.path & "\Recursos\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ItemName.Visible = False
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If left$(stxtbuffer, 1) = "/" Then
            If UCase(left$(stxtbuffer, 8)) = "/CAMBIARCONTRASEÑA " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                stxtbuffer = "/CAMBIARCONTRASEÑA " & j
                
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If

           Call SendData(stxtbuffer)
           
ElseIf ModoHabla = H.Gritar Then

            Call SendData("-" & stxtbuffer)

       ' Hablar en Global

       ' ElseIf ModoHabla = h.cGlobal Then

            'Call SendData(":" & stxtbuffer)

        'Whisper

      '  ElseIf ModoHabla = h.cPrivado Then

           ' Call SendData("\" & NamePrivado & " " & stxtbuffer)

        ElseIf ModoHabla = H.Clan Then

            Call SendData("/CMSG " & stxtbuffer & "")

        'Say

        ElseIf ModoHabla = H.cNormal Then

            Call SendData(";" & stxtbuffer)

        'DENUNCIAS

        ElseIf ModoHabla = H.cDenuncia Then

            Call SendData("/DENUNCIAR " & stxtbuffer & "")

            'PARTY

        ElseIf ModoHabla = H.cParty Then

            Call SendData("/PMSG " & stxtbuffer & "")

            End If

        stxtbuffer = ""

        SendTxt.Text = ""

        KeyCode = 0

        SendTxt.Visible = False
        
        End If

End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Second.Enabled = True
    
    Call SendData("gIvEmEvAlcOde")
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0

    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub




'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Second.Enabled = True
    
    Call SendData("gIvEmEvAlcOde")
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Public Sub PonerListaAmigos(ByVal Rdata As String)

 

Dim j As Integer, k As Integer

For j = 0 To ChatContacts.ListCount - 1

    Me.ChatContacts.RemoveItem 0

Next j

k = CInt(ReadField(1, Rdata, 44))

 

For j = 1 To k

    ChatContacts.AddItem ReadField(1 + j, Rdata, 44)

Next j

 

End Sub

Private Sub ChatContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then

PopUpMenu cmdA

End If

End Sub

Private Sub cmdAddC_Click()

Dim nickadd As String

nickadd = InputBox$("Ingrese un nombre", "Agregar Contacto")

Call SendData("ADF" & nickadd & "," & ChatContacts.ListIndex + 1)

Call SendData("ACLIST")

End Sub

Private Sub cmdElimC_Click()

Call SendData("ELF" & ChatContacts.ListIndex + 1)

Call SendData("ACLIST")

End Sub
Private Sub picMacro_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    frmMacro.lblF.Caption = "Macro F" & Index + 1
    frmMacro.numF.Caption = Index + 1
    If Macros(Index).comando <> "" Then
        frmMacro.optComando.value = True
        frmMacro.txtComando.Text = Macros(Index).comando
    ElseIf Macros(Index).Hechizo > 0 Then
        frmMacro.optHechi.value = True
    ElseIf Macros(Index).Equipar > 0 Then
        frmMacro.optEquipar.value = True
    ElseIf Macros(Index).Usar > 0 Then
        frmMacro.optUsar.value = True
    Else
        frmMacro.optComando.value = True
    End If
    frmMacro.Show
ElseIf Button = vbLeftButton Then
    UsarMacro (Index)
End If
   
End Sub
