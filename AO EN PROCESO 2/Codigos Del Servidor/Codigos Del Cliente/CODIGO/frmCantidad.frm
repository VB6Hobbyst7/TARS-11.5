VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   1680
   ClientTop       =   4455
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Todo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MouseIcon       =   "frmCantidad.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1035
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&ceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      MouseIcon       =   "frmCantidad.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1035
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la cantidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   585
      TabIndex        =   0
      Top             =   165
      Width           =   2415
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Giosaca

Option Explicit

Private Sub Command1_Click()
frmCantidad.Visible = False
SendData "TI" & Inventario.SelectedItem & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End Sub

Private Sub Command2_Click()

frmCantidad.Visible = False
If Inventario.SelectedItem <> FLAGORO Then
    SendData "TI" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
Else
    SendData "TI" & Inventario.SelectedItem & "," & UserGLD
End If

frmCantidad.Text1.Text = "0"

End Sub

Private Sub Form_Deactivate()
'Unload Me
End Sub

Private Sub text1_Change()
On Error GoTo ErrHandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        If Inventario.SelectedItem <> FLAGORO Or Val(Text1.Text) > UserGLD Then
            Text1.Text = "1"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
