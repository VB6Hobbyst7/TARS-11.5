Attribute VB_Name = "Mod_General"
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

Public bK As Long
Public bRK As Long


Public iplst As String
Public banners As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long
Public sHKeys() As String
Public Function DirPath(ByVal path As String) As String
'•Parra: Nuevo Engine v2.0
    Select Case path
        Case "Graficos"
            DirPath = App.path & "\Recursos\GRAFICOS\"
            Exit Function
        
        Case "Sound"
            DirPath = App.path & "\Recursos\WAV\"
            Exit Function
        
        Case "Midi"
            DirPath = App.path & "\Recursos\MIDI\"
            Exit Function
        
        Case "Maps"
            DirPath = App.path & "\Recursos\MAPAS\"
            Exit Function
    End Select
End Function

Public Function DirGraficos() As String
    DirGraficos = App.path & "\Recursos\" & "GRAFICOS" & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\Recursos\" & "WAV" & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\Recursos\" & "MIDI" & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\Recursos\" & "MAPAS" & "\"
End Function

Public Function SumaDigitos(ByVal numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (numero Mod 10)
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (numero Mod 10) - 1
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function Complex(ByVal numero As Integer) As Integer
    If numero Mod 2 <> 0 Then
        Complex = numero * SumaDigitos(numero)
    Else
        Complex = numero * SumaDigitosMenos(numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(numero)
    AuxInteger2 = SumaDigitosMenos(numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\Recursos\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.path & "\Recursos\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.path & "\Recursos\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\Recursos\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.y).CharIndex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm

    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini

    'Unload the connect form
    Unload frmConnect
    
    frmMain.Label8.Caption = UserName
    'Load main form
    frmMain.Visible = True
End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.x, UserPos.y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.y)
    End Select
    
    If LegalOk Then
        Call SendData("M" & Direccion)
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            engine.Char_Move_by_Head UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
              
                Call MoveTo(NORTH)
                
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
              
                Call MoveTo(EAST)
                
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
              
                Call MoveTo(SOUTH)
                
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
             
                Call MoveTo(WEST)
            
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0) Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0
            If kp Then Call RandomMove
          
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
    
        Case E_Heading.EAST
            x = 1
    
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

Sub SwitchMap(ByVal map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    handle = FreeFile()
    
    Open DirPath("Maps") & "Mapa" & map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            MapData(x, y).luz = 0
            MapData(x, y).particle_group = 0
            MapData(x, y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(x, y).Graphic(1).grhindex
            InitGrh MapData(x, y).Graphic(1), MapData(x, y).Graphic(1).grhindex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(x, y).Graphic(2).grhindex
                InitGrh MapData(x, y).Graphic(2), MapData(x, y).Graphic(2).grhindex
            Else
                MapData(x, y).Graphic(2).grhindex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(x, y).Graphic(3).grhindex
                InitGrh MapData(x, y).Graphic(3), MapData(x, y).Graphic(3).grhindex
            Else
                MapData(x, y).Graphic(3).grhindex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(x, y).Graphic(4).grhindex
                InitGrh MapData(x, y).Graphic(4), MapData(x, y).Graphic(4).grhindex
            Else
                MapData(x, y).Graphic(4).grhindex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(x, y).Trigger
            Else
                MapData(x, y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, y).CharIndex > 0 Then
                Call EraseChar(MapData(x, y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(x, y).ObjGrh.grhindex = 0
        Next x
    Next y
    
    Close handle
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = map
End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.path & "\Recursos\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
On Error GoTo errorH
    Dim f As String
    Dim c As Integer
    Dim i As Long
    
    f = App.path & "\Recursos\init\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).desc = GetVar(f, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(f, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(f, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(f, "S" & i, "PJ"))
    Next i
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    End
End Sub

Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer

        CurServerPasRecPort = "7667"

End Function

Public Function CurServerIp() As String

        CurServerIp = "127.0.0.1"

End Function

Public Function CurServerPort() As Integer

        CurServerPort = "7666"

End Function


Sub Main()

On Error Resume Next

    Call WriteClientVer

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8

        
    frmCargando.Show
    frmCargando.Refresh
    
    frmMain.Socket1.Startup
        
    Call InicializarNombres
    


UserMap = 1
    
    LoadGrhData
    CargarCabezas
    CargarCascos
    CargarCuerpos
    CargarArrayLluvia
    CargarFxs
    Call engine.Engine_Init
    Call engine.setup_ambient
    Call CargarParticulas
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores

    
    Unload frmCargando
    
    'Inicializamos el sonido
     Call Audio.Initialize(frmMain.hWnd, App.path & "\" & "WAV" & "\", App.path & "\" & "MIDI" & "\")
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    
        Call Audio.PlayMIDI(MIdi_Inicio & ".mid")


    'frmPres.Picture = LoadPicture(App.path & "\Graficos\bosquefinal.jpg")
    'frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
        Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font

    
engine.Start
    
Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.source
    End
End Sub
Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean

    HayAgua = MapData(x, y).Graphic(1).grhindex >= 1505 And _
                MapData(x, y).Graphic(1).grhindex <= 1520 And _
                MapData(x, y).Graphic(2).grhindex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.path & "\Recursos\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.Resistencia) = "Resistencia magica"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub

Public Function CalculateK(ByVal Gold As Long) As String

'**************************************************************

'Author: Standelf

'Last Modify Date: 11/12/2008

'It calculates the value of the currency in AO

'**************************************************************

    If Val(Gold) < 1000000 Then

        CalculateK = "[" & FormatNumber((Val(Gold) / 1000), 2) & "K]"

    Else

        CalculateK = "[" & FormatNumber((Val(Gold) / 1000000), 2) & "KK]"

    End If

End Function
