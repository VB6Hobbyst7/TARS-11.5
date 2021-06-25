Attribute VB_Name = "General"
'Argentum Online 0.11.5

Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader

Sub DarCuerpoDesnudo(ByVal Userindex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(Userindex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(Userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 21
                    Else
                        UserList(Userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 39
                    Else
                        UserList(Userindex).Char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(Userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 32
                    Else
                        UserList(Userindex).Char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 40
                    Else
                        UserList(Userindex).Char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(Userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 53
                    Else
                        UserList(Userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 60
                    Else
                        UserList(Userindex).Char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(Userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 53
                    Else
                        UserList(Userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 60
                    Else
                        UserList(Userindex).Char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(Userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 21
                    Else
                        UserList(Userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(Userindex).CharMimetizado.Body = 39
                    Else
                        UserList(Userindex).Char.Body = 39
                    End If
      End Select
    
End Select

UserList(Userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal x As Integer, ByVal y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & x & "," & y & "," & b)

End Sub


Function HayAgua(Map As Integer, x As Integer, y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
    If MapData(Map, x, y).Graphic(1) >= 1505 And _
       MapData(Map, x, y).Graphic(1) <= 1520 And _
       MapData(Map, x, y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function




Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer


For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(SendTarget.ToMap, 0, d.Map, 1, d.Map, d.x, d.y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista



End Sub

Sub EnviarSpawnList(ByVal Userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.ToIndex, Userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar
Call CargarPremiosList

Prision.Map = 66
Libertad.Map = 66

Prision.x = 75
Prision.y = 47
Libertad.x = 75
Libertad.y = 65


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")



ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty
ReDim Guilds(1 To MAX_GUILDS) As clsClan



IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


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
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"
ListaClases(18) = "Nigromante"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Resistencia Magica"

frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList
Call CargarForbidenWords
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tLluvia.Enabled = True
    .tPiqueteC.Enabled = True
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .tLluviaEvent.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(file, FileType) <> ""
End Function

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
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

Public Function Tilde(Data As String) As String
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub

Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
Cifra = str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
End Function

Public Function Intemperie(ByVal Userindex As Integer) As Boolean
    
    If MapInfo(UserList(Userindex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger <> 1 And _
           MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger <> 2 And _
           MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub EfectoLluvia(ByVal Userindex As Integer)
End Sub


Public Sub TiempoInvocacion(ByVal Userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(Userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(Userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal Userindex As Integer)

Dim modifi As Integer

If UserList(Userindex).Counters.Frio < IntervaloFrio Then
  UserList(Userindex).Counters.Frio = UserList(Userindex).Counters.Frio + 1
Else
  If MapInfo(UserList(Userindex).Pos.Map).Terreno = Nieve Then
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||¡¡Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
    modifi = Porcentaje(UserList(Userindex).Stats.MaxHP, 5)
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP - modifi
    If UserList(Userindex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||¡¡Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(Userindex).Stats.MinHP = 0
            
            If Userindex = GranPoder Then
                Call SendData(SendTarget.ToAll, Userindex, 0, "||" & UserList(Userindex).name & " ha muerto." & FONTTYPE_GUILD)
                Call OtorgarGranPoder(0)
            End If
            
            Call UserDie(Userindex)
    End If
    Call SendData(SendTarget.ToIndex, Userindex, 0, "ASH" & UserList(Userindex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(Userindex).Stats.MaxSta, 5)
    Call QuitarSta(Userindex, modifi)
    Call SendData(SendTarget.ToIndex, Userindex, 0, "ASS" & UserList(Userindex).Stats.MinSta)
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(Userindex).Counters.Frio = 0
  
  
End If

End Sub

Public Sub EfectoMimetismo(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(Userindex).Counters.Mimetismo = UserList(Userindex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(Userindex).Char.Body = UserList(Userindex).CharMimetizado.Body
    UserList(Userindex).Char.Head = UserList(Userindex).CharMimetizado.Head
    UserList(Userindex).Char.CascoAnim = UserList(Userindex).CharMimetizado.CascoAnim
    UserList(Userindex).Char.ShieldAnim = UserList(Userindex).CharMimetizado.ShieldAnim
    UserList(Userindex).Char.WeaponAnim = UserList(Userindex).CharMimetizado.WeaponAnim
        
    
    UserList(Userindex).Counters.Mimetismo = 0
    UserList(Userindex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, Userindex, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
End If
            
End Sub

Public Sub EfectoInvisibilidad(ByVal Userindex As Integer)

Dim TiempoTranscurrido As Long

If UserList(Userindex).Counters.Invisibilidad < IntervaloInvisible Then

     UserList(Userindex).Counters.Invisibilidad = UserList(Userindex).Counters.Invisibilidad + 1

     TiempoTranscurrido = (UserList(Userindex).Counters.Invisibilidad * frmMain.GameTimer.Interval)

     If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then

         If TiempoTranscurrido = 40 Then

             Call SendData(SendTarget.ToIndex, Userindex, 0, "INVI" & ((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000))

         Else

             Call SendData(SendTarget.ToIndex, Userindex, 0, "INVI" & (((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000) - (TiempoTranscurrido / 1000)))

         End If

     End If

Else

     UserList(Userindex).Counters.Invisibilidad = 0

     UserList(Userindex).flags.Invisible = 0

     If UserList(Userindex).flags.Oculto = 0 Then

         Call SendData(SendTarget.ToIndex, Userindex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)

         Call SendData(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, "NOVER" & UserList(Userindex).Char.CharIndex & ",0")

         Call SendData(SendTarget.ToIndex, Userindex, 0, "INVI0")

     End If

End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Ceguera > 0 Then
    UserList(Userindex).Counters.Ceguera = UserList(Userindex).Counters.Ceguera - 1
Else
    If UserList(Userindex).flags.Ceguera = 1 Then
        UserList(Userindex).flags.Ceguera = 0
        Call SendData(SendTarget.ToIndex, Userindex, 0, "NSEGUE")
    End If
    If UserList(Userindex).flags.Estupidez = 1 Then
        UserList(Userindex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, Userindex, 0, "NESTUP")
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal Userindex As Integer)

If UserList(Userindex).Counters.Paralisis > 0 Then
    UserList(Userindex).Counters.Paralisis = UserList(Userindex).Counters.Paralisis - 1
Else
    UserList(Userindex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(SendTarget.ToIndex, Userindex, 0, "PARADOK")
End If

End Sub

Public Sub RecStamina(Userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger = 1 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger = 2 And _
   MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger = 4 Then Exit Sub

Dim massta As Integer
If UserList(Userindex).Stats.MinSta < UserList(Userindex).Stats.MaxSta Then
   If UserList(Userindex).Counters.STACounter < Intervalo Then
       UserList(Userindex).Counters.STACounter = UserList(Userindex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(Userindex).Counters.STACounter = 0
       massta = RandomNumber(1, Porcentaje(UserList(Userindex).Stats.MaxSta, 5))
       UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + massta
       If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then
            UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(Userindex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(Userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(Userindex).Counters.Veneno = UserList(Userindex).Counters.Veneno + 1
Else
  Call SendData(SendTarget.ToIndex, Userindex, 0, "||Estas envenenado, si no te curas moriras." & FONTTYPE_VENENO)
  UserList(Userindex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP - N
  If UserList(Userindex).Stats.MinHP < 1 Then Call UserDie(Userindex)
    If Userindex = GranPoder And UserList(Userindex).Stats.MinHP <= 0 Then
            Call SendData(SendTarget.ToAll, Userindex, 0, "||" & UserList(Userindex).name & " ha muerto." & FONTTYPE_GUILD)
            Call OtorgarGranPoder(0)
    End If
  Call SendData(SendTarget.ToIndex, Userindex, 0, "ASH" & UserList(Userindex).Stats.MinHP)
End If

End Sub

Public Sub DuracionPociones(Userindex As Integer)

'Controla la duracion de las pociones
If UserList(Userindex).flags.DuracionEfecto > 0 Then
   UserList(Userindex).flags.DuracionEfecto = UserList(Userindex).flags.DuracionEfecto - 1
   If UserList(Userindex).flags.DuracionEfecto = 0 Then
        UserList(Userindex).flags.TomoPocion = False
        UserList(Userindex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(Userindex).Stats.UserAtributos(loopX) = UserList(Userindex).Stats.UserAtributosBackUP(loopX)
              Call SendData(ToIndex, Userindex, UserList(Userindex).Pos.Map, "PZ" & UserList(Userindex).Stats.UserAtributos(Fuerza))
              Call SendData(ToIndex, Userindex, UserList(Userindex).Pos.Map, "PM" & UserList(Userindex).Stats.UserAtributos(Agilidad))
        Next
   End If
End If

End Sub

Public Sub HambreYSed(Userindex As Integer, fenviarAyS As Boolean)
'Sed
If UserList(Userindex).Stats.MinAGU > 0 Then
    If UserList(Userindex).Counters.AGUACounter < IntervaloSed Then
          UserList(Userindex).Counters.AGUACounter = UserList(Userindex).Counters.AGUACounter + 1
    Else
          UserList(Userindex).Counters.AGUACounter = 0
          UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU - 10
                            
          If UserList(Userindex).Stats.MinAGU <= 0 Then
               UserList(Userindex).Stats.MinAGU = 0
               UserList(Userindex).flags.Sed = 1
          End If
                            
          fenviarAyS = True
                            
    End If
End If

'hambre
If UserList(Userindex).Stats.MinHam > 0 Then
   If UserList(Userindex).Counters.COMCounter < IntervaloHambre Then
        UserList(Userindex).Counters.COMCounter = UserList(Userindex).Counters.COMCounter + 1
   Else
        UserList(Userindex).Counters.COMCounter = 0
        UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam - 10
        If UserList(Userindex).Stats.MinHam <= 0 Then
               UserList(Userindex).Stats.MinHam = 0
               UserList(Userindex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(Userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub

Sub PasarSegundo()
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1

                Call SendData(SendTarget.ToIndex, i, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, i, 0, "FINOK")
                
                Call CloseSocket(i)
                Exit Sub
            End If
        
        'ANTIEMPOLLOS
        ElseIf UserList(i).flags.EstaEmpo = 1 Then
             UserList(i).EmpoCont = UserList(i).EmpoCont + 1
             If UserList(i).EmpoCont = 30 Then
                 
                 'If FileExist(CharPath & UserList(Z).Name & ".chr", vbNormal) Then
                 'esto siempre existe! sino no estaria logueado ;p
                 
                 'TmpP = val(GetVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant"))
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant", TmpP + 1)
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "P" & TmpP + 1, LCase$(UserList(Z).Name) & ": CARCEL " & 30 & "m, MOTIVO: Empollando" & " " & Date & " " & Time)

                 'Call Encarcelar(Z, 30, "El sistema anti empollo")
                 Call SendData(SendTarget.ToIndex, i, 0, "!! Fuiste expulsado por permanecer muerto sobre un item")
                 'Call SendData(SendTarget.ToAdmins, Z, 0, "|| " & UserList(Z).Name & " Fue encarcelado por empollar" & FONTTYPE_INFO)
                 UserList(i).EmpoCont = 0
                 Call CloseSocket(i)
                 Exit Sub
             ElseIf UserList(i).EmpoCont = 15 Then
                 Call SendData(SendTarget.ToIndex, i, 0, "|| LLevas 15 segundos bloqueando el item, muévete o serás desconectado." & FONTTYPE_WARNING)
             End If
         End If
    Next i
    
    'revisamos auto reiniciares
'    If IntervaloAutoReiniciar <> -1 Then
'        IntervaloAutoReiniciar = IntervaloAutoReiniciar - 1
'
'        If IntervaloAutoReiniciar <= 1200 Then
'            Select Case IntervaloAutoReiniciar
'
'                Case 1200, 600, 240, 120, 180, 60, 30
'                    Call SendData(SendTarget.ToAll, 0, 0, "|| Servidor> El servidor se reiniciará por mantenimiento automático en " & IntervaloAutoReiniciar & " segundos. Tomen las debidas precauciones" & FONTTYPE_SERVER)
'                Case 300
'                    Call SendData(SendTarget.ToAll, 0, 0, "!! El servidor se reiniciará por mantenimiento automático en " & IntervaloAutoReiniciar & " segundos. Tomen las debidas precauciones")
'                Case Is < 30
'                    Call SendData(SendTarget.ToAll, 0, 0, "|| Servidor> El servidor se reiniciará en " & IntervaloAutoReiniciar & " segundos." & FONTTYPE_TALK)
'            End Select
'
'            If IntervaloAutoReiniciar = 0 Then
'                Call ReiniciarServidor(True)
'            End If
'        End If
'    End If
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Grabando Personajes" & FONTTYPE_SERVER)
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Personajes Grabados" & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Function SendFriendList(ByVal Userindex As Integer) As String
Dim tStr As String
Dim tIntx As Integer
Dim CantAmigos As Byte
CantAmigos = 20
    tStr = CantAmigos & ","
    For tIntx = 1 To CantAmigos
      If NameIndex(GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A" & tIntx)) <= 0 Then
        tStr = tStr & GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A" & tIntx) & "(OFF),"
      Else
        tStr = tStr & GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A" & tIntx) & "(ON),"
      End If
    Next tIntx
    SendFriendList = tStr
End Function

Sub ListAmigosON(ByVal Userindex As Integer, ByVal Amigo As Integer)
Dim tStr As String
Dim nombresitoamigo As String
nombresitoamigo = UserList(Amigo).name
     nombresitoamigo = UCase(nombresitoamigo)
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A1") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A2") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A3") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A4") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A5") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A6") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A7") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A8") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A9") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A10") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A11") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A12") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A13") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A14") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A15") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A16") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A17") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A18") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A19") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A20") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha conectado." & FONTTYPE_VENENOXX)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
End Sub

Sub ListAmigosOFF(ByVal Userindex As Integer, ByVal Amigo As Integer)
Dim tStr As String
Dim nombresitoamigo As String
nombresitoamigo = UserList(Amigo).name
nombresitoamigo = UCase(nombresitoamigo)
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A1") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A2") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A3") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A4") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A5") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A6") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A7") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A8") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A9") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A10") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A11") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A12") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A13") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A14") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A15") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A16") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A17") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A18") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A19") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
     If nombresitoamigo = GetVar(App.Path & "\Accounts\" & UserList(Userindex).Accounted & ".act", "AMIGOS", "A20") Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & nombresitoamigo & " se ha desconectado." & FONTTYPE_FRIENDOFF)
            tStr = SendFriendList(Userindex)
            Call SendData(SendTarget.ToIndex, Userindex, 0, "ALS" & SendFriendList(Userindex))
        Exit Sub
     End If
End Sub

Public Sub SwapObjects(ByVal Userindex As Integer)

Dim tmpUserObj As UserOBJ

    With UserList(Userindex)

        'Cambiamos si alguno es una herramienta

        If .Invent.HerramientaEqpSlot = ObjSlot1 Then

            .Invent.HerramientaEqpSlot = ObjSlot2

        ElseIf .Invent.HerramientaEqpSlot = ObjSlot2 Then

            .Invent.HerramientaEqpSlot = ObjSlot1

        End If

        'Cambiamos si alguno es un armor

        If .Invent.ArmourEqpSlot = ObjSlot1 Then

            .Invent.ArmourEqpSlot = ObjSlot2

        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then

            .Invent.ArmourEqpSlot = ObjSlot1

        End If

        'Cambiamos si alguno es un barco

        If .Invent.BarcoSlot = ObjSlot1 Then

            .Invent.BarcoSlot = ObjSlot2

        ElseIf .Invent.BarcoSlot = ObjSlot2 Then

            .Invent.BarcoSlot = ObjSlot1

        End If

        'Cambiamos si alguno es un casco

        If .Invent.CascoEqpSlot = ObjSlot1 Then

            .Invent.CascoEqpSlot = ObjSlot2

        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then

            .Invent.CascoEqpSlot = ObjSlot1

        End If

        'Cambiamos si alguno es un escudo

        If .Invent.EscudoEqpSlot = ObjSlot1 Then

            .Invent.EscudoEqpSlot = ObjSlot2

        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then

            .Invent.EscudoEqpSlot = ObjSlot1

        End If

        'Cambiamos si alguno es munición

        If .Invent.MunicionEqpSlot = ObjSlot1 Then

            .Invent.MunicionEqpSlot = ObjSlot2

        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then

            .Invent.MunicionEqpSlot = ObjSlot1

        End If

        'Cambiamos si alguno es un arma

        If .Invent.WeaponEqpSlot = ObjSlot1 Then

            .Invent.WeaponEqpSlot = ObjSlot2

        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then

            .Invent.WeaponEqpSlot = ObjSlot1

        End If

        'Hacemos el intercambio propiamente dicho

        tmpUserObj = .Invent.Object(ObjSlot1)

        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)

        .Invent.Object(ObjSlot2) = tmpUserObj

        'Actualizamos los 2 slots que cambiamos solamente

        Call UpdateUserInv(False, Userindex, ObjSlot1)

        Call UpdateUserInv(False, Userindex, ObjSlot2)

    End With

End Sub
