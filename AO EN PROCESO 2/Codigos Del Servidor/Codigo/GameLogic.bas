Attribute VB_Name = "Extra"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, x, y) Then
    
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).OBJType = eOBJType.otTeleport
    End If
    
    If MapData(Map, x, y).TileExit.Map > 0 Then
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, x, y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                    If nPos.x <> 0 And nPos.y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                Dim veces As Byte
                veces = 0
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.x <> 0 And nPos.y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                If nPos.x <> 0 And nPos.y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Function InRangoVision(ByVal UserIndex As Integer, x As Integer, y As Integer) As Boolean

If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
    If y > UserList(UserIndex).Pos.y - MinYBorder And y < UserList(UserIndex).Pos.y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, x As Integer, y As Integer) As Boolean

If x > Npclist(NpcIndex).Pos.x - MinXBorder And x < Npclist(NpcIndex).Pos.x + MinXBorder Then
    If y > Npclist(NpcIndex).Pos.y - MinYBorder And y < Npclist(NpcIndex).Pos.y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.x, nPos.y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.y - LoopC To Pos.y + LoopC
        For tX = Pos.x - LoopC To Pos.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.x = tX
                nPos.y = tY
                '¿Hay objeto?
                
                tX = Pos.x + LoopC
                tY = Pos.y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
End If

End Sub

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.x, nPos.y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.y - LoopC To Pos.y + LoopC
        For tX = Pos.x - LoopC To Pos.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.x = tX
                nPos.y = tY
                '¿Hay objeto?
                
                tX = Pos.x + LoopC
                tY = Pos.y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
End If

End Sub

Function NameIndex(ByRef name As String) As Integer

Dim UserIndex As Integer
'¿Nombre valido?
If name = "" Then
    NameIndex = 0
    Exit Function
End If

name = UCase$(Replace(name, "+", " "))

UserIndex = 1
Do Until UCase$(UserList(UserIndex).name) = name
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = UserIndex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim UserIndex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = UserIndex

Exit Function

End Function


Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).name) = UCase$(name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

x = Pos.x
y = Pos.y

If Head = eHeading.NORTH Then
    nX = x
    nY = y - 1
End If

If Head = eHeading.SOUTH Then
    nX = x
    nY = y + 1
End If

If Head = eHeading.EAST Then
    nX = x + 1
    nY = y
End If

If Head = eHeading.WEST Then
    nX = x - 1
    nY = y
End If

'Devuelve valores
Pos.x = nX
Pos.y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal PuedeAgua As Boolean = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
            LegalPos = False
Else
  
  If Not PuedeAgua Then
        LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
                   (MapData(Map, x, y).UserIndex = 0) And _
                   (MapData(Map, x, y).NpcIndex = 0) And _
                   (Not HayAgua(Map, x, y))
  Else
        LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
                   (MapData(Map, x, y).UserIndex = 0) And _
                   (MapData(Map, x, y).NpcIndex = 0) And _
                   (HayAgua(Map, x, y))
  End If
   
End If

End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).UserIndex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, x, y)
 Else
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).UserIndex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> eTrigger.POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(SendTarget.toindex, index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim OBJType As Integer

'¿Posicion valida?
If InMapBounds(Map, x, y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = y
    '¿Es un obj?
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = x
        UserList(UserIndex).flags.TargetObjY = y
        FoundSomething = 1
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
        If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & ObjData(UserList(UserIndex).flags.TargetObj).name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.Amount & "" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & ObjData(UserList(UserIndex).flags.TargetObj).name & FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If y + 1 <= YMaxMapSize Then
        If MapData(Map, x, y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).UserIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, x, y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, x, y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, x, y).UserIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, x, y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
            
            If UserList(TempCharIndex).DescRM = "" Then
                If EsNewbie(TempCharIndex) Then
                    Stat = " <Newbie>"
                End If
                
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Sagrada Orden> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Horda Infernal> " & "<" & TituloCaos(TempCharIndex) & ">"
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & ">"
                End If
                
                If Len(UserList(TempCharIndex).Desc) > 1 Then
Stat = "Ves a " & UserList(TempCharIndex).name & Stat & " - " & UserList(TempCharIndex).Desc & " - " & "[" & UserList(TempCharIndex).Clase & " | " & UserList(TempCharIndex).Raza & " |" & " Nivel:" & UserList(TempCharIndex).Stats.ELV & " | "
Else
Stat = "Ves a " & UserList(TempCharIndex).name & Stat & " [" & UserList(TempCharIndex).Clase & " | " & UserList(TempCharIndex).Raza & " |" & " Nivel:" & UserList(TempCharIndex).Stats.ELV & " | "
End If
 
If UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.05) Then
                    Stat = Stat & " Muerto]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.1) Then
                    Stat = Stat & " Casi muerto]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.25) Then
                    Stat = Stat & " Muy Malherido]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.5) Then
                    Stat = Stat & " Malherido]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.75) Then
                    Stat = Stat & " Herido]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP) Then
                    Stat = Stat & " Levemente Herido]"
                Else
                    Stat = Stat & " Intacto]"
                End If
                
                    Dim Alianza As String
                    Dim Horda As String
                If UserList(TempCharIndex).Faccion.RecompensasReal = 0 Then
                    Alianza = "~125~177~230~1~0"
                Else
                    Alianza = "~0~0~255~1~0"
                End If
                If UserList(TempCharIndex).Faccion.RecompensasCaos = 0 Then
                    Horda = "~255~75~75~1~0"
                Else
                    Horda = "~255~0~0~1~0"
                End If
                
                If UserList(TempCharIndex).flags.JerarquiaDios = 1 Then
                    Stat = Stat & " [Sirviente de " & UserList(TempCharIndex).flags.SirvienteDeDios & "]"
                ElseIf UserList(TempCharIndex).flags.JerarquiaDios = 2 Then
                    Stat = Stat & " [Soldado de " & UserList(TempCharIndex).flags.SirvienteDeDios & "]"
                ElseIf UserList(TempCharIndex).flags.JerarquiaDios = 3 Then
                    Stat = Stat & " [Guerrero de " & UserList(TempCharIndex).flags.SirvienteDeDios & "]"
                ElseIf UserList(TempCharIndex).flags.JerarquiaDios = 4 Then
                    Stat = Stat & " [Caballero de " & UserList(TempCharIndex).flags.SirvienteDeDios & "]"
                ElseIf UserList(TempCharIndex).flags.JerarquiaDios = 5 Then
                    Stat = Stat & " [Campeon de " & UserList(TempCharIndex).flags.SirvienteDeDios & "]"
                End If
                
                If UserList(TempCharIndex).flags.PertAlCons > 0 Then
                    Stat = Stat & " [Maestro del Orden]" & FONTTYPE_CONSEJOVesA
                ElseIf UserList(TempCharIndex).flags.PertAlConsCaos > 0 Then
                    Stat = Stat & " [Maestro del Infierno]" & FONTTYPE_CONSEJOCAOSVesA
                Else
                    If UserList(TempCharIndex).flags.Privilegios > 3 Then
                        Stat = Stat & " <Administrador> ~255~255~255~1~0"
                    ElseIf UserList(TempCharIndex).flags.Privilegios > 2 Then
                        Stat = Stat & " <Dios> ~255~255~255~1~0"
                    ElseIf UserList(TempCharIndex).flags.Privilegios > 1 Then
                        Stat = Stat & " <Semi-Dios> ~255~255~255~1~0"
                    ElseIf UserList(TempCharIndex).flags.Privilegios > 0 Then
                        Stat = Stat & " <Consejero> ~255~255~255~1~0"
                        
                    ElseIf UserList(TempCharIndex).Faccion.Alineacion = 1 Then
                        Stat = Stat & " <Criminal>" & Horda
                    ElseIf UserList(TempCharIndex).Faccion.Alineacion = 2 Then
                        Stat = Stat & " <Ciudadano>" & Alianza
                    ElseIf UserList(TempCharIndex).Faccion.Alineacion = 0 Then
                        Stat = Stat & " <Neutral>" & " ~124~124~124~1~0"
                    End If
                End If
            Else
                Stat = UserList(TempCharIndex).DescRM & " " & FONTTYPE_INFOBOLD
            End If
            
            If Len(Stat) > 0 Then _
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & Stat)

            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            If UserList(UserIndex).flags.Privilegios >= PlayerType.SemiDios Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ")"
            Else
                If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                    estatus = "(Dudoso) "
                ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                        estatus = "(Agonizando) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                        estatus = "(Casi muerto) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                        estatus = "(Sano) "
                    Else
                        estatus = "(Intacto) "
                    End If
                ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
                    estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                Else
                    estatus = "!error!"
                End If
            End If
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(UserIndex)
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "|| " & estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "|| " & estatus & Npclist(TempCharIndex).name & "." & FONTTYPE_INFO)
                End If
                
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If


End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim x As Integer
Dim y As Integer

x = Pos.x - Target.x
y = Pos.y - Target.y

'NE
If Sgn(x) = -1 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'NW
If Sgn(x) = 1 And Sgn(y) = 1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SW
If Sgn(x) = 1 And Sgn(y) = -1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SE
If Sgn(x) = -1 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'Sur
If Sgn(x) = 0 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(x) = 0 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(x) = 1 And Sgn(y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(x) = -1 And Sgn(y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(x) = 0 And Sgn(y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
MostrarCantidad = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
