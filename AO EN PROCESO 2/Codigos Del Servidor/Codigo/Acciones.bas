Attribute VB_Name = "Acciones"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio, Jonatan Ezequiel Salguero
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
On Error Resume Next

'¿Posicion valida?
If InMapBounds(Map, x, y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '¿Es un obj?
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
        
        If UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Dios) = "MIFRIT" Or UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Dios) = "POSEIDON" Or UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Dios) = "EREBROS" Or UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Dios) = "TERRASKE" Then
          If UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Dios) = UCase$(UserList(UserIndex).flags.SirvienteDeDios) Then
           If UserList(UserIndex).flags.JerarquiaDios = 5 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "GODS" & UserList(UserIndex).flags.AlmasOfrecidas & "," & GetVar(App.Path & "\Dioses\" & "Configuracion.ini", "INIT", "AlmasNecesarias") * 4 & "," & UserList(UserIndex).flags.SirvienteDeDios)
           Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "GODS" & UserList(UserIndex).flags.AlmasOfrecidas & "," & GetVar(App.Path & "\Dioses\" & "Configuracion.ini", "INIT", "AlmasNecesarias") * UserList(UserIndex).flags.JerarquiaDios & "," & UserList(UserIndex).flags.SirvienteDeDios)
           End If
          Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Esta no es la estatua de tu dios." & FONTTYPE_INFO)
            Exit Sub
          End If
        End If
        
        Select Case ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(Map, x, y, UserIndex)
            Case eOBJType.otCarteles 'Es un cartel
                Call AccionParaCartel(Map, x, y, UserIndex)
            Case eOBJType.otForos 'Foro
                Call AccionParaForo(Map, x, y, UserIndex)
            Case eOBJType.otLeña    'Leña
                If MapData(Map, x, y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                    Call AccionParaRamita(Map, x, y, UserIndex)
                End If
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, y, UserIndex)
            
        End Select
    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, x, y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x, y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, x, y).NpcIndex > 0 Then     'Acciones NPCs
        'Set the target NPC
        UserList(UserIndex).flags.TargetNPC = MapData(Map, x, y).NpcIndex
        
        If Npclist(MapData(Map, x, y).NpcIndex).Comercia = 1 Then
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(UserIndex)
        
        ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = eNPCType.Banquero Then
            If Distancia(Npclist(MapData(Map, x, y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub
            End If
            
                        'A depositar de una
            Call IniciarDeposito(UserIndex)
        
        ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = eNPCType.NpcBargomaud Then
            If Distancia(Npclist(MapData(Map, x, y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del npc." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < 50 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes ser nivel 50.~255~0~0~0~1")
              Exit Sub
            End If
            
            Call WarpUserChar(UserIndex, 1, 50, 54, True)
            
        ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = eNPCType.NpcDioses Then
            If Distancia(Npclist(MapData(Map, x, y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del npc." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes ser nivel 50 + 10.~255~0~0~0~1")
              Exit Sub
            End If
            
            If TieneObjetos(1274, 1, UserIndex) = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya eres sirviente de un dios." & FONTTYPE_INFO)
              Exit Sub
            End If
            
            Dim ElContenedor As Obj
            ElContenedor.ObjIndex = 1274
            ElContenedor.Amount = 1
            
            If Not MeterItemEnInventario(UserIndex, ElContenedor) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Necesitas tener al menos un espacio en el inventario." & FONTTYPE_INFO)
              Exit Sub
            Else
                
                Dim RandomDios As Byte
                RandomDios = RandomNumber(1, 4)
                
                If RandomDios = 1 Then
                    UserList(UserIndex).flags.SirvienteDeDios = "Mifrit"
                ElseIf RandomDios = 2 Then
                    UserList(UserIndex).flags.SirvienteDeDios = "Poseidon"
                ElseIf RandomDios = 3 Then
                    UserList(UserIndex).flags.SirvienteDeDios = "Erebros"
                ElseIf RandomDios = 4 Then
                    UserList(UserIndex).flags.SirvienteDeDios = "Terraske"
                End If
                    
                UserList(UserIndex).flags.JerarquiaDios = 1
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has obtenido un contenedor de las almas." & FONTTYPE_ORO)
            End If
        
        ElseIf Npclist(MapData(Map, x, y).NpcIndex).NPCtype = eNPCType.Revividor Then
            If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, x, y).NpcIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
                Exit Sub
            End If
           
           'Revivimos si es necesario
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call RevivirUsuario(UserIndex)
            End If
            
            'curamos totalmente
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
            Call SendUserStatsBox(UserIndex)
        End If
    Else
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
    End If
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.Map = Map
Pos.x = x
Pos.y = y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
    Exit Sub
End If

'¿Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = ""
        auxcad = ""
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(SendTarget.ToIndex, UserIndex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, x, y) > 2) Then
    If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, x, y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).IndexAbierta
                    
                    Call ModAreas.SendToAreaByPos(Map, x, y, "HO" & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).GrhIndex & "," & x & "," & y)
                     
                    'Desbloquea
                    MapData(Map, x, y).Blocked = 0
                    MapData(Map, x - 1, y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, x, y, 0)
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, x - 1, y, 0)
                    
                      
                    'Sonido
                    SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, x, y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).IndexCerrada
                
                Call ModAreas.SendToAreaByPos(Map, x, y, "HO" & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).GrhIndex & "," & x & "," & y)
                
                
                MapData(Map, x, y).Blocked = 1
                MapData(Map, x - 1, y).Blocked = 1
                
                
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, x - 1, y, 1)
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, x, y, 1)
                
                SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "MCAR" & _
        ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.Map = Map
Pos.x = x
Pos.y = y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, x, y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||En zona segura no puedes hacer fogatas." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(UserIndex).Pos.Map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.Amount = 1
        
        Call SendData(ToIndex, UserIndex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, x, y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.x = x
        Fogatita.y = y
        Call TrashCollector.Add(Fogatita)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||La ley impide realizar fogatas en las ciudades." & FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
    Call SubirSkill(UserIndex, Supervivencia)
End If

End Sub

Sub OtorgarGranPoder(UserIndex As Integer)
Dim LoopC As Integer
Dim EncontroIdeal As Boolean
If LastUser = 0 Then Exit Sub
If UserIndex = 0 Then
    Do While EncontroIdeal = False And LoopC < 500
        LoopC = LoopC + 1
        UserIndex = RandomNumber(1, LastUser)
        If UserList(UserIndex).flags.UserLogged = True And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Privilegios = User Then
            EncontroIdeal = True
            Exit Do
        End If
    Loop
    If Not EncontroIdeal Then
        UserIndex = 0
        GranPoder = 0
    End If
End If
If UserIndex > 0 Then
    If UserList(UserIndex).flags.Muerto <> 0 Then Call OtorgarGranPoder(0)
    GranPoder = UserIndex
    Call SendData(SendTarget.ToAll, UserIndex, 0, "||Los dioses le otorgan el Gran Poder a " & UserList(UserIndex).name & " en el mapa " & UserList(UserIndex).Pos.Map & "." & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
End Sub
