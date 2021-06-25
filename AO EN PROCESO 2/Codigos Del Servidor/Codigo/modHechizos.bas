Attribute VB_Name = "modHechizos"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio, Jonatan Ezequiel Salguero
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(Userindex).flags.Invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim da�o As Integer

If Hechizos(Spell).SubeHP = 1 Then

    da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP + da�o
    If UserList(Userindex).Stats.MinHP > UserList(Userindex).Stats.MaxHP Then UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
    
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    Call SendUserStatsBox(val(Userindex))

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(Userindex).flags.Privilegios = PlayerType.User Then
    
    da�o = da�o - Porcentaje(da�o, Int(((UserList(Userindex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(Userindex).Clase)))

    Call SubirSkill(Userindex, Resistencia)
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(Userindex).Invent.HerramientaEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        If da�o < 0 Then da�o = 0
        
        Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP - da�o
        
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
        Call SendUserStatsBox(val(Userindex))
        
        'Muere
        If UserList(Userindex).Stats.MinHP < 1 Then
            UserList(Userindex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (Userindex)
            End If
            
            If Userindex = GranPoder Then
                    Call SendData(SendTarget.ToAll, Userindex, 0, "||" & UserList(Userindex).name & " ha sido asesinado." & FONTTYPE_GUILD)
                    Call OtorgarGranPoder(0)
            End If
            
            Call UserDie(Userindex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(Userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(Userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(Userindex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(Userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(Userindex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(Userindex).flags.Paralizado = 1
          UserList(Userindex).Counters.Paralisis = IntervaloParalizado

#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.ToIndex, Userindex, 0, "PARADOK")
        Else
#End If
            Call SendData(SendTarget.ToIndex, Userindex, 0, "PARADOK")
#If SeguridadAlkon Then
        End If
#End If
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim da�o As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "CFX" & Npclist(TargetNPC).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - da�o
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal Userindex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(Userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal Userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, Userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(Userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(Userindex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(Userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, Userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(Userindex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal Userindex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(Userindex).Char.CharIndex
    Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "||" & vbCyan & "�" & S & "�" & ind)
    Exit Sub
End Sub

Function PuedeLanzar(ByVal Userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(Userindex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(Userindex).flags.TargetMap
    wp2.x = UserList(Userindex).flags.TargetX
    wp2.y = UserList(Userindex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(Userindex).Clase) = "MAGO" Or UCase$(UserList(Userindex).Clase) = "NIGROMANTE" Then
            If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Tu B�culo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||No puedes lanzar este conjuro sin la ayuda de un b�culo." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(Userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(Userindex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(Userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Est�s muy cansado para lanzar este hechizo." & FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||No tenes suficientes puntos de magia para lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call SendData(SendTarget.ToIndex, Userindex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal Userindex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(Userindex).flags.TargetX
    PosCasteadaY = UserList(Userindex).flags.TargetY
    PosCasteadaM = UserList(Userindex).flags.TargetMap
    
    H = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).Userindex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(Userindex)
    End If

End Sub

Sub HechizoInvocacion(ByVal Userindex As Integer, ByRef b As Boolean)

If UserList(Userindex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(Userindex).Pos.Map).Pk = False Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||En zona segura no puedes invocar criaturas." & FONTTYPE_INFO)
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(Userindex).flags.TargetMap
TargetPos.x = UserList(Userindex).flags.TargetX
TargetPos.y = UserList(Userindex).flags.TargetY

H = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(Userindex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(Userindex).NroMacotas = UserList(Userindex).NroMacotas + 1
            
            index = FreeMascotaIndex(Userindex)
            
            UserList(Userindex).MascotasIndex(index) = ind
            UserList(Userindex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = Userindex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(Userindex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal Userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(Userindex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(Userindex, b)
    
End Select

If b Then
    Call SubirSkill(Userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call SendUserStatsBox(Userindex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal Userindex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(Userindex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(Userindex, b)
End Select

If b Then
    Call SubirSkill(Userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call SendUserStatsBox(Userindex)
    Call SendUserStatsBox(UserList(Userindex).flags.TargetUser)
    UserList(Userindex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal Userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(Userindex).flags.TargetNPC, uh, b, Userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(Userindex).flags.TargetNPC, Userindex, b)
End Select

If b Then
    Call SubirSkill(Userindex, Magia)
    UserList(Userindex).flags.TargetNPC = 0
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call SendUserStatsBox(Userindex)
End If

End Sub


Sub LanzarHechizo(index As Integer, Userindex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(Userindex).Stats.UserHechizos(index)

If PuedeLanzar(Userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(Userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.y - UserList(Userindex).Pos.y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(Userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case TargetType.uNPC
            If UserList(Userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.y - UserList(Userindex).Pos.y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(Userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(Userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.y - UserList(Userindex).Pos.y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(Userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            ElseIf UserList(Userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.y - UserList(Userindex).Pos.y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(Userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(Userindex, uh)
    End Select
    
End If

If UserList(Userindex).Counters.Trabajando Then _
    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando - 1

If UserList(Userindex).Counters.Ocultando Then _
    UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal Userindex As Integer, ByRef b As Boolean)



Dim H As Integer, TU As Integer
H = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
TU = UserList(Userindex).flags.TargetUser


If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||�Est� muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Criminal(TU) And Not Criminal(Userindex) Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call VolverCriminal(Userindex)
        End If
    End If
    
    UserList(TU).flags.Invisible = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call InfoHechizo(Userindex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(Userindex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios >= PlayerType.Consejero Then
        Exit Sub
    End If
    
    If UserList(Userindex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(Userindex)
        .CharMimetizado.Body = .Char.Body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.Body = UserList(TU).Char.Body
        .Char.Head = UserList(TU).Char.Head
        .Char.CascoAnim = UserList(TU).Char.CascoAnim
        .Char.ShieldAnim = UserList(TU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(TU).Char.WeaponAnim
    
        Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(Userindex)
   b = True
End If


If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(Userindex, TU) Then Exit Sub
        If Userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(Userindex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(Userindex, TU) Then Exit Sub
        If Userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(Userindex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(Userindex, TU) Then Exit Sub
            
            If Userindex <> TU Then
                Call UsuarioAtacadoPorUsuario(Userindex, TU)
            End If
            
            Call InfoHechizo(Userindex)
            b = True
            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, TU, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToIndex, Userindex, 0, "|| �El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                Call SendCryptedData(SendTarget.ToIndex, TU, 0, "PARADOK")
            Else
#End If
                Call SendData(SendTarget.ToIndex, TU, 0, "PARADOK")
#If SeguridadAlkon Then
            End If
#End If
            
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
        If Criminal(TU) And Not Criminal(Userindex) Then
            If UserList(Userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(Userindex)
            End If
        End If
        
        UserList(TU).flags.Paralizado = 0
        'no need to crypt this
        Call SendData(SendTarget.ToIndex, TU, 0, "PARADOK")
        Call InfoHechizo(Userindex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If Not UserList(TU).flags.Estupidez = 0 Then
                UserList(TU).flags.Estupidez = 0
                'no need to crypt this
                Call SendData(SendTarget.ToIndex, TU, 0, "NESTUP")
                Call InfoHechizo(Userindex)
                b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        If Criminal(TU) And Not Criminal(Userindex) Then
            If UserList(Userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(Userindex)
            End If
        End If

        'revisamos si necesita vara
        If UCase$(UserList(Userindex).Clase) = "MAGO" Or UCase$(UserList(Userindex).Clase) = "NIGROMANTE" Then
            If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Necesitas un mejor b�culo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UCase$(UserList(Userindex).Clase) = "BARDO" Then
            If UserList(Userindex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Necesitas un instrumento m�gico para devolver la vida" & FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        'Pablo Toxic Waste
        UserList(TU).Stats.MinAGU = UserList(TU).Stats.MinAGU - 25
        UserList(TU).Stats.MinHam = UserList(TU).Stats.MinHam - 25
        'Juan Maraxus
        If UserList(TU).Stats.MinAGU <= 0 Then
                UserList(TU).Stats.MinAGU = 0
                UserList(TU).flags.Sed = 1
        End If
        If UserList(TU).Stats.MinHam <= 0 Then
                UserList(TU).Stats.MinHam = 0
                UserList(TU).flags.Hambre = 1
        End If
        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> Userindex Then
                UserList(Userindex).Reputacion.NobleRep = UserList(Userindex).Reputacion.NobleRep + 500
                If UserList(Userindex).Reputacion.NobleRep > MAXREP Then _
                    UserList(Userindex).Reputacion.NobleRep = MAXREP
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||�Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
            End If
        End If
        UserList(TU).Stats.MinMAN = 0
        Call EnviarHambreYsed(TU)
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(Userindex)
        Call RevivirUsuario(TU)
    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(Userindex, TU) Then Exit Sub
        If Userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(Userindex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3
#If SeguridadAlkon Then
        Call SendCryptedData(SendTarget.ToIndex, TU, 0, "CEGU")
#Else
        Call SendData(SendTarget.ToIndex, TU, 0, "CEGU")
#End If
        Call InfoHechizo(Userindex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(Userindex, TU) Then Exit Sub
        If Userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(Userindex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.ToIndex, TU, 0, "DUMB")
        Else
#End If
            Call SendData(SendTarget.ToIndex, TU, 0, "DUMB")
#If SeguridadAlkon Then
        End If
#End If
        Call InfoHechizo(Userindex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal Userindex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(Userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
            Exit Sub
        Else
            UserList(Userindex).Reputacion.NobleRep = 0
            UserList(Userindex).Reputacion.PlebeRep = 0
            UserList(Userindex).Reputacion.AsesinoRep = UserList(Userindex).Reputacion.AsesinoRep + 200
            If UserList(Userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(Userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
        
   Call InfoHechizo(Userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(Userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
            Exit Sub
        Else
            UserList(Userindex).Reputacion.NobleRep = 0
            UserList(Userindex).Reputacion.PlebeRep = 0
            UserList(Userindex).Reputacion.AsesinoRep = UserList(Userindex).Reputacion.AsesinoRep + 200
            If UserList(Userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(Userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
    
    Call InfoHechizo(Userindex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(Userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(Userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(Userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
                Exit Sub
            Else
                UserList(Userindex).Reputacion.NobleRep = 0
                UserList(Userindex).Reputacion.PlebeRep = 0
                UserList(Userindex).Reputacion.AsesinoRep = UserList(Userindex).Reputacion.AsesinoRep + 500
                If UserList(Userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(Userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Call InfoHechizo(Userindex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = Userindex Then
            Call InfoHechizo(Userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(SendTarget.ToIndex, Userindex, 0, "||Este hechizo solo afecta NPCs que tengan amo." & FONTTYPE_WARNING)
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(Userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
                Exit Sub
            Else
                UserList(Userindex).Reputacion.NobleRep = 0
                UserList(Userindex).Reputacion.PlebeRep = 0
                UserList(Userindex).Reputacion.AsesinoRep = UserList(Userindex).Reputacion.AsesinoRep + 500
                If UserList(Userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(Userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(Userindex)
        b = True
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByRef b As Boolean)

Dim da�o As Long


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
    
    Call InfoHechizo(Userindex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + da�o
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Has curado " & da�o & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Npclist(NpcIndex).NPCtype = 2 And UserList(Userindex).flags.Seguro Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Debes sacarte el seguro para atacar guardias del imperio." & FONTTYPE_FIGHT)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(Userindex).Clase) = "MAGO" Or UCase$(UserList(Userindex).Clase) = "NIGROMANTE" Then
            If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                da�o = (da�o * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta da�o segun el staff-
                'Da�o = (Da�o* (80 + BonifB�culo)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 80% del original
            End If
        End If
    End If
    If UserList(Userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If


    Call InfoHechizo(Userindex)
    b = True
    Call NpcAtacado(NpcIndex, Userindex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
    SendData SendTarget.ToIndex, Userindex, 0, "||Le has causado " & da�o & " puntos de da�o a la criatura!" & FONTTYPE_FIGHT
    Call CalcularDarExp(Userindex, NpcIndex, da�o)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, Userindex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal Userindex As Integer)


    Dim H As Integer
    H = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, Userindex)
    
    If UserList(Userindex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CFX" & UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, UserList(Userindex).Pos.Map, "TW" & Hechizos(H).WAV)
    ElseIf UserList(Userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, Npclist(UserList(Userindex).flags.TargetNPC).Pos.Map, "CFX" & Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, UserList(Userindex).Pos.Map, "TW" & Hechizos(H).WAV)
    End If
    
    If UserList(Userindex).flags.TargetUser > 0 Then
        If Userindex <> UserList(Userindex).flags.TargetUser Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(Userindex).flags.TargetUser).name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToIndex, UserList(Userindex).flags.TargetUser, 0, "||" & UserList(Userindex).name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(Userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal Userindex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim da�o As Integer
Dim tempChr As Integer
    
    
H = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
tempChr = UserList(Userindex).flags.TargetUser
      
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| le tiro el hechizo " & H & " a " & UserList(tempChr).Name & FONTTYPE_VENENO)
'End If
      
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(Userindex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + da�o
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has restaurado " & da�o & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(Userindex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - da�o
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has quitado " & da�o & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(Userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + da�o
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If Userindex <> tempChr Then
      Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has restaurado " & da�o & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
      Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call InfoHechizo(Userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - da�o
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has quitado " & da�o & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Not Criminal(Userindex) Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(Userindex, UserList(Userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(Userindex)
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call InfoHechizo(Userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    If Criminal(tempChr) And Not Criminal(Userindex) Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(Userindex, UserList(Userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(Userindex)
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call InfoHechizo(Userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    
    If Criminal(tempChr) And Not Criminal(Userindex) Then
        If UserList(Userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(Userindex, UserList(Userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
    
    Call InfoHechizo(Userindex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + da�o
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If Userindex = tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    da�o = da�o - Porcentaje(da�o, Int(((UserList(tempChr).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(tempChr).Clase)))
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| danio, minhp, maxhp " & da�o & " " & Hechizos(H).MinHP & " " & Hechizos(H).MaxHP & FONTTYPE_VENENO)
'End If
    
    
    da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| da�o, ELV " & da�o & " " & UserList(UserIndex).Stats.ELV & FONTTYPE_VENENO)
'End If
    
    
    If Hechizos(H).StaffAffected Then
        If UCase$(UserList(Userindex).Clase) = "MAGO" Or UCase$(UserList(Userindex).Clase) = "NIGROMANTE" Then
            If GranPoder = Userindex Then da�o = da�o * 2
            If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                If GranPoder = Userindex Then da�o = da�o * 2
                da�o = (da�o * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 70% del original
            End If
        End If
    End If
    
    If UserList(Userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If da�o < 0 Then da�o = 0
    
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call SubirSkill(tempChr, Resistencia)
    Call InfoHechizo(Userindex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - da�o
    
    Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has quitado " & da�o & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_FIGHT)
    Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    If UserList(tempChr).Stats.MinHP < 1 Then
        If tempChr = GranPoder Then
            Call SendData(SendTarget.ToAll, tempChr, 0, "||" & UserList(tempChr).name & " ha sido asesinado." & FONTTYPE_GUILD)
            Call OtorgarGranPoder(Userindex)
        End If
        
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, Userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, Userindex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(Userindex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + da�o
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has restaurado " & da�o & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call InfoHechizo(Userindex)
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has quitado " & da�o & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - da�o
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(Userindex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + da�o
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has restaurado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    End If
    
    If Userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
    End If
    
    Call InfoHechizo(Userindex)
    
    If Userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Le has quitado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||Te has quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - da�o
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(Userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(Userindex, Slot, UserList(Userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(Userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(Userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(Userindex, LoopC, UserList(Userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(Userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(Userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.ToIndex, Userindex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(SendTarget.ToIndex, Userindex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal Userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
        UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(Userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, Userindex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
        UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(Userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, Userindex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, Userindex, CualHechizo)

End Sub


Public Sub DisNobAuBan(ByVal Userindex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'Si estamos en la arena no hacemos nada
    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.x, UserList(Userindex).Pos.y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(Userindex).Reputacion.NobleRep = UserList(Userindex).Reputacion.NobleRep - NoblePts
    If UserList(Userindex).Reputacion.NobleRep < 0 Then
        UserList(Userindex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(Userindex).Reputacion.BandidoRep = UserList(Userindex).Reputacion.BandidoRep + BandidoPts
    If UserList(Userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(Userindex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.ToIndex, Userindex, 0, "PN")
    If Criminal(Userindex) Then If UserList(Userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)
End Sub

Function ResistenciaClase(Clase As String) As Integer

Dim Cuan As Integer

Select Case UCase$(Clase)

    Case "MAGO"

        Cuan = 1

    Case "DRUIDA"

        Cuan = 2

    Case "CLERIGO"

        Cuan = 1

    Case "BARDO"

        Cuan = 1

    Case Else

        Cuan = 0

End Select

ResistenciaClase = Cuan

End Function
