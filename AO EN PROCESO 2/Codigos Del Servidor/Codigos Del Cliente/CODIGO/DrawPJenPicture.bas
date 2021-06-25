Attribute VB_Name = "DrawPJenPicture"

Sub DibujaPJ(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Index As Integer)
On Error Resume Next
Dim iGrhIndex As Integer
If Grh.grhindex <= 0 Then Exit Sub
iGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

Call engine.GrhRenderToHdc(iGrhIndex, frmCuent.PJ(Index).hdc, X, Y, True)

frmCuent.PJ(Index).Refresh

End Sub

Sub dibujamuerto(Index As Integer)

End Sub

Sub DibujarTodo(ByVal Index As Integer, Body As Integer, Head As Integer, Casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, Muerto As Integer)

Dim Grh As Grh
Dim Pos As Integer
Dim loopc As Integer

Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer

frmCuent.Nombre(Index).Caption = Nombre

frmCuent.Label1(Index).font = frmMain.font
frmCuent.Label1(Index).font = frmMain.font

frmCuent.Label1(Index).Caption = LVL
frmCuent.Label2(Index).Caption = Clase

XBody = 12
YBody = 15
BBody = 17

If Muerto = 1 Then
    Body = 8
    Head = 500
    Arma = 2
    Shield = 2
    Weapon = 2
    XBody = 10
    YBody = 35
    BBody = 16
    Call dibujamuerto(Index)
End If

Grh = BodyData(Body).Walk(3)
    
Call DibujaPJ(Grh, XBody, YBody, Index)

If Muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y
If Muerto = 1 Then YYY = -9

Pos = YYY + GrhData(GrhData(Grh.grhindex).Frames(Grh.FrameCounter)).pixelHeight
Grh = HeadData(Head).Head(3)
    
Call DibujaPJ(Grh, BBody, Pos, Index)

If Casco <> 2 And Casco > 0 Then
Call DibujaPJ(CascoAnimData(Casco).Head(3), BBody, Pos, Index)

End If

If Weapon <> 2 And Weapon > 0 Then
Call DibujaPJ(WeaponAnimData(Weapon).WeaponWalk(3), XBody, BBody, Index)
End If

If Shield <> 2 And Shield > 0 Then
Call DibujaPJ(ShieldAnimData(Shield).ShieldWalk(3), XBody, BBody, Index)
End If

End Sub




