Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2009 Juan Andres Dalmasso (CHOTS)
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

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Integer
    y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type


'Lista de cuerpos
Public Type FxData
    fX As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    BodyNum As Integer
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    Clan As String

    'CHOTS | Optimizamos los colores
    redColor As Byte
    greenColor As Byte
    blueColor As Byte
    nameCenter As Long
    clanCenter As Long
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public Userindex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public lastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()
On Error Resume Next
Dim N As Integer, i As Integer, Numheads As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Ropas.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #N, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #N, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #N, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
Next i

Close #N

End Sub
Sub CargarFxs()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #N, , MisFxs(i)
    Call InitGrh(FxData(i).fX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #N

End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, i As Integer
Dim Nu As Integer

N = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #N, , bLluvia(i)
Next i

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)

Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = (0 - (HWindowY - cy))
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.x + cx
tY = UserPos.y + cy

End Sub






Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If charindex > LastChar Then LastChar = charindex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(charindex).iHead = Head
charlist(charindex).iBody = Body
charlist(charindex).Head = HeadData(Head)
charlist(charindex).Body = BodyData(Body)
charlist(charindex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(charindex).Arma.WeaponAttack = 0

charlist(charindex).Escudo = ShieldAnimData(Escudo)
charlist(charindex).Casco = CascoAnimData(Casco)

charlist(charindex).Heading = Heading

'Reset moving stats
charlist(charindex).Moving = 0
charlist(charindex).MoveOffset.x = 0
charlist(charindex).MoveOffset.y = 0

'Update position
charlist(charindex).Pos.x = x
charlist(charindex).Pos.y = y

'Make active
charlist(charindex).Active = 1

'Plot on map
MapData(x, y).charindex = charindex

End Sub

Sub ResetCharInfo(ByVal charindex As Integer)

    charlist(charindex).Active = 0
    charlist(charindex).Criminal = 0
    charlist(charindex).fX = 0
    charlist(charindex).FxLoopTimes = 0
    charlist(charindex).invisible = False

    charlist(charindex).Moving = 0
    charlist(charindex).muerto = False
    charlist(charindex).Nombre = ""
    charlist(charindex).Clan = ""
    charlist(charindex).pie = False
    charlist(charindex).Pos.x = 0
    charlist(charindex).Pos.y = 0
    charlist(charindex).UsandoArma = False

    'CHOTS | Optimizamos colores
    charlist(charindex).redColor = 0
    charlist(charindex).greenColor = 0
    charlist(charindex).blueColor = 0
    charlist(charindex).nameCenter = 0
    charlist(charindex).clanCenter = 0

End Sub


Sub EraseChar(ByVal charindex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(charindex).Active = 0

'Update lastchar
If charindex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0

Call ResetCharInfo(charindex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurría debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal charindex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

x = charlist(charindex).Pos.x
y = charlist(charindex).Pos.y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = x + addX
nY = y + addY

MapData(nX, nY).charindex = charindex
charlist(charindex).Pos.x = nX
charlist(charindex).Pos.y = nY
MapData(x, y).charindex = 0

charlist(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(charindex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Debug.Print UserCharIndex
    Call EraseChar(charindex)
End If

End Sub

Public Sub DoFogataFx()

    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If

End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim x As Integer, y As Integer

For y = UserPos.y - MinYBorder + 1 To UserPos.y + MinYBorder - 1
  For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            
            If MapData(x, y).charindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next x
Next y

EstaPCarea = False

End Function


Sub DoPasosFx(ByVal charindex As Integer)
Static pie As Boolean

If Not UserNavegando Then
    If Not charlist(charindex).muerto And EstaPCarea(charindex) Then
        charlist(charindex).pie = Not charlist(charindex).pie
        If charlist(charindex).pie Then
            Call Audio.PlayWave(SND_PASOS1)
        Else
            Call Audio.PlayWave(SND_PASOS2)
        End If
    End If
Else
    Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub


Sub MoveCharbyPos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



x = charlist(charindex).Pos.x
y = charlist(charindex).Pos.y

MapData(x, y).charindex = 0

addX = nX - x
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).charindex = charindex


charlist(charindex).Pos.x = nX
charlist(charindex).Pos.y = nY

charlist(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(charindex).fX
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(charindex).fX = 0
    charlist(charindex).FxLoopTimes = 0
End If

If Not EstaPCarea(charindex) Then Dialogos.QuitarDialogo (charindex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(charindex)
End If

End Sub

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

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.x = x
    UserPos.x = tX
    AddtoUserPos.y = y
    UserPos.y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.x - 8 To UserPos.x + 8
    For k = UserPos.y - 6 To UserPos.y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer




'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        'CHOTS | Anti hack invi
        If GrhData(Grh).Speed > 2 Then GrhData(Grh).Speed = 1
        If FPSFast = True Then GrhData(Grh).Speed = GrhData(Grh).Speed * 4
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '¿Hay un personaje?
If MapData(x, y).charindex > 0 Then
LegalPos = False
Exit Function
End If
   
    If Not UserNavegando Then
        If HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast x, y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = x
    .Top = y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub


Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

Dim iGrhIndex As Integer
Dim SourceRect As RECT
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).fX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast x, y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).fX = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(x < 0, Abs(x), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(y < 0, Abs(y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = x
    .Top = y
    .Right = x + GrhData(iGrhIndex).pixelWidth
    .Bottom = y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast x, y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(x + x, y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hWnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC hdc, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
On Error Resume Next


If UserCiego Then Exit Sub

Dim y        As Integer 'Keeps track of where on map we are
Dim x        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim tempChar As char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer
Dim currentCharIndex As Integer

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)


'Draw floor layer
ScreenY = 8
For y = (minY + 8) To maxY - 8
    ScreenX = 8
    For x = minX + 8 To maxX - 8
        If x > 100 Or y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(x, y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                            .FrameCounter = 1
                    End If
                End If
            End If

            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With

        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight

        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
        '******************************************
        'Layer 2 **********************************
        If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(x, y).Graphic(2), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, _
                    1)
        End If
        '******************************************
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
    If y > 100 Then Exit For
Next y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For y = minY + 8 To maxY - 1
    If y > 0 And y < 101 Then 'In map bounds
        ScreenX = 5
        For x = minX + 5 To maxX - 5
            If x > 0 And x < 101 Then 'In map bounds
        iPPx = 32 * ScreenX - 32 + PixelOffsetX
        iPPy = 32 * ScreenY - 32 + PixelOffsetY

        'Object Layer **********************************
        If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
            Call DDrawTransGrhtoSurface( _
                                BackBufferSurface, _
                                MapData(x, y).ObjGrh, _
                                iPPx, iPPy, 1, 1)
 
        End If
        '***********************************************
        'Char layer ************************************
        If MapData(x, y).charindex <> 0 Then
            currentCharIndex = MapData(x, y).charindex
            tempChar = charlist(currentCharIndex)
            PixelOffsetXTemp = PixelOffsetX
            PixelOffsetYTemp = PixelOffsetY

            Moved = 0
            'If needed, move left and right
            If tempChar.MoveOffset.x <> 0 Then
                tempChar.Body.Walk(tempChar.Heading).Started = 1
                tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 1
                tempChar.Escudo.ShieldWalk(tempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + tempChar.MoveOffset.x
                If FPSFast = True Then
                    tempChar.MoveOffset.x = tempChar.MoveOffset.x - (2 * Sgn(tempChar.MoveOffset.x))
                Else
                    tempChar.MoveOffset.x = tempChar.MoveOffset.x - (8 * Sgn(tempChar.MoveOffset.x))
                End If
                Moved = 1
            End If
            'If needed, move up and down
            If tempChar.MoveOffset.y <> 0 Then
                tempChar.Body.Walk(tempChar.Heading).Started = 1
                tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 1
                tempChar.Escudo.ShieldWalk(tempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + tempChar.MoveOffset.y
                If FPSFast = True Then
                    tempChar.MoveOffset.y = tempChar.MoveOffset.y - (2 * Sgn(tempChar.MoveOffset.y))
                Else
                    tempChar.MoveOffset.y = tempChar.MoveOffset.y - (8 * Sgn(tempChar.MoveOffset.y))
                End If
                Moved = 1
            End If
            'If done moving stop animation
            If Moved = 0 And tempChar.Moving = 1 Then
                tempChar.Moving = 0
                tempChar.Body.Walk(tempChar.Heading).FrameCounter = 1
                tempChar.Body.Walk(tempChar.Heading).Started = 0
                tempChar.Arma.WeaponWalk(tempChar.Heading).FrameCounter = 1
                tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 0
                tempChar.Escudo.ShieldWalk(tempChar.Heading).FrameCounter = 1
                tempChar.Escudo.ShieldWalk(tempChar.Heading).Started = 0
            End If
            
            '[ANIM ATAK]
            If tempChar.Arma.WeaponAttack > 0 Then
                tempChar.Arma.WeaponAttack = tempChar.Arma.WeaponAttack - 1
                If tempChar.Arma.WeaponAttack = 0 Then
                    tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 0
                End If
            End If
            '[/ANIM ATAK]
            
            'Dibuja solamente players
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
            If tempChar.Head.Head(tempChar.Heading).GrhIndex <> 0 Then
                If Not tempChar.invisible Or (charlist(UserCharIndex).priv > 2 And charlist(UserCharIndex).priv < 5) Or MismoClan(currentCharIndex) = True Or UserCharIndex = currentCharIndex Then
                    '[CUERPO]'
                    Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), _
                            (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                            (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                            1, 1)
                    '[CABEZA]'
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            tempChar.Head.Head(tempChar.Heading), _
                            iPPx + tempChar.Body.HeadOffset.x, _
                            iPPy + tempChar.Body.HeadOffset.y, _
                            1, 0)
                    '[Casco]'
                    If tempChar.Casco.Head(tempChar.Heading).GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface( _
                                BackBufferSurface, _
                                tempChar.Casco.Head(tempChar.Heading), _
                                iPPx + tempChar.Body.HeadOffset.x, _
                                iPPy + tempChar.Body.HeadOffset.y, _
                                1, 0)
                    End If
                    '[ARMA]'
                    If tempChar.Arma.WeaponWalk(tempChar.Heading).GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface( _
                                BackBufferSurface, _
                                tempChar.Arma.WeaponWalk(tempChar.Heading), _
                                iPPx, iPPy, 1, 1)
                    End If
                    '[Escudo]'
                    If tempChar.Escudo.ShieldWalk(tempChar.Heading).GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface( _
                                BackBufferSurface, _
                                tempChar.Escudo.ShieldWalk(tempChar.Heading), _
                                iPPx, iPPy, 1, 1)
                    End If
                    
                    Call DrawNames(tempChar, iPPx, iPPy)
                    
                End If  'end if ~in

                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + tempChar.Body.HeadOffset.x), _
                            (iPPy + tempChar.Body.HeadOffset.y), _
                            currentCharIndex)
                End If
                
                
            Else 'No tiene cabeza
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + tempChar.Body.HeadOffset.x), _
                            (iPPy + tempChar.Body.HeadOffset.y), _
                            currentCharIndex)
                End If

                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        tempChar.Body.Walk(tempChar.Heading), _
                        iPPx, iPPy, 1, 1)
                
                'CHOTS, Bysnack | No tiene cabeza, entonces o es un GM invi o un user navegando. En caso de ser el segundo, mostramos nombres.
                If tempChar.priv = 0 Or tempChar.priv = 5 Or tempChar.priv = 6 Then
                    If (Not tempChar.invisible) Or (MismoClan(currentCharIndex) = True) Or (UserCharIndex = currentCharIndex) Then
                        Call DrawNames(tempChar, iPPx, iPPy)
                    End If
                End If

            End If

            'Refresh charlist
            charlist(currentCharIndex) = tempChar

            'BlitFX (TM)
            If charlist(currentCharIndex).fX <> 0 Then

                Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    FxData(tempChar.fX).fX, _
                    iPPx + FxData(tempChar.fX).OffsetX, _
                    iPPy + FxData(tempChar.fX).OffsetY, _
                    1, 1, currentCharIndex)

            End If
        End If '<-> If MapData(X, Y).CharIndex <> 0 Then
        '*************************************************
        'Layer 3 *****************************************
        If MapData(x, y).Graphic(3).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(x, y).Graphic(3), _
                    ((32 * (ScreenX)) - 32) + PixelOffsetX, _
                    ((32 * (ScreenY)) - 32) + PixelOffsetY, _
                    1, 1)
        End If
        '************************************************
        End If
        ScreenX = ScreenX + 1
    Next x
    End If
    ScreenY = ScreenY + 1
    If y >= 100 Or y < 1 Then Exit For
Next y

If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5
    For y = minY + 5 To maxY - 1
        ScreenX = 5
        For x = minX + 5 To maxX
            'Check to see if in bounds
            If x < 101 And x > 0 And y < 101 And y > 0 Then
                If MapData(x, y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(x, y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                End If
            End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
End If

End Sub
Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.y <= y
        
End If
End Function

Function PixelPos(ByVal x As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * x) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
    AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
    
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.x = MinXBorder
UserPos.y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.status, "Cargando Gráficos....", 0, 0, 0, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer
    
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            If FPSFast = True Then
                OffsetCounterX = (OffsetCounterX - (2 * Sgn(AddtoUserPos.x)))
            Else
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.x)))
            End If
           If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
               UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.y <> 0 Then
            If FPSFast = True Then
                OffsetCounterY = OffsetCounterY - (2 * Sgn(AddtoUserPos.y))
            Else
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.y))
            End If
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = 0
           End If
        End If
        
        '****** Update screen ******
        Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        
        Call DialogosClanes.Draw(Dialogos)
        
        Call DrawBackBufferSurface
        
        frmMain.MiniMap.Refresh
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, index As Integer)
ReDim Preserve Grh(1 To index) As Grh
Grh(index).FrameCounter = 1
Grh(index).GrhIndex = GrhIndex
Grh(index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(index).Started = 1
End Sub

Sub CargarAnimsExtra()

'CHOTS | Minimapa
Dim DDm As DDSURFACEDESC2
DDm.lHeight = 101
DDm.lWidth = 101
DDm.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
DDm.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
Set SupMiniMap = DirectDraw.CreateSurface(DDm)
Set SupBMiniMap = DirectDraw.CreateSurface(DDm)
'CHOTS | Minimapa

Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function


#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub
#End If

Private Function MismoClan(ByVal charindex As Integer) As Boolean
    MismoClan = ((charlist(UserCharIndex).Clan = charlist(charindex).Clan) And charlist(UserCharIndex).Clan <> "")
End Function

Sub SetNameColor(ByVal charindex As Integer)
Dim nameRed As Byte, nameGreen As Byte, nameBlue As Byte

If charlist(charindex).invisible = True Then
    'CHOTS | Invi
    nameRed = 255
    nameGreen = 255
    nameBlue = 85
Else
    Select Case charlist(charindex).priv
        Case 1:
            'CHOTS | Consejero
            nameRed = 173
            nameGreen = 251
            nameBlue = 156
        Case 2:
            'CHOTS | Ot
            nameRed = 128
            nameGreen = 128
            nameBlue = 255
        Case 3:
            'CHOTS | Semidios
            nameRed = 0
            nameGreen = 180
            nameBlue = 0
        Case 4:
            'CHOTS | Dios
            nameRed = 0
            nameGreen = 255
            nameBlue = 0
        Case 0:
            If charlist(charindex).Criminal Then
                'CHOTS | Criminal
                nameRed = 255
                nameGreen = 0
                nameBlue = 0
            Else
                'CHOTS | Ciudadano
                nameRed = 0
                nameGreen = 128
                nameBlue = 255
            End If
        Case 5:
            'CHOTS | Consejo Real
            nameRed = 0
            nameGreen = 200
            nameBlue = 255
        Case 6:
            'CHOTS | Consejo Caos
            nameRed = 255
            nameGreen = 100
            nameBlue = 0
        Case Else:
            nameRed = 0
            nameGreen = 0
            nameBlue = 0
    End Select
End If

charlist(charindex).redColor = nameRed
charlist(charindex).greenColor = nameGreen
charlist(charindex).blueColor = nameBlue
charlist(charindex).nameCenter = (frmMain.TextWidth(charlist(charindex).Nombre) / 2) - 16
If charlist(charindex).Clan <> "" Then
    charlist(charindex).clanCenter = (frmMain.TextWidth(charlist(charindex).Clan) / 2) - 16
End If

End Sub

Sub DrawNames(tempChar As char, iPPx As Integer, iPPy As Integer)
    If Nombres And tempChar.Nombre <> "" Then
        Call Dialogos.DrawText(iPPx - tempChar.nameCenter, iPPy + 30, tempChar.Nombre, RGB(tempChar.redColor, tempChar.greenColor, tempChar.blueColor))

        If tempChar.Clan <> "" Then
            Call Dialogos.DrawText(iPPx - tempChar.clanCenter, iPPy + 45, tempChar.Clan, RGB(tempChar.redColor, tempChar.greenColor, tempChar.blueColor))
        End If
    End If
End Sub
