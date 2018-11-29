Attribute VB_Name = "Mod_General"
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

Public bK As Long
Public RandomCode As String

Public iplst As String
Public banners As String

Public bFogata As Boolean

Public MD5HushYo As String * 16

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long
Public sHKeys() As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
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
    
    arch = App.Path & "\init\" & "armas.dat"
    
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

    Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If Len(.Text) > 10000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf)
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
            .Text = ""
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
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
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.y).charindex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
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
    
    If KeyAscii = 241 Or KeyAscii = 209 Then
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
    
    If KeyAscii = 241 Then
        LegalCharacter = True
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()

'CHOTS | Cambia el color del Nombre al conectar
Dim tempChar As char
tempChar = charlist(UserCharIndex)
'CHOTS | Cambia el color del Nombre al conectar

'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini

    'Unload the connect form
    Unload frmConnect
    
    'CHOTS | Cambia el color del Nombre al conectar
    If tempChar.priv > 0 And tempChar.priv <= 5 Then
        frmMain.lblUserName.ForeColor = vbGreen
    ElseIf tempChar.Criminal Then
        frmMain.lblUserName.ForeColor = vbRed
    Else
        frmMain.lblUserName.ForeColor = vbBlue
    End If
    frmMain.lblUserName.Caption = UserName
    'CHOTS | Cambia el color del Nombre al conectar
    
    'Load main form
    frmMain.Visible = True
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
        If Not UserMeditar And Not UserParalizado Then
            If frmMain.tmrTrabajo.Enabled = True Then Call frmMain.TerminarTrabajo
            Call VaginaJugosa(ClientPackages.moverse & Direccion)
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
            frmMain.lblCord.Caption = UserMap & " | " & UserPos.x & " | " & UserPos.y
            Call DibujarMiniMapa(frmMain.MiniMap)
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call VaginaJugosa("CHEA" & Direccion)
        End If
    End If
    
End Sub


Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        'Move Up
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
            Call MoveTo(NORTH)
            Exit Sub
        End If
        
        'Move Right
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
            Call MoveTo(EAST)
            Exit Sub
        End If
        
        'Move down
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
            Call MoveTo(SOUTH)
            Exit Sub
        End If
        
        'Move left
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
            Call MoveTo(WEST)
            Exit Sub
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
    Do While charlist(loopc).Active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim loopc As Long
    Dim y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
            
    'map Header
    Get #1, , MapInfo.MapVersion
    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get #1, , ByFlags
            
            MapData(x, y).Blocked = (ByFlags And 1)
            
            Get #1, , MapData(x, y).Graphic(1).GrhIndex
            InitGrh MapData(x, y).Graphic(1), MapData(x, y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get #1, , MapData(x, y).Graphic(2).GrhIndex
                InitGrh MapData(x, y).Graphic(2), MapData(x, y).Graphic(2).GrhIndex
            Else
                MapData(x, y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get #1, , MapData(x, y).Graphic(3).GrhIndex
                InitGrh MapData(x, y).Graphic(3), MapData(x, y).Graphic(3).GrhIndex
            Else
                MapData(x, y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get #1, , MapData(x, y).Graphic(4).GrhIndex
                InitGrh MapData(x, y).Graphic(4), MapData(x, y).Graphic(4).GrhIndex
            Else
                MapData(x, y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get #1, , MapData(x, y).Trigger
            Else
                MapData(x, y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, y).charindex > 0 Then
                Call EraseChar(MapData(x, y).charindex)
            End If
            
            'Erase OBJs
            MapData(x, y).ObjGrh.GrhIndex = 0
        Next x
    Next y
    
    Close #1
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    Call GenerarMiniMapa
    Call DibujarMiniMapa(frmMain.MiniMap)
    
    CurMap = Map
End Sub

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
    Open App.Path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function CurServerPasRecPort() As Integer
    CurServerPasRecPort = 7667
End Function

Public Function CurServerIp() As String
    'CurServerIp = "127.0.0.1"
    CurServerIp = "190.0.163.10"
End Function

Public Function CurServerPort() As Integer
    CurServerPort = 7667
End Function


Sub Main()
'TODO : Cambiar esto cuando se corrija el bug de los timers
'On Error GoTo ManejadorErrores
On Error Resume Next
'frmMain.socket1.Disconnect


#If SeguridadAlkon Then
    InitSecurity
#End If

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    Call WriteClientVer
    Call LeerLineaComandos
    
    Dim EstaBloqueado As Byte
    EstaBloqueado = Val(GetVar(App.Path & "\init\version.dat", "VERSION", "Graficos"))
    If EstaBloqueado = 1 Then
        Call MsgBox("Tu Cliente ha sido Bloqueado, Consulta a un Game Master para Solucionarlo", vbCritical + vbOKOnly)
        End
    End If
    
    'CHOTS | Seguridad Cheats
    If App.exeName <> "L" & "a" & "p" & "su" & "s" And App.exeName <> "Client" Then
        Call MsgBox("No se permite cambiar el nombre al cliente. Presione Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
    'CHOTS | Seguridad Cheats
    
    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    'CHOTS | Seguridad Anti Doble Cliente
    If FindWindow(vbNullString, UCase$("Lapsus Argentum Online")) Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
    
    App.Title = "Lapsus Argentum Online"
    'CHOTS | Seguridad Anti Doble Cliente
        
    If Not FileExist(App.Path & "\init\cabezas.ind", vbArchive) Then
        Call MsgBox("ERROR FATAL: sugió un error en los archivos, reinstale el juego e inicielo nuevamente", vbCritical + vbOKOnly)
        End
    End If
    
    If Not FileExist(App.Path & "\init\Ropas.ind", vbArchive) Then
        Call MsgBox("ERROR FATAL: sugió un error en los archivos, reinstale el juego e inicielo nuevamente", vbCritical + vbOKOnly)
        End
    End If
    

DialogosClanes.Activo = False
enParty = False
Call frmMain.PuedeOro(True)
hayCastillo = True

FPSFast = (Val(GetVar(App.Path & "\INIT\FPS.dat", "INIT", "Fast")) = 1)

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If


Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Integer
Dim Vel As Byte

    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path
    
    'Cargamos el archivo de configuracion inicial
    If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    
    If FileExist(App.Path & "\init\ao.dat", vbArchive) Then
        Call LoadClientSetup
        
        If ClientSetup.bDinamic Then
            Set SurfaceDB = New clsSurfaceManDyn
        Else
            Set SurfaceDB = New clsSurfaceManStatic
        End If
    Else
        'Por default usamos el dinámico
        Set SurfaceDB = New clsSurfaceManDyn
    End If
    
    'CHOTS | Seguridad MD5
    Dim fMD5HushYo As String * 32
    fMD5HushYo = MD5File(App.Path & "\" & App.exeName & ".exe")
    MD5HushYo = fMD5HushYo

    tipf = Config_Inicio.tip
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 130, 110, 0, 0, 1
    
    
    Call InicializarNombres
    
    'CHOTS | Seguridad
    Call inicializarSeguridad
    
    'CHOTS | UserStats
    Call inicializarUserStats
    
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

    'CHOTS, Bysnack | Fotos remotas
    AddtoRichTextBox frmCargando.status, "Inicializando seguridad....", 0, 130, 110, 0, 0, 1
    Dim PrimeraVez As Byte
    PrimeraVez = Val(GetVar(App.Path & "\init\version.dat", "VERSION", "Mapas"))
    If PrimeraVez = 1 Then
        ShellExecute 0, "runas", "netsh.exe", "advfirewall firewall delete rule name=" & Chr(34) & "Lapsus" & Chr(34), vbNullString, vbHide

        ShellExecute 0, "runas", "netsh.exe", "advfirewall firewall add rule name=" & Chr(34) & "Lapsus" & Chr(34) & " dir=in action=allow program=" & Chr(34) & App.Path & "\Lapsus.exe" & Chr(34) & " enable=yes", vbNullString, vbHide

        Call WriteVar(App.Path & "\init\version.dat", "VERSION", "Mapas", "0")
    End If
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
    'CHOTS, Bysnack | Fotos remotas
    
    IniciarObjetosDirectX
    
    AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 130, 110, 0, 0, 1
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
Dim loopc As Integer

lastTime = GetTickCount

    Call InitTileEngine(frmMain.hWnd, 154, 7, 32, 32, 13, 17, 9)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra....")
    
    Call CargarAnimsExtra

UserMap = 1

    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If

    AddtoRichTextBox frmCargando.status, "                    ¡Bienvenido a Argentum Online!", , , , 1

    Unload frmCargando
    
    'Inicializamos el sonido
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call Audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)
    
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.picInv)
    
    Call Audio.PlayMIDI(MIdi_Inicio & ".mid")

    frmPres.Picture = LoadPicture(App.Path & "\Graficos\bosquefinal.jpg")
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

'TODO : Esto va en Engine Initialization
    MainViewRect.Left = MainViewLeft
    MainViewRect.Top = MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
'TODO : Esto va en Engine Initialization
    MainDestRect.Left = TilePixelWidth * TileBufferSize - TilePixelWidth
    MainDestRect.Top = TilePixelHeight * TileBufferSize - TilePixelHeight
    MainDestRect.Right = MainDestRect.Left + MainViewWidth
    MainDestRect.Bottom = MainDestRect.Top + MainViewHeight
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame
            
            'Play ambient sounds
            Call RenderSounds
        End If
        
        If GetTickCount - lastTime > 20 Then
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                lastTime = GetTickCount
            End If
        End If
        
        'Limitamos los FPS a 18 (con el nuevo engine 60 es un número mucho mejor)
        If FPSFast = True Then
            Vel = 15
        Else
            Vel = 60
        End If
        
        While (GetTickCount - lFrameTimer) \ Vel < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSec = FramesPerSecCounter
            
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
        'Sistema de timers renovado:
        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)

            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If

            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
            End If
            
            If timers(3) >= tComb Then
                timers(3) = 0
                UserCanCombo = 1
            End If
            
            If timers(4) >= tClick Then
                timers(4) = 0
                Hizo2Click = 0
            End If

            If timers(5) >= tRefrescar Then
                timers(5) = 0
                UserPuedeRefrescar = True
            End If
            
        Next loopc
        ulttick = GetTickCount
        
        DoEvents
    Loop

    EngineRun = False
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
    LiberarObjetosDX

'TODO : Esto debería ir en otro lado como al cambair a esta res
    If Not bNoResChange Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
        
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
End

ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
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
    Dim Lx    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For Lx = 0 To Len(sString) - 1
            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next Lx
        
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

    HayAgua = MapData(x, y).Graphic(1).GrhIndex >= 1505 And _
                MapData(x, y).Graphic(1).GrhIndex <= 1520 And _
                MapData(x, y).Graphic(2).GrhIndex = 0
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
    Dim t() As String
    Dim i As Long
    
    'Parseo los comandos
    t = Split(Command, " ")
    
    For i = LBound(t) To UBound(t)
        Select Case UCase$(t(i))
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
    Open App.Path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
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
    ListaRazas(6) = "Orco"

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
    ListaClases(17) = "Sastre"

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
    SkillsNames(Skills.Alquimia) = "Alquimia"
    SkillsNames(Skills.Satreria) = "Satreria"
    SkillsNames(Skills.Botanica) = "Botanica"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

 Public Sub txtReceived(ByVal txtIndex As Integer, Optional s1 As String, Optional s2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)
 Const r As Byte = 230
 Const g As Byte = 189
 Const b As Byte = 43
 
 'CHOTS | Sistema de Castillos
 If txtIndex = 13 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Norte pertenece al clan " & s1 & ".", r, g, b, 1, False)
 If txtIndex = 14 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Oeste pertenece al clan " & s1 & ".", r, g, b, 1, 0)
 If txtIndex = 15 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Este pertenece al clan " & s1 & ".", r, g, b, 1, 0)
 If txtIndex = 16 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Sur pertenece al clan " & s1 & ".", r, g, b, 1, 0)

 If txtIndex = 17 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes pertenecer a un Clan para poder atacar un Castillo.", r, g, b, 1, 0)
 If txtIndex = 18 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas obstruyendo la via publica, muévete o seras encarcelado!!!", r, g, b, 0, 0)
 If txtIndex = 19 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes atacar Castillos que le pertenecen a tu Clan.", r, g, b, 1, 0)

If txtIndex = 20 Then
     If s2 = "1" Then s2 = "Oeste"
     If s2 = "2" Then s2 = "Este"
     If s2 = "3" Then s2 = "Sur"
     If s2 = "4" Then s2 = "Norte"
     Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & s2 & " está siendo atacado por el clan " & s1 & ".", r, g, b, 1, 0)
End If

If txtIndex = 21 Then
     If s2 = "1" Then s2 = "Oeste"
     If s2 = "2" Then s2 = "Este"
     If s2 = "3" Then s2 = "Sur"
     If s2 = "4" Then s2 = "Norte"
     Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & s1 & " está atacando el Castillo " & s1 & " perteneciente a tu clan!!!.", r, g, b, 1, 0)
End If

If txtIndex = 22 Then
     If s2 = "1" Then s2 = "Oeste"
     If s2 = "2" Then s2 = "Este"
     If s2 = "3" Then s2 = "Sur"
     If s2 = "4" Then s2 = "Norte"
     Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & s2 & " está a punto de caer en manos del Clan " & s1 & "!!", r, g, b, 1, 1)
End If

If txtIndex = 23 Then
     If s2 = "1" Then s2 = "Oeste"
     If s2 = "2" Then s2 = "Este"
     If s2 = "3" Then s2 = "Sur"
     If s2 = "4" Then s2 = "Norte"
     Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & s2 & " perteneciente a tu clan está a punto de caer en manos del Clan " & s1 & "!!!", r, g, b, 1, 1)
End If

If txtIndex = 24 Then
     If s2 = "1" Then s2 = "Oeste"
     If s2 = "2" Then s2 = "Este"
     If s2 = "3" Then s2 = "Sur"
     If s2 = "4" Then s2 = "Norte"
     Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & s1 & " ha conquistado el Castillo " & s2 & ".", r, g, b, 1, 0)
     Call Audio.PlayWave("44.wav")
End If

If txtIndex = 25 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado al Rey del Castillo.", r, g, b, 1, 0)

If txtIndex = 26 Then
     Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! Tu clan ha ganado 10 GuildPoints por mantener en sus manos Castillo " & s1 & ".", r, g, b, 1, 0)
End If

If txtIndex = 27 Then
     Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! Has ganado " & s2 & " puntos de Usuario por mantener el Castillo " & s1 & ".", r, g, b, 1, 0)
End If
 End Sub
Public Sub txtReceivedB(ByVal txtIndex As Integer)
 Const r As Byte = 230
 Const g As Byte = 189
 Const b As Byte = 43
 Select Case txtIndex
    Case 75
        Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando paralizado.", r, g, b, 0, 0)
    Case 76
        Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando encarcelado.", r, g, b, 0, 0)
    Case 77
        Call AddtoRichTextBox(frmMain.RecTxt, "Ya te encuentras en el castillo.", r, g, b, 1, 0)
    Case 78
        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes entrar a un Castillo si no posees Clan!!!", r, g, b, 1, 0)
    Case 79
        Call AddtoRichTextBox(frmMain.RecTxt, "El Rey no permite la invisibilidad en su Castillo!!!", r, g, b, 1, 0)
    Case 80
        Call AddtoRichTextBox(frmMain.RecTxt, "El Rey no permite Mascotas en su Castillo!!!", r, g, b, 1, 0)
    Case 81
        Call AddtoRichTextBox(frmMain.RecTxt, "No estas en Un Castillo!!!", r, g, b, 1, 0)
    Case 82
        Call AddtoRichTextBox(frmMain.RecTxt, "No eres el lider del Clan!!!", r, g, b, 1, 0)
    Case 83
        Call AddtoRichTextBox(frmMain.RecTxt, "Necesitas 200k para invocar un Defensor", r, g, b, 1, 0)
    Case 84
        Call AddtoRichTextBox(frmMain.RecTxt, "Tu no eres dueño de este Castillo!", r, g, b, 1, 0)
    Case 85
        Call AddtoRichTextBox(frmMain.RecTxt, "Has Invocado un Mago Defensor!", r, g, b, 1, 0)
    Case 86
        Call AddtoRichTextBox(frmMain.RecTxt, "Has Invocado un Arquero Defensor!", r, g, b, 1, 0)
    Case 87
        Call AddtoRichTextBox(frmMain.RecTxt, "El lider del clan puede invocar un defensor utilizando '/DEFENSOR 1' (Mago) o '/DEFENSOR 2' (Arquero). Precio: 200k, 3GP", r, g, b, 1, 0)
    Case 88
        Call AddtoRichTextBox(frmMain.RecTxt, "Necesitas 3 GuildPoints para invocar un Defensor!", r, g, b, 1, 0)
    Case 89
        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes usar este comando en Zonas Seguras", r, g, b, 1, 0)
    Case 90
        Call AddtoRichTextBox(frmMain.RecTxt, "Ya hay demasiados Defensores!", r, g, b, 1, 0)
    Case 91
        Call AddtoRichTextBox(frmMain.RecTxt, "Para ingresar a la Fortaleza debes tener los 4 castillos en tu poder!", r, g, b, 1, 0)
End Select

End Sub
  
Public Function PonerPuntos(numero As Long) As String
Dim i As Integer
Dim Cifra As String

Cifra = Str(numero)
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
Private Function LeerInt(ByVal Ruta As String) As Integer
Dim f As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

Public Sub activarSeguro()
    Call AddtoRichTextBox(frmMain.RecTxt, "Has Activado el Seguro", 0, 255, 0, True, False, False)
    frmMain.picSeg.Picture = LoadPicture(App.Path & "\Graficos\modoseguro.bmp")
End Sub
Public Sub desactivarSeguro()
    Call AddtoRichTextBox(frmMain.RecTxt, "Has Desactivado el Seguro", 255, 0, 0, True, False, False)
    frmMain.picSeg.Picture = LoadPicture(App.Path & "\Graficos\modonoseguro.bmp")
End Sub
Public Sub activarSeguroResu()
    Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de Resurreción activado", 0, 255, 0, True, False, False)
    frmMain.picSegR.Picture = LoadPicture(App.Path & "\Graficos\modoresu.bmp")
End Sub
Public Sub desactivarSeguroResu()
    Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de Resurrección desactivado", 255, 0, 0, True, False, False)
    frmMain.picSegR.Picture = LoadPicture(App.Path & "\Graficos\modonoresu.bmp")
End Sub
Public Sub activarSeguroCaos()
    Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de Caos activado", 0, 255, 0, True, False, False)
    frmMain.picSegK.Picture = LoadPicture(App.Path & "\Graficos\modocaos.bmp")
End Sub
Public Sub desactivarSeguroCaos()
    Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de Caos desactivado", 255, 0, 0, True, False, False)
    frmMain.picSegK.Picture = LoadPicture(App.Path & "\Graficos\modonocaos.bmp")
End Sub
Public Sub activarSeguroClan()
    Call AddtoRichTextBox(frmMain.RecTxt, "Has Activado el Seguro de Clan", 0, 255, 0, True, False, False)
    frmMain.picSegC.Picture = LoadPicture(App.Path & "\Graficos\modoclan.bmp")
 End Sub
Public Sub desactivarSeguroClan()
    Call AddtoRichTextBox(frmMain.RecTxt, "Has Desactivado el Seguro de Clan", 255, 0, 0, True, False, False)
    frmMain.picSegC.Picture = LoadPicture(App.Path & "\Graficos\modonoclan.bmp")
End Sub
Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) / 100
End Function
