Attribute VB_Name = "Mod_ErrorLOG"
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

Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub


Sub PeneVenoso(ByVal semen As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim x As Integer
    Dim y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, index As Integer, m As Integer
    Dim t() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    sData = UCase$(semen)
    
    Select Case sData
        'CHOTS | Optimizaciones de mensajes
        Case "Z1"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje1, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z2"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje2, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z3"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje3, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z4"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje4, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z5"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje5, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z6"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje6, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z7"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje7, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z8"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje8, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z9"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje9, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z10"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje10, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z11"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje11, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z12"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje12, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z13"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje13, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z14"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje14, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z15"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje15, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z16"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje16, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z17"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje17, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z18"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje18, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z19"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje19, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z20"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje20, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z21"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje21, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z22"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje22, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z23"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje23, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z24"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje24, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z25"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje25, 255, 0, 0, True, False, False)
        Exit Sub
        Case "Z26"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje26, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z27"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje27, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z28"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje28, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z29"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje29, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z30"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje30, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z31"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje31, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z32"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje32, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z33"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje33, 220, 177, 0, False, False, False)
        Exit Sub
        Case "Z34"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje34, 220, 177, 0, False, False, False)
        Exit Sub
        Case "Z35"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje35, 0, 255, 0, False, False, False)
        Exit Sub
        Case "Z36"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje36, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z37"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje37, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z38"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje38, 0, 255, 0, True, False, False)
        Exit Sub
        Case "Z39"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje39, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z40"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje40, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z41"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje41, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z42"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje42, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z43"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje43, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z44"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje44, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z45"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje45, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z46"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje46, 255, 255, 255, True, False, False)
        Exit Sub
        Case "Z47"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje47, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z48"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje48, 31, 51, 223, True, True, False)
        Exit Sub
        Case "Z49"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje49, 35, 51, 223, True, True, False)
        Exit Sub
        Case "Z50"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje50, 255, 0, 0, True, False, False)
        Exit Sub
        Case "Z51"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje51, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z52"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje52, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z53"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje53, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z54"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje54, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z55"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje55, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z56"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje56, 255, 0, 0, True, False, False)
        Exit Sub
        Case "Z57"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje57, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z58"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje58, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z59"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje59, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z60"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje60, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z61"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje61, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z62"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje62, 255, 0, 0, True, False, False)
        Exit Sub
        Case "Z63"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje63, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z64"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje64, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z65"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje65, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z66"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje66, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z67"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje67, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z68"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje68, 31, 51, 223, True, True, False)
        Exit Sub
        Case "Z69"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje69, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z70"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje70, 32, 51, 223, True, True, False)
        Exit Sub
        Case "Z71"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje71, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z72"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje72, 32, 51, 223, True, True, False)
        Exit Sub
        Case "Z73"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje73, 255, 255, 255, True, False, False)
        Exit Sub
        Case "Z74"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje74, 255, 255, 255, True, False, False)
        Exit Sub
        Case "Z75"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje75, 255, 255, 255, True, False, False)
        Exit Sub
        Case "Z76"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje76, 255, 255, 255, True, False, False)
        Exit Sub
        Case "Z77"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje77, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z78"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje78, 172, 172, 172, True, False, False)
        Exit Sub
        Case "Z79"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje79, 65, 190, 156, True, False, False)
        Exit Sub
        Case "Z80"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje80, 172, 172, 172, True, False, False)
        Exit Sub
        Case "Z81"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje81, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z82"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje82, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z83"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje83, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z84"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje84, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z85"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje85, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z86"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje86, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z87"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje87, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z88"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje88, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z89"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje89, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z90"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje90, 32, 51, 223, True, True, False)
        Exit Sub
        Case "Z91"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje91, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z92"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje92, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z93"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje93, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z94"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje94, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z95"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje95, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z96"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje96, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z97"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje97, 0, 74, 149, True, False, False)
        Exit Sub
        Case "Z98"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje98, 65, 190, 156, False, False, False)
        Exit Sub
        Case "Z99"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje99, 255, 255, 0, True, False, False)
        Exit Sub
        Case "Z100"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje100, 65, 190, 156, False, False, False)
        Exit Sub
        'CHOTS | Optimizaciones de mensajes
 
        Case "BUENO"
             TimerPing(2) = GetTickCount()
             Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 0, 0, True, False, False)
        Exit Sub

        Case ServerPackages.login
            Call GenerarMiniMapa
            Call DibujarMiniMapa(frmMain.MiniMap)
            Logged = True
            UserCiego = False
            EngineRun = True
            UserDescansar = False
            Nombres = True
            Hizo2Click = 0
            If frmCrearPersonaje1.Visible Then
                'Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje1
                Unload frmConnect
                frmMain.Show
                frmMain.SendCMSTXT.Visible = False
                frmMain.SendTxt.Visible = False
            End If
            Call SetConnected
            bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            DialogosClanes.PushBackText ("Bienvenido a Lapsus2017!+" & vbWhite)
            DialogosClanes.PushBackText ("No te olvides de visitar nuestra Web: www.Lapsus2017.com+" & vbGreen)
            DialogosClanes.PushBackText ("Te recordamos que al jugar, estás aceptando el Reglamento+" & vbYellow)
            DialogosClanes.PushBackText ("Te recomendamos leer el Manual o tipea /AYUDA para consultas rápidas...+" & vbRed)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case ServerPackages.logout
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            'CHOTS | No enviamos mas el Z84
            Call AddtoRichTextBox(frmMain.RecTxt, Mensaje84, 65, 190, 156, False, False, False)
            frmMain.Visible = False
            Logged = False
            UserParalizado = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Call Audio.StopWave
            frmMain.IsPlaying = PlayLoop.plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
            
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim ii As Integer
            ii = 1
            Do While ii <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(ii) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(ii)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show , frmMain
            Exit Sub
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
            Exit Sub
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "BORROK"
            Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
            frmBorrar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmBorrar
            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
         Case "ALQ"
            frmAlquimia.Show , frmMain 'CHOTS | Alquimia
            Exit Sub
        Case "SAS"
            frmSastre.Show , frmMain 'CHOTS | Sastre
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call activarSeguro
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call desactivarSeguro
            Exit Sub
        Case "SEGRON" 'CHOTS | Activa seguro Resu
            Call activarSeguroResu
            Exit Sub
        Case "SEGROFF" 'CHOTS | Desactiva seguro Resu
            Call desactivarSeguroResu
            Exit Sub
        Case "SEGKON" 'CHOTS | Activa seguro Caos
            Call activarSeguroCaos
            Exit Sub
        Case "SEGKOFF" 'CHOTS | Desactiva seguro Caos
            Call desactivarSeguroCaos
            Exit Sub
        Case "SEGCON" '  <--- Activa el seguro clan
            Call activarSeguroClan
            Exit Sub
        Case "SEGCOFF" ' <--- Desactiva el seguro clan
            desactivarSeguroClan
            Exit Sub
            
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
        Case ServerPackages.moverChar
            semen = Right$(semen, Len(semen) - 1)

#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveCharDecrypt(semen, charindex, x, y)
#Else
            charindex = Val(ReadField(1, semen, Asc(",")))
            x = Val(ReadField(2, semen, Asc(",")))
            y = Val(ReadField(3, semen, Asc(",")))
#End If

            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(semen, 1, 1)) * 64 + (Asc(Mid$(semen, 2, 1)) And &HFC&) / 4

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(charindex).fX >= 40 And charlist(charindex).fX <= 49 Then   'si esta meditando
                charlist(charindex).fX = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(charindex).priv = 0 Or charlist(charindex).priv = 5 Or charlist(charindex).priv = 6 Then
                Call DoPasosFx(charindex)
            End If

            Call MoveCharbyPos(charindex, x, y)
            
            Call RefreshAllChars
            Exit Sub
        Case ServerPackages.moverNpc
            semen = Right$(semen, Len(semen) - 1)
            
#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveNPCDecrypt(semen, charindex, x, y, Left$(sData, 1) <> "*")
#Else
            charindex = Val(ReadField(1, semen, Asc(",")))
            x = Val(ReadField(2, semen, Asc(",")))
            y = Val(ReadField(3, semen, Asc(",")))
#End If
            
            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(semen, 1, 1)) * 64 + (Asc(Mid$(semen, 2, 1)) And &HFC&) / 4
            
'            If charlist(CharIndex).Body.Walk(1).GrhIndex = 4747 Then
'                Debug.Print "hola"
'            End If
            
            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(charindex).fX >= 40 And charlist(charindex).fX <= 49 Then   'si esta meditando
                charlist(charindex).fX = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(charindex).priv = 0 Or charlist(charindex).priv = 5 Or charlist(charindex).priv = 6 Then
                Call DoPasosFx(charindex)
            End If
            
            Call MoveCharbyPos(charindex, x, y)
            'Call MoveCharbyPos(CharIndex, Asc(Mid$(semen, 3, 1)), Asc(Mid$(semen, 4, 1)))
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case Left$(sData, 2)
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M":
                    SetMana (k)
                    Exit Sub
                Case "H":
                    SetHp (k)
                    Exit Sub
                Case "S":
                    SetStamina (k)
                    Exit Sub
                Case "G":
                    SetGold (k)
                    Exit Sub
                Case "E":
                    SetExp (k)
                    Exit Sub
            End Select
            Exit Sub
        Case ServerPackages.cargarMapa
            semen = Right$(semen, Len(semen) - 2)
            UserMap = ReadField(1, semen, 44)
            'Obtiene la version del mapa
            
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
'                If tempint = Val(ReadField(2, semen, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            Call Audio.StopWave(RainBufferIndex)
                            RainBufferIndex = 0
                            frmMain.IsPlaying = PlayLoop.plNone
                        End If
                    End If
'                Else
'                    'vers incorrecta
'                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
'                    Call LiberarObjetosDX
'                    Call UnloadAllForms
'                    End
'                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        
        Case ServerPackages.updatePos
            semen = Right$(semen, Len(semen) - 2)
            MapData(UserPos.x, UserPos.y).charindex = 0
            UserPos.x = CInt(ReadField(1, semen, 44))
            UserPos.y = CInt(ReadField(2, semen, 44))
            MapData(UserPos.x, UserPos.y).charindex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            frmMain.lblCord.Caption = UserMap & " | " & UserPos.x & " | " & UserPos.y
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            semen = Right$(semen, Len(semen) - 2)
            i = Val(ReadField(1, semen, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            semen = Right$(semen, Len(semen) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & semen & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            semen = Right$(semen, Len(semen) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & semen & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            semen = Right$(semen, Len(semen) - 2)
            i = Val(ReadField(1, semen, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, semen, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            semen = Right$(semen, Len(semen) - 2)
            i = Val(ReadField(1, semen, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, semen, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, semen, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case ServerPackages.dialogo
        'CHOTS | Modificado por CHOTS para mejorar el ancho de banda
            semen = Right$(semen, Len(semen) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, semen, 176))
            Dim r As Byte
            Dim g As Byte
            Dim b As Byte
            Dim N As Byte
            Dim c As Byte
            Dim num As Byte
            
            num = Val(ReadField(2, semen, 126))
            
            Select Case num
            Case 1 'CHOTS | Talk
                r = 255
                g = 255
                b = 255
                N = 0
                c = 0
            Case 2 'CHOTS | Fight
                r = 255
                g = 0
                b = 0
                N = 1
                c = 0
            Case 3 'CHOTS | Warning
                r = 32
                g = 51
                b = 223
                N = 1
                c = 1
            Case 4 'CHOTS | Info
                r = 65
                g = 190
                b = 156
                N = 0
                c = 0
            Case 5 'CHOTS | Gema
                r = 255
                g = 0
                b = 255
                N = 1
                c = 0
            Case 6 'CHOTS | Apu
                r = 255
                g = 128
                b = 0
                N = 1
                c = 0
            Case 7 'CHOTS | Dios
                r = 0
                g = 240
                b = 0
                N = 1
                c = 0
            Case 8 'CHOTS | Semi
                r = 255
                g = 255
                b = 128
                N = 1
                c = 0
            Case 9 'CHOTS | Infon
                r = 65
                g = 190
                b = 156
                N = 1
                c = 0
            Case 10 'CHOTS | Ejecucion
                r = 130
                g = 130
                b = 130
                N = 1
                c = 0
            Case 11 'CHOTS | Party
                r = 255
                g = 180
                b = 255
                N = 0
                c = 0
            Case 12 'CHOTS | Veneno
                r = 0
                g = 255
                b = 0
                N = 0
                c = 0
            Case 13 'CHOTS | Guild
                r = 255
                g = 255
                b = 255
                N = 1
                c = 0
            Case 14 'CHOTS | Server
                r = 0
                g = 185
                b = 0
                N = 0
                c = 0
            Case 15 'CHOTS | Guildmsj
                r = 228
                g = 199
                b = 27
                N = 0
                c = 0
            Case 16 'CHOTS | Consejo
                r = 130
                g = 130
                b = 255
                N = 1
                c = 0
            Case 17 'CHOTS | Consejocaos
                r = 255
                g = 60
                b = 0
                N = 1
                c = 0
            Case 18 'CHOTS | Consejovesa
                r = 0
                g = 200
                b = 255
                N = 1
                c = 0
            Case 19 'CHOTS | Consejocaosvesa
                r = 255
                g = 100
                b = 0
                N = 1
                c = 0
            Case 20 'CHOTS | Oro
                r = 255
                g = 255
                b = 0
                N = 1
                c = 0
            Case 21 'CHOTS | Celestenegrita
                r = 0
                g = 128
                b = 255
                N = 1
                c = 0
            Case 22 'CHOTS | Azul
                r = 0
                g = 0
                b = 255
                N = 1
                c = 0
            Case 23 'CHOTS | GM
                r = 0
                g = 255
                b = 0
                N = 1
                c = 0
            Case 24 'CHOTS | Troforo
                r = 233
                g = 198
                b = 1
                N = 1
                c = 0
            Case 25 'CHOTS | Trofplata
                r = 196
                g = 198
                b = 196
                N = 1
                c = 0
            Case 26 'CHOTS | Casamiento
                r = 255
                g = 55
                b = 155
                N = 1
                c = 0
            Case 27 'CHOTS | Duelos
                r = 128
                g = 64
                b = 64
                N = 1
                c = 0
            Case 28 'CHOTS | Hogar
                r = 128
                g = 128
                b = 255
                N = 1
                c = 0
            Case 29 'CHOTS | Invocaciones
                r = 172
                g = 172
                b = 172
                N = 1
                c = 0
            Case 30 'CHOTS | Torneos auto
                r = 0
                g = 74
                b = 149
                N = 1
                c = 0
            Case 31 'CHOTS | Monturas
                r = 2
                g = 134
                b = 45
                N = 1
                c = 0
            Case 32 'CHOTS | Guerras
                r = 235
                g = 235
                b = 188
                N = 1
                c = 0
            Case 33 'BysNacK | Chat privado
                r = 175
                g = 15
                b = 150
                N = 1
                c = 0
            Case Else
                r = 65
                g = 190
                b = 156
                N = 0
                c = 0
        End Select

            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, semen, 176), iuser, Val(ReadField(1, semen, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, semen, 126), r, g, b, N, c
                End If
            End If

            Exit Sub
        Case ServerPackages.dialogoConsola
        'CHOTS | Modificado por CHOTS para mejorar el ancho de banda
            semen = Right$(semen, Len(semen) - 2)
            
            iuser = Val(ReadField(3, semen, 176))
            
            num = Val(ReadField(2, semen, 126))
            
            Select Case num
            Case 1 'CHOTS | Talk
                r = 255
                g = 255
                b = 255
                N = 0
                c = 0
            Case 2 'CHOTS | Fight
                r = 255
                g = 0
                b = 0
                N = 1
                c = 0
            Case 3 'CHOTS | Warning
                r = 32
                g = 51
                b = 223
                N = 1
                c = 1
            Case 4 'CHOTS | Info
                r = 65
                g = 190
                b = 156
                N = 0
                c = 0
            Case 5 'CHOTS | Gema
                r = 255
                g = 0
                b = 255
                N = 1
                c = 0
            Case 6 'CHOTS | Apu
                r = 255
                g = 128
                b = 0
                N = 1
                c = 0
            Case 7 'CHOTS | Dios
                r = 0
                g = 240
                b = 0
                N = 1
                c = 0
            Case 8 'CHOTS | Semi
                r = 255
                g = 255
                b = 128
                N = 1
                c = 0
            Case 9 'CHOTS | Infon
                r = 65
                g = 190
                b = 156
                N = 1
                c = 0
            Case 10 'CHOTS | Ejecucion
                r = 130
                g = 130
                b = 130
                N = 1
                c = 0
            Case 11 'CHOTS | Party
                r = 255
                g = 180
                b = 255
                N = 0
                c = 0
            Case 12 'CHOTS | Veneno
                r = 0
                g = 255
                b = 0
                N = 0
                c = 0
            Case 13 'CHOTS | Guild
                r = 255
                g = 255
                b = 255
                N = 1
                c = 0
            Case 14 'CHOTS | Server
                r = 0
                g = 185
                b = 0
                N = 0
                c = 0
            Case 15 'CHOTS | Guildmsj
                r = 228
                g = 199
                b = 27
                N = 0
                c = 0
            Case 16 'CHOTS | Consejo
                r = 130
                g = 130
                b = 255
                N = 1
                c = 0
            Case 17 'CHOTS | Consejocaos
                r = 255
                g = 60
                b = 0
                N = 1
                c = 0
            Case 18 'CHOTS | Consejovesa
                r = 0
                g = 200
                b = 255
                N = 1
                c = 0
            Case 19 'CHOTS | Consejocaosvesa
                r = 255
                g = 100
                b = 0
                N = 1
                c = 0
            Case 20 'CHOTS | Oro
                r = 255
                g = 255
                b = 0
                N = 1
                c = 0
            Case 21 'CHOTS | Celestenegrita
                r = 0
                g = 128
                b = 255
                N = 1
                c = 0
            Case 22 'CHOTS | Azul
                r = 0
                g = 0
                b = 255
                N = 1
                c = 0
            Case 23 'CHOTS | GM
                r = 0
                g = 255
                b = 0
                N = 1
                c = 0
            Case 24 'CHOTS | Troforo
                r = 233
                g = 198
                b = 1
                N = 1
                c = 0
            Case 25 'CHOTS | Trofplata
                r = 196
                g = 198
                b = 196
                N = 1
                c = 0
            Case 26 'CHOTS | Casamiento
                r = 255
                g = 55
                b = 155
                N = 1
                c = 0
            Case 27 'CHOTS | Duelos
                r = 128
                g = 64
                b = 64
                N = 1
                c = 0
            Case 28 'CHOTS | Hogar
                r = 128
                g = 128
                b = 255
                N = 1
                c = 0
            Case 29 'CHOTS | Invocaciones
                r = 172
                g = 172
                b = 172
                N = 1
                c = 0
            Case 30 'CHOTS | Torneos auto
                r = 0
                g = 74
                b = 149
                N = 1
                c = 0
            Case 31 'CHOTS | Monturas
                r = 2
                g = 134
                b = 45
                N = 1
                c = 0
            Case 32 'CHOTS | Guerras
                r = 235
                g = 235
                b = 188
                N = 1
                c = 0
            Case 33 'BysNacK | Chat privado
                r = 175
                g = 15
                b = 150
                N = 1
                c = 0
            Case Else
                r = 65
                g = 190
                b = 156
                N = 0
                c = 0
        End Select

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, semen, 126), r, g, b, N, c
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, semen, 126)
                End If
            End If

            Exit Sub
        
        Case "!G"
            'CHOTS | Guerras
            semen = Right$(semen, Len(semen) - 2)
            DialogosClanes.PushBackText semen
            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                semen = Right$(semen, Len(semen) - 2)
                frmMensaje.msg.Caption = semen
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            semen = Right$(semen, Len(semen) - 2)
            Userindex = Val(semen)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            semen = Right$(semen, Len(semen) - 2)
            UserCharIndex = Val(semen)
            UserPos = charlist(UserCharIndex).Pos
            frmMain.lblCord.Caption = UserMap & " | " & UserPos.x & " | " & UserPos.y
            Exit Sub
        Case ServerPackages.crearNpc
            semen = Right$(semen, Len(semen) - 2)
            charindex = ReadField(4, semen, 44)
            x = ReadField(5, semen, 44)
            y = ReadField(6, semen, 44)
            
            charlist(charindex).fX = Val(ReadField(9, semen, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(10, semen, 44))
            charlist(charindex).Nombre = ReadField(12, semen, 44)
            'CHOTS | Agrego el clan aca
            If InStr(charlist(charindex).Nombre, "<") > 0 And InStr(charlist(charindex).Nombre, ">") > 0 Then
                charlist(charindex).Clan = mid(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<"))
                charlist(charindex).Nombre = Left(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<") - 1)
            Else
                charlist(charindex).Clan = ""
            End If
            charlist(charindex).Criminal = Val(ReadField(13, semen, 44))
            charlist(charindex).priv = Val(ReadField(14, semen, 44))
            'CHOTS | Optimizamos colores
            Call SetNameColor(charindex)
            Call MakeChar(charindex, ReadField(1, semen, 44), ReadField(2, semen, 44), ReadField(3, semen, 44), x, y, Val(ReadField(7, semen, 44)), Val(ReadField(8, semen, 44)), Val(ReadField(11, semen, 44)))
            charlist(charindex).BodyNum = ReadField(1, semen, 44)
            Call RefreshAllChars
            Exit Sub
        Case ServerPackages.crearChar
            semen = Right$(semen, Len(semen) - 2)
            charindex = ReadField(4, semen, 44)
            x = ReadField(5, semen, 44)
            y = ReadField(6, semen, 44)
            
            charlist(charindex).fX = Val(ReadField(9, semen, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(10, semen, 44))
            charlist(charindex).Nombre = ReadField(12, semen, 44)
            'CHOTS | Agrego el clan aca
            If InStr(charlist(charindex).Nombre, "<") > 0 And InStr(charlist(charindex).Nombre, ">") > 0 Then
                charlist(charindex).Clan = mid(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<"))
                charlist(charindex).Nombre = Left(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<") - 1)
            Else
                charlist(charindex).Clan = ""
            End If
            charlist(charindex).Criminal = Val(ReadField(13, semen, 44))
            charlist(charindex).priv = Val(ReadField(14, semen, 44))
            'CHOTS | Optimizamos colores
            Call SetNameColor(charindex)
            Call MakeChar(charindex, ReadField(1, semen, 44), ReadField(2, semen, 44), ReadField(3, semen, 44), x, y, Val(ReadField(7, semen, 44)), Val(ReadField(8, semen, 44)), Val(ReadField(11, semen, 44)))
            Call RefreshAllChars
            Exit Sub
            
        Case ServerPackages.borrarChar
            semen = Right$(semen, Len(semen) - 2)
            Call EraseChar(Val(semen))
            Call Dialogos.QuitarDialogo(Val(semen))
            Call RefreshAllChars
            Exit Sub
        Case ServerPackages.moverPersonaje
            semen = Right$(semen, Len(semen) - 2)
            charindex = Val(ReadField(1, semen, 44))
            
            If charlist(charindex).fX >= 40 And charlist(charindex).fX <= 49 Then   'si esta meditando
                charlist(charindex).fX = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            If charlist(charindex).priv = 0 Or charlist(charindex).priv = 5 Or charlist(charindex).priv = 6 Then
                Call DoPasosFx(charindex)
            End If
            
            Call MoveCharbyPos(charindex, ReadField(2, semen, 44), ReadField(3, semen, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            semen = Right$(semen, Len(semen) - 2)
            
            charindex = Val(ReadField(1, semen, 44))
            charlist(charindex).muerto = Val(ReadField(3, semen, 44)) = 500
            charlist(charindex).Body = BodyData(Val(ReadField(2, semen, 44)))
            charlist(charindex).Head = HeadData(Val(ReadField(3, semen, 44)))
            charlist(charindex).Heading = Val(ReadField(4, semen, 44))
            charlist(charindex).fX = Val(ReadField(7, semen, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(8, semen, 44))
            tempint = Val(ReadField(5, semen, 44))
            If tempint <> 0 Then charlist(charindex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, semen, 44))
            If tempint <> 0 Then charlist(charindex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, semen, 44))
            If tempint <> 0 Then charlist(charindex).Casco = CascoAnimData(tempint)

            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            semen = Right$(semen, Len(semen) - 2)
            x = Val(ReadField(2, semen, 44))
            y = Val(ReadField(3, semen, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(x, y).ObjGrh.GrhIndex = Val(ReadField(1, semen, 44))
            InitGrh MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            semen = Right$(semen, Len(semen) - 2)
            x = Val(ReadField(1, semen, 44))
            y = Val(ReadField(2, semen, 44))
            MapData(x, y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            'Dim b As Byte
            semen = Right$(semen, Len(semen) - 2)
            MapData(Val(ReadField(1, semen, 44)), Val(ReadField(2, semen, 44))).Blocked = Val(ReadField(3, semen, 44))
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            semen = Right$(semen, Len(semen) - 2)
            currentMidi = Val(ReadField(1, semen, 45))
                If currentMidi <> 0 Then
                    semen = Right$(semen, Len(semen) - Len(ReadField(1, semen, 45)))
                    If Len(semen) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(semen, Len(semen) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
                semen = Right$(semen, Len(semen) - 2)
                 Call Audio.PlayWave(semen & ".wav")
            Exit Sub
        Case "PT" 'CHOTS | Puntos de Clan
            semen = Right$(semen, Len(semen) - 2)
            Dim puntos As Integer
            Dim Miembros As Byte
            Dim NombresMiembros As String
            
            puntos = Val(ReadField(1, semen, 44))
            Miembros = Val(ReadField(2, semen, 44))
            
            
            frmPuntos.lblPuntos.Caption = puntos
            If Miembros <> 0 Then
                
                NombresMiembros = ReadField(3, semen, 44)
                For i = 1 To Miembros
                    frmPuntos.lstClan.AddItem ReadField(i, NombresMiembros, Asc("@"))
                Next i
                frmPuntos.lstClan.listIndex = 0
            Else
                frmPuntos.lstClan.Enabled = False
                frmPuntos.transferir.Enabled = False
            End If
            frmPuntos.Show
            Exit Sub
        Case "GL" 'Lista de guilds
            semen = Right$(semen, Len(semen) - 2)
            Call frmGuildAdm.ParseGuildList(semen)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
            
    Case "MÑ"
            semen = Right$(semen, Len(semen) - 2)
            Dim EspiadoMaxMan As Integer
            Dim EspiadoMinMan As Integer
            EspiadoMaxMan = Val(ReadField(2, semen, 44))
            EspiadoMinMan = Val(ReadField(1, semen, 44))
            frmEspia.man.Width = (EspiadoMinMan / EspiadoMaxMan) * 4080
            frmEspia.lblMan.Caption = EspiadoMinMan & "/" & EspiadoMaxMan
            Exit Sub

    Case "MB" 'CHOTS | Toma Poción Azul
            Call AddMinManaPercentage(5)
            Exit Sub

    Case "MN"
            semen = Right$(semen, Len(semen) - 2)
            SetMana (Val(semen))
            Exit Sub
        Case "CÑ"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)

        Case ServerPackages.validarCliente
            semen = Right$(semen, Len(semen) - 3)
            RandomCode = semen
            UseNum = CByte(Right$(RandomCode, 1))
            UseAcum = RandomCode
            RandomCode = RandomCodeEncrypt(RandomCode)
            CargarCabezas

            If EstadoLogin = E_MODO.BorrarPj Then
                Call VaginaJugosa(ClientPackages.borrarPersonaje & frmBorrar.txtNombre.Text & "," & frmBorrar.txtEmail.Text & "," & frmBorrar.txtPasswd.Text & "," & RandomCode)
            ElseIf EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Then
                Call login(RandomCode)
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje1.Show vbModal
            ElseIf EstadoLogin = E_MODO.RecuperarPass Then
                Call VaginaJugosa(ClientPackages.recuperarPersonaje & frmRecuperar.txtNombre.Text & "," & frmRecuperar.txtEmail.Text & "," & RandomCode)
            End If
            Exit Sub
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
            Else
                If bLluvia(UserMap) Then
                    'Stop playing the rain sound
                    Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = 0
                    If bTecho Then
                        Call Audio.PlayWave("lluviainend.wav", LoopStyle.Disabled)
                    Else
                        Call Audio.PlayWave("lluviaoutend.wav", LoopStyle.Disabled)
                    End If
                    frmMain.IsPlaying = PlayLoop.plNone
                End If
                bRain = False
            End If
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            semen = Right$(semen, Len(semen) - 3)
            Call Dialogos.QuitarDialogo(Val(semen))
            Exit Sub
        Case "GUE" 'CHOTS | Crea Guerra
            semen = Right$(semen, Len(semen) - 3)
            Dim numeroSala As Byte
            Dim nombreSala As String
            numeroSala = Val(ReadField(1, semen, 44))
            nombreSala = ReadField(2, semen, 44)
            frmGuerra.Show
            Call frmGuerra.setNumeroSala(numeroSala)
            frmGuerra.lblNombreSala.Caption = nombreSala
            
            Exit Sub
        Case "CXF"                  ' >>>>> Mostrar FX sobre Personaje :: "CFX"
            semen = Right$(semen, Len(semen) - 3)
            charindex = Val(ReadField(1, semen, 44))
            charlist(charindex).fX = Val(ReadField(2, semen, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(3, semen, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim n1 As String, n2 As String
            semen = Right$(semen, Len(semen) - 3)
            n1 = ReadField(2, semen, 176)
            n2 = ReadField(1, semen, 176)
            frmMSG.CrearGMmSg n1, n2
            frmMSG.Show , frmMain
            Exit Sub
        Case "ENP" 'CHOTS | Esta en Party
            enParty = Not enParty
            Exit Sub
        Case "TAU" 'CHOTS | Crea Torneo AUTO
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "[TORNEO AUTOMÁTICO 1vs1]", 0, 74, 149, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Nivel mínimo: 46", 0, 74, 149, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Nivel máximo: 54", 0, 74, 149, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Cupo máximo: " & semen, 0, 74, 149, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Tipea /PARTICIPAR para inscribirte", 0, 74, 149, True, False, False)
            Exit Sub
        Case "DRR"
            semen = Right$(semen, Len(semen) - 3)
            Amarilla = Val(ReadField(1, semen, 44))
            Verde = Val(ReadField(2, semen, 44))
            frmMain.lblAgi.Caption = Amarilla
            frmMain.lblfuerza.Caption = Verde
            Exit Sub
        Case "CHV" 'CHOTS | Recibe Vestimentas
            semen = Right$(semen, Len(semen) - 3)
            ArmorMin = Val(ReadField(1, semen, 44))
            ArmorMax = Val(ReadField(2, semen, 44))
            frmMain.LblArmor.Caption = ArmorMin & "/" & ArmorMax
            Exit Sub
        Case "CHA" 'CHOTS | Recibe Armas
            semen = Right$(semen, Len(semen) - 3)
            ArmaMin = Val(ReadField(1, semen, 44))
            ArmaMax = Val(ReadField(2, semen, 44))
            frmMain.LblArma.Caption = ArmaMin & "/" & ArmaMax
            Exit Sub
        Case "CHE" 'CHOTS | Recibe Escu
            semen = Right$(semen, Len(semen) - 3)
            EscuMin = Val(ReadField(1, semen, 44))
            EscuMax = Val(ReadField(2, semen, 44))
            frmMain.LblEscudo.Caption = EscuMin & "/" & EscuMax
            Exit Sub
        Case "CHC" 'CHOTS | Recibe CASCO
            semen = Right$(semen, Len(semen) - 3)
            CascMin = Val(ReadField(1, semen, 44))
            CascMax = Val(ReadField(2, semen, 44))
            frmMain.LblCasc.Caption = CascMin & "/" & CascMax
            Exit Sub
        Case "CHD" 'CHOTS | Recibe Def Mag
            semen = Right$(semen, Len(semen) - 3)
            MagMin = Val(ReadField(1, semen, 44))
            MagMax = Val(ReadField(2, semen, 44))
            frmMain.lblmagica.Caption = MagMin & "/" & MagMax
            Exit Sub
        Case "ARX" 'CHOTS | Papio
            frmMain.LblArmor.Caption = "0/0"
            frmMain.LblArma.Caption = "0/0"
            frmMain.LblEscudo.Caption = "0/0"
            frmMain.LblCasc.Caption = "0/0"
            frmMain.lblmagica.Caption = "0/0"
            Exit Sub
        Case "ARM"
            semen = Right$(semen, Len(semen) - 3)
            ArmaMin = Val(ReadField(1, semen, 44))
            ArmaMax = Val(ReadField(2, semen, 44))
            ArmorMin = Val(ReadField(3, semen, 44))
            ArmorMax = Val(ReadField(4, semen, 44))
            EscuMin = Val(ReadField(5, semen, 44))
            EscuMax = Val(ReadField(6, semen, 44))
            CascMin = Val(ReadField(7, semen, 44))
            CascMax = Val(ReadField(8, semen, 44))
            MagMin = Val(ReadField(9, semen, 44))
            MagMax = Val(ReadField(10, semen, 44))
            frmMain.LblArmor.Caption = ArmorMin & "/" & ArmorMax
            frmMain.LblArma.Caption = ArmaMin & "/" & ArmaMax
            frmMain.LblEscudo.Caption = EscuMin & "/" & EscuMax
            frmMain.LblCasc.Caption = CascMin & "/" & CascMax
            frmMain.lblmagica.Caption = MagMin & "/" & MagMax
            Exit Sub
        Case "GEM"
            semen = Right$(semen, Len(semen) - 3)
            frmGema.habLbl.Caption = ReadField(1, semen, 44)
            frmGema.crLbl.Caption = ReadField(2, semen, 44) & " - " & ReadField(3, semen, 44)
            frmGema.Show
            Exit Sub
        Case "RUN" 'CHOTS | Abre el cambiador de Runas
            frmRunas.Show
            frmRunas.lstObjetos.listIndex = 0
            Exit Sub
        Case "TRD" 'CHOTS | Abre el cambiador de Premios
            frmTrade.Show vbModal, frmMain
            frmTrade.lstObjetos.listIndex = 0
            Exit Sub
        Case "PST" 'CHOTS | Abre el cambiador de Puntos
            semen = Right$(semen, Len(semen) - 3)
            frmPts.Show
            frmPts.lblPts.Caption = Val(semen)
            Exit Sub
        Case "UON" 'CHOTS | Users Online, Records, miembros del clan
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "Usuarios Online: " & Val(ReadField(1, semen, 44)), 65, 190, 156, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Record de Usuarios: " & Val(ReadField(2, semen, 44)), 65, 190, 156, False, False)
            If ReadField(3, semen, 44) <> "" Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Compañeros de tu clan conectados:" & ReadField(3, semen, 44), 228, 199, 27, False, False)
            End If
        Case "LEV" 'CHOTS | Sube de nivel
            semen = Right$(semen, Len(semen) - 3)
            Call Audio.PlayWave("6.wav")
            Call AddtoRichTextBox(frmMain.RecTxt, "Has subido de Nivel!", 65, 190, 156, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & ReadField(1, semen, 64) & " skillpoints", 65, 190, 156, False, False)
            SkillPoints = SkillPoints + Val(ReadField(1, semen, 64))
            
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & ReadField(2, semen, 64) & " puntos de Stamina", 65, 190, 156, False, False)
            Call AddMaxSta(Val(ReadField(2, semen, 64)))
            
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & ReadField(3, semen, 64) & " puntos de Maná", 65, 190, 156, False, False)
            Call AddMaxMana(Val(ReadField(3, semen, 64)))
            
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & ReadField(4, semen, 64) & " puntos de Vida", 65, 190, 156, True, False)
            Call AddMaxHp(Val(ReadField(4, semen, 64)))
            
            Call SetElu(Val(ReadField(7, semen, 64)), False)
            Call SetExp(Val(ReadField(6, semen, 64)), False)
            Call AddLevel(1)

            Call AddtoRichTextBox(frmMain.RecTxt, "Tu golpe mínimo aumento en " & ReadField(5, semen, 64) & " puntos; tu golpe máximo aumento en " & ReadField(5, semen, 64) & " puntos", 65, 190, 156, False, False)
            Exit Sub
        Case "MAÑ" 'CHOTS | Mata un user
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "Has matado a " & ReadField(1, semen, 44) & "!", 255, 0, 0, True, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & ReadField(2, semen, 44) & " puntos de experiencia", 255, 0, 0, True, False)
            Exit Sub
        Case "CNC" 'CHOTS | Conecta
            semen = Right$(semen, Len(semen) - 3)
            Call SetHp(Val(ReadField(2, semen, 64)), Val(ReadField(1, semen, 64)))
            Call SetMana(Val(ReadField(4, semen, 64)), Val(ReadField(3, semen, 64)))
            Call SetStamina(Val(ReadField(6, semen, 64)), Val(ReadField(5, semen, 64)))

            ArmaMin = Val(ReadField(11, semen, 64))
            ArmaMax = Val(ReadField(12, semen, 64))
            ArmorMin = Val(ReadField(13, semen, 64))
            ArmorMax = Val(ReadField(14, semen, 64))
            EscuMin = Val(ReadField(15, semen, 64))
            EscuMax = Val(ReadField(16, semen, 64))
            CascMin = Val(ReadField(17, semen, 64))
            CascMax = Val(ReadField(18, semen, 64))
            MagMin = Val(ReadField(19, semen, 64))
            MagMax = Val(ReadField(20, semen, 64))

            Call SetSed(Val(ReadField(21, semen, 64)))
            Call SetHambre(Val(ReadField(22, semen, 64)))

            Amarilla = Val(ReadField(23, semen, 64))
            Verde = Val(ReadField(24, semen, 64))

            Call SetElu(Val(ReadField(9, semen, 64)), False)
            Call SetLevel(Val(ReadField(8, semen, 64)), False)
            Call SetExp(Val(ReadField(10, semen, 64)), True)
            Call SetGold(Val(ReadField(7, semen, 64)))
            
            frmMain.LblArmor.Caption = ArmorMin & "/" & ArmorMax
            frmMain.LblArma.Caption = ArmaMin & "/" & ArmaMax
            frmMain.LblEscudo.Caption = EscuMin & "/" & EscuMax
            frmMain.LblCasc.Caption = CascMin & "/" & CascMax
            frmMain.lblmagica.Caption = MagMin & "/" & MagMax
        
            frmMain.lblAgi.Caption = Amarilla
            frmMain.lblfuerza.Caption = Verde
            
            Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Norte pertenece al clan " & ReadField(25, semen, 64) & ".", 230, 189, 43, True, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Oeste pertenece al clan " & ReadField(26, semen, 64) & ".", 230, 189, 43, True, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Este pertenece al clan " & ReadField(27, semen, 64) & ".", 230, 189, 43, True, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Sur pertenece al clan " & ReadField(28, semen, 64) & ".", 230, 189, 43, True, False)
            SkillPoints = Val(ReadField(29, semen, 64))
        
            Exit Sub
        Case "EXT"                  ' >>>>> Actualiza Estadisticas de Usuario :: "EST"
            semen = Right$(semen, Len(semen) - 3)

            Call SetHp(Val(ReadField(2, semen, 44)), Val(ReadField(1, semen, 44)))
            Call SetMana(Val(ReadField(4, semen, 44)), Val(ReadField(3, semen, 44)))
            Call SetStamina(Val(ReadField(6, semen, 44)), Val(ReadField(5, semen, 44)))

            Call SetElu(Val(ReadField(9, semen, 44)), False)
            Call SetLevel(Val(ReadField(8, semen, 44)), False)
            Call SetExp(Val(ReadField(10, semen, 44)), True)
            Call SetGold(Val(ReadField(7, semen, 44)))
        
            Exit Sub
        Case "MUE" '"Muere o Renace"
            semen = Right$(semen, Len(semen) - 3)
            SetHp (Val(ReadField(1, semen, 44)))
            SetStamina (Val(ReadField(2, semen, 44)))
            Exit Sub
        
        Case "CHW" 'CHOTS | Has Ganado x puntos de experienca
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & semen & " puntos de experiencia!", 255, 0, 0, True, False)
            Exit Sub
            
        Case "CHO" 'CHOTS | Has Ganado x monedas de oro
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & semen & " Monedas de Oro!", 255, 0, 0, True, False)
            Exit Sub
            
        Case "VHÑ" 'CHOTS | Espiar HP
            semen = Right$(semen, Len(semen) - 3)
            Dim EspiadoMaxHp As Integer
            Dim EspiadoMinHp As Integer
            EspiadoMinHp = Val(ReadField(1, semen, 44))
            EspiadoMaxHp = Val(ReadField(2, semen, 44))
            
            frmEspia.hp.Width = (EspiadoMinHp / EspiadoMaxHp) * 4080
            frmEspia.lblHp.Caption = EspiadoMinHp & "/" & EspiadoMaxHp
            Exit Sub

        Case "VHP" '"VID"
            semen = Right$(semen, Len(semen) - 3)
            SetHp (Val(semen))
            Exit Sub
            
        Case "FIX" 'CHOTS | Fixture
            semen = Right$(semen, Len(semen) - 3)
            Call frmFixture.Show
            frmFixture.cargarLabels (semen)
            Exit Sub
            
        Case ServerPackages.recibeDados
            semen = Right$(semen, Len(semen) - 3)
            With frmCrearPersonaje1
                If .Visible Then
                    .lbFuerza.Caption = 10 + Val(mid$(semen, 1, 1))
                    .lbAgilidad.Caption = 10 + Val(mid$(semen, 2, 1))
                    .lbInteligencia.Caption = 10 + Val(mid$(semen, 3, 1))
                    .lbCarisma.Caption = 10 + Val(mid$(semen, 4, 1))
                    .lbConstitucion.Caption = 10 + Val(mid$(semen, 5, 1))
                End If
                
                Call .SetDadosFinal
            End With
            Exit Sub

        Case "VHR" 'CHOTS | Toma Poción roja
            Call AddMinHp(30)
            Exit Sub
        Case "NPÑ" 'CHOTS | Optimización de clicks a NPCs
            semen = Right$(semen, Len(semen) - 3)
            Dim npc As String
            Dim estado As String
            Dim estadoIndex As Byte
            Dim maestro As String
            Dim Aconsola As String
            npc = ReadField(1, semen, 44)
            estadoIndex = CByte(ReadField(2, semen, 44))
            maestro = ReadField(3, semen, 44)
            
            Select Case estadoIndex
                Case 0: estado = "Intacto"
                Case 1: estado = "Sano"
                Case 2: estado = "Levemente Herido"
                Case 3: estado = "Herido"
                Case 4: estado = "Malherido"
                Case 5: estado = "Muy Malherido"
                Case 6: estado = "Casi Muerto"
                Case 7: estado = "Agonizando"
                Case 8: estado = "Dudoso"
                Case Else: estado = "Error"
            End Select
            
            Aconsola = "(" & estado & ") " & npc
            
            If Len(maestro) >= 1 Then
                Aconsola = Aconsola & " es mascota de " & maestro & "."
            Else
                Aconsola = Aconsola & "."
            End If
                
            Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 65, 190, 156, False, False, False)
            Exit Sub
            
            
Case "NPZ" 'CHOTS | Optimización de clicks a NPCs
            semen = Right$(semen, Len(semen) - 3)
            Dim minvida As String
            Dim Maxvida As String
            npc = ReadField(1, semen, 44)
            minvida = ReadField(2, semen, 44)
            Maxvida = ReadField(3, semen, 44)
            maestro = ReadField(4, semen, 44)
            Aconsola = "(" & minvida & "/" & Maxvida & ") " & npc
            
            If Len(maestro) >= 1 Then
                Aconsola = Aconsola & " es mascota de " & maestro & "."
            Else
                Aconsola = Aconsola & "."
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 65, 190, 156, False, False, False)
            Exit Sub
            
            
Case "VES" 'CHOTS | Optimización de clicks a usuarios
            semen = Right$(semen, Len(semen) - 3)
            Dim Nick As String
            Dim Newb As Byte
            Dim Facc As Byte
            Dim Tit As Byte
            Dim Clan As String
            Dim Casado As Byte
            Dim Pareja As String
            Dim desc As String
            Dim Pert As Byte
            Dim status As Byte
            
            Aconsola = "Ves a "
            
            Nick = ReadField(1, semen, 44)
            Newb = Val(ReadField(2, semen, 44))
            Facc = Val(ReadField(3, semen, 44))
            Tit = Val(ReadField(4, semen, 44))
            Clan = ReadField(5, semen, 44)
            Casado = Val(ReadField(6, semen, 44))
            Pareja = ReadField(7, semen, 44)
            desc = ReadField(8, semen, 44)
            Pert = Val(ReadField(9, semen, 44))
            status = Val(ReadField(10, semen, 44))
            
            Aconsola = Aconsola & Nick & " "
            Aconsola = Aconsola & IIf(Newb = 1, "<NEWBIE> ", "")
            Aconsola = Aconsola & IIf(Facc = 1, "<Armada Real> ", IIf(Facc = 2, "<Legion Oscura> ", ""))
            
            If Tit > 0 Then
                If Facc = 1 Then
                    Aconsola = Aconsola & "<" & TituloReal(Tit) & "> "
                Else
                    Aconsola = Aconsola & "<" & TituloCaos(Tit) & "> "
                End If
            End If
            
            If Clan <> 0 Then Aconsola = Aconsola & "<" & Clan & "> "
            
            If Casado = 1 Then
                Aconsola = Aconsola & "<Casado con " & Pareja & "> "
            ElseIf Casado = 2 Then
                Aconsola = Aconsola & "<Casada con " & Pareja & "> "
            Else
                Aconsola = Aconsola
            End If
            
            If desc <> 0 Then Aconsola = Aconsola & "- " & desc & " "
            
            If Pert <> 0 Then
            
                If Pert = 1 Then
                    Aconsola = Aconsola & "[Consejo de Banderbill]"
                    Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 0, 200, 255, True, False)
                ElseIf Pert = 2 Then
                    Aconsola = Aconsola & "[Consejo de Arghal]"
                    Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 255, 100, 0, True, False)
                End If
                
            Else
            
                If status = 0 Then
                    Aconsola = Aconsola & "<CIUDADANO>"
                    Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 0, 0, 255, True, False)
                ElseIf status = 1 Then
                    Aconsola = Aconsola & "<CRIMINAL>"
                    Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 255, 0, 0, True, False)
                Else
                    Aconsola = Aconsola & "<GAME MASTER>"
                    Call AddtoRichTextBox(frmMain.RecTxt, Aconsola, 0, 255, 0, True, False)
                End If
                
            End If
            
            
    Exit Sub
            
                
                

    Case "STT" '"STA"
            semen = Right$(semen, Len(semen) - 3)
            SetStamina (Val(semen))
            Exit Sub
    
    Case "PRE" 'CHOTS | Recibe mensaje de castillos
            If Not hayCastillo Then Exit Sub
            Dim Texto As Integer
            Dim CClan As String
            Dim puntoss As String
            Dim Casti As String
            semen = Right$(semen, Len(semen) - 3)
            Texto = Int(ReadField(1, semen, 44))
            If Texto >= 75 Then
                Call txtReceivedB(Texto)
            ElseIf Texto < 17 Then
                CClan = (ReadField(2, semen, 44))
                Call txtReceived(Texto, CClan)
            ElseIf Texto < 20 Or Texto = 25 Then
                Call txtReceived(Texto)
            ElseIf Texto = 26 Then
                Casti = (ReadField(2, semen, 44))
                Call txtReceived(Texto, Casti)
            ElseIf Texto = 27 Then
                Casti = ReadField(2, semen, 44)
                puntoss = ReadField(3, semen, 44)
                Call txtReceived(Texto, Casti, puntoss)
            Else
                CClan = ReadField(2, semen, 44)
                Casti = ReadField(3, semen, 44)
                Call txtReceived(Texto, CClan, Casti)
            End If
            Exit Sub
    
    Case "OLD" '"ORO"
            semen = Right$(semen, Len(semen) - 3)
            AddGold (Val(semen))
            Exit Sub
            
    Case "SKI" 'CHOTS | Sube Skill
            Dim Skill As String
            Dim Cant As Byte
            semen = Right$(semen, Len(semen) - 3)
            Skill = Val(ReadField(1, semen, 44))
            Cant = CByte(Val(ReadField(2, semen, 44)))
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & Cant & " pts.", 65, 190, 156, False, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, Mensaje25, 255, 0, 0, True, False, False)
            Exit Sub
    
    Case "MIN" 'CHOTS | Extrae Minerales
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has extraído algunos minerales! " & "(" & semen & ")", 65, 190, 156, False, False, False)
            Exit Sub
            
            
    Case "RMN" 'CHOTS | Recupera Maná
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has recuperado " & semen & " puntos de mana!", 65, 190, 156, False, False, False)
            Exit Sub
            
            
    Case "LEÑ" 'CHOTS | Extrae Leña
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has conseguido algo de leña! " & "(" & semen & ")", 65, 190, 156, False, False, False)
            Exit Sub
            
    Case "CON" 'CHOTS | Conecta clan
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, semen & " Conectó", 255, 255, 255, True, False, False)
            Exit Sub
            
    Case "VOT" 'CHOTS | Hay Elecciones
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, semen & " Conectó", 255, 255, 255, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Hoy es la votacion para elegir un nuevo líder para el clan!!.", 255, 255, 255, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", 255, 255, 255, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Para votar escribe /VOTO NICKNAME.", 255, 255, 255, True, False, False)
            Call AddtoRichTextBox(frmMain.RecTxt, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", 255, 255, 255, True, False, False)
            Exit Sub
            
    Case "DES" 'CHOTS | Desconecta clan
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, semen & " Desconectó", 255, 255, 255, True, False, False)
            Exit Sub
            
    Case "CHL" 'CHOTS | Extrae Chalas
            semen = Right$(semen, Len(semen) - 3)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has conseguido algunas raíces! " & "(" & semen & ")", 65, 190, 156, False, False, False)
            Exit Sub
            
    Case "ESP" '"EXP"
            semen = Right$(semen, Len(semen) - 3)
            Call SetExp(Val(semen))
            Exit Sub
    
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            semen = Right$(semen, Len(semen) - 3)
            UsingSkill = Val(semen)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Botanica
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case CapturarNpc
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_CAPTURARNPC, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "IIH" 'CHOTS | Inicializa Inventario y Hechizos
            
            frmMain.hlst.Clear
            
            For slot = 1 To MAX_INVENTORY_SLOTS
                Call Inventario.SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, "(none)")
            Next slot
            
            For slot = 1 To MAXHECHI
                UserHechizos(slot) = 0
                frmMain.hlst.AddItem "(None)"
            Next slot
            
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            semen = Right$(semen, Len(semen) - 3)
            Dim CHOTS As Integer
            CHOTS = ReadField(2, semen, 44)
            slot = ReadField(1, semen, 44)
            If CHOTS <> 0 Then
                Call Inventario.SetItem(slot, ReadField(2, semen, 44), ReadField(4, semen, 44), ReadField(5, semen, 44), Val(ReadField(6, semen, 44)), Val(ReadField(7, semen, 44)), _
                                    Val(ReadField(8, semen, 44)), Val(ReadField(9, semen, 44)), Val(ReadField(10, semen, 44)), Val(ReadField(11, semen, 44)), ReadField(3, semen, 44))
            Else
                Call Inventario.SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, "(none)")
            End If
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBÑ"                 ' CHOTS | Inicializa inventario del banco
            semen = Right$(semen, Len(semen) - 3)
            For i = 1 To MAX_BANCOINVENTORY_SLOTS
                UserBancoInventory(i).OBJIndex = 0
                UserBancoInventory(slot).Name = "(none)"
                UserBancoInventory(slot).Amount = 0
                UserBancoInventory(slot).GrhIndex = 0
                UserBancoInventory(slot).OBJType = 0
                UserBancoInventory(slot).MaxHit = 0
                UserBancoInventory(slot).MinHit = 0
                UserBancoInventory(slot).Def = 0
            Next i
            
            Exit Sub
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            semen = Right$(semen, Len(semen) - 3)
            slot = ReadField(1, semen, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, semen, 44)
            
            If Val(ReadField(2, semen, 44)) = 0 Then
               UserBancoInventory(slot).Name = "(none)"
               UserBancoInventory(slot).Amount = 0
               UserBancoInventory(slot).GrhIndex = 0
               UserBancoInventory(slot).OBJType = 0
               UserBancoInventory(slot).MaxHit = 0
               UserBancoInventory(slot).MinHit = 0
               UserBancoInventory(slot).Def = 0
            Else
                UserBancoInventory(slot).Name = ReadField(3, semen, 44)
                UserBancoInventory(slot).Amount = ReadField(4, semen, 44)
                UserBancoInventory(slot).GrhIndex = Val(ReadField(5, semen, 44))
                UserBancoInventory(slot).OBJType = Val(ReadField(6, semen, 44))
                UserBancoInventory(slot).MaxHit = Val(ReadField(7, semen, 44))
                UserBancoInventory(slot).MinHit = Val(ReadField(8, semen, 44))
                UserBancoInventory(slot).Def = Val(ReadField(9, semen, 44))
            End If
            
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            semen = Right$(semen, Len(semen) - 3)
            slot = ReadField(1, semen, 44)
            If UCase$(ReadField(2, semen, 44)) = "N" Then
                UserHechizos(slot) = 0
                If slot > frmMain.hlst.ListCount Then
                    frmMain.hlst.AddItem "(None)"
                Else
                    frmMain.hlst.List(slot - 1) = "(None)"
                End If
            Else
                UserHechizos(slot) = ReadField(2, semen, 44)
                If slot > frmMain.hlst.ListCount Then
                    frmMain.hlst.AddItem ReadField(3, semen, 44)
                Else
                    frmMain.hlst.List(slot - 1) = ReadField(3, semen, 44)
                End If
            End If
            Exit Sub
        Case "LAH"
            semen = Right$(semen, Len(semen) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, semen, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, semen, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            semen = Right$(semen, Len(semen) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, semen, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, semen, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "LGL"
            semen = Right$(semen, Len(semen) - 3)
            
            For m = 0 To UBound(ObjDruida)
                ObjDruida(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, semen, 44)
                ObjDruida(m) = Val(ReadField(i + 1, semen, 44))
                If cad$ <> "" Then frmAlquimia.lstPociones.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            semen = Right$(semen, Len(semen) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, semen, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, semen, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "OBS" 'sastre
            semen = Right$(semen, Len(semen) - 3)
            
            For m = 0 To UBound(ObjSastre)
                ObjSastre(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, semen, 44)
                ObjSastre(m) = Val(ReadField(i + 1, semen, 44))
                If cad$ <> "" Then frmSastre.lstRopas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "SPL"
            semen = Right(semen, Len(semen) - 3)
            For i = 1 To Val(ReadField(1, semen, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, semen, 44)
            Next i
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "FPZ"
               Call VaginaJugosa("FPS" & FramesPerSec)
               Exit Sub
        Case "FPP"
               Call VaginaJugosa("FPI" & tAt & "," & tUs & "," & tComb & "," & tClick)
               Exit Sub
        Case "ERR"
            semen = Right$(semen, Len(semen) - 3)
            frmConnect.MousePointer = 1
            frmCrearPersonaje1.MousePointer = 1
            If Not frmCrearPersonaje1.Visible Then
#If UsarWrench = 1 Then
                frmMain.Socket1.Disconnect
#Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
#End If
            End If
            'If Not frmCrearPersonaje3.Visible = True Then
            'frmConnect.Label1.Caption = semen
            'frmConnect.Timer1.Enabled = True
            'Else
            MsgBox (semen)
            'End If
            frmConnect.MousePointer = 1
            Exit Sub
    End Select
    
    
    Select Case Left$(sData, 4)
        Case "SEGS" ' CHOTS | Manda todos los seguros cuando conectas
            semen = Right$(semen, Len(semen) - 4)
            '1=seguro
            '2=clan
            
            If Val(ReadField(1, semen, 44)) = 1 Then
                activarSeguro
            Else
                desactivarSeguro
            End If
            
            If Val(ReadField(2, semen, 44)) = 1 Then
                activarSeguroClan
            Else
                desactivarSeguroClan
            End If
            
            desactivarSeguroCaos
            desactivarSeguroResu
            
            Exit Sub
        Case "MATA" ' CHOTS | Matar Procesos
            Dim Procesoo As String
            semen = Right$(semen, Len(semen) - 4)
            Procesoo = ReadField(1, semen, 44)
            Call KillProcess(Procesoo)
            Exit Sub
        Case "PCGN" ' CHOTS | Poner Procesos en frm
            Dim Proceso As String
            Dim Nombre As String
            semen = Right$(semen, Len(semen) - 4)
            Proceso = ReadField(1, semen, 44)
            Nombre = ReadField(2, semen, 44)
            Call FrmProcesos.Show
            FrmProcesos.List1.AddItem Proceso
            FrmProcesos.Caption = "Procesos de " & Nombre
            FrmProcesos.Label1.Caption = Nombre
            Exit Sub
        Case "PCSS" ' CHOTS | Poner Prosesos en frm
            Dim Proseso As String
            Dim Nonbre As String
            Dim Peso As String
            Dim verssion As String
            
            semen = Right$(semen, Len(semen) - 4)
            Proseso = ReadField(1, semen, 44)
            Peso = ReadField(2, Proseso, 64)
            verssion = ReadField(3, Proseso, 64)
            Proseso = ReadField(1, Proseso, 64)
            Nonbre = ReadField(2, semen, 44)
            Call frmProsesos.Show
            
            With frmProsesos.FlxGd
                
                .ColAlignment(-1) = 1       'all Left alligned
                .TextMatrix(0, 1) = "Ruta"
                .TextMatrix(0, 2) = "Peso"
                .TextMatrix(0, 3) = "Version"
                
                .Row = 1
                .Col = 1
                .CellBackColor = &HC0FFFF   'lt. yellow
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = "Proc " + Str(.Row)
                .ColWidth(1) = 7000
                .TextMatrix(.Row, 1) = Proseso
                .TextMatrix(.Row, 2) = KiloBytes(Val(Peso))
                .TextMatrix(.Row, 3) = verssion
                .Refresh
            End With
            
            frmProsesos.Caption = "Procesos de " & Nonbre
            Exit Sub
        Case "PCCC" ' CHOTS | Poner Captions en frm
            Dim Caption As String
            Dim Nomvre As String
            semen = Right$(semen, Len(semen) - 4)
            Caption = ReadField(1, semen, 44)
            Nomvre = ReadField(2, semen, 44)
            Call frmCaptions.Show
            frmCaptions.List1.AddItem Caption
            frmCaptions.Caption = "Captions de " & Nomvre
            Exit Sub
        Case "PCCP" ' CHOTS | Ver Captions
            frmCaptions.List1.Clear
            frmCaptions.Caption = ""
            semen = Right$(semen, Len(semen) - 4)
            charindex = Val(ReadField(1, semen, 44))
            Call frmCaptions.Listar(charindex)
            Exit Sub
        Case "PCGR" ' CHOTS | Ver procesos
            FrmProcesos.List1.Clear
            FrmProcesos.Caption = ""
            semen = Right$(semen, Len(semen) - 4)
            charindex = Val(ReadField(1, semen, 44))
            Call enumProc(charindex)
            Exit Sub
        Case "PCSC" ' CHOTS | Ver prosesos
            frmProsesos.FlxGd.Clear
            frmProsesos.Caption = ""
            semen = Right$(semen, Len(semen) - 4)
            charindex = Val(ReadField(1, semen, 44))
            Call PROC(charindex)
            Exit Sub
        Case "PCFT" ' CHOTS | Ver Foto
            semen = Right$(semen, Len(semen) - 4)
            charindex = Val(ReadField(1, semen, 44))
            Call frmScreenshots.TakeAndUploadScreenshot(charindex)
            Exit Sub
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, semen, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim rr As RECT
            BackBufferSurface.BltColorFill rr, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            semen = Right$(semen, Len(semen) - 4)
            UserAtributos(1) = ReadField(1, semen, 44)
            UserAtributos(2) = ReadField(2, semen, 44)
            UserAtributos(3) = ReadField(3, semen, 44)
            UserAtributos(4) = ReadField(4, semen, 44)
            UserAtributos(5) = ReadField(5, semen, 44)
            
            frmCrearPersonaje1.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje1.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje1.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje1.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje1.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            semen = Right$(semen, Len(semen) - 4)
            Call InitCartel(ReadField(1, semen, 176), CInt(ReadField(2, semen, 176)))
            Exit Sub
        Case "NPCÑ"              ' CHOTS | Inicializa Inventario del NPC
            semen = Right(semen, Len(semen) - 4)
            For i = 1 To MAX_INVENTORY_SLOTS
                NPCInventory(i).Name = "Nada"
                NPCInventory(i).Amount = 0
                NPCInventory(i).Valor = 0
                NPCInventory(i).GrhIndex = 0
                NPCInventory(i).OBJIndex = 0
                NPCInventory(i).OBJType = 0
                NPCInventory(i).MaxHit = 0
                NPCInventory(i).MinHit = 0
                NPCInventory(i).Def = 0
                NPCInventory(i).C1 = 0
                NPCInventory(i).C2 = 0
                NPCInventory(i).C3 = 0
                NPCInventory(i).C4 = 0
                NPCInventory(i).C5 = 0
                NPCInventory(i).C6 = 0
                NPCInventory(i).C7 = 0
            Next i
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            semen = Right(semen, Len(semen) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, semen, 44)
            If UCase$(NPCInventory(NPCInvDim).Name) = "N" Then
                NPCInventory(NPCInvDim).Amount = 0
                NPCInventory(NPCInvDim).Valor = 0
                NPCInventory(NPCInvDim).GrhIndex = 0
                NPCInventory(NPCInvDim).OBJIndex = 0
                NPCInventory(NPCInvDim).OBJType = 0
                NPCInventory(NPCInvDim).MaxHit = 0
                NPCInventory(NPCInvDim).MinHit = 0
                NPCInventory(NPCInvDim).Def = 0
                NPCInventory(NPCInvDim).C1 = 0
                NPCInventory(NPCInvDim).C2 = 0
                NPCInventory(NPCInvDim).C3 = 0
                NPCInventory(NPCInvDim).C4 = 0
                NPCInventory(NPCInvDim).C5 = 0
                NPCInventory(NPCInvDim).C6 = 0
                NPCInventory(NPCInvDim).C7 = 0
                frmComerciar.List1(0).AddItem "Nada"
                Exit Sub
            End If
            NPCInventory(NPCInvDim).Amount = ReadField(2, semen, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, semen, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, semen, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, semen, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, semen, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, semen, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, semen, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, semen, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, semen, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, semen, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, semen, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, semen, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, semen, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, semen, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, semen, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            semen = Right$(semen, Len(semen) - 4)
            SetSed (Val(ReadField(1, semen, 44)))
            SetHambre (Val(ReadField(2, semen, 44)))
            Exit Sub
        Case "XEST" 'CHOTS | Full estadisticas
            semen = Right$(semen, Len(semen) - 4)

            ' CHOTS | Leo todas las stats y dps abro el frmEstadisticas
            ' Atrib, Fama, Skills, Stats

            'ATR, son siempre 5
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, semen, 44))
            Next i

            'FAMA
            UserReputacion.AsesinoRep = Val(ReadField(6, semen, 44))
            UserReputacion.BandidoRep = Val(ReadField(7, semen, 44))
            UserReputacion.BurguesRep = Val(ReadField(8, semen, 44))
            UserReputacion.LadronesRep = Val(ReadField(9, semen, 44))
            UserReputacion.NobleRep = Val(ReadField(10, semen, 44))
            UserReputacion.PlebeRep = Val(ReadField(11, semen, 44))
            UserReputacion.Promedio = Val(ReadField(12, semen, 44))

            'ESKILS, son 24
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(12 + i, semen, 44))
            Next i

            'MEST
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(37, semen, 44))
                .CriminalesMatados = Val(ReadField(38, semen, 44))
                .UsuariosMatados = Val(ReadField(39, semen, 44))
                .NpcsMatados = Val(ReadField(40, semen, 44))
                .Clase = ReadField(41, semen, 44)
                .PenaCarcel = Val(ReadField(42, semen, 44))
            End With

            frmEstadisticas2.Iniciar_Labels
            frmEstadisticas2.Show , frmMain
            Alocados = SkillPoints
            frmEstadisticas2.puntos.Caption = SkillPoints
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            semen = Right$(semen, Len(semen) - 4)
            SkillPoints = SkillPoints + Val(semen)
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            semen = Right$(semen, Len(semen) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & semen, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            semen = Right$(semen, Len(semen) - 4)
            frmMSG.List1.AddItem semen
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            semen = Right$(semen, Len(semen) - 4)
            frmForo.List.AddItem ReadField(1, semen, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, semen, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
            
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
            
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, semen, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(charindex)
            Else
                Call MI(CualMI).ResetInvisible(charindex)
            End If
#End If

            'Exit Sub
            
    End Select

    Select Case Left(sData, 6)
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "RECUPS" 'CHOTS | Recuperar Personaje
            semen = Right$(semen, Len(semen) - 6)
            Call MsgBox("Su Nueva Password es:" & vbNewLine & semen)
            Unload frmRecuperar
            Exit Sub
        Case "RECUBP" 'CHOTS | Borrar Personaje
            semen = Right$(semen, Len(semen) - 6)
            With frmBorrar
                .lblPreg.Caption = "¿" & " " & semen & " " & "?"
                .lblPreg.Visible = True
                .txtResp.Visible = True
                .Label6.Visible = True
                .Label7.Visible = True
                .MousePointer = vbDefault
                .txtEmail.Enabled = False
                .txtNombre.Enabled = False
                .Command3.Visible = False
                .Command2.Visible = True
            End With
            Exit Sub
        Case "RECUPR" 'CHOTS | Recuperar Personaje
            semen = Right$(semen, Len(semen) - 6)
            With frmRecuperar
                .lblPreg.Caption = "¿" & " " & semen & " " & "?"
                .lblPreg.Visible = True
                .txtResp.Visible = True
                .Label5.Visible = True
                .Label2.Visible = True
                .MousePointer = vbDefault
                .txtEmail.Enabled = False
                .txtNombre.Enabled = False
                .Command1.Visible = False
                .Command2.Visible = True
            End With
            Exit Sub
        Case "LSTQUE" 'CHOTS | Sistema de Quest
            semen = Right(semen, Len(semen) - 6)
            frmQuest.Show , frmMain
            Exit Sub
        Case "LSTCRI"
            semen = Right(semen, Len(semen) - 6)
            For i = 1 To Val(ReadField(1, semen, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, semen, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
        'BYSNACK | Retos
        Case "PANRET"
            frmRetos.Show , frmMain
        Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
        Case "GUILDNE"
            semen = Right(semen, Len(semen) - 7)
            Call frmGuildNews.ParseGuildNews(semen)
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            semen = Right(semen, Len(semen) - 7)
            Call frmUserRequest.recievePeticion(semen)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            semen = Right(semen, Len(semen) - 7)
            Call frmUserRequest.recievePeticion(semen)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            semen = Right(semen, Len(semen) - 7)
            Call frmPeaceProp.ParseAllieOffers(semen)
        Case "PEACEPR"  'lista de prop de paz
            semen = Right(semen, Len(semen) - 7)
            Call frmPeaceProp.ParsePeaceOffers(semen)
            Exit Sub
        Case "CHRINFO"
            semen = Right(semen, Len(semen) - 7)
            Call frmCharInfo.parseCharInfo(semen)
            Exit Sub
        Case "LEADERI"
            semen = Right(semen, Len(semen) - 7)
            Call frmGuildLeader.ParseLeaderInfo(semen)
            Exit Sub
        Case "CLANDET"
            semen = Right(semen, Len(semen) - 7)
            Call frmGuildBrief.ParseGuildInfo(semen)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"
            semen = Right(semen, Len(semen) - 7)
            Call frmUserRequest.recievePeticion(semen)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                semen = Right(semen, Len(semen) - 7)
                
                If ReadField(2, semen, 44) = "0" Then
                    frmComerciar.List1(0).listIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).listIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                ii = 1
                Do While ii <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                semen = Right(semen, Len(semen) - 7)
                
                If ReadField(2, semen, 44) = "0" Then
                        frmBancoObj.List1(0).listIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).listIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
        Exit Sub
        
        Case "ABESPIA" 'CHOTS | Sistema de Espías
            semen = Right$(semen, Len(semen) - 7)
            frmEspia.lblEspiado.Caption = "Espiando a: " & ReadField(1, semen, 44)
            frmEspia.hp.Width = (Val(ReadField(2, semen, 44)) / Val(ReadField(3, semen, 44))) * 4080
            frmEspia.lblHp.Caption = Val(ReadField(2, semen, 44)) & "/" & Val(ReadField(3, semen, 44))
            frmEspia.man.Width = (Val(ReadField(4, semen, 44)) / Val(ReadField(5, semen, 44))) * 4080
            frmEspia.lblMan.Caption = Val(ReadField(4, semen, 44)) & "/" & Val(ReadField(5, semen, 44))
            frmEspia.Show , frmMain
        Exit Sub
        
        Case "ABDENU"
            frmMain.tmrDenu.Enabled = True
        Exit Sub
        
        Case "ABBLOCK"
            Call WriteVar(App.Path & "\init\version.dat", "VERSION", "Graficos", "1")
            Call MsgBox("Tu Cliente ha sido Bloqueado")
            End
        Exit Sub

        Case "ABCENTI" 'CHOTS | Sistema de Centinela
            semen = Right$(semen, Len(semen) - 7)
            Call frmMain.MostrarCentinela(semen)
        Exit Sub

        Case "PANTOR"
            Call FrmConsolaTorneo.Show(vbModeless, frmMain)
        Exit Sub
        
        Case "LISTUSU"
            semen = Right$(semen, Len(semen) - 7)
            t = Split(semen, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(t) To UBound(t)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem t(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0
            End If
            Exit Sub
    End Select
    
    
    '[Alejo]
    Select Case UCase$(Left$(semen, 9))
        Case "COMUSUINV"
            semen = Right$(semen, Len(semen) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, semen, 44)
            OtroInventario(1).Name = ReadField(3, semen, 44)
            OtroInventario(1).Amount = ReadField(4, semen, 44)
            OtroInventario(1).Equipped = ReadField(5, semen, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, semen, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, semen, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, semen, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, semen, 44))
            OtroInventario(1).Def = Val(ReadField(10, semen, 44))
            OtroInventario(1).Valor = Val(ReadField(11, semen, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
            Exit Sub
    End Select
    
    'CHOTS | Acá lee el NOVER
    If Len(semen) > 5 Then
        semen = Right$(semen, Len(semen) - 5)
        charindex = Val(ReadField(1, semen, 44))
        charlist(charindex).invisible = (Val(ReadField(2, semen, 44)) = 1)
        Call SetNameColor(charindex)
    End If
    'CHOTS | Acá lee el NOVER

End Sub

Sub VaginaJugosa(ByVal flujoVaginal As String)
    'No enviamos nada si no estamos conectados
    #If UsarWrench = 1 Then
        If Not frmMain.Socket1.Connected Then Exit Sub
    #Else
        If frmMain.Winsock1.State <> sckConnected Then Exit Sub
    #End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(flujoVaginal, 5))
    
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()
    
    flujoVaginal = flujoVaginal & ENDC

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(flujoVaginal) > 8000 Then
        Exit Sub
    ElseIf Len(flujoVaginal) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If
    
    flujoVaginal = ChotsEncrypt(flujoVaginal)

    #If UsarWrench = 1 Then
        Call frmMain.Socket1.Write(flujoVaginal, Len(flujoVaginal))
    #Else
        Call frmMain.Winsock1.SendData(flujoVaginal)
    #End If

End Sub

