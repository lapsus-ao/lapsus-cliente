Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
Public CustomKeys As New clsCustomKeys
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Cont As Byte
Public SupBMiniMap As DirectDrawSurface7
Public SupMiniMap As DirectDrawSurface7
'Objetos públicos
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New cDialogos
Public Audio As New clsAudio
Public Inventario As New clsGrapchicalInventory
Public SurfaceDB As clsSurfaceManager   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

Public TimerPing(1 To 2) As Long 'CHOTS | /PING

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String

Public Const passC As String = "wwn"

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public vecClan() As String 'CHOTS | Vector de Clanes

Public hayCastillo As Boolean 'CHOTS | Textos de Castillos

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long

Public FPSFast As Boolean

Public ArmaMin As Integer
Public ArmaMax As Integer
Public ArmorMin As Integer
Public ArmorMax As Integer
Public EscuMin As Integer
Public EscuMax As Integer
Public CascMin As Integer
Public MagMin As Integer
Public MagMax As Integer
Public CascMax As Integer
Public Verde As Integer
Public Amarilla As Integer
Public MyGuildName As String
Public CurServer As Integer
Public CreandoClan As Boolean
Public ClanName As String
Public Site As String
Public UseNum As Byte
Public UseAcum As Integer

Public UserCiego As Boolean
Public UserEstupido As Boolean
Public NoRes As Boolean 'no cambiar la resolucion
Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 992
Public Const tComb = 397
Public Const tUs = 799
Public Const tClick = 500
Public Const tRefrescar = 1000

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer
Public ObjSastre(0 To 100) As Integer
Public ObjDruida(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean

'CHOTS | Reducción de mensajes
Public Const Mensaje1 As String = "Estás muy cansado para lanzar este hechizo."
Public Const Mensaje2 As String = "No tenes suficientes puntos de magia para lanzar este hechizo."
Public Const Mensaje3 As String = "No tenes suficiente mana."
Public Const Mensaje4 As String = "No podes lanzar hechizos porque estas muerto."
Public Const Mensaje5 As String = "En zona segura no puedes invocar criaturas."
Public Const Mensaje6 As String = "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos"
Public Const Mensaje7 As String = "No podes atacar a ese npc."
Public Const Mensaje8 As String = "Debes quitarte el seguro para de poder atacar guardias"
Public Const Mensaje9 As String = "El npc es inmune a este hechizo."
Public Const Mensaje10 As String = "Este hechizo solo afecta NPCs que tengan amo."
Public Const Mensaje11 As String = "¡Has vuelto a ser visible!"
Public Const Mensaje12 As String = "¡¡Estás muerto!!"
Public Const Mensaje13 As String = "Estas muy lejos del usuario."
Public Const Mensaje14 As String = "Usuario inexistente."
Public Const Mensaje15 As String = "/salir cancelado."
Public Const Mensaje16 As String = "Has Terminado de meditar."
Public Const Mensaje17 As String = "No podes moverte porque estas paralizado."
Public Const Mensaje18 As String = "¡¡No podes atacar a nadie porque estas muerto!!."
Public Const Mensaje19 As String = "No podés usar asi esta arma."
Public Const Mensaje20 As String = "¡¡Estás muerto!! Los muertos no pueden tomar objetos."
Public Const Mensaje21 As String = "Escribe /SEG para quitar el seguro"
Public Const Mensaje22 As String = "No podes atacarte a vos mismo."
Public Const Mensaje23 As String = "Comienzas a Meditar"
Public Const Mensaje24 As String = "No puedo Cargar mas Objetos"
Public Const Mensaje25 As String = "¡Has Ganado 100 puntos de Experiencia!"
Public Const Mensaje26 As String = "Target Invalido"
Public Const Mensaje27 As String = "Estas demasiado lejos."
Public Const Mensaje28 As String = "Ya Estas Oculto."
Public Const Mensaje29 As String = "¡Primero selecciona el hechizo que quieres lanzar!"
Public Const Mensaje30 As String = "¡Primero tenes que seleccionar un personaje, hace click izquierdo sobre el."
Public Const Mensaje31 As String = "Primero hace click izquierdo sobre el personaje."
Public Const Mensaje32 As String = "El sacerdote no puede curarte debido a que estas demasiado lejos."
Public Const Mensaje33 As String = "Has Muerto! Tipea '/HOGAR' para volver a tu Ciudad de origen!"
Public Const Mensaje34 As String = "Tipea /HOGAR para volver a tu Ciudad"
Public Const Mensaje35 As String = "Estas envenenado, si no te curas moriras."
Public Const Mensaje36 As String = "Has Sanado."
Public Const Mensaje37 As String = "Te estas concentrando, en 3 segundos comenzarás a meditar."
Public Const Mensaje38 As String = "AntiCheat> Tu Cliente es Valido, Gracias por jugar Twist AO!!"
Public Const Mensaje39 As String = "Estas obstruyendo la via publica, muévete o seras encarcelado!!!"
Public Const Mensaje40 As String = "Has Sido Resucitado!!"
Public Const Mensaje41 As String = "Has Sido Curado!!"
Public Const Mensaje42 As String = "Tu Clase, Genero o Raza, no puede usar este Objeto."
Public Const Mensaje43 As String = "Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro."
Public Const Mensaje44 As String = "No puedes lanzar este conjuro sin la ayuda de un báculo."
Public Const Mensaje45 As String = "Mapa Exclusivo para Newbies."
Public Const Mensaje46 As String = "Recuperas tu Fuerza y Agilidad Original."
Public Const Mensaje47 As String = "El Usuario esta Offline."
Public Const Mensaje48 As String = "No Puedes Salir estando Paralizado."
Public Const Mensaje49 As String = "El NPC no está interesado en Comprar ese Objeto!"
Public Const Mensaje50 As String = "Has matado a la Criatura!"
Public Const Mensaje51 As String = "Pierdes el control de tus mascotas."
Public Const Mensaje52 As String = "¡Te has escondido entre las sombras!"
Public Const Mensaje53 As String = "¡No has logrado esconderte!"
Public Const Mensaje54 As String = "Has construido el/los objeto/s!"
Public Const Mensaje55 As String = "Has obtenido un lingote!!!"
Public Const Mensaje56 As String = "¡No has logrado apuñalar a tu enemigo!"
Public Const Mensaje57 As String = "¡No has obtenido raíces!"
Public Const Mensaje58 As String = "¡No has obtenido leña!"
Public Const Mensaje59 As String = "¡No has obtenido minerales!"
Public Const Mensaje60 As String = "Este hechizo actua solo sobre usuarios."
Public Const Mensaje61 As String = "Este hechizo solo afecta a los npcs."
Public Const Mensaje62 As String = "No podes atacarte a vos mismo."
Public Const Mensaje63 As String = "Debes desequiparte tu escudo para poder usar esta arma."
Public Const Mensaje64 As String = "Debes desequiparte tu arma para poder usar este escudo."
Public Const Mensaje65 As String = "¡Debes aproximarte al agua para usar el barco!"
Public Const Mensaje66 As String = "No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo."
Public Const Mensaje67 As String = "¡Primero selecciona el hechizo que quieres lanzar!"
Public Const Mensaje68 As String = "¡No puedes salir estando en Duelo!"
Public Const Mensaje69 As String = "El usuario no tiene intenciones de regresar a la vida!"
Public Const Mensaje70 As String = "Debes desactivar tu Seguro de Caos para atacar Legionarios"
Public Const Mensaje71 As String = "Necesitas un instrumento mágico para devolver la vida"
Public Const Mensaje72 As String = "No se permite la invisibilidad en este mapa"
Public Const Mensaje73 As String = "No puedes salir del clan estando en un Castillo"
Public Const Mensaje74 As String = "No puedes echar a alguien que se encuentra en un Castillo"
Public Const Mensaje75 As String = "No puedes echar a alguien si te encuentras en un Castillo"
Public Const Mensaje76 As String = "No puedes cerrar el clan estando en un Castillo"
Public Const Mensaje77 As String = "Denuncia enviada, espere.."
Public Const Mensaje78 As String = "Se ha Invocado una criatura en la sala de Invocaciones!!!"
Public Const Mensaje79 As String = "No tienes Suficientes Puntos de Usuario!!!"
Public Const Mensaje80 As String = "No puedes Invocar un Elemental en la Sala de Invocaciones"
Public Const Mensaje81 As String = "Has viajado! Tu hambre y sed han disminuído por el cansancio"
Public Const Mensaje82 As String = "¡¡Estás muriendo de frío, abrígate o morirás!!"
Public Const Mensaje83 As String = "¡¡Has muerto de frío!!."
Public Const Mensaje84 As String = "Gracias por jugar Lapsus AO"
Public Const Mensaje85 As String = "El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM."
Public Const Mensaje86 As String = "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes."
Public Const Mensaje87 As String = "Debes esperar un minuto para enviar otra denuncia"
Public Const Mensaje88 As String = "¡Has pescado un lindo Pez!"
Public Const Mensaje89 As String = "¡No has pescado nada!"
Public Const Mensaje90 As String = "Debes abandonar el Dungeon Newbie."
Public Const Mensaje91 As String = "TorneosAuto> Los torneos automáticos han sido DESHABILITADOS!"
Public Const Mensaje92 As String = "TorneosAuto> Los torneos automáticos han sido HABILITADOS!"
Public Const Mensaje93 As String = "Has entrado al torneo!"
Public Const Mensaje94 As String = "El cupo ha sido alcanzado!"
Public Const Mensaje95 As String = "Tipea /FIXTURE para ver los enfrentamientos"
Public Const Mensaje96 As String = "TorneosAuto> En 5 Minutos dará comienzo un torneo automático"
Public Const Mensaje97 As String = "TorneosAuto> En 1 Minuto dará comienzo un torneo automático"
Public Const Mensaje98 As String = "Debes tener una ropa de montura!"
Public Const Mensaje99 As String = "Has Ganado 50 puntos de usuario!"
Public Const Mensaje100 As String = "Solamente el Druida conoce el arte de la Captura!"

'CHOTS | Reducción de mensajes


'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

Public Const LoopAdEternum = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 25

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Enum Skills
     Suerte = 1
     Magia = 2
     Robar = 3
     Tacticas = 4
     Armas = 5
     Meditar = 6
     Apuñalar = 7
     Ocultarse = 8
     Supervivencia = 9
     Talar = 10
     Comerciar = 11
     Defensa = 12
     Pesca = 13
     Mineria = 14
     Carpinteria = 15
     Herreria = 16
     Liderazgo = 17 ' NOTA: Solia decir "Curacion"
     Domar = 18
     Proyectiles = 19
     Wresterling = 20
     Navegacion = 21
     Alquimia = 22
     Satreria = 23
     Botanica = 24
End Enum

Public Const FundirMetal As Integer = 88
Public Const CapturarNpc As Integer = 89 'CHOTS | Monturas

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."
Public IsSeguro As Boolean
Public IsSeguroC As Boolean
Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_CAPTURARNPC As String = "Haz click sobre el Npc..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public Nombres As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserPreg As String
Public UserResp As String
Public UserFichas As Long 'CHOTS | Casinooou
Public UserCanAttack As Integer
Public UserCanCombo As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public pausa As Boolean
Public enParty As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As String

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 3
Public Const NUMSKILLS As Byte = 24
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 17
Public Const NUMRAZAS As Byte = 6

Public UserSkills(1 To NUMSKILLS) As Integer
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Integer
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String
Public CityDesc(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public Logged As Boolean
Public NoPuedeUsar As Boolean
Public Hizo2Click As Byte

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer

Public Enum E_MODO
    Normal = 1
    BorrarPj = 2
    CrearNuevoPj = 3
    Dados = 4
    RecuperarPass = 5
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
'    FXMEDITARCHICO = 4
'    FXMEDITARMEDIANO = 5
'    FXMEDITARGRANDE = 6
'    FXMEDITARXGRANDE = 16
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum


'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

