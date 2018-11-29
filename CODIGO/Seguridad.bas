Attribute VB_Name = "Seguridad"
'Lapsus2017
'Copyright (C) 2017 Dalmasso, Juan Andres
'
'Modulo de seguridad de LapsusAO
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'
'ATENCION: El valor de las variables publicas sera cambiado con cada nueva version

Public Const SECURITY_ENABLED As Boolean = True

Public Type cSecurityParameters
   multiplicator As Double
   keyA As Byte
   keyB As Byte
   primeExp As String
   primeMod As String
   commaString As String
End Type

'CHOTS | Packages sent by client
Public Type cPackageNamesClient
    getValCode As String
    sendSecretKey As String
    login As String
    register As String
    tirarDados As String
    borrarPersonaje As String
    recuperarPersonaje As String
    confirmarBorradoPersonaje As String
    confirmarRecuperarPersonaje As String
    hablar As String
    gritar As String
    moverse As String
    atacar As String
    agarrarObjeto As String
    lanzarHechizo As String
    leftClick As String
    rightClick As String
    trabajoClick As String
    usarSkill As String
    usarItem As String
    equiparItem As String
    tirarItem As String
End Type

'CHOTS | Packages sent by server
Public Type cPackageNamesServer
    validarCliente As String
    login As String
    logout As String
    moverChar As String
    moverNpc As String
    cargarMapa As String
    updatePos As String
    dialogo As String
    dialogoConsola As String
    crearNpc As String
    crearChar As String
    borrarChar As String
    moverPersonaje As String
    recibeDados As String
End Type

'CHOTS | Initialize vars

Public SecurityParameters As cSecurityParameters
Public ClientPackages As cPackageNamesClient
Public ServerPackages As cPackageNamesServer

'CHOTS | Initialize vars
Public Sub inicializarSeguridad()
With SecurityParameters
    .multiplicator = 0.114
    .keyA = 3
    .keyB = 106
    .primeExp = "23291"
    .primeMod = "31547"
    .commaString = "__.__"
End With

With ClientPackages
    .getValCode = "p$#e&&814rf_."
    .login = "OLOSNF"
    .register = "NLOSNF"
    .tirarDados = "TIRKAK"
    .borrarPersonaje = "BORO"
    .recuperarPersonaje = "RECU"
    .confirmarBorradoPersonaje = "BORR"
    .confirmarRecuperarPersonaje = "RECO"
    .hablar = ";"
    .gritar = "-"
    .moverse = "Ñ"
    .atacar = "AQ"
    .agarrarObjeto = "AH"
    .lanzarHechizo = "HV"
    .leftClick = "LC"
    .rightClick = "RC"
    .trabajoClick = "WLC"
    .usarSkill = "UX"
    .usarItem = "USX"
    .equiparItem = "EQUI"
    .tirarItem = "TT"
End With


With ServerPackages
    .validarCliente = "VAY"
    .login = "LOGLAP"
    .logout = "FINLA"
    .moverChar = "+"
    .moverNpc = "*"
    .cargarMapa = "CM"
    .updatePos = "PU"
    .dialogo = "||"
    .dialogoConsola = "|+"
    .crearNpc = "BC"
    .crearChar = "ÑC"
    .borrarChar = "BP"
    .moverPersonaje = "MP"
    .recibeDados = "DAD"
End With
End Sub

Public Function ChotsEncrypt(ByVal data As String) As String

If Not SECURITY_ENABLED Then
    ChotsEncrypt = data
    Exit Function
End If

ChotsEncrypt = DyeCifro(data)

End Function


Public Function ChotsDecrypt(ByVal data As String) As String

If Not SECURITY_ENABLED Then
    ChotsDecrypt = data
    Exit Function
End If

ChotsDecrypt = DyeDecifro(data)

End Function

Public Function EncryptStr(ByVal s As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""

For i = 1 To Len(s)
    C1 = Asc(mid(s, i, 1))
    If i > Len(p) Then
        C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
    Else
        C2 = Asc(mid(p, i, 1))
    End If
    C1 = C1 + C2 + 64
    If C1 > 255 Then C1 = C1 - 256
        r = r + Chr(C1)
Next i

EncryptStr = r

End Function

Public Function Nover(Longitud As Integer) As String
Nover = vbNullString
Dim i As Integer
If Longitud <= 1 Then Exit Function

For i = 1 To Longitud
    Nover = Nover & Chr(RandomNumber(160, 255))
Next i

End Function

Function ENCRYPT(ByVal STRG As String) As String
Dim asd As Long
Dim suma As Long
If Val(STRG) <> 5 Then
    For asd = 1 To Len(STRG)
        suma = suma + Asc(mid$(STRG, asd, 1))
    Next
    For asd = 1 To Asc(mid$(STRG, 1, 1))
        If ENCRYPT = "" Then
            ENCRYPT = MD5String(CStr(suma * SecurityParameters.multiplicator))
        Else
            ENCRYPT = MD5String(ENCRYPT)
        End If
    Next

End If
ENCRYPT = ENCRYPT
End Function

Function RandomCodeEncrypt(ByVal RandomCode As String) As String
    RandomCodeEncrypt = RC4_EncryptString(RandomCode, mpModExp(RandomCode, SecurityParameters.primeExp, SecurityParameters.primeMod))
    RandomCodeEncrypt = CommaReplace(RandomCodeEncrypt)
End Function

Function CommaReplace(ByVal Text As String) As String
    CommaReplace = Replace(Text, ",", SecurityParameters.commaString)
End Function

Public Sub IncrementarUseNum()
'CHOTS | Secuencia: 7>4>6>2>9>1>5>3>0>8>7...

    If Logged And (Inventario.SelectedItem > 0) And (Inventario.SelectedItem <= MAX_INVENTORY_SLOTS) Then
        Select Case Val(UseNum)
            Case 0
                UseNum = 8
            Case 1
                UseNum = 5
            Case 2
                UseNum = 9
            Case 3
                UseNum = 0
            Case 4
                UseNum = 6
            Case 5
                UseNum = 3
            Case 6
                UseNum = 2
            Case 7
                UseNum = 4
            Case 8
                UseNum = 7
            Case 9
                UseNum = 1
        End Select

        If UseAcum > 30000 Then
            UseAcum = UseAcum - 30000
        End If

        UseAcum = UseAcum + (UseNum * 200)
    End if

End Sub
