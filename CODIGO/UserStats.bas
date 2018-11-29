Attribute VB_Name = "UserStats"
'Lapsus2017
'Copyright (C) 2017 Dalmasso, Juan Andres
'
'Modulo UserStats
'Encargado de gestionar Stamina, Mana, Vida, Hambre y Sed del personaje
'Recibe los datos, los encripta antes de almacenarlos en memoria y renderiza la barrita
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'

Public Const BARRITA_LENGTH As Byte = 95
Public Const BARRITAEXP_LENGTH As Byte = 166

Public Type cUserStats
    MaxHP As Integer
    MinHP As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxSTA As Integer
    MinSTA As Integer
    MaxAGU As Integer
    MinAGU As Integer
    MaxHAM As Integer
    MinHAM As Integer
    Gold As Long
    Level As Integer
    Elu As Long
    EXP As Long
End Type

Public Type cUserStatsMultipliers
    MaxHP As Byte
    MinHP As Byte
    MaxMAN As Byte
    MinMAN As Byte
    MaxSTA As Byte
    MinSTA As Byte
    MaxAGU As Byte
    MinAGU As Byte
    MaxHAM As Byte
    MinHAM As Byte
End Type

Public MyUserStats As cUserStats
Public StatsMultipliers As cUserStatsMultipliers

'CHOTS | Initialize vars
Public Sub inicializarUserStats()
With MyUserStats
    .MaxHP = 0
    .MinHP = 0
    .MaxMAN = 0
    .MinMAN = 0
    .MaxSTA = 0
    .MinSTA = 0
    .MaxAGU = 100
    .MinAGU = 0
    .MaxHAM = 100
    .MinHAM = 0
    .Gold = 0
    .Level = 1
    .Elu = 150
    .EXP = 0
End With

With StatsMultipliers
    .MaxHP = 5
    .MinHP = 2
    .MaxMAN = 3
    .MinMAN = 4
    .MaxSTA = 3
    .MinSTA = 2
    .MaxAGU = 9
    .MinAGU = 7
    .MaxHAM = 6
    .MinHAM = 8
End With

End Sub

'CHOTS | Funciones de renderizado
Public Sub RenderHpBar()
    frmMain.Hpshp.Width = ((MyUserStats.MinHP / StatsMultipliers.MinHP) / (MyUserStats.MaxHP / StatsMultipliers.MaxHP)) * BARRITA_LENGTH
    frmMain.HpBar.Caption = (MyUserStats.MinHP / StatsMultipliers.MinHP) & "/" & (MyUserStats.MaxHP / StatsMultipliers.MaxHP)
End Sub

Public Sub RenderManaBar()
    If MyUserStats.MaxMAN > 0 Then
        frmMain.MANShp.Width = ((MyUserStats.MinMAN / StatsMultipliers.MinMAN) / (MyUserStats.MaxMAN / StatsMultipliers.MaxMAN)) * BARRITA_LENGTH
    Else
        frmMain.MANShp.Width = 0
    End If
    frmMain.ManaBar.Caption = (MyUserStats.MinMAN / StatsMultipliers.MinMAN) & "/" & (MyUserStats.MaxMAN / StatsMultipliers.MaxMAN)
End Sub

Public Sub RenderStaBar()
    frmMain.STAShp.Width = ((MyUserStats.MinSTA / StatsMultipliers.MinSTA) / (MyUserStats.MaxSTA / StatsMultipliers.MaxSTA)) * BARRITA_LENGTH
    frmMain.StaBar.Caption = (MyUserStats.MinSTA / StatsMultipliers.MinSTA) & "/" & (MyUserStats.MaxSTA / StatsMultipliers.MaxSTA)
End Sub

Public Sub RenderHambreBar()
    frmMain.COMIDAsp.Width = ((MyUserStats.MinHAM / StatsMultipliers.MinHAM) / (MyUserStats.MaxHAM / StatsMultipliers.MaxHAM)) * BARRITA_LENGTH
    frmMain.HamBar.Caption = (MyUserStats.MinHAM / StatsMultipliers.MinHAM) & "/" & (MyUserStats.MaxHAM / StatsMultipliers.MaxHAM)
End Sub

Public Sub RenderSedBar()
    frmMain.AGUAsp.Width = ((MyUserStats.MinAGU / StatsMultipliers.MinAGU) / (MyUserStats.MaxAGU / StatsMultipliers.MaxAGU)) * BARRITA_LENGTH
    frmMain.AguBar.Caption = (MyUserStats.MinAGU / StatsMultipliers.MinAGU) & "/" & (MyUserStats.MaxAGU / StatsMultipliers.MaxAGU)
End Sub

Public Sub RenderGoldBar()
    frmMain.GldLbl.Caption = PonerPuntos(MyUserStats.Gold)
End Sub

Public Sub RenderExpBar()
    frmMain.ExpShp.Width = (MyUserStats.EXP / (MyUserStats.Elu + 1)) * BARRITAEXP_LENGTH
    frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(MyUserStats.EXP) * CDbl(100) / CDbl(MyUserStats.Elu + 1), 2) & "%]"
    frmMain.lblExpe.Caption = PonerPuntos(MyUserStats.EXP) & " / " & PonerPuntos(MyUserStats.Elu)
    frmMain.LvlLbl.Caption = MyUserStats.Level
End Sub


'CHOTS | Funciones de seteo
Public Sub SetHp(Optional ByVal MinHP As Integer = -1, Optional ByVal MaxHP As Integer = -1)
    If MinHP >= 0 Then
        MyUserStats.MinHP = MinHP * StatsMultipliers.MinHP
    End If

    If MaxHP >= 0 Then
        MyUserStats.MaxHP = MaxHP * StatsMultipliers.MaxHP
    End If

    If MyUserStats.MinHP = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0
    End If

    Call RenderHpBar
End Sub

Public Sub SetMana(Optional ByVal MinMAN As Integer = -1, Optional ByVal MaxMAN As Integer = -1)
    If MinMAN >= 0 Then
        MyUserStats.MinMAN = MinMAN * StatsMultipliers.MinMAN
    End If

    If MaxMAN >= 0 Then
        MyUserStats.MaxMAN = MaxMAN * StatsMultipliers.MaxMAN
    End If

    Call RenderManaBar
End Sub

Public Sub SetStamina(Optional ByVal MinSTA As Integer = -1, Optional ByVal MaxSTA As Integer = -1)
    If MinSTA >= 0 Then
        MyUserStats.MinSTA = MinSTA * StatsMultipliers.MinSTA
    End If

    If MaxSTA >= 0 Then
        MyUserStats.MaxSTA = MaxSTA * StatsMultipliers.MaxSTA
    End If

    Call RenderStaBar
End Sub

Public Sub SetHambre(ByVal MinHAM As Integer)
    MyUserStats.MinHAM = MinHAM * StatsMultipliers.MinHAM

    MyUserStats.MaxHAM = 100 * StatsMultipliers.MaxHAM

    Call RenderHambreBar
End Sub

Public Sub SetSed(ByVal MinAGU As Integer)
    MyUserStats.MinAGU = MinAGU * StatsMultipliers.MinAGU

    MyUserStats.MaxAGU = 100 * StatsMultipliers.MaxAGU

    Call RenderSedBar
End Sub

Public Sub SetGold(ByVal Gold As Long)
    MyUserStats.Gold = Gold

    Call RenderGoldBar
End Sub

Public Sub SetLevel(ByVal Level As Integer, Optional ByVal Render As Boolean = True)
    MyUserStats.Level = Level

    If Render = True Then Call RenderExpBar
End Sub

Public Sub SetElu(ByVal Elu As Long, Optional ByVal Render As Boolean = True)
    MyUserStats.Elu = Elu

    If Render = True Then Call RenderExpBar
End Sub

Public Sub SetExp(ByVal EXP As Long, Optional ByVal Render As Boolean = True)
    MyUserStats.EXP = EXP

    If Render = True Then Call RenderExpBar
End Sub

'CHOTS | Funciones de adicion
Public Sub AddMinHp(ByVal Value As Integer)
    MyUserStats.MinHP = MyUserStats.MinHP + (Value * StatsMultipliers.MinHP)

    If (MyUserStats.MinHP / StatsMultipliers.MinHP) > (MyUserStats.MaxHP / StatsMultipliers.MaxHP) Then
        MyUserStats.MinHP = (MyUserStats.MaxHP / StatsMultipliers.MaxHP) * StatsMultipliers.MinHP
    End If

    Call RenderHpBar
End Sub

Public Sub AddMaxMana(ByVal Value As Integer)
    MyUserStats.MaxMAN = MyUserStats.MaxMAN + (Value * StatsMultipliers.MaxMAN)

    Call RenderManaBar
End Sub

Public Sub AddMaxHp(ByVal Value As Integer)
    MyUserStats.MaxHP = MyUserStats.MaxHP + (Value * StatsMultipliers.MaxHP)
    
    MyUserStats.MinHP = (MyUserStats.MaxHP / StatsMultipliers.MaxHP) * StatsMultipliers.MinHP

    Call RenderHpBar
End Sub

Public Sub AddMaxSta(ByVal Value As Integer)
    MyUserStats.MaxSTA = MyUserStats.MaxSTA + (Value * StatsMultipliers.MaxSTA)

    Call RenderManaBar
End Sub

Public Sub AddMinManaPercentage(ByVal percentage As Integer)
    MyUserStats.MinMAN = ((MyUserStats.MinMAN / StatsMultipliers.MinMAN) + Porcentaje(MyUserStats.MaxMAN / StatsMultipliers.MaxMAN, percentage)) * StatsMultipliers.MinMAN

    If (MyUserStats.MinMAN / StatsMultipliers.MinMAN) > (MyUserStats.MaxMAN / StatsMultipliers.MaxMAN) Then
        MyUserStats.MinMAN = (MyUserStats.MaxMAN / StatsMultipliers.MaxMAN) * StatsMultipliers.MinMAN
    End If

    Call RenderManaBar
End Sub

Public Sub AddLevel(ByVal Value As Integer)
    MyUserStats.Level = MyUserStats.Level + Value

    Call RenderExpBar
End Sub

Public Sub AddGold(ByVal Gold As Long)
    Dim exOro As Double
    Dim oroGanado As Double
    exOro = MyUserStats.Gold
    MyUserStats.Gold = Gold
    oroGanado = Gold - exOro
    If oroGanado > 0 Then frmMain.gldLbl2.Caption = "+" & oroGanado
    frmMain.tmrOro2.Enabled = True

    Call RenderGoldBar
End Sub
