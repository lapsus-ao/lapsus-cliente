VERSION 5.00
Begin VB.Form FrmConsolaTorneo 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Summon automático"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   105
      TabIndex        =   19
      Top             =   2625
      Width           =   3165
      Begin VB.CheckBox Check9 
         BackColor       =   &H00000000&
         Caption         =   "Activado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1785
         TabIndex        =   26
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox TxtY 
         Height          =   285
         Left            =   1260
         TabIndex        =   22
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TxtX 
         Height          =   285
         Left            =   735
         TabIndex        =   21
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TxtMap 
         Height          =   285
         Left            =   105
         TabIndex        =   20
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Sala"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1260
         TabIndex        =   25
         Top             =   210
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   735
         TabIndex        =   24
         Top             =   210
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Facción / Alineación"
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   105
      TabIndex        =   18
      Top             =   1575
      Width           =   3165
      Begin VB.CheckBox Check13 
         BackColor       =   &H00000000&
         Caption         =   "Armada REAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   30
         Top             =   210
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00000000&
         Caption         =   "Armada CAOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   29
         Top             =   525
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00000000&
         Caption         =   "Ciudadano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1680
         TabIndex        =   28
         Top             =   210
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00000000&
         Caption         =   "Criminal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1680
         TabIndex        =   27
         Top             =   525
         Value           =   1  'Checked
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaskColor       =   &H000000FF&
      TabIndex        =   15
      Top             =   3990
      Width           =   4635
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Comenzar torneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaskColor       =   &H000000FF&
      TabIndex        =   14
      Top             =   3570
      Width           =   4635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Clases válidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   3360
      TabIndex        =   7
      Top             =   525
      Width           =   1380
      Begin VB.CheckBox Check7 
         BackColor       =   &H00000000&
         Caption         =   "Druida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   2205
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00000000&
         Caption         =   "Cazador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   16
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00000000&
         Caption         =   "Asesino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   1890
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00000000&
         Caption         =   "Bardo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   1575
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Clérigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   1260
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Paladín"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   945
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Mago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   9
         Top             =   630
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Guerrero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   8
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.TextBox Txt_Cupo 
      Height          =   285
      Left            =   1365
      TabIndex        =   6
      Top             =   1260
      Width           =   1905
   End
   Begin VB.TextBox Txt_LvlMax 
      Height          =   285
      Left            =   1365
      TabIndex        =   4
      Top             =   630
      Width           =   1905
   End
   Begin VB.TextBox Txt_LvlMin 
      Height          =   285
      Left            =   1365
      TabIndex        =   3
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label Cup 
      BackStyle       =   0  'Transparent
      Caption         =   "Cupo máximo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel mínimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel máximo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIGURACIÓN DE TORNEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   105
      Width           =   4950
   End
End
Attribute VB_Name = "FrmConsolaTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()

If Not CheckDatos Then Exit Sub

VaginaJugosa "/TOR " & Txt_LvlMin & " " & Txt_LvlMax & " " & Txt_Cupo & " " & Check1.Value & " " & Check2.Value & " " & Check3.Value & " " & Check4.Value & " " & Check5.Value & " " & Check6.Value & " " & Check7.Value & " " & Check8.Value & " " & Check9.Value & " " & TxtMap & " " & TxtX & " " & TxtY & " " & Check10.Value & " " & Check11.Value & " " & Check12.Value & " " & Check13.Value

Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Function CheckDatos() As Boolean
CheckDatos = True

If Txt_LvlMax = "" Then
CheckDatos = False
MsgBox "Falta completa el nivel máximo."
Exit Function
End If

If Txt_LvlMin = "" Then
MsgBox "Falta completa el nivel mínimo."
CheckDatos = False
Exit Function
End If

If Txt_Cupo = "" Then
MsgBox "Falta completa el cupo."
CheckDatos = False
Exit Function
End If

If Not IsNumeric(Txt_LvlMax) Then
CheckDatos = False
MsgBox "Nivel máximo no numérico."
Exit Function
End If

If Not IsNumeric(Txt_LvlMin) Then
MsgBox "Nivel mínimo no numérico."
CheckDatos = False
Exit Function
End If

If Not IsNumeric(Txt_Cupo) Then
MsgBox "Cupo no numérico."
CheckDatos = False
Exit Function
End If

End Function

Private Sub Label7_Click()
TxtMap.text = 81
TxtX.text = 37
TxtY.text = 43
Check9.Value = 1
End Sub
