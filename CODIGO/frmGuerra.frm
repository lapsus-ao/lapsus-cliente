VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGuerra 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   4215
   ClientTop       =   5970
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4710
   Begin MSComctlLib.Slider sldPuntos 
      Height          =   195
      Left            =   1575
      TabIndex        =   3
      Top             =   1665
      Width           =   2600
      _ExtentX        =   4577
      _ExtentY        =   344
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   100
      Min             =   100
      Max             =   1000
      SelStart        =   100
      TickFrequency   =   100
      Value           =   100
   End
   Begin MSComctlLib.Slider sldUsers 
      Height          =   200
      Left            =   1575
      TabIndex        =   2
      Top             =   1960
      Width           =   2600
      _ExtentX        =   4577
      _ExtentY        =   344
      _Version        =   393216
      Min             =   2
      Max             =   5
      SelStart        =   2
      Value           =   2
   End
   Begin VB.TextBox txtClan 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C000&
      Height          =   190
      Left            =   1575
      TabIndex        =   0
      Top             =   1360
      Width           =   2565
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   360
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3360
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2vs2 - 100 puntos"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblNombreSala 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   1575
      TabIndex        =   1
      Top             =   1065
      Width           =   2565
   End
End
Attribute VB_Name = "frmGuerra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private numeroSala As Byte

Public Sub setNumeroSala(ByVal numero As Byte)
numeroSala = numero
End Sub

Private Sub Form_Load()
    frmGuerra.Picture = LoadPicture(DirGraficos & "GUERRAS.jpg")
    Call setCaptions
End Sub

Private Sub Image1_Click()
    Call SendData("/CREARGUERRA " & numeroSala & "," & sldUsers.Value & "," & sldPuntos.Value & "," & txtClan.Text)
    Unload Me
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub sldPuntos_Change()
    Call setCaptions
End Sub
Private Sub sldUsers_Change()
    Call setCaptions
End Sub

Private Sub setCaptions()
lblPuntos.Caption = sldUsers.Value & "vs" & sldUsers.Value & " - " & PonerPuntos(sldPuntos.Value) & " puntos"
End Sub

