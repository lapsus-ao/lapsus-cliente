VERSION 5.00
Begin VB.Form FrmMapa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   3930
   ClientTop       =   1155
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7965
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "55: Montaña"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   6
      Left            =   255
      TabIndex        =   6
      Top             =   5700
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "48: Sala de Invocacion"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   5
      Left            =   255
      TabIndex        =   5
      Top             =   5340
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "52: Dungeon  Aqua"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   4
      Left            =   255
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "41: Dungeon Dragon"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   3
      Left            =   255
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "33: Dungeon Verill"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   2
      Top             =   4965
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "21: Barco Pirata"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   4785
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "19: Dungeon Marabell"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   4605
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7560
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\Mapa.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

