VERSION 5.00
Begin VB.Form frmGema 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label habLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "habilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label crLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clase - Raza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2535
   End
   Begin VB.Image lblSalir 
      Height          =   255
      Left            =   1560
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image lblLiberar 
      Height          =   255
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmGema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\Gemas.jpg")
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub lblLiberar_Click()
Call SendData("GEMHAB")
frmGema.Visible = False
End Sub

Private Sub lblSalir_Click()
frmGema.Visible = False
End Sub

