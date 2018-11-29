VERSION 5.00
Begin VB.Form frmCrearPersonaje1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPasswdCheck 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtCorreo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   5400
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtPreg 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   5400
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtResp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   5400
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2340
      Width           =   1815
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0063
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":006D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0080
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":0096
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00C9
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":00D9
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3150
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   0
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label lblMaxAgi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3900
      TabIndex        =   46
      Top             =   6960
      Width           =   330
   End
   Begin VB.Label lblMaxFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3900
      TabIndex        =   45
      Top             =   6600
      Width           =   330
   End
   Begin VB.Image boton 
      Height          =   1050
      Index           =   2
      Left            =   240
      MouseIcon       =   "frmCrearPersonaje.frx":0102
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2820
      TabIndex        =   44
      Top             =   6600
      Width           =   330
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2820
      TabIndex        =   43
      Top             =   6960
      Width           =   330
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2820
      TabIndex        =   42
      Top             =   8070
      Width           =   330
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2820
      TabIndex        =   41
      Top             =   7320
      Width           =   330
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2820
      TabIndex        =   40
      Top             =   7665
      Width           =   330
   End
   Begin VB.Label lblCarisma2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3510
      TabIndex        =   39
      Top             =   7665
      Width           =   330
   End
   Begin VB.Label lblInteligencia2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3510
      TabIndex        =   38
      Top             =   7320
      Width           =   330
   End
   Begin VB.Label lblConstitucion2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3510
      TabIndex        =   37
      Top             =   8070
      Width           =   330
   End
   Begin VB.Label lblAgilidad2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3510
      TabIndex        =   36
      Top             =   6960
      Width           =   330
   End
   Begin VB.Label lblFuerza2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3510
      TabIndex        =   35
      Top             =   6600
      Width           =   330
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   3795
      TabIndex        =   34
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   2745
      TabIndex        =   33
      Top             =   5190
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   41
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":0254
      MousePointer    =   99  'Custom
      Top             =   5250
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   40
      Left            =   3120
      MouseIcon       =   "frmCrearPersonaje.frx":03A6
      MousePointer    =   99  'Custom
      Top             =   5250
      Width           =   135
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   2745
      TabIndex        =   32
      Top             =   5400
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   2745
      TabIndex        =   31
      Top             =   5625
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   2745
      TabIndex        =   30
      Top             =   5850
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   3120
      MouseIcon       =   "frmCrearPersonaje.frx":04F8
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":064A
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   44
      Left            =   3120
      MouseIcon       =   "frmCrearPersonaje.frx":079C
      MousePointer    =   99  'Custom
      Top             =   5625
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   45
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":08EE
      MousePointer    =   99  'Custom
      Top             =   5625
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   46
      Left            =   3120
      MouseIcon       =   "frmCrearPersonaje.frx":0A40
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   47
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":0B92
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   135
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2745
      TabIndex        =   29
      Top             =   1095
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2745
      TabIndex        =   28
      Top             =   885
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2745
      TabIndex        =   27
      Top             =   1305
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2745
      TabIndex        =   26
      Top             =   1515
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2745
      TabIndex        =   25
      Top             =   1725
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2745
      TabIndex        =   24
      Top             =   1935
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2745
      TabIndex        =   23
      Top             =   2145
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2745
      TabIndex        =   22
      Top             =   2355
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2745
      TabIndex        =   21
      Top             =   2595
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2745
      TabIndex        =   20
      Top             =   2820
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2745
      TabIndex        =   19
      Top             =   3030
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   2745
      TabIndex        =   18
      Top             =   3240
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2745
      TabIndex        =   17
      Top             =   3465
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   2745
      TabIndex        =   16
      Top             =   3675
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   2745
      TabIndex        =   15
      Top             =   3900
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   2745
      TabIndex        =   14
      Top             =   4125
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   2745
      TabIndex        =   13
      Top             =   4335
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   2745
      TabIndex        =   12
      Top             =   4530
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   2745
      TabIndex        =   11
      Top             =   4740
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   2745
      TabIndex        =   10
      Top             =   4980
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":0CE4
      MousePointer    =   99  'Custom
      Top             =   5025
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":0E36
      MousePointer    =   99  'Custom
      Top             =   5025
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":0F88
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":10DA
      MousePointer    =   99  'Custom
      Top             =   4620
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":122C
      MousePointer    =   99  'Custom
      Top             =   4620
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":137E
      MousePointer    =   99  'Custom
      Top             =   4395
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":14D0
      MousePointer    =   99  'Custom
      Top             =   4395
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":1622
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1774
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":18C6
      MousePointer    =   99  'Custom
      Top             =   3975
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1A18
      MousePointer    =   99  'Custom
      Top             =   3975
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1B6A
      MousePointer    =   99  'Custom
      Top             =   3750
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1CBC
      MousePointer    =   99  'Custom
      Top             =   3525
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1E0E
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":1F60
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":20B2
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":2204
      MousePointer    =   99  'Custom
      Top             =   2670
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":2356
      MousePointer    =   99  'Custom
      Top             =   2445
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":24A8
      MousePointer    =   99  'Custom
      Top             =   2220
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":25FA
      MousePointer    =   99  'Custom
      Top             =   2010
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":274C
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":289E
      MousePointer    =   99  'Custom
      Top             =   1575
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":29F0
      MousePointer    =   99  'Custom
      Top             =   1365
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":2B42
      MousePointer    =   99  'Custom
      Top             =   1125
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":2C94
      MousePointer    =   99  'Custom
      Top             =   915
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":2DE6
      MousePointer    =   99  'Custom
      Top             =   915
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":2F38
      MousePointer    =   99  'Custom
      Top             =   3750
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":308A
      MousePointer    =   99  'Custom
      Top             =   3525
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":31DC
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":332E
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3480
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":35D2
      MousePointer    =   99  'Custom
      Top             =   2670
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3724
      MousePointer    =   99  'Custom
      Top             =   2445
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3876
      MousePointer    =   99  'Custom
      Top             =   2220
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   11
      Left            =   2520
      MouseIcon       =   "frmCrearPersonaje.frx":39C8
      MousePointer    =   99  'Custom
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3B1A
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3C6C
      MousePointer    =   99  'Custom
      Top             =   1575
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3DBE
      MousePointer    =   99  'Custom
      Top             =   1365
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   2640
      MouseIcon       =   "frmCrearPersonaje.frx":3F10
      MousePointer    =   99  'Custom
      Top             =   1125
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   3150
      MouseIcon       =   "frmCrearPersonaje.frx":4062
      MousePointer    =   99  'Custom
      Top             =   4830
      Width           =   180
   End
   Begin VB.Image boton 
      Height          =   375
      Index           =   1
      Left            =   6720
      MouseIcon       =   "frmCrearPersonaje.frx":41B4
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1875
   End
   Begin VB.Image boton 
      Height          =   330
      Index           =   0
      Left            =   9240
      MouseIcon       =   "frmCrearPersonaje.frx":4306
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   2040
   End
End
Attribute VB_Name = "frmCrearPersonaje1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public SkillPoints As Byte


Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

CheckData = True


End Function

Private Sub boton_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
                UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            If CheckDatos() Then
#If SeguridadAlkon Then
    UserPassword = md5.GetMD5String(txtPasswd.Text)
    Call md5.MD5Reset
#Else
    UserPassword = txtPasswd.Text
#End If
    UserEmail = txtCorreo.Text
    UserPreg = txtPreg.Text
    UserResp = txtResp.Text
    
    If Not CheckMailString(UserEmail) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 11

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call login(RandomCode)
    End If
End If
        End If
        
        
        
    Case 1
        frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call tirarDados
      
End Select



End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub Form_Load()
frmCrearPersonaje1.Picture = LoadPicture(App.Path & "\Graficos\crearPersonaje.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

SkillPoints = 10
puntos.Caption = SkillPoints
Call tirarDados
End Sub



Private Sub lstRaza_Click()
Call SetDadosFinal
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tirarDados()

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call VaginaJugosa(ClientPackages.tirarDados)
    End If

End Sub

Private Sub Command1_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

puntos.Caption = SkillPoints
End Sub

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

If txtPreg.Text = txtResp.Text Then
    MsgBox "La pregunta y Respuesta secreta son iguales, por favor vuelva a ingresarlos."
    Exit Function
End If

If Len(txtResp.Text) < 3 Then
    MsgBox "La respuesta debe tener al menos 3 letras."
    Exit Function
End If

If Not TextoValido(txtPasswd.Text) Then
    MsgBox "Tu password no es seguro"
    Exit Function
End If

If Not TextoValido(txtPreg.Text) Then
    MsgBox "Tu pregunta no es segura"
    Exit Function
End If

If Not TextoValido(txtResp.Text) Then
    MsgBox "Tu respuesta no es segura"
    Exit Function
End If

If Not MailValido(txtCorreo.Text) Then
    MsgBox "Tu mail no es seguro"
    Exit Function
End If

If puntos.Caption > 0 Then
    MsgBox "Aún tienes puntos que asignar"
    Exit Function
End If

CheckDatos = True

End Function

Function TextoValido(ByVal Texto As String) As Boolean
   
   If UCase$(Texto) = "ASD" Or _
      UCase$(Texto) = "ASDASD" Or _
      UCase$(Texto) = "ASDASD123" Or _
      UCase$(Texto) = "123456" Or _
      UCase$(Texto) = "ASD123" Or _
      UCase$(Texto) = "QWERTY" Or _
      UCase$(Texto) = "123" Or _
      UCase$(Texto) = "AAAAAA" Or _
      UCase$(Texto) = "AAA" Then
      TextoValido = False
      Exit Function
   End If

   TextoValido = True
End Function

Function MailValido(ByVal mail As String) As Boolean
   
   If UCase$(mail) = "A@A.A" Or _
      UCase$(mail) = "ASD@ASD.ASD" Or _
      UCase$(mail) = "ASD@ASD.COM" Or _
      UCase$(mail) = "A@A.COM" Or _
      UCase$(mail) = "A@ASD.COM" Or _
      UCase$(mail) = "A@ASD.ASD" Or _
      UCase$(mail) = "ASDASD@ASD.COM" Or _
      UCase$(mail) = "ASD123@ASD.ASD" Or _
      UCase$(mail) = "AAA@AAA.AAA" Or _
      UCase$(mail) = "AAA@AAA.COM" Then
      MailValido = False
      Exit Function
   End If

   MailValido = True
End Function

Private Sub botonn_Click()
    frmCrearPersonaje1.Visible = False
End Sub



Private Sub txtPreg_GotFocus()
    MsgBox ("ATENCION! ROBO DE PERSONAJES" & vbNewLine & "Lapsus Corp recomienda seleccionar una pregunta y respuesta que sólo usted sepa" & vbNewLine & "Es la única manera que tendrá usted de recuperar su personaje (y que tendrán los usuarios ajenos de robárselo)")
End Sub

Public Sub SetDadosFinal()
'CHOTS | Label Final de Atributos
With frmCrearPersonaje1
    Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
        Case "ENANO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 2
            .lblAgilidad2.Caption = .lbAgilidad.Caption - 2
            .lblInteligencia2.Caption = .lbInteligencia.Caption - 4
            .lblCarisma2.Caption = .lbCarisma.Caption - 1
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 3
        Case "GNOMO"
            .lblFuerza2.Caption = .lbFuerza.Caption - 2
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 2
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 5
            .lblCarisma2.Caption = .lbCarisma.Caption
            .lblConstitucion2.Caption = .lbConstitucion.Caption
        Case "ELFO"
            .lblFuerza2.Caption = .lbFuerza.Caption - 1
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 3
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 3
            .lblCarisma2.Caption = .lbCarisma.Caption + 2
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 1
        Case "ELFO OSCURO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 1
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 1
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 3
            .lblCarisma2.Caption = .lbCarisma.Caption - 2
            .lblConstitucion2.Caption = .lbConstitucion.Caption
        Case "HUMANO"
            .lblFuerza2.Caption = .lbFuerza.Caption
            .lblAgilidad2.Caption = .lbAgilidad.Caption
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 1
            .lblCarisma2.Caption = .lbCarisma.Caption + 1
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 2
        Case "ORCO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 3
            .lblAgilidad2.Caption = .lbAgilidad.Caption - 3
            .lblInteligencia2.Caption = .lbInteligencia.Caption - 7
            .lblCarisma2.Caption = .lbCarisma.Caption - 3
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 4
        Case Else
            .lblFuerza2.Caption = .lbFuerza.Caption
            .lblAgilidad2.Caption = .lbAgilidad.Caption
            .lblInteligencia2.Caption = .lbInteligencia.Caption
            .lblCarisma2.Caption = .lbCarisma.Caption
            .lblConstitucion2.Caption = .lbConstitucion.Caption
    End Select
    .lblMaxFuerza.Caption = .lblFuerza2.Caption * 2
    .lblMaxAgi.Caption = .lblAgilidad2.Caption * 2

    If UCase$(.lstRaza.List(.lstRaza.listIndex)) = "HUMANO" Then
        If Val(.lblMaxFuerza.Caption) > 35 Then .lblMaxFuerza.Caption = 35
        If Val(.lblMaxAgi.Caption) > 35 Then .lblMaxAgi.Caption = 35
    Else
        If Val(.lblMaxFuerza.Caption) > 36 Then .lblMaxFuerza.Caption = 36
        If Val(.lblMaxAgi.Caption) > 36 Then .lblMaxAgi.Caption = 36
    End If
End With
'CHOTS | Label Final de Atributos
End Sub
