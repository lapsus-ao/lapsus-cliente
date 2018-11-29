VERSION 5.00
Begin VB.Form frmOldPersonaje 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   251
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2265
      TabIndex        =   0
      Top             =   705
      Width           =   4530
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2265
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   4530
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   510
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   4920
      MouseIcon       =   "frmOldPersonaje.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   1365
      MouseIcon       =   "frmOldPersonaje.frx":0152
      MousePointer    =   99  'Custom
      Top             =   3105
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   3120
      MouseIcon       =   "frmOldPersonaje.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   960
   End
End
Attribute VB_Name = "frmOldPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
