VERSION 5.00
Begin VB.Form frmEspia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Espía"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Dejar de Espiar"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblMan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "220 / 230"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   4095
   End
   Begin VB.Label lblHp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "220 / 230"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   4095
   End
   Begin VB.Shape man 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      Height          =   360
      Left            =   135
      Top             =   1215
      Width           =   4080
   End
   Begin VB.Shape hp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   360
      Left            =   135
      Top             =   495
      Width           =   4080
   End
   Begin VB.Label lblEspiado 
      Caption         =   "Espiando a: TheCheater"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      Top             =   480
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "frmEspia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStop_Click()
Call VaginaJugosa("/ESPIAR IOPUJA")
Unload frmEspia
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
