VERSION 5.00
Begin VB.Form frmPts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puntos de Usuario"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstItems 
      Height          =   1230
      ItemData        =   "frmPts.frx":0000
      Left            =   230
      List            =   "frmPts.frx":0019
      TabIndex        =   1
      Top             =   210
      Width           =   4380
   End
   Begin VB.Image cerrar 
      Height          =   375
      Left            =   2760
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image cambiar 
      Height          =   375
      Left            =   2760
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblReq 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   3
      Top             =   2020
      Width           =   975
   End
   Begin VB.Label lblPts 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   2
      Top             =   1700
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmPts.frx":008C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1060
      Left            =   210
      TabIndex        =   0
      Top             =   2500
      Width           =   4410
   End
End
Attribute VB_Name = "frmPts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sistema de Puntos de usuario de LapsusAO AO
'Programado por CHOTS para LapsusAO 2010

Private Sub cambiar_Click()
Call VaginaJugosa("CPT" & lstItems.listIndex)
Unload Me
End Sub

Private Sub cerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmPts.Picture = LoadPicture(DirGraficos & "Puntos.jpg")
End Sub

Private Sub lstItems_Click()

Select Case lstItems.listIndex
    Case 0, 1
        lblReq.Caption = 250
    Case 3, 6
        lblReq.Caption = 100
    Case 2, 5
        lblReq.Caption = 500
    Case 4
        lblReq.Caption = 200
End Select

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
