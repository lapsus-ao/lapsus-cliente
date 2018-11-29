VERSION 5.00
Begin VB.Form frmRunas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar premios"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstObjetos 
      Height          =   1230
      ItemData        =   "frmRunas.frx":0000
      Left            =   210
      List            =   "frmRunas.frx":000A
      TabIndex        =   0
      Top             =   250
      Width           =   4400
   End
   Begin VB.Image cerrar 
      Height          =   495
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
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trofeos de Plata: 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmRunas.frx":002E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   210
      TabIndex        =   1
      Top             =   2520
      Width           =   4410
   End
End
Attribute VB_Name = "frmRunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sistema de Canjes LapsusAO AO
'Obtenido de LOD AO
'Reprogramado y Adaptado por CHOTS Para SilvAO 2008
'Reprogramado y Adaptado por CHOTS Para LapsusAO 2010

Private Sub cambiar_Click()
Call SendData("CRN" & lstObjetos.listIndex)
Unload Me
End Sub

Private Sub cerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()

lstObjetos.listIndex = 0
lblOro.Caption = "Trofeos de Plata: 3"
frmRunas.Picture = LoadPicture(DirGraficos & "Runas.jpg")

End Sub

Private Sub lstObjetos_Click()

Select Case lstObjetos.listIndex
    Case 0
        lblOro.Caption = "Trofeos de Plata: 3"
    Case 1
        lblOro.Caption = "Runas: 3"
End Select

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
