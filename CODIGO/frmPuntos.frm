VERSION 5.00
Begin VB.Form frmPuntos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puntos de Usuario"
   ClientHeight    =   3450
   ClientLeft      =   8580
   ClientTop       =   4155
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   2865
   Begin VB.ListBox lstClan 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image transferir 
      Height          =   375
      Left            =   480
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblPuntos 
      BackStyle       =   0  'Transparent
      Caption         =   "83"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   330
      Width           =   855
   End
End
Attribute VB_Name = "frmPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmPuntos.Picture = LoadPicture(DirGraficos & "Transferir.jpg")
End Sub

Private Sub transferir_Click()
Dim puntos As Integer
puntos = 0
puntos = Val(InputBox("Ingrese la cantidad de Puntos a Transferir:", , 0))
If puntos < 1 Then Exit Sub
Call SendData("TPT" & lstClan.text & "," & puntos)
Unload Me
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
