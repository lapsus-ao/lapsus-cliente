VERSION 5.00
Begin VB.Form frmTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar premios"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFoto 
      FillColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4170
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   270
      Width           =   540
   End
   Begin VB.ListBox lstObjetos 
      Height          =   1230
      ItemData        =   "frmTrade.frx":0000
      Left            =   210
      List            =   "frmTrade.frx":0037
      TabIndex        =   0
      Top             =   270
      Width           =   3885
   End
   Begin VB.Image Image1 
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
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmTrade.frx":0174
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
      TabIndex        =   1
      Top             =   2540
      Width           =   4410
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sistema de Canjes LapsusAO AO
'Obtenido de LOD AO
'Reprogramado y Adaptado por CHOTS Para SilvAO 2008
'Reprogramado y Adaptado por CHOTS Para LapsusAO 2010

Private Sub Command2_Click()
End Sub

Private Sub cambiar_Click()
Call SendData("CTR" & lstObjetos.listIndex)
Unload Me
End Sub

Private Sub Form_Load()
lstObjetos.listIndex = 0
lblOro.Caption = "Trofeos de Oro: 30"
frmTrade.Picture = LoadPicture(DirGraficos & "Ermitanio.jpg")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
Private Sub lstObjetos_Click()

Select Case lstObjetos.listIndex
    Case 0
        lblOro.Caption = "Trofeos de Oro: 30"
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16
        lblOro.Caption = "Trofeos de Oro: 20"
    Case 11, 12, 13, 14
        lblOro.Caption = "Trofeos de Oro: 10"
End Select

Select Case lstObjetos.listIndex
    Case 0
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/4811.bmp")
    Case 1
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27106.bmp")
    Case 2
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27100.bmp")
    Case 3
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27090.bmp")
    Case 4
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27092.bmp")
    Case 5
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27094.bmp")
    Case 6
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27098.bmp")
    Case 7
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27108.bmp")
    Case 8
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27110.bmp")
    Case 9
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27088.bmp")
    Case 10
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27112.bmp")
    Case 11
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27104.bmp")
    Case 12
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/13191.bmp")
    Case 13
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27102.bmp")
    Case 14
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/27096.bmp")
    Case 15
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/1046.bmp")
    Case 16
        picFoto.Picture = LoadPicture(App.Path & "/Graficos/1045.bmp")
End Select

End Sub

