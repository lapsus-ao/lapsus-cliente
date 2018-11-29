VERSION 5.00
Begin VB.Form frmQuest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest"
   ClientHeight    =   3660
   ClientLeft      =   7650
   ClientTop       =   3960
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4875
   Begin VB.ListBox lstQuest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmQuest.frx":0000
      Left            =   240
      List            =   "frmQuest.frx":0002
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4680
      Top             =   0
      Width           =   255
   End
   Begin VB.Image iniciar 
      Height          =   495
      Left            =   2760
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matar 20 Lobos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2415
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misil M�gico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
frmQuest.Picture = LoadPicture(DirGraficos & "Quest.jpg")

'CHOTS | Cargamos las Quests
frmQuest.lstQuest.Clear
frmQuest.lstQuest.AddItem "�Lobo, est�s?"
frmQuest.lstQuest.AddItem "Viejos feos como monos"
frmQuest.lstQuest.AddItem "Todo un palo"
frmQuest.lstQuest.AddItem "Roto y mal parado"
frmQuest.lstQuest.AddItem "Cruz Diablo"
frmQuest.lstQuest.AddItem "Brujas de alma sencilla"
frmQuest.lstQuest.AddItem "Aliento de drag�n"
frmQuest.lstQuest.AddItem "Golem de paternal"
frmQuest.lstQuest.AddItem "Vencedores vencidos"
frmQuest.lstQuest.AddItem "Cordero atado"

lblRec.Caption = "Misil M�gico"
lblDesc.Caption = "Necesito deshacerme de esos mugrientos lobos que abundan en estas tierras. Solo 20 bastar�n para hacerme feliz"
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub iniciar_Click()
Call VaginaJugosa("REAQ" & lstQuest.listIndex + 1)
Unload frmQuest
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub lstQuest_Click()

Select Case lstQuest.listIndex
    Case 0
        lblRec.Caption = "Misil M�gico"
        lblDesc.Caption = "Necesito deshacerme de esos mugrientos lobos que abundan en estas tierras. Solo 20 bastar�n para hacerme feliz"
    Case 1
        lblRec.Caption = "10.000 monedas de oro; 1 Punto de Usuario; Inmovilizar"
        lblDesc.Caption = "El Bosque Dorck apesta a Orcos, si pudieras matar 15 de ellos te recompensar�"
    Case 2
        lblRec.Caption = "15.000 monedas de oro; 3 Puntos de Usuario; Rel�mpago"
        lblDesc.Caption = "�Animales con garrotes? S�, de esos hablo. Mata 10 ogros y luego ven por tu recompensa"
    Case 3
        lblRec.Caption = "30.000 monedas de oro; 5 Puntos de Usuario"
        lblDesc.Caption = "La m�s repugnante evolucion para la m�s repugnante criatura. Si te deshaces de 15 Lord Orco te dar� unas monedas"
    Case 4
        lblRec.Caption = "50.000 monedas de oro; 10 Puntos de Usuario"
        lblDesc.Caption = "Criaturas del infierno, tambi�n llamados Demonios. Dungeon Marabel estar�a mejor si aniquilas 5"
    Case 5
        lblRec.Caption = "75.000 monedas de oro; 15 Puntos de Usuario"
        lblDesc.Caption = "Una vez me enamor� de una Bruja, fue mi peor decisi�n. Mata 10 y luego vuelve conmigo"
    Case 6
        lblRec.Caption = "100.000 monedas de oro; 20 Puntos de Usuario"
        lblDesc.Caption = "Si matas 15 de estos peque�os monstruos que habitan el Dungeon Dragon, te dar� un saco de monedas que te har�n feliz"
    Case 7
        lblRec.Caption = "125.000 monedas de oro; 30 Puntos de Usuario"
        lblDesc.Caption = "Moles de oro, tambi�n llamados Golems. �Puedes creer que existan? Pues yo s�, mata 3 y luego ven a verme"
    Case 8
        lblRec.Caption = "150.000 monedas de oro; 50 Puntos de Usuario"
        lblDesc.Caption = "Los ogros se han equipado, y quieren venganza! Si matas 10 de estos Ogros Armados me har�as un gran favor"
    Case 9
        lblRec.Caption = "100 Puntos de Usuario"
        lblDesc.Caption = "Corderito... no es bueno mantener al lobo hambriento. Si consigues aniquilar a 100 de estos lobos sueltos te recompensar�"
End Select

End Sub
