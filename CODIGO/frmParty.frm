VERSION 5.00
Begin VB.Form frmParty 
   Caption         =   "Sistema de Party"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLider 
      Caption         =   "Lider"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdHablar 
      Caption         =   "Enviar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtHablar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar Party"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdStat 
      Caption         =   "Estado Party"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Echar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear Party"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblHablar 
      Alignment       =   2  'Center
      Caption         =   "Hablar por Party"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   3000
      Y1              =   240
      Y2              =   3840
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CHOTS | Formulario Party

Private Sub cmdAccept_Click()
Call SendData("/ACCEPTPARTY " & txtUser.Text)
End Sub

Private Sub cmdCerrar_Click()
Call SendData("/SALIRPARTY")
Unload frmParty
End Sub

Private Sub cmdCrear_Click()
Call SendData("/CREARPARTUSA")
Unload frmParty
End Sub

Private Sub cmdHablar_Click()
Call SendData("/PMSG " & txtHablar.Text)
Unload frmParty
End Sub

Private Sub cmdKick_Click()
Call SendData("/ECHARPARTY " & txtUser.Text)
Unload frmParty
End Sub

Private Sub cmdLider_Click()
Call SendData("/PARTYLIDER " & txtUser.Text)
End Sub

Private Sub cmdStat_Click()
Call SendData("/ONLINEPARTY")
Unload frmParty
End Sub

Private Sub Command3_Click()
Unload frmParty
End Sub

Private Sub Form_Load()

If enParty = True Then
    cmdCrear.Enabled = False
    cmdStat.Enabled = True
    cmdKick.Enabled = True
    cmdAccept.Enabled = True
    cmdHablar.Enabled = True
    txtUser.Enabled = True
    txtHablar.Enabled = True
    cmdCerrar.Enabled = True
Else
    cmdCrear.Enabled = True
    cmdStat.Enabled = False
    cmdKick.Enabled = False
    cmdAccept.Enabled = False
    cmdHablar.Enabled = False
    txtUser.Enabled = False
    txtHablar.Enabled = False
    cmdCerrar.Enabled = False
End If

End Sub
