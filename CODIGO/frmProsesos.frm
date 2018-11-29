VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProsesos 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Foto"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid FlxGd 
      Height          =   6015
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   4
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmProsesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim Nombre As String
Nombre = ReadField(3, frmProsesos.Caption, Asc(" "))
Call frmMain.Capturar_Guardar("Prosesos de " & Nombre)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Unload Me
End Sub
