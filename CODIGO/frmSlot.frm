VERSION 5.00
Begin VB.Form frmSlot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Casino Arghal"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCambNum3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   1200
   End
   Begin VB.Timer tmrCambNum2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SPIN"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Caption         =   "Cobrar"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Timer tmrCambNum1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   240
   End
   Begin VB.Timer tmrFin 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8040
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Slot"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4455
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   4695
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   8
         Left            =   2880
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   7
         Left            =   1800
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   6
         Left            =   720
         TabIndex        =   18
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   5
         Left            =   2880
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   4
         Left            =   1800
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   2
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Su Crédito"
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Monto:"
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Apuesta: "
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Gana:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblEstado 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   1560
      TabIndex        =   8
      Top             =   5880
      Width           =   5655
   End
   Begin VB.Label Label8 
      Caption         =   "American Slot - LapsusAO"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label9 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
End
Attribute VB_Name = "frmSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public contador As Integer
Public resultado1 As Byte
Public resultado2 As Byte
Public resultado3 As Byte
Public apuestaa As Long
Private Sub Command1_Click()
'Timba = 2
apuestaa = Val(Text1.Text)

If apostar(apuestaa) = True Then
 tmrCambNum1.Enabled = True 'CHOTS | Empieza a "girar" el Slot
 tmrCambNum2.Enabled = True 'CHOTS | Empieza a "girar" el Slot
 tmrCambNum3.Enabled = True 'CHOTS | Empieza a "girar" el Slot
 tmrFin.Enabled = True 'CHOTS | Comienza el conteo para "parar" el giro
 Command1.Enabled = False
 Label6.Caption = "Apuesta: " & apuestaa
 Text1.Enabled = False
End If



End Sub
Private Sub setInterval()
 tmrCambNum1.Interval = RandomNumber(50, 150)
 tmrCambNum2.Interval = RandomNumber(50, 150)
 tmrCambNum3.Interval = RandomNumber(50, 150)
End Sub

Private Sub Command3_Click()

Call cobrar(UserFichas)

End Sub

Private Sub Form_Load()
'Timba = 2
contador = 0
Call setInterval
End Sub

Private Sub Text1_Change()

If Val(Text1) > 10000 Then 'CHOTS | Para evitar que pierdan su Sueldo por viciosos
 Text1.Text = 10000
End If

If Val(Text1) <= 0 Then 'CHOTS | Para evitar que hackeen gatos!
 Text1.Text = 1
End If

If Val(Text1) > UserFichas Then 'CHOTS | Por cuestiones de comodidad
 Text1.Text = UserFichas
End If

End Sub

Private Sub tmrCambNum1_Timer()

 lblEstado.Caption = "Girando Slot..."

  lblSlot(3).Caption = RandomNumber(0, 9)
  lblSlot(0).Caption = (lblSlot(3).Caption) - 1
  lblSlot(6).Caption = (lblSlot(3).Caption) + 1
  
  If lblSlot(0).Caption = -1 Then
   lblSlot(0).Caption = 9
  End If
  
  If lblSlot(6).Caption = -1 Then
   lblSlot(6).Caption = 9
  End If
  
  If lblSlot(0).Caption = 10 Then
   lblSlot(0).Caption = 0
  End If
  
  If lblSlot(6).Caption = 10 Then
   lblSlot(6).Caption = 0
  End If
  
End Sub

Private Sub tmrCambNum2_Timer()

 lblEstado.Caption = "Girando Slot..."

  lblSlot(4).Caption = RandomNumber(0, 9)
  lblSlot(1).Caption = (lblSlot(4).Caption) - 1
  lblSlot(7).Caption = (lblSlot(4).Caption) + 1
  
  If lblSlot(1).Caption = -1 Then
   lblSlot(1).Caption = 9
  End If
  
  If lblSlot(7).Caption = -1 Then
   lblSlot(7).Caption = 9
  End If
  
  If lblSlot(1).Caption = 10 Then
   lblSlot(1).Caption = 0
  End If
  
  If lblSlot(7).Caption = 10 Then
   lblSlot(7).Caption = 0
  End If
  
End Sub
  
Private Sub tmrCambNum3_Timer()

 lblEstado.Caption = "Girando Slot..."

  lblSlot(5).Caption = RandomNumber(0, 9)
  lblSlot(2).Caption = (lblSlot(5).Caption) - 1
  lblSlot(8).Caption = (lblSlot(5).Caption) + 1
  
  If lblSlot(2).Caption = -1 Then
   lblSlot(2).Caption = 9
  End If
  
  If lblSlot(8).Caption = -1 Then
   lblSlot(8).Caption = 9
  End If
  
  If lblSlot(2).Caption = 10 Then
   lblSlot(2).Caption = 0
  End If
  
  If lblSlot(8).Caption = 10 Then
   lblSlot(8).Caption = 0
  End If
  
End Sub

Private Sub tmrFin_Timer()

Call setInterval

contador = contador + 1


Select Case contador
 Case 3
  tmrCambNum1.Enabled = False
  
 Case 4
  tmrCambNum2.Enabled = False
  
 Case 5
  tmrCambNum3.Enabled = False
  contador = 0
  tmrFin.Enabled = False
  Command1.Enabled = True
  Text1.Enabled = True
  resultado1 = Val(lblSlot(3).Caption)
  resultado2 = Val(lblSlot(4).Caption)
  resultado3 = Val(lblSlot(5).Caption)
  Call verSiGano(resultado1, resultado2, resultado3, apuestaa)
  
End Select


End Sub
