VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administraci�n del Clan"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Propuestas de alianzas"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5460
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4950
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3930
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":069A
         Left            =   120
         List            =   "frmGuildLeader.frx":069C
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":069E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":07F0
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0942
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0A94
         Left            =   120
         List            =   "frmGuildLeader.frx":0A96
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         Caption         =   "Abrir elecciones"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0A98
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0BEA
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1170
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   840
         ItemData        =   "frmGuildLeader.frx":0D3C
         Left            =   120
         List            =   "frmGuildLeader.frx":0D3E
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Alignment       =   2  'Center
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1620
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2009 Juan Andres Dalmasso (CHOTS)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub
Private Sub cmdElecciones_Click()
    Call SendData("ABREELEC")
    Unload Me
End Sub

Private Sub Command1_Click()

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.listIndex))

'Unload Me

End Sub

Private Sub Command2_Click()

frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.listIndex))

'Unload Me

End Sub

Private Sub Command3_Click()

Dim k$

k$ = Replace(txtguildnews, vbCrLf, "�")

Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()

frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & Trim$(ReadField(1, guildslist.List(guildslist.listIndex), Asc("("))))

'Unload Me

End Sub

Private Sub Command5_Click()

Call frmGuildDetails.Show(vbModal, frmGuildLeader)

'Unload Me

End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
    Call SendData("ENVALPRO")
End Sub


Private Sub Command8_Click()
Unload Me
frmMain.SetFocus
End Sub


Public Sub ParseLeaderInfo(ByVal data As String)

If Me.Visible Then Exit Sub

Dim r%, t%

r% = Val(ReadField(1, data, Asc("�")))

'Dim j As Integer, k As Integer
'For j = 0 To guildslist.ListCount - 1
'    Me.guildslist.RemoveItem 0
'Next j
'k = CInt(ReadField(1, Rdata, 44))

'For j = 1 To k

For t% = 1 To r%
    'guildslist.AddItem ReadField(1 + T%, Data, Asc("�"))
    guildslist.AddItem ReadField(1, ReadField(1 + t%, data, Asc("�")), Asc("@")) & " (" & ReadField(2, ReadField(1 + t%, data, Asc("�")), Asc("@")) & ")"
Next t%

r% = Val(ReadField(t% + 1, data, Asc("�")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(t% + 1 + k%, data, Asc("�"))
Next k%

txtguildnews = Replace(ReadField(t% + k% + 1, data, Asc("�")), "�", vbCrLf)

t% = t% + k% + 2

r% = Val(ReadField(t%, data, Asc("�")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(t% + k%, data, Asc("�"))
Next k%

Me.Show , frmMain

End Sub


Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

