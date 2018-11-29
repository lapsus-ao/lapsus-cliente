VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
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
   ScaleHeight     =   3450
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildAdm.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmGuildAdm.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ListBox GuildsList 
         Height          =   2010
         ItemData        =   "frmGuildAdm.frx":02A4
         Left            =   120
         List            =   "frmGuildAdm.frx":02A6
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub Command1_Click()

'If GuildsList.ListIndex = 0 Then Exit Sub
Call VaginaJugosa("CLANDETAILS" & Trim$(ReadField(1, guildslist.List(guildslist.listIndex), Asc("("))))

End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Public Sub ParseGuildList(ByVal Rdata As String)
'CHOTS | Recibe la lista de clanes y la ordena de Acuerdo a sus GuildPoints

Dim j As Integer, k As Integer, l As Integer, N As Integer
For j = 0 To guildslist.ListCount - 1
    Me.guildslist.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))


Dim aux As String

ReDim vecClan(1 To k) As String

For j = 1 To k
    vecClan(j) = ReadField(1 + j, Rdata, 44)
Next j

For l = 1 To (k - 1)
    For N = (l + 1) To k
        If Val(ReadField(2, vecClan(l), Asc("@"))) < Val(ReadField(2, vecClan(N), Asc("@"))) Then
            aux = vecClan(l)
            vecClan(l) = vecClan(N)
            vecClan(N) = aux
        End If
    Next N
Next l

For j = 1 To k
    If ReadField(2, vecClan(j), Asc("@")) <> -1 Then guildslist.AddItem ReadField(1, vecClan(j), Asc("@")) & " (Puntos: " & ReadField(2, vecClan(j), Asc("@")) & ")"
Next j

Me.Show vbModal, frmMain
'CHOTS | Recibe la lista de clanes y la ordena de Acuerdo a sus GuildPoints
End Sub

