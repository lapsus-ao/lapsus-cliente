VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2520
      TabIndex        =   2
      Text            =   "1"
      Top             =   2880
      Width           =   405
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4220
   End
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4220
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   3315
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   90
      Top             =   2850
      Width           =   990
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   3120
      Top             =   0
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   240
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "herre.jpg")
End Sub

Private Sub Image1_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Image2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image4_Click()

On Error Resume Next
Dim tortugazeta As Long
tortugazeta = Text1.text
If tortugazeta <= 0 Then
    Unload Me
    Exit Sub
End If
If tortugazeta > 10000 Then
    Unload Me
    Exit Sub
End If
If lstArmas.Visible Then
    Call VaginaJugosa("CNS" & ArmasHerrero(lstArmas.listIndex) & "," & tortugazeta)
Else
    Call VaginaJugosa("CNS" & ArmadurasHerrero(lstArmaduras.listIndex) & "," & tortugazeta)
End If

Unload Me
End Sub

