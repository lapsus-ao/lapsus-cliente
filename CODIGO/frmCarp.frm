VERSION 5.00
Begin VB.Form frmCarp 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   3225
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4710
   ControlBox      =   0   'False
   FillColor       =   &H80000001&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.ListBox lstArmas 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2200
      Left            =   240
      TabIndex        =   0
      Top             =   470
      Width           =   4240
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   120
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3360
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmCarp"
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

Private Sub Image1_Click()
On Error Resume Next
Dim tetat As Integer
tetat = Text1.text
If tetat <= 0 Then
    Unload Me
    Exit Sub
End If
If tetat > 10000 Then
    Unload Me
    Exit Sub
End If
Call VaginaJugosa("CNC" & ObjCarpintero(lstArmas.listIndex) & "," & tetat)
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "carpi.JPG")
End Sub

