VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   0  'None
   Caption         =   "Sastrería"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.ListBox lstRopas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      ItemData        =   "frmSastre.frx":0000
      Left            =   240
      List            =   "frmSastre.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   4250
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3360
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "frmSastre"
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
Unload Me
End Sub

Private Sub Image2_Click()
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
Call VaginaJugosa("CND" & ObjSastre(lstRopas.listIndex) & "," & tetat)
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "vesti.JPG")
End Sub

