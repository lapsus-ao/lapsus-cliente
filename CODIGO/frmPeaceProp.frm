VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
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
   ScaleHeight     =   2895
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3720
      MouseIcon       =   "frmPeaceProp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      MouseIcon       =   "frmPeaceProp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalles"
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmPeaceProp.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmPeaceProp.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmPeaceProp.frx":0548
      Left            =   120
      List            =   "frmPeaceProp.frx":054A
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPeaceProp"
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

Private tipoprop As TIPO_PROPUESTA
Private Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub ParsePeaceOffers(ByVal s As String)

Dim t%, r%

t% = Val(ReadField(1, s, 44))

For r% = 1 To t%
    Call lista.AddItem(ReadField(r% + 1, s, 44))
Next r%


tipoprop = PAZ

Me.Show vbModeless, frmMain

End Sub

Public Sub ParseAllieOffers(ByVal s As String)

Dim t%, r%

t% = Val(ReadField(1, s, 44))

For r% = 1 To t%
    Call lista.AddItem(ReadField(r% + 1, s, 44))
Next r%

tipoprop = ALIANZA
Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("PEACEDET" & lista.List(lista.listIndex))
Else
    Call SendData("ALLIEDET" & lista.List(lista.listIndex))
End If
End Sub

Private Sub Command3_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("ACEPPEAT" & lista.List(lista.listIndex))
Else
    Call SendData("ACEPALIA" & lista.List(lista.listIndex))
End If
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
If tipoprop = PAZ Then
    Call SendData("RECPPEAT" & lista.List(lista.listIndex))
Else
    Call SendData("RECPALIA" & lista.List(lista.listIndex))
End If
Me.Hide
Unload Me
End Sub
