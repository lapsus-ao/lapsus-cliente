VERSION 5.00
Begin VB.Form frmPasswdSinPadrinos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5010
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
   Moveable        =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   420
      Left            =   105
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3495
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   3885
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3540
      Width           =   1080
   End
   Begin VB.TextBox txtPasswdCheck 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   765
      TabIndex        =   7
      Top             =   2910
      Width           =   3510
   End
   Begin VB.TextBox txtPasswd 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   765
      TabIndex        =   5
      Top             =   2295
      Width           =   3510
   End
   Begin VB.TextBox txtCorreo 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   765
      TabIndex        =   3
      Top             =   1710
      Width           =   3510
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Verifiaci�n del password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   6
      Top             =   2670
      Width           =   3555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   2040
      Width           =   3555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Direcci�n de correo electronico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   1455
      Width           =   3555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmPasswdSinPadrinos.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�CUIDADO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1965
      TabIndex        =   0
      Top             =   105
      Width           =   1035
   End
End
Attribute VB_Name = "frmPasswdSinPadrinos"
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

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub
