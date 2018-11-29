VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7200
   ClientLeft      =   5355
   ClientTop       =   2745
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   477.019
   ScaleMode       =   0  'User
   ScaleWidth      =   639.002
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox MP3Files 
      Height          =   480
      Left            =   180
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9630
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   8280
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox Status 
         Height          =   2175
         Left            =   2160
         TabIndex        =   2
         Top             =   4440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3836
         _Version        =   393217
         BackColor       =   -2147483647
         BorderStyle     =   0
         TextRTF         =   $"frmCargando.frx":0000
      End
   End
End
Attribute VB_Name = "frmCargando"
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
Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long


Private Sub Form_Load()
Dim result As Long
Call SetWindowLong(status.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
status.SelAlignment = rtfCenter
LOGO.Picture = LoadPicture(DirGraficos & "cargando.jpg")
Call Audio.PlayMIDI("7.mid")
End Sub


 
