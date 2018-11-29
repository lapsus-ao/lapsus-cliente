VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PasswordTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   4815
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3120
      Width           =   2880
   End
   Begin VB.TextBox NombreTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4815
      TabIndex        =   0
      Top             =   2280
      Width           =   2880
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5280
      MouseIcon       =   "frmConnect.frx":000C
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v2.1.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   5880
      TabIndex        =   2
      Top             =   1440
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   5280
      MouseIcon       =   "frmConnect.frx":015E
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5370
      MouseIcon       =   "frmConnect.frx":02B0
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   5280
      MouseIcon       =   "frmConnect.frx":0402
      MousePointer    =   99  'Custom
      Top             =   4875
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   5280
      MouseIcon       =   "frmConnect.frx":0554
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   1890
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 130, 110, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        Call WriteVar(App.Path & "\init\version.dat", "VERSION", "Graficos", "0")
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 130, 110, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub

Private Sub Form_Load()
    EngineRun = False
    FONDO.Picture = LoadPicture(App.Path & "\Graficos\Conectar.jpg")
    version.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub Image1_Click()

Call Audio.PlayWave(SND_CLICK)
        
Call Audio.PlayMIDI("7.mid")
        
EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        Me.MousePointer = 11

Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub


Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

#If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    #Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
    #End If
            If frmConnect.MousePointer = 11 Then
                Exit Sub
            End If
           
            UserName = NombreTXT.Text
        Dim aux As String
        aux = PasswordTXT.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
            If CheckUserData(False) = True Then
                'SendNewChar = False
                EstadoLogin = Normal
                Me.MousePointer = 11
    #If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
    #Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
                frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If
            End If
End Sub

Private Sub Image3_Click()
    frmBorrar.Show
End Sub

Private Sub Image4_Click()
    frmRecuperar.Show
End Sub

Private Sub Image5_Click()
    Shell ("cmd /c start http://www.lapsus2017.com"), vbHide
End Sub

Private Sub NombreTXT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image2_Click
    End If
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image2_Click
    End If
End Sub
