VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "1"
      Top             =   6600
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   435
      ScaleHeight     =   600
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   750
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6780
      Width           =   465
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   1
      Left            =   3855
      TabIndex        =   1
      Top             =   1800
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   0
      Left            =   615
      TabIndex        =   0
      Top             =   1800
      Width           =   2490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2580
      TabIndex        =   9
      Top             =   6645
      UseMnemonic     =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3855
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6165
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   615
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6150
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   7
      Top             =   975
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   75
   End
End
Attribute VB_Name = "frmBancoObj"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then VaginaJugosa ("FINBAN")
End Sub

Private Sub cantidad_Change()
If Val(cantidad.text) < 0 Then
    cantidad.text = 1
End If

If Val(cantidad.text) > MAX_INVENTORY_OBJS Then
    cantidad.text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_Click()
VaginaJugosa ("FINBAN")
End Sub



Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\boveda.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\Bot�nComprar.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Bot�nvender.jpg")

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\Bot�nComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Bot�nvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(index).List(List1(index).listIndex) = "Nada" Or _
   List1(index).listIndex < 0 Then Exit Sub

Select Case index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).listIndex
        
        VaginaJugosa ("RETI" & "," & List1(0).listIndex + 1 & "," & cantidad.text)
        
   Case 1
        LastIndex2 = List1(1).listIndex
        If Not Inventario.Equipped(List1(1).listIndex + 1) Then
            VaginaJugosa ("DEPO" & "," & List1(1).listIndex + 1 & "," & cantidad.text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
List1(0).Clear

List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\Bot�nComprarApretado.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Bot�nvender.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Bot�nvenderapretado.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\Bot�nComprar.jpg")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub list1_Click(index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case index
    Case 0
        Label1(0).Caption = UserBancoInventory(List1(0).listIndex + 1).Name
        Label1(2).Caption = UserBancoInventory(List1(0).listIndex + 1).Amount
        Select Case UserBancoInventory(List1(0).listIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserBancoInventory(List1(0).listIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        Call DrawGrhtoHdc(Picture1.hwnd, Picture1.hDC, UserBancoInventory(List1(0).listIndex + 1).GrhIndex, SR, DR)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).listIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).listIndex + 1)
        Select Case Inventario.OBJType(List1(1).listIndex + 1)
            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).listIndex + 1)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).listIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(List1(1).listIndex + 1)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        Call DrawGrhtoHdc(Picture1.hwnd, Picture1.hDC, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select
Picture1.Refresh

End Sub
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\Bot�nComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Bot�nvender.jpg")
    Image1(1).Tag = 1
End If
End Sub
