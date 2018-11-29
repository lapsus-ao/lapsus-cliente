VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picture2v2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      Picture         =   "frmRetos.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   4215
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtPareja 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1790
         Width           =   2025
      End
      Begin VB.TextBox txtRival2 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1390
         Width           =   2025
      End
      Begin VB.TextBox txtPuntos2 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   195
         Left            =   850
         TabIndex        =   9
         Text            =   "0"
         Top             =   650
         Width           =   3210
      End
      Begin VB.TextBox txtOro2 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   240
         Left            =   850
         TabIndex        =   8
         Text            =   "0"
         Top             =   170
         Width           =   3210
      End
      Begin VB.TextBox txtRival1 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   860
         TabIndex        =   7
         Top             =   1035
         Width           =   2025
      End
      Begin VB.CheckBox chkItems2 
         BackColor       =   &H80000007&
         Height          =   195
         Left            =   3750
         MaskColor       =   &H8000000A&
         TabIndex        =   6
         Top             =   1140
         Width           =   255
      End
      Begin VB.Image imgComenzar 
         Height          =   255
         Left            =   2880
         MouseIcon       =   "frmRetos.frx":7897
         MousePointer    =   99  'Custom
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image imgCancelar 
         Height          =   255
         Left            =   3000
         MouseIcon       =   "frmRetos.frx":7BA1
         MousePointer    =   99  'Custom
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.PictureBox panel1v1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      Picture         =   "frmRetos.frx":7EAB
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4245
      Begin VB.CheckBox chkItems 
         BackColor       =   &H80000007&
         Height          =   195
         Left            =   990
         MaskColor       =   &H8000000A&
         TabIndex        =   4
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox txtOponente 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   870
         TabIndex        =   3
         Top             =   1080
         Width           =   3225
      End
      Begin VB.TextBox txtPuntos 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   240
         Left            =   860
         TabIndex        =   2
         Text            =   "0"
         Top             =   630
         Width           =   3210
      End
      Begin VB.TextBox txtOro 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Text            =   "0"
         Top             =   180
         Width           =   3210
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   2880
         MouseIcon       =   "frmRetos.frx":E4FA
         MousePointer    =   99  'Custom
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   3000
         MouseIcon       =   "frmRetos.frx":E804
         MousePointer    =   99  'Custom
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmRetos.frx":EB0E
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmRetos.frx":EE18
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image img1v1 
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmRetos.frx":F122
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim textval As String
Dim numval As String
Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Graficos\Retos.jpg")
End Sub

Private Sub Image2_Click()
    panel1v1.Visible = False
End Sub

Private Sub Image3_Click()
If Val(txtPuntos.Text) > 32767 Or Val(txtOro.Text) > 999999999 Then
    MsgBox "El maximo de oro para jugar es de 999.999.999 monedas, y el máximo de puntos es de 32.767", , "Error"
Else
    VaginaJugosa "/RETA1 " & txtOponente.Text & "@" & txtPuntos.Text & "@" & chkItems.Value & "@" & txtOro.Text & "@" & "" & "@" & ""
    Unload Me
End If
End Sub

Private Sub Image4_Click()
    picture2v2.Visible = True
End Sub

Private Sub Image5_Click()
    Unload Me
End Sub

Private Sub img1v1_Click()
    panel1v1.Visible = True
End Sub

Private Sub imgCancelar_Click()
    picture2v2.Visible = False
End Sub

Private Sub imgComenzar_Click()
If Val(txtPuntos2.Text) > 32767 Or Val(txtOro2.Text) > 999999999 Then
    MsgBox "El maximo de oro para jugar es de 999.999.999 monedas, y el máximo de puntos es de 32.767", , "Error"
Else
    VaginaJugosa "/RETA1 " & txtRival1.Text & "@" & txtPuntos2.Text & "@" & chkItems2.Value & "@" & txtOro2.Text & "@" & txtPareja.Text & "@" & txtRival2.Text
    Unload Me
End If
End Sub

Private Sub txtOro_Change()
  textval = txtOro.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtOro.Text = CStr(numval)
  End If
End Sub
Private Sub txtPuntos_Change()
  textval = txtPuntos.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtPuntos.Text = CStr(numval)
  End If
End Sub
Private Sub txtOro2_Change()
  textval = txtOro2.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtOro2.Text = CStr(numval)
  End If
End Sub
Private Sub txtPuntos2_Change()
  textval = txtPuntos2.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtPuntos2.Text = CStr(numval)
  End If
End Sub
