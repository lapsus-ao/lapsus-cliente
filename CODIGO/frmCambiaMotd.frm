VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """ZMOTD"""
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkItalic 
      Caption         =   "Cursiva"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Negrita"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdMarron 
      Caption         =   "Marron"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdVerde 
      Caption         =   "Verde"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdMorado 
      Caption         =   "Morado"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAmarillo 
      Caption         =   "Amarillo"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdGris 
      Caption         =   "Gris"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdBlanco 
      Caption         =   "Blanco"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdRojo 
      Caption         =   "Rojo"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAzul 
      BackColor       =   &H00FF0000&
      Caption         =   "Azul"
      Height          =   375
      Left            =   600
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox txtMotd 
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No olvides agregar los colores al final de cada línea (ver tabla de abajo)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmCambiaMotd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
