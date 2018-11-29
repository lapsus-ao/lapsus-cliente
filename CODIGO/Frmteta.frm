VERSION 5.00
Begin VB.Form frmEstadisticas2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   24
      Left            =   5745
      TabIndex        =   40
      Top             =   6650
      Width           =   135
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   23
      Left            =   5790
      TabIndex        =   39
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   22
      Left            =   5750
      TabIndex        =   38
      Top             =   6380
      Width           =   135
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   21
      Left            =   6075
      TabIndex        =   37
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   6
      Left            =   1420
      TabIndex        =   36
      Top             =   5620
      Width           =   1065
   End
   Begin VB.Label label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   35
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   47
      Left            =   5400
      Top             =   6600
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   46
      Left            =   6000
      Top             =   6720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   45
      Left            =   5400
      Top             =   6120
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   44
      Left            =   6000
      Top             =   6120
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   43
      Left            =   5280
      Top             =   6360
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   42
      Left            =   6000
      Top             =   6360
      Width           =   345
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2450
      TabIndex        =   34
      Top             =   6900
      Width           =   120
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Index           =   1
      Left            =   2535
      TabIndex        =   33
      Top             =   6240
      Width           =   120
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   8160
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   5
      Left            =   1080
      TabIndex        =   32
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   31
      Top             =   4920
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   3
      Left            =   1700
      TabIndex        =   30
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   1
      Left            =   1600
      TabIndex        =   28
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   0
      Left            =   2080
      TabIndex        =   27
      Top             =   3490
      Width           =   1065
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   26
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   1400
      TabIndex        =   25
      Top             =   1350
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   24
      Top             =   1770
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   23
      Top             =   2450
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   2150
      TabIndex        =   22
      Top             =   2100
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   41
      Left            =   5760
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   40
      Left            =   6240
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   21
      Top             =   6555
      Width           =   135
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   7800
      TabIndex        =   20
      Top             =   240
      Width           =   120
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   39
      Left            =   5520
      Top             =   4200
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   38
      Left            =   6120
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   20
      Left            =   5900
      TabIndex        =   19
      Top             =   4220
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   37
      Left            =   6600
      Top             =   3960
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   36
      Left            =   7200
      Top             =   3960
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   19
      Left            =   7000
      TabIndex        =   18
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   18
      Left            =   6600
      TabIndex        =   17
      Top             =   4515
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   35
      Left            =   6240
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   34
      Left            =   6840
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   17
      Left            =   5880
      TabIndex        =   16
      Top             =   3650
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   315
      Index           =   33
      Left            =   5520
      Top             =   3600
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   315
      Index           =   32
      Left            =   6000
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   16
      Left            =   5640
      TabIndex        =   15
      Top             =   5350
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   31
      Left            =   5280
      Top             =   5400
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   30
      Left            =   5880
      Top             =   5400
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   15
      Left            =   6015
      TabIndex        =   14
      Top             =   5880
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   29
      Left            =   5640
      Top             =   5880
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   28
      Left            =   6240
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   14
      Left            =   5640
      TabIndex        =   13
      Top             =   5085
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   27
      Left            =   5280
      Top             =   5040
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   26
      Left            =   5760
      Top             =   5040
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   13
      Left            =   5480
      TabIndex        =   12
      Top             =   6915
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   25
      Left            =   5160
      Top             =   6960
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   24
      Left            =   5640
      Top             =   6960
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   6600
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   22
      Left            =   7200
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   12
      Left            =   6960
      TabIndex        =   11
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   5520
      Top             =   3120
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   20
      Left            =   6120
      Top             =   3120
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   11
      Left            =   5910
      TabIndex        =   10
      Top             =   3120
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   5880
      Top             =   5640
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   18
      Left            =   6480
      Top             =   5640
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   10
      Left            =   6220
      TabIndex        =   9
      Top             =   5640
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   17
      Left            =   5880
      Top             =   2880
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   345
      Index           =   16
      Left            =   6495
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   9
      Left            =   6300
      TabIndex        =   8
      Top             =   2850
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   5400
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   14
      Left            =   6090
      Top             =   2625
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   8
      Left            =   5850
      TabIndex        =   7
      Top             =   2595
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   13
      Left            =   5400
      Top             =   2310
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   12
      Left            =   6000
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   7
      Left            =   5790
      TabIndex        =   6
      Top             =   2295
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   11
      Left            =   5280
      Top             =   2040
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   10
      Left            =   5760
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   6
      Left            =   5600
      TabIndex        =   5
      Top             =   1995
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   6480
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   8
      Left            =   7080
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   5
      Left            =   6860
      TabIndex        =   4
      Top             =   1770
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   7
      Left            =   6480
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   6
      Left            =   7080
      Top             =   1560
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   4
      Left            =   6880
      TabIndex        =   3
      Top             =   1500
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   5
      Left            =   5160
      Top             =   1200
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   4
      Left            =   5760
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   3
      Left            =   5550
      TabIndex        =   2
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   5160
      Top             =   960
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   2
      Left            =   5760
      Top             =   960
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   2
      Left            =   5550
      TabIndex        =   1
      Top             =   960
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   5160
      Top             =   720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   0
      Left            =   5760
      Top             =   720
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   5550
      TabIndex        =   0
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmEstadisticas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.Picture = LoadPicture(App.Path & "\Graficos\Estadisticas.jpg")
   ReDim flags(1 To NUMSKILLS)
End Sub


Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = UserAtributos(i)
Next
For i = 1 To NUMSKILLS
    Text1(i).Caption = UserSkills(i)
Next


Label4(0).Caption = UserReputacion.AsesinoRep
Label4(1).Caption = UserReputacion.BandidoRep
Label4(2).Caption = UserReputacion.BurguesRep
Label4(3).Caption = UserReputacion.LadronesRep
Label4(4).Caption = UserReputacion.NobleRep
Label4(5).Caption = UserReputacion.PlebeRep

If UserReputacion.Promedio < 0 Then
    Label4(6).ForeColor = vbRed
    Label4(6).Caption = "CRIMINAL"
Else
    Label4(6).ForeColor = vbBlue
    Label4(6).Caption = "CIUDADANO"
End If

With UserEstadisticas
    label6(0).Caption = .CriminalesMatados
    label6(1).Caption = .CiudadanosMatados
    label6(2).Caption = .UsuariosMatados
    'label6(3).Caption = .NpcsMatados
    label6(4).Caption = .Clase
    'Label6(5).Caption = "Tiempo restante en carcel: " & .PenaCarcel
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Command1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If Alocados > 0 Then
        indice = index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
            Text1(indice).Caption = Val(Text1(indice).Caption) + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = index \ 2 + 1
        If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) - 1
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If

puntos.Caption = Alocados
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()

    Dim i As Integer
    Dim cad As String
    Dim sumaCHOTS As Byte
    sumaCHOTS = 0
    
    For i = 1 To NUMSKILLS
        cad = cad & flags(i) & ","
        sumaCHOTS = sumaCHOTS + Val(flags(i))
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    If sumaCHOTS > 0 Then VaginaJugosa "SKSE" & cad
    SkillPoints = Alocados
    Unload Me

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Dim i As Integer
    Dim cad As String
    Dim sumaCHOTS As Byte
    sumaCHOTS = 0
    
    For i = 1 To NUMSKILLS
        cad = cad & flags(i) & ","
        sumaCHOTS = sumaCHOTS + Val(flags(i))
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    If sumaCHOTS > 0 Then VaginaJugosa "SKSE" & cad
    SkillPoints = Alocados
    Unload Me
End If
End Sub
