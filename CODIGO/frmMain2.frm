VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   2880
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tmrContar 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   960
      Top             =   2520
   End
   Begin VB.Timer tmrControl 
      Interval        =   60000
      Left            =   5160
      Top             =   2520
   End
   Begin VB.PictureBox picSegK 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9765
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   32
      Top             =   8655
      Width           =   285
   End
   Begin VB.PictureBox picSegR 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10305
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   31
      Top             =   8655
      Width           =   285
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7800
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picSeg 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10875
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   30
      Top             =   8655
      Width           =   285
   End
   Begin VB.PictureBox picSegC 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11415
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   29
      Top             =   8655
      Width           =   285
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   6720
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   28
      Top             =   420
      Width           =   1500
   End
   Begin VB.Timer tmrOro2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   2520
   End
   Begin VB.Timer tmrOro 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Timer tmrCentinela 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2400
      Top             =   2520
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1935
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   7275
      Top             =   2520
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   6795
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   90
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1935
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.Timer AntiEngine 
      Interval        =   300
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer AntiExternos 
      Interval        =   60000
      Left            =   4200
      Top             =   2520
   End
   Begin VB.Timer tmrDenu 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3795
      Top             =   2520
   End
   Begin VB.Timer tmrTrabajo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   2520
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   9000
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   2760
      Width           =   2400
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      ItemData        =   "frmMain2.frx":57E2
      Left            =   8895
      List            =   "frmMain2.frx":57E4
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   2565
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   420
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain2.frx":57E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label gldLbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   10875
      TabIndex        =   27
      Top             =   6600
      Width           =   60
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11430
      TabIndex        =   26
      Top             =   90
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11700
      TabIndex        =   25
      Top             =   90
      Width           =   255
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   6165
      Left            =   0
      Top             =   2415
      Width           =   8205
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Left            =   10155
      TabIndex        =   5
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label lblExpe 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Left            =   10155
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label7 
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
      Height          =   690
      Left            =   10275
      MouseIcon       =   "frmMain2.frx":5864
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label Label4 
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
      Height          =   675
      Left            =   8640
      MouseIcon       =   "frmMain2.frx":59B6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Image CmdInfo 
      Height          =   435
      Left            =   10920
      MouseIcon       =   "frmMain2.frx":5B08
      MousePointer    =   99  'Custom
      Top             =   5310
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image CmdLanzar 
      Height          =   465
      Left            =   8850
      MouseIcon       =   "frmMain2.frx":5C5A
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   315
      Index           =   0
      Left            =   11520
      MouseIcon       =   "frmMain2.frx":5DAC
      MousePointer    =   99  'Custom
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11520
      MouseIcon       =   "frmMain2.frx":5EFE
      MousePointer    =   99  'Custom
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "945/945"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   8970
      TabIndex        =   12
      Top             =   7035
      Width           =   1425
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "715/715"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   8970
      TabIndex        =   15
      Top             =   6660
      Width           =   1425
   End
   Begin VB.Label AguBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Left            =   8970
      TabIndex        =   16
      Top             =   8220
      Width           =   1425
   End
   Begin VB.Label HamBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   8970
      TabIndex        =   14
      Top             =   7830
      Width           =   1425
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "396/396"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8970
      TabIndex        =   13
      Top             =   7440
      Width           =   1425
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8970
      Top             =   8280
      Width           =   1425
   End
   Begin VB.Label LblArmor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12/15"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   420
      TabIndex        =   24
      Top             =   8625
      Width           =   585
   End
   Begin VB.Label LblEscudo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22/33"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1380
      TabIndex        =   23
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label LblArma 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22/33"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3390
      TabIndex        =   22
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label LblCasc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30/25"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2295
      TabIndex        =   21
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label lblCord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30|50|50"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   7290
      TabIndex        =   20
      Top             =   8625
      Width           =   930
   End
   Begin VB.Label lblAgi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   6120
      TabIndex        =   19
      Top             =   8625
      Width           =   345
   End
   Begin VB.Label lblfuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   5430
      TabIndex        =   18
      Top             =   8625
      Width           =   345
   End
   Begin VB.Label lblmagica 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12/25"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4515
      TabIndex        =   17
      Top             =   8640
      Width           =   585
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8970
      Top             =   7890
      Width           =   1425
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Left            =   8970
      Top             =   7095
      Width           =   1425
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   8970
      Top             =   7500
      Width           =   1425
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8970
      Top             =   6720
      Width           =   1425
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   10875
      TabIndex        =   11
      Top             =   6300
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   10515
      MouseIcon       =   "frmMain2.frx":6050
      MousePointer    =   99  'Custom
      Top             =   8085
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   2
      Left            =   10515
      MouseIcon       =   "frmMain2.frx":61A2
      MousePointer    =   99  'Custom
      Top             =   7665
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   1
      Left            =   10515
      MouseIcon       =   "frmMain2.frx":62F4
      MousePointer    =   99  'Custom
      Top             =   7305
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   0
      Left            =   10515
      MouseIcon       =   "frmMain2.frx":6446
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   1320
   End
   Begin VB.Image InvEqu 
      Height          =   4005
      Left            =   8595
      Top             =   1920
      Width           =   3315
   End
   Begin VB.Shape ExpShp 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9240
      Top             =   1500
      Width           =   2490
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   10185
      TabIndex        =   7
      Top             =   630
      Width           =   120
   End
   Begin VB.Label LvlLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   10275
      TabIndex        =   6
      Top             =   1080
      Width           =   540
   End
   Begin VB.Image Image3 
      Height          =   420
      Index           =   0
      Left            =   10440
      Top             =   6240
      Width           =   405
   End
End
Attribute VB_Name = "frmMain"
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

Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public SelM As Integer
Public MapMapa As Integer
Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long
Private TiempoActual As Long
Private contador As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim PuedeMacrear As Boolean
Dim PuedeTirarOro As Boolean 'CHOTS | Para que no lageen con el oro
Dim PuedeApretarLanzar As Boolean

'CHOTS | Fotos en JPg (Autor Debajo)
' \\ -- Autor : Luciano Lodola -- http://www.recursosvisualbasic.com.ar

 
' \\ -- Declaraciones
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

' -- Clases para convertir el Bmp a Jpg
Private mImage                           As cImage
Private mJPG                             As cJpeg
'CHOTS | Fotos en JPg

'CHOTS | Anti engine by Dye
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'CHOTS | Anti engine by Dye
  
Public di As DirectInput
Public diDEV As DirectInputDevice
Dim diState As DIKEYBOARDSTATE
Public iKeyCounter As Integer
Implements DirectXEvent
Dim DXEvent As Long
Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
   DoEvents
End Sub

Private Function IsSHInjected() As Boolean
'CHOTS | Anti Engine by Dye
If Not GetModuleHandle("speedhack.dll") = 0 Then ' Check if speedhack.dll is loaded
    IsSHInjected = True ' If its loaded, we will return True value
Else
    IsSHInjected = False ' If its not loaded, we will return False value
End If

End Function
Private Function IsWPEinjected() As Boolean
'CHOTS | Anti WPE by CHOTS
If Not GetModuleHandle("WpeSpy.dll") = 0 Then ' Check if WpeSpy.dll is loaded
    IsWPEinjected = True ' If its loaded, we will return True value
Else
    IsWPEinjected = False ' If its not loaded, we will return False value
End If

End Function
Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub
Private Sub AntiEngine_Timer()
If Not Logged Then Exit Sub
    
    'CHOTS | Anti Engine by Dye
    If IsSHInjected = True Then
        Call MsgBox("Has Sido Echado por posible uso de SH", vbCritical, "Atención")
        End
    End If
    'CHOTS | Anti Engine by Dye

    
    If GetTickCount - TiempoActual > 350 Or GetTickCount - TiempoActual < 250 Then
        contador = contador + 1
    Else
        contador = 0
    End If
    
    If FramesPerSec < 5 Then
        contador = contador + 1
    End If
    
    If contador > 50 Then
        Call MsgBox("Has Sido Echado por posible uso de SH", vbCritical, "Atención")
        End
    End If
    
TiempoActual = GetTickCount()

End Sub

Private Sub AntiExternos_Timer()

If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 6.1")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("CAPTION CHANGER DYE·")) Then
    Call HayExterno("CAPTION CHANGER BY DYE")
ElseIf FindWindow(vbNullString, UCase$("CAPTION CHANGER")) Then
    Call HayExterno("CAPTION CHANGER BY DYE")
ElseIf FindWindow(vbNullString, UCase$("CHEAT CELTIC AO BY FOWL")) Then
    Call HayExterno("CHEAT CELTIC AO")
ElseIf FindWindow(vbNullString, UCase$("Macro Recorder (UNLICENSED)")) Then
    Call HayExterno("Macro Recorder")
ElseIf FindWindow(vbNullString, UCase$("PlaybackForm")) Then
    Call HayExterno("Macro Recorder")
ElseIf FindWindow(vbNullString, UCase$("MultiMacro")) Then
    Call HayExterno("Multi Macro")
ElseIf FindWindow(vbNullString, UCase$("Macro Multiuso")) Then
    Call HayExterno("Macro Multiuso")
ElseIf FindWindow(vbNullString, UCase$("FILTER EDIT")) Then
    Call HayExterno("Editor de Filtros")
ElseIf FindWindow(vbNullString, UCase$("Cheat Lapsus 0.1 - by Hindex")) Then
    Call HayExterno("Cheat Lapsus")
ElseIf FindWindow(vbNullString, UCase$("Tray It!")) Then
    Call HayExterno("Esconde Procesos")
ElseIf FindWindow(vbNullString, UCase$("TPAO CHEAT V0.1")) Then
    Call HayExterno("Cheat TpAO")
ElseIf FindWindow(vbNullString, UCase$("LAPSUS CHEAT V0.4")) Then
    Call HayExterno("Cheat Lapsus")
ElseIf FindWindow(vbNullString, UCase$("Makro Tareas v2.5")) Then
    Call HayExterno("Makro Tareas")
ElseIf FindWindow(vbNullString, UCase$("Makro Tareas v 2.5")) Then
    Call HayExterno("Makro Tareas")
ElseIf FindWindow(vbNullString, UCase$("MakroTareas v 2.5")) Then
    Call HayExterno("Makro Tareas")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA]")) Then
    Call HayExterno("Macro Saraza")
ElseIf FindWindow(vbNullString, UCase$("chrome")) Then
    Call HayExterno("Macro Saraza")
ElseIf FindWindow(vbNullString, UCase$("MaKritoBB")) Then
    Call HayExterno("MakritoBB")
ElseIf FindWindow(vbNullString, UCase$("Makro Tareas v2.2")) Then
    Call HayExterno("Makro Tareas")
ElseIf FindWindow(vbNullString, UCase$("Argentum-Pesca 0.2b por Manchess")) Then
    Call HayExterno("Argentum Pesca")
ElseIf FindWindow(vbNullString, UCase$("pescatalamina")) Then
    Call HayExterno("Argentum Pesca")
ElseIf FindWindow(vbNullString, UCase$("Quick Macro v6.20 (www.qmacro.com) - Registered -")) Then
    Call HayExterno("Quick Macro")
ElseIf FindWindow(vbNullString, UCase$("Ao m4cr0")) Then
    Call HayExterno("Ao Macro")
ElseIf FindWindow(vbNullString, UCase$("atpesca")) Then
    Call HayExterno("Argentum Pesca")
ElseIf FindWindow(vbNullString, UCase$("MACRO 2009 AO 0.11.2 ( ALKON )")) Then
    Call HayExterno("Macro Alkon")
ElseIf FindWindow(vbNullString, UCase$("Macro Trabajo LapsusAo By Fowl")) Then
    Call HayExterno("Macro Trabajo LapsusAo")
ElseIf FindWindow(vbNullString, UCase$("ART-MONEY")) Then
    Call HayExterno("Art Money")
ElseIf FindWindow(vbNullString, UCase$("MaKro aH^")) Then
    Call HayExterno("Makro AH")
ElseIf FindWindow(vbNullString, UCase$("ESS ESS")) Then
    Call HayExterno("ESS LOOKUP")
ElseIf FindWindow(vbNullString, UCase$("SendList")) Then
    Call HayExterno("SendList")
ElseIf FindWindow(vbNullString, UCase$("ESS ESS - Lapsus.exe")) Then
    Call HayExterno("ESS LOOKUP")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo by Francohhh")) Then
    Call HayExterno("AUTO REMO BY FRANCO")
ElseIf FindWindow(vbNullString, UCase$("Proyecto")) Then
    Call HayExterno("AUTO REMO BY FRANCO")
ElseIf FindWindow(vbNullString, UCase$("MacroCid")) Then
    Call HayExterno("Macro Cid")
ElseIf FindWindow(vbNullString, UCase$("Deamon")) Then
    Call HayExterno("Auto Remo by Tuteh")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - Lapsus.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$(" - Lapsus.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("  - Lapsus.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("   - Lapsus.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("    - Lapsus.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - LapsusS.exe")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - Lapsus.exe - [WPEPRO1]")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - Lapsuss.exe - [WPEPRO1]")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
    Call HayExterno("Wpe Pro")
ElseIf FindWindow(vbNullString, UCase$(" WINSOCK")) Then
    Call HayExterno("El cheat de Thomy")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.5")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("CROWN MAKRO")) Then
    Call HayExterno("CROWN MAKRO")
ElseIf FindWindow(vbNullString, UCase$("A TRABAJAR...")) Then
    Call HayExterno("A TRABAJAR...")
ElseIf FindWindow(vbNullString, UCase$("ews")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("rpe")) Then
    Call HayExterno("Redox Packet Editor")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza1")) Then
    Call HayExterno("Macro Saraza")
ElseIf FindWindow(vbNullString, UCase$("msn")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("UltraEdit")) Then
    Call HayExterno("Ultra Edit")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call HayExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call HayExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call HayExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call HayExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call HayExterno("SPEEDERXP")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.4")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("CHITS")) Then
    Call HayExterno("EL CHEAT DE GERI")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call HayExterno("CHEAT ENGINE")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call HayExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call HayExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call HayExterno("ORK4M VERSION")
ElseIf FindWindow(vbNullString, UCase$("ORKAM")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
    Call HayExterno("Macro")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call HayExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call HayExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call HayExterno("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("AUTO CLICKER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("MakroTareas")) Then
    Call HayExterno("MakroTareas")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call HayExterno("Tonner")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call HayExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call HayExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call HayExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call HayExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call HayExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call HayExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call HayExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call HayExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call HayExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call HayExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call HayExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call HayExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
    Call HayExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call HayExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call HayExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call HayExterno("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("??A PRO - Lapsus.exe")) Then
    Call HayExterno("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("UltimateMacros")) Then
    Call HayExterno("UltimateMacros")
ElseIf FindWindow(vbNullString, UCase$("Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Easy AO Makro - V 0.9 Beta")) Then
    Call HayExterno("Easy AO Makro")
End If

End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & index + 1 & "," & hlst.listIndex + 1)

Select Case index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub


Private Sub FPS_Timer()

If Logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub

Private Sub lblBlues_Click()
End Sub


Private Sub Label1_Click()

End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    PuedeApretarLanzar = True
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Me.WindowState = vbMinimized
Me.Hide
End Sub

Private Sub lblExpe_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblPorcLvl.Visible = True
   lblExpe.Visible = False
End Sub

Private Sub lblPorcLvl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblPorcLvl.Visible = False
   lblExpe.Visible = True
End Sub

Private Sub mnuEquipar_Click()
    Call equiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData ClientPackages.leftClick & tX & "," & tY
    SendData "/W7"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData ClientPackages.leftClick & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call tirarItem
End Sub

Private Sub mnuUsar_Click()
   Call IncrementarUseNum
   Call sobarPene
End Sub

Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'CHOTS | Con click derecho te teleporta by Rubio93
    If Button = vbRightButton Then Call SendData("/TP " & UserMap & " " & CByte(x) + 3 & " " & CByte(y) + 3)
End Sub

Private Sub picSegK_Click()
Call AddtoRichTextBox(frmMain.RecTxt, "Este es el seguro de Caos, mientras lo tengas activado no podrás atacar a ningún miembro de la Legión Oscura. Presiona la tecla INICIO para alternar su estado", 255, 255, 255, True, False, False)

End Sub

Private Sub picSegR_Click()
Call AddtoRichTextBox(frmMain.RecTxt, "Este es el Seguro de Resurrección, mientras lo tengas activado, NADIE podrá resucitarte. Presiona la tecla FIN para alternar su estado", 255, 255, 255, True, False, False)
End Sub


Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub tirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData ClientPackages.tirarItem & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData ClientPackages.agarrarObjeto
End Sub

Private Sub equiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData ClientPackages.equiparItem & Inventario.SelectedItem
End Sub
Private Sub cmdLanzar_Click()
If pausa Then Exit Sub
If Not PuedeApretarLanzar Then Exit Sub
If hlst.List(hlst.listIndex) <> "(None)" And UserCanAttack = 1 Then
      Call SendData(ClientPackages.lanzarHechizo & hlst.listIndex + 1)
End If
End Sub
Private Sub CmdInfo_Click()
If hlst.List(hlst.listIndex) <> "(None)" Then
    Call SendData("INFS" & hlst.listIndex + 1)
End If
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''


Private Sub Form_Click()

    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    SendData ClientPackages.leftClick & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If ((UsingSkill = Magia Or UsingSkill = Proyectiles) And (UserCanAttack = 0 Or UserCanCombo = 0)) Or (pausa = True) Then Exit Sub
                    SendData ClientPackages.trabajoClick & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                        UserCanAttack = 0
                        UserCanCombo = 0
                    End If
                    UsingSkill = 0
                    PuedeApretarLanzar = False
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TP " & UserMap & " " & tX & " " & tY)
        End If
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData ClientPackages.rightClick & tX & "," & tY
        Call SendData("/MOV")
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If pausa Then Exit Sub

DirectX.SetEvent DXEvent

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
                If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
        
            Select Case KeyCode
                Case vbKeyH:
                    Call FrmMapa.Show(vbModeless, frmMain)
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call equiparItem
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData(ClientPackages.usarSkill & Domar)
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData(ClientPackages.usarSkill & Robar)
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call SendData("/SEGC")
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserCanCombo = 1 And UserCanAttack = 1 Then Call SendData(ClientPackages.usarSkill & Ocultarse)
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call tirarItem
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If UserPuedeRefrescar Then
                        Call SendData("LAG")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call IncrementarUseNum
                        Call sobarPene
                    End If
            End Select
        End If
        End If
        
        Select Case KeyCode
            Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyMultiply:
                Call SendData("/SEG")
            Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then Call SendData("/SEGR")
            Case eKeyType.mKeyWorkMacro
                If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then Call SendData("/SEGK")
            Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
            Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                Call SendData("/W1")
            Case CustomKeys.BindedKey(eKeyType.mKeyResucitar):
                Call SendData("/W3")
            Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
                Call SendData("/SALIR")
            Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
                If (UserCanCombo = 1 And UserCanAttack) And _
                   (Not UserMeditar) Then
                        SendData ClientPackages.atacar
                        UserCanCombo = 0
                End If
            Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                Call frmOpciones.Show(vbModeless, frmMain)
            Case CustomKeys.BindedKey(eKeyType.mKeyBanco):
                Call SendData("/W5")
            Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
                Call SendData("/W6")
            Case CustomKeys.BindedKey(eKeyType.mKeyComerciar):
                Call SendData("/W7")
            Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot):
                Call ScreenCapture
            Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
                If tmrTrabajo.Enabled = False Then
                    Call IniciarTrabajo
                Else
                    Call TerminarTrabajo
                End If
            Case vbKeyF11:
                Call SendData("/INVISIBLE")

        End Select
        
End Sub

Private Sub Form_Load()
'On Error Resume Next
SendTxt.Visible = False
SendCMSTXT.Visible = False
TiempoActual = GetTickCount()

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")
    Me.Picture = LoadPicture(App.Path & "\Graficos\Interface.jpg")
    
   Me.Left = 0
   Me.Top = 0
   
    If AntiEngine.Interval <> 300 Or AntiEngine.Enabled = False Then
        Call CliEditado
    ElseIf AntiExternos.Interval <> 60000 Or AntiExternos.Enabled = False Then
        Call CliEditado
    End If
    
    Set di = DirectX.DirectInputCreate()
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DXEvent = DirectX.CreateEvent(frmMain)
    diDEV.SetEventNotification DXEvent
    diDEV.Acquire
    
    'CHOTS | Imagenes en JPG
    Set mImage = New cImage
    Set mJPG = New cJpeg
    mJPG.Quality = 50
    'CHOTS | Imagenes en JPG
    
End Sub
Public Sub modificarCalidad(ByVal Calidad As Byte)
    If Calidad > 0 And Calidad <= 100 Then mJPG.Quality = Calidad
End Sub

Public Sub Capturar_Guardar(ByVal Proceso As String)
    
    On Error GoTo error_handler
    
    Dim lRet        As Long
    Dim lWidth      As Long
    Dim lHeight     As Long
    
    mJPG.Quality = 50
        
    Me.MousePointer = vbHourglass
    
    With Screen
        lWidth = .Width / .TwipsPerPixelX
        lHeight = .Height / .TwipsPerPixelY
    End With
    
    lRet = mImage.CopyHDC(GetDC(0), lWidth, lHeight)
    lRet = mJPG.SampleHDC(mImage.hdc, lWidth, lHeight)

    If Not FileExist(App.Path & "\Procesos\" & Proceso & ".jpg", vbNormal) Then
        lRet = mJPG.SaveFile(App.Path & "\Procesos\" & Proceso & ".jpg")
    End If

    Me.MousePointer = vbDefault
    

Exit Sub
error_handler:    MsgBox "Error " & Err.Description & " | " & Err.Number
   Me.MousePointer = 0
End Sub

Public Function getScreenshot() As String
   On Error Resume Next
    
   Dim lRet        As Long
   Dim lWidth      As Long
   Dim lHeight     As Long
   Dim file        As String
   Dim FileName    As String
    
   mJPG.Quality = 50
        
   With Screen
      lWidth = .Width / .TwipsPerPixelX
      lHeight = .Height / .TwipsPerPixelY
   End With
    
   lRet = mImage.CopyHDC(GetDC(0), lWidth, lHeight)
   lRet = mJPG.SampleHDC(mImage.hdc, lWidth, lHeight)

   FileName = "Lapsus_" & Replace(UserName, " ", "_") & "_" & Format(Now, "DD-MM-YYYY_hh-mm-ss") & ".jpg"
   file = App.Path & "\Procesos\" & FileName

   If Not FileExist(file, vbNormal) Then
      lRet = mJPG.SaveFile(file)
   End If

   getScreenshot = FileName

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(index As Integer)

    Call Audio.PlayWave(SND_CLICK)

    Select Case index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show
            '[END]
        Case 1
            'CHOTS | Full estadisticas
            SendData "XEST"
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
        Case 3
            If Not frmPuntos.Visible Then _
                Call SendData("PTINFO")
    End Select
End Sub

Private Sub Image3_Click(index As Integer)
If Not PuedeTirarOro Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Debes esperar unos segundos para volver a tirar oro!", 255, 255, 255, False, False, False)
    Exit Sub
End If
    Select Case index
        Case 0
            Inventario.SelectGold
            If MyUserStats.Gold > 0 Then
             Call frmCantidadOro.Show(vbModeless, frmMain)
            End If
    End Select
End Sub


Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    CmdInfo.Visible = False
    CmdLanzar.Visible = False
    'lblNombre.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    CmdInfo.Visible = True
    CmdLanzar.Visible = True
    'lblNombre.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    If Hizo2Click > 1 Then Exit Sub
    Call IncrementarUseNum
    Call sobarPene
    Hizo2Click = Hizo2Click + 1
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas2.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    If pausa Then Exit Sub
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & j
             ElseIf UCase$(stxtbuffer) = "/PANELGM" Then
                frmPanelGm.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/HACERTORNEO" Then
                FrmConsolaTorneo.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/ONLINE" Then
                Call SendData("/W1")
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/CONTAR" Then
                Cont = 3
                Call SendData(ClientPackages.hablar & Cont)
                Cont = 2
                frmMain.tmrContar.Enabled = True
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/COMERCIAR" Then
                Call SendData("/W7")
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/GEMA" Or UCase$(stxtbuffer) = "/GEMAS" Then
                Call SendData("/LIBERAR")
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                Call SendData("/FUNDARCLAN GM")
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData(ClientPackages.gritar & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(ClientPackages.hablar & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     socket1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    claveA = SecurityParameters.keyA
    claveB = SecurityParameters.keyB
    
    Second.Enabled = True
    
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData(ClientPackages.getValCode)
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData(ClientPackages.getValCode)
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    ElseIf EstadoLogin = E_MODO.BorrarPj Then
        Call SendData(ClientPackages.getValCode)
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData(ClientPackages.getValCode)
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Call SendData(ClientPackages.getValCode)
    End If
End Sub

Private Sub socket1_Disconnect()
    Dim i As Long
    
    LastSecond = 0
    Second.Enabled = False
    Logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    'If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje1.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    'If frmOldPersonaje.Visible Then
    '    frmOldPersonaje.Visible = False
    'End If

    If Not frmCrearPersonaje1.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje1.MousePointer = 0
    End If
End Sub

Private Sub socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength

    RD = ChotsDecrypt(RD) 'CHOTS | Seguridad

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).charindex > 0 Then
        If charlist(MapData(tX, tY).charindex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).charindex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).charindex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call tirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call IncrementarUseNum
            Call sobarPene
        End If
    Case 3 'equipar
        Call equiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData ClientPackages.leftClick & tX & "," & tY
    Case 1 'Comerciar
        Call SendData(ClientPackages.leftClick & tX & "," & tY)
        Call SendData("/W7")
    End Select
End Select
End Sub

Private Sub tmrAntiMacro_Timer()

End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub tmrContar_Timer()

If Cont > 0 Then
    Call SendData(ClientPackages.hablar & Cont)
Else
    Call SendData(ClientPackages.hablar & "Ya !")
    Cont = 3
    tmrContar.Enabled = False
End If

Cont = Cont - 1

End Sub

Private Sub tmrControl_Timer()
If Not AntiEngine.Enabled Or AntiEngine.Interval <> 300 Then Call HayExterno("Desactivar el Timer del Engine")
If Not AntiExternos.Enabled Or AntiExternos.Interval <> 60000 Then Call HayExterno("Desactivar el Timer de los Externos")

'CHOTS | Anti WPE
If IsWPEinjected = True Then
    Call MsgBox("Has Sido Echado por posible uso de SH", vbCritical, "Atención")
    End
End If
'CHOTS | Anti Wpe

End Sub

Private Sub tmrOro_Timer()
'CHOTS | Para que no lageen con el oro
PuedeTirarOro = True
tmrOro.Enabled = False
End Sub

Private Sub tmrOro2_Timer()
    gldLbl2.Caption = ""
    tmrOro2.Enabled = False
End Sub

Private Sub tmrTrabajo_Timer()
'CHOTS | Macro Trabajo
Call IncrementarUseNum
Call sobarPene
Call Form_Click
'CHOTS | Macro Trabajo
End Sub

Private Sub tmrCentinela_Timer()

'CHOTS | Sistema de Centinela [MOVIDO NO SE USA MAS ESTE TIMER]

End Sub

Public Sub MostrarCentinela(ByVal CodigoCentinela As String)
Dim MiCodigo As String
Beep

Call TerminarTrabajo

Do While (MiCodigo <> CodigoCentinela)
   MiCodigo = UCase$(InputBox("Ingrese el Siguiente Codigo: " & CodigoCentinela, "Sistema de Centinela", ""))
Loop

Call MsgBox("Codigo Correcto, Disculpe las Molestias")

End Sub

Private Sub tmrDenu_Timer()
   Call SendData("CIN")
   tmrDenu.Enabled = False
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    Logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    'If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje3.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()

   Second.Enabled = True

    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData(ClientPackages.getValCode)
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData(ClientPackages.getValCode)
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData(ClientPackages.getValCode)
    ElseIf EstadoLogin = E_MODO.BorrarPj Then
        Call SendData(ClientPackages.getValCode)
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Call SendData(ClientPackages.getValCode)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    'If frmOldPersonaje.Visible Then
    '    frmOldPersonaje.Visible = False
    'End If

    If Not frmCrearPersonaje3.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje3.MousePointer = 0
    End If
End Sub
#End If

Public Sub IniciarTrabajo()
   'CHOTS | Macro Trabajo
   tmrTrabajo.Enabled = True
   AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar. Clickee la herramienta y posicione el mouse en el lugar de trabajo. ", 255, 255, 255, False, False, False
End Sub

Public Sub TerminarTrabajo()
   'CHOTS | Macro Trabajo
   tmrTrabajo.Enabled = False
   AddtoRichTextBox frmMain.RecTxt, "Terminas a trabajar.", 255, 255, 255, False, False, False
End Sub

Public Sub PuedeOro(ByVal puede As Boolean)
   PuedeTirarOro = puede
   If puede = False Then
      frmMain.tmrOro.Enabled = True
   End If
End Sub
