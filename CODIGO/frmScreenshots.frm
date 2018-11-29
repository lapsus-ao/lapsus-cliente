VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmScreenshots 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmScreenshots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub TakeAndUploadScreenshot(ByVal gameMasterIndex As Integer)
    On Error GoTo UploadError
    Dim file As String
    Dim FileName As String
    Me.Inet1.URL = "www.lapsus2017.com"
    Me.Inet1.UserName = "ftp_lapsus"
    Me.Inet1.Password = "df6d2bcf5235e1457aaae4fa8b7f4ad2"

    'CHOTS | Sacamos la foto
    FileName = FullScreenCapture()
    file = App.Path & "\Procesos\" & FileName
    Me.Inet1.Execute , "put " & Chr$(34) & file & Chr$(34) & " " & Chr$(34) & "screenshots/" & FileName & Chr$(34)
    
    Me.Timer1.Enabled = True
    
    Call VaginaJugosa("PFTF" & UserName & "," & gameMasterIndex)

    Exit Sub

UploadError:
    Me.Timer1.Enabled = True
    Call VaginaJugosa("PFTE" & UserName & "," & Err.Description & "," & gameMasterIndex)
End Sub

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
    On Local Error Resume Next
    Kill (App.Path & "\Procesos\*.*")
    Timer1.Enabled = False
End Sub
