VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_CANTIDADDIALOGOS As Byte
Private Const p_SEGUNDOS = 10
Private Type t_GuildDlg
    Texto As String
    Segundos As Byte
End Type
Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

Public Property Let CantidadDialogos(ByVal v As Byte)
    If v > 0 Then
    ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
    End If
    p_CANTIDADDIALOGOS = v
End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v
End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo
End Property

Public Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw(ByRef C As cDialogos)
Dim i As Byte
Dim Y As Long

'If Not p_Activo Then Exit Sub

Y = 250
'500 600 700 800 900
    For i = 1 To p_CANTIDADDIALOGOS
        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            If Left$(p_Dialogos(i).Texto, 1) = Chr(3) Then
                Call C.DrawText(260, Y + (i * 10), mid(p_Dialogos(i).Texto, 2), vbGreen)
            Else

                Call C.DrawText2(260, Y + (i * 11), ReadField(1, p_Dialogos(i).Texto, Asc("+")), ReadField(2, p_Dialogos(i).Texto, Asc("+")))
            End If
        End If
    Next i
End Sub

Public Sub PassTimer()
Dim i As Byte

For i = 1 To p_CANTIDADDIALOGOS
    If p_Dialogos(i).Segundos > 0 Then
        p_Dialogos(i).Segundos = p_Dialogos(i).Segundos - 1
    Else
        p_Dialogos(i).Texto = vbNullString
    End If
Next i
    
End Sub

Public Sub PushBackText(ByVal s As String)
Dim i As Byte
Dim Vacio As Boolean
Dim Multiplicador As Integer
If Left$(s, 1) = Chr(3) Then
    Multiplicador = 7
Else
    Multiplicador = 1
End If

If p_CANTIDADDIALOGOS > 0 Then
    i = p_CANTIDADDIALOGOS
    Vacio = True
    While i > 0 And Vacio
        Vacio = p_Dialogos(i).Texto = vbNullString
        If Vacio Then i = i - 1
    Wend
    If i = p_CANTIDADDIALOGOS Then
        'hay q scrollear, estamos llenos
        i = 1
        While i < p_CANTIDADDIALOGOS
            p_Dialogos(i) = p_Dialogos(i + 1)
            i = i + 1
        Wend
        p_Dialogos(i).Texto = s
        p_Dialogos(i).Segundos = p_SEGUNDOS * Multiplicador
    Else
        p_Dialogos(i + 1).Texto = s
        p_Dialogos(i + 1).Segundos = p_SEGUNDOS * Multiplicador
    End If
End If

End Sub

