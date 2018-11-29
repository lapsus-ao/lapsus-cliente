Attribute VB_Name = "Anti_Cheat"
Option Explicit

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Function HayExterno(ByVal Chit As String, Optional echa As Boolean = True)
    Call SendData("BANEAME" & Chit)
    
    If echa Then
        Call MsgBox("Programa externo detectado. Argentum Online se cerrará. Tu Nombre ha quedado en los Logs.")
        End
    End If
    
End Function
Public Function CliEditado()
    Call MsgBox("No se admite editar el cliente en este servidor")
    End
End Function

Public Function KiloBytes(ByVal Bytes As Long) As String
Dim Tamanio As Double
Tamanio = Bytes / 1024

If Tamanio < 1024 Then
    KiloBytes = Round(Tamanio, 2) & " Kb"
    Exit Function
Else
    If (Tamanio / 1024) < 1024 Then
        KiloBytes = Round((Tamanio / 1024), 2) & " Mb"
        Exit Function
    End If

End If
End Function
