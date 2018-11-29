Attribute VB_Name = "Dye"
'MÓDULO PROGRAMADO POR JOSÉ IGNACIO PARODI (DYE)
'PARA TWISTEROS AO 2010
'REPROGRAMADO Y ADAPTADO POR CHOTS
'PARA LAPSUS AO 2.1
'24/11/2010
Option Explicit
Public claveA As String
Public claveB As String

Public Function DyeCifro(ByVal datos As String) As String
Dim Buffer() As Byte
Dim OutBuffer() As Byte

Dim i As Long

Buffer = StrConv(datos, vbFromUnicode)
ReDim OutBuffer(Len(datos) - 1) As Byte
OutBuffer(0) = Buffer(0) Xor claveA
For i = 1 To (Len(datos) - 1)
     OutBuffer(i) = Buffer(i) Xor OutBuffer(i - 1)
Next i

claveA = OutBuffer(Len(datos) - 1)
DyeCifro = StrConv(OutBuffer, vbUnicode)
End Function

Public Function DyeDecifro(ByRef data As String) As String
Dim Buffer() As Byte
Dim OutBuffer() As Byte
Dim i As Long

ReDim Buffer(Len(data) - 1) As Byte
ReDim OutBuffer(Len(data) - 1) As Byte

Buffer = StrConv(data, vbFromUnicode)

OutBuffer(0) = Buffer(0) Xor claveB

For i = 1 To (Len(data) - 1)
    OutBuffer(i) = Buffer(i) Xor Buffer(i - 1)
Next i

claveB = Buffer(Len(data) - 1)
DyeDecifro = StrConv(OutBuffer, vbUnicode)
End Function
