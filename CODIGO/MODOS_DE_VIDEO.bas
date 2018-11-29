Attribute VB_Name = "Mod_MODOS_DE_VIDEO"
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

'Testea si la maquina soporta un modo de video ;-)
Function SoportaDisplay(DD As DirectDraw7, DDSDaTestear As DDSURFACEDESC2) As Boolean
Dim ddsd As DDSURFACEDESC2
Dim DDEM As DirectDrawEnumModes

Set DDEM = DD.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd)

Dim loopc As Integer
Dim flag As Boolean
loopc = 1
   
Do While loopc <> DDEM.GetCount And Not flag

    DDEM.GetItem loopc, ddsd
    flag = ddsd.lHeight = DDSDaTestear.lHeight _
    And ddsd.lWidth = DDSDaTestear.lWidth _
    And ddsd.ddpfPixelFormat.lRGBBitCount = _
    DDSDaTestear.ddpfPixelFormat.lRGBBitCount
    loopc = loopc + 1
Loop
SoportaDisplay = flag
End Function

Function ModosDeVideoIguales(dd1 As DDSURFACEDESC2, dd2 As DDSURFACEDESC2) As Boolean
ModosDeVideoIguales = _
    dd1.lHeight = dd2.lHeight _
    And dd1.lWidth = dd2.lWidth _
    And dd1.ddpfPixelFormat.lRGBBitCount = _
    dd2.ddpfPixelFormat.lRGBBitCount
End Function

Public Sub DibujarMiniMapa(ByRef Pic As PictureBox)
'CHOTS | Sistema de Minimapa by Standelf

    Dim DR As RECT
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 100
    DR.Right = 100
    SupMiniMap.BltFast 1, 1, SupBMiniMap, DR, DDBLTFAST_WAIT
   
       DR.Left = UserPos.x - 4
       DR.Top = UserPos.y - 4
       DR.Bottom = UserPos.y - 2
       DR.Right = UserPos.x - 2
       SupMiniMap.BltColorFill DR, &HFFFF00
  
       DR.Left = 0
       DR.Top = 0
       DR.Bottom = 100
       DR.Right = 100
       SupMiniMap.BltToDC Pic.hDC, DR, DR
End Sub
  
   Public Sub GenerarMiniMapa()
   On Local Error Resume Next
       'CHOTS | Sistema de Minimapa by Standelf
       Dim x As Integer
       Dim y As Integer
       Dim i As Integer
       Dim DR As RECT
       Dim SR As RECT
       Dim aux As Integer
  
       SR.Left = 0
       SR.Top = 0
       SR.Bottom = 100
       SR.Right = 100
       'SupBMiniMap.BltColorFill SR, vbBlack
  
       For x = MinYBorder To MaxXBorder
           For y = MinYBorder To MaxYBorder
               If MapData(x, y).Graphic(1).GrhIndex > 0 Then
                   With MapData(x, y).Graphic(1)
                       i = GrhData(.GrhIndex).Frames(1)
                   End With
  
                   SR.Left = GrhData(i).sX
                   SR.Top = GrhData(i).sY
                   SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                   SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                   DR.Left = x - 5
                   DR.Top = y - 5
                   DR.Bottom = y - 3
                   DR.Right = x - 3
                   SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
               End If
  
               If MapData(x, y).Graphic(2).GrhIndex > 0 Then
                   With MapData(x, y).Graphic(2)
                       i = GrhData(.GrhIndex).Frames(1)
                   End With
  
                   SR.Left = GrhData(i).sX
                   SR.Top = GrhData(i).sY
                   SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                   SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                   DR.Left = x - 5
                   DR.Top = y - 5
                   DR.Bottom = y - 3
                   DR.Right = x - 3
                   SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
             End If

             If MapData(x, y).Graphic(3).GrhIndex > 0 Then
                 With MapData(x, y).Graphic(3)
                     i = GrhData(.GrhIndex).Frames(1)
                 End With

                 SR.Left = GrhData(i).sX
                 SR.Top = GrhData(i).sY
                 SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                 SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                 DR.Left = x - 5
                 DR.Top = y - 5
                 DR.Bottom = y - 3
                 DR.Right = x - 3
                 SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
             End If

         Next
     Next

End Sub
