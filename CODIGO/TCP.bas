Attribute VB_Name = "Mod_TCP"
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

Public Function PuedoQuitarFoco() As Boolean
    PuedoQuitarFoco = True
End Function

Sub login(ByVal Valcode As String)

    If EstadoLogin = Normal Then
        Dim a As String
        MsgBox (MD5HushYo)
        a = UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & Valcode & "," & MD5HushYo
        SendData (ClientPackages.login & a)
        
    ElseIf EstadoLogin = CrearNuevoPj Then ' hay que cambiar esto
        SendData (ClientPackages.register & UserName & "," & UserPassword _
                & "," & App.Major & "." & App.Minor & "." & App.Revision _
                & "," & UserRaza & "," & UserSexo & "," & UserClase _
                & "," & UserSkills(1) & "," & UserSkills(2) _
                & "," & UserSkills(3) & "," & UserSkills(4) _
                & "," & UserSkills(5) & "," & UserSkills(6) _
                & "," & UserSkills(7) & "," & UserSkills(8) _
                & "," & UserSkills(9) & "," & UserSkills(10) _
                & "," & UserSkills(11) & "," & UserSkills(12) _
                & "," & UserSkills(13) & "," & UserSkills(14) _
                & "," & UserSkills(15) & "," & UserSkills(16) _
                & "," & UserSkills(17) & "," & UserSkills(18) _
                & "," & UserSkills(19) & "," & UserSkills(20) _
                & "," & UserSkills(21) & "," & UserSkills(22) _
                & "," & UserSkills(23) & "," & UserSkills(24) _
                & "," & UserEmail & "," & UserHogar & "," & UserPreg & "," & UserResp & "," & Valcode & "," & MD5HushYo)

    End If
End Sub


