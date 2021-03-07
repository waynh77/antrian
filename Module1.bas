Attribute VB_Name = "Module1"
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public con As New ADODB.Connection
'Public nmuser As String
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Function nama(a As String) As String
  Select Case a
    Case "1": nama = "Satu "
    Case "2": nama = "Dua "
    Case "3": nama = "Tiga "
    Case "4": nama = "Empat "
    Case "5": nama = "Lima "
    Case "6": nama = "Enam "
    Case "7": nama = "Tujuh "
    Case "8": nama = "Delapan "
    Case "9": nama = "Sembilan "
    Case "0": nama = ""
  End Select
End Function
Public Function SayNe(nNumber As Double) 'As String
Dim Z, s, a, c, x
Dim ulang As Double
Dim i As Byte
Dim tampung(5) As String
Dim n As String
  n = LTrim(RTrim(nNumber))
  ulang = (Len(n) - 1) \ 3 + 1
  For i = 1 To ulang
     If Len(n) > 3 Then
       c = Mid(n, Len(n) - 2, 3)
       n = Mid(n, 1, Len(n) - 3)
       tampung(i) = c
     Else
       tampung(i) = n
     End If
  Next i
  
  Z = ""
  If n = "0" Then
    Z = "Nol"
  Else
  
    i = ulang
    Do
      a = ""
      x = ""
      s = tampung(i)
      While Len(s) < 3
        s = "0" + s
      Wend
      
      'digit ratusan
      
      If Mid(s, 1, 1) <> "0" Then
        If Mid(s, 1, 1) = "1" Then
          a = a + "Seratus "
        Else
          a = a + nama(Mid(s, 1, 1)) + "Ratus "
        End If
      End If
  
      'digit 11-19
      If Mid(s, 2, 1) = "1" Then
        If (Mid(s, 3, 1) <> "1") And (Mid(s, 3, 1) <> "0") Then a = a + nama(Mid(s, 3, 1)) + "Belas "
        If Mid(s, 3, 1) = "1" Then a = a + "Sebelas "
        If Mid(s, 3, 1) = "0" Then a = a + "Sepuluh "
      End If
  
      'digit puluhan
      If (Mid(s, 2, 1) <> "1") And (s <> "000") And (Mid(s, 2, 1) <> "0") Then
        a = a + nama(Mid(s, 2, 1)) + "Puluh " '{+nama(mid(s,3,1))}
      End If
      
      If (Mid(s, 3, 1) <> "0") And (Mid(s, 2, 1) <> "1") Then
        a = a + nama(Mid(s, 3, 1))
      End If
      'perkecualian untuk seribu
      If (i = 2) Then
        If s = "001" Then a = "Se"
      End If
      
      If s <> "000" Then
        If i = 1 Then x = ""
        If i = 2 Then x = "Ribu "
        If i = 3 Then x = "Juta "
        If i = 4 Then x = "Miliar "
        If i = 5 Then x = "Triliun "
      End If
      If a = "Se" Then x = LCase(x)
      Z = Z + a + x
      i = i - 1
    Loop Until i = 0
  End If
  'SayN = Z
End Function

Public Function SayN(nNumber As Double) As String
Dim Z, s, a, c, x
Dim ulang As Double
Dim i As Byte
Dim tampung(5) As String
Dim n As String
  n = LTrim(RTrim(nNumber))
  ulang = (Len(n) - 1) \ 3 + 1
  
  For i = 1 To ulang
     If Len(n) > 3 Then
       c = Mid(n, Len(n) - 2, 3)
       n = Mid(n, 1, Len(n) - 3)
       tampung(i) = c
     Else
       tampung(i) = n
     End If
  Next i
  
  Z = ""
  If n = "0" Then
    Z = "Nol "
  Else
  
    i = ulang
    Do
      a = ""
      x = ""
      s = tampung(i)
      While Len(s) < 3
        s = "0" + s
      Wend
      
      'digit ratusan
      
      If Mid(s, 1, 1) <> "0" Then
        If Mid(s, 1, 1) = "1" Then
          a = a + "Seratus "
        Else
          a = a + nama(Mid(s, 1, 1)) + "Ratus "
        End If
      End If
  
      'digit 11-19
      If Mid(s, 2, 1) = "1" Then
        If (Mid(s, 3, 1) <> "1") And (Mid(s, 3, 1) <> "0") Then a = a + nama(Mid(s, 3, 1)) + "Belas "
        If Mid(s, 3, 1) = "1" Then a = a + "Sebelas "
        If Mid(s, 3, 1) = "0" Then a = a + "Sepuluh "
      End If
  
      'digit puluhan
      If (Mid(s, 2, 1) <> "1") And (s <> "000") And (Mid(s, 2, 1) <> "0") Then
        a = a + nama(Mid(s, 2, 1)) + "Puluh " '{+nama(mid(s,3,1))}
      End If
      
      If (Mid(s, 3, 1) <> "0") And (Mid(s, 2, 1) <> "1") Then
        a = a + nama(Mid(s, 3, 1))
      End If
      'perkecualian untuk seribu
      If (i = 2) Then
        If s = "001" Then a = "Se"
      End If
      
      If s <> "000" Then
        If i = 1 Then x = ""
        If i = 2 Then x = "Ribu "
        If i = 3 Then x = "Juta "
        If i = 4 Then x = "Miliar "
        If i = 5 Then x = "Triliun "
      End If
      If a = "Se" Then x = LCase(x)
      Z = Z + a + x
      i = i - 1
    Loop Until i = 0
  End If
  SayN = Z
End Function

