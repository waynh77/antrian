VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "ANTRIAN"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11850
   LinkTopic       =   "Form4"
   ScaleHeight     =   8505
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   12120
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   18000
      Top             =   360
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   18615
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   9975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SERATUS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         TabIndex        =   13
         Top             =   2280
         Width           =   10335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ANTRIAN "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   10440
         TabIndex        =   12
         Top             =   120
         Width           =   4455
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   14760
         X2              =   14760
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   10440
         X2              =   10440
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   15000
         TabIndex        =   11
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   10920
         TabIndex        =   10
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   15000
         TabIndex        =   9
         Top             =   1560
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   18615
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   15000
         TabIndex        =   7
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   10920
         TabIndex        =   6
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TELLER"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   15000
         TabIndex        =   5
         Top             =   120
         Width           =   3375
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   10440
         X2              =   10440
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   14760
         X2              =   14760
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ANTRIAN "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   10440
         TabIndex        =   4
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SERATUS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         TabIndex        =   3
         Top             =   2280
         Width           =   10335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9975
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10:10:10"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   12360
      TabIndex        =   16
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANTRIAN NO. C S"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Width           =   10095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANTRIAN NO. PADA TELLER"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_SYNC = &H0         '  play synchronously (default)
Dim jreng As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim kntr As String
Dim tipedeh As Integer
Public Sub Delay(detik As Long)
Dim TimeO As Long
TimeO = (GetTickCount / 1000) + detik
Do
DoEvents
Loop Until TimeO < (GetTickCount / 1000)
End Sub
Private Function nama(a As String) As String
  Select Case a
    Case "1":
    DoEvents
    jreng = PlaySound(App.Path & "\satu.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "2":
    DoEvents
    jreng = PlaySound(App.Path & "\dua.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1
    
    Case "3":
    DoEvents
    jreng = PlaySound(App.Path & "\tiga.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "4":
    DoEvents
    jreng = PlaySound(App.Path & "\empat.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "5":
    DoEvents
    jreng = PlaySound(App.Path & "\lima.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "6":
    DoEvents
    jreng = PlaySound(App.Path & "\enam.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1
    
    Case "7":
    DoEvents
    jreng = PlaySound(App.Path & "\tujuh.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "8":
    DoEvents
    jreng = PlaySound(App.Path & "\delapan.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1

    Case "9":
    DoEvents
    jreng = PlaySound(App.Path & "\sembilan.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 1
    
    Case "0": nama = ""
  End Select
End Function
Private Function SayNex(nNumber As Double) 'As String
Dim Z, s, a, c, x
Dim ulang As Double
Dim i As Byte
Dim tampung(5) As String
Dim n As String
Dim rsx As New ADODB.Recordset
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
          'a = a + "Seratus "
            DoEvents
            jreng = PlaySound(App.Path & "\seratus.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
            Delay 1
          
        Else
          a = a + nama(Mid(s, 1, 1)) '+ "Ratus "
          
        DoEvents
        jreng = PlaySound(App.Path & "\ratus.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
        Delay 1
          
        End If
      End If
  
      'digit 11-19
      If Mid(s, 2, 1) = "1" Then
        If (Mid(s, 3, 1) <> "1") And (Mid(s, 3, 1) <> "0") Then
        a = a + nama(Mid(s, 3, 1)) '+ "Belas "
        DoEvents
        jreng = PlaySound(App.Path & "\belas.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
        Delay 1
        End If
        If Mid(s, 3, 1) = "1" Then
        'a = a + "Sebelas "
        DoEvents
        jreng = PlaySound(App.Path & "\sebelas.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
        Delay 1
        End If
        If Mid(s, 3, 1) = "0" Then
        'a = a + "Sepuluh "
        DoEvents
        jreng = PlaySound(App.Path & "\sepuluh.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
        Delay 1
        End If
      End If
  
      'digit puluhan
      If (Mid(s, 2, 1) <> "1") And (s <> "000") And (Mid(s, 2, 1) <> "0") Then
        a = a + nama(Mid(s, 2, 1)) '+ "Puluh " '{+nama(mid(s,3,1))}
        DoEvents
        jreng = PlaySound(App.Path & "\puluh.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
        Delay 1
        
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
  Set rsx = con.Execute("select suara From konter where nama='" & kntr & "' and jenis=" & tipedeh)
  If rsx.EOF = False Then
    If tipedeh = 1 Then
    DoEvents
    jreng = PlaySound(App.Path & "\teler.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 2
    DoEvents
    jreng = PlaySound(App.Path & "\" & rsx.Fields(0), SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Else
    DoEvents
    jreng = PlaySound(App.Path & "\cs.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    Delay 3
    DoEvents
    jreng = PlaySound(App.Path & "\" & rsx.Fields(0), SND_FILENAME Or SND_ASYNC, SND_ASYNC)
    
    End If
'    Delay 3
  End If
  'SayN = Z
  Timer2.Enabled = True
End Function

Private Sub Timer1_Timer()
Label15.Caption = Format(Now, "hh:mm:ss")


End Sub

Private Sub Timer2_Timer()
Dim rs As New ADODB.Recordset
Dim hitung As Double
Set rs = con.Execute("Select konter2 from konter_teler where datevalue(tgl)=datevalue(now) and jenis=1")
If rs.EOF = False Then
hitung = rs.Fields(0)
If hitung < 10 Then
Label2.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label2.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label2.Caption = hitung
End If
Label3.Caption = SayN(hitung)
End If
Set rs = con.Execute("Select konter2 from konter_teler where datevalue(tgl)=datevalue(now) and jenis=2")
If rs.EOF = False Then
hitung = rs.Fields(0)
If hitung < 10 Then
Label13.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label13.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label13.Caption = hitung
End If
Label12.Caption = SayN(hitung)
End If
Call jadidah

End Sub
Sub jadidah()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim jns As Integer
Dim ang As Double
Timer2.Enabled = False
Set rs = con.Execute("select * from h_call where datevalue(tgl)=datevalue(now) and status=0")
If rs.EOF = False Then
jns = rs.Fields!jenis
If jns = 1 Then
Label9.Caption = rs.Fields!antrian
Label10.Caption = rs.Fields!konter
tipedeh = 1
kntr = Label10.Caption
ang = rs.Fields!antrian
DoEvents
jreng = PlaySound(App.Path & "\antrian.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
Delay 3

DoEvents
SayNex (ang)

Else
Label5.Caption = rs.Fields!antrian
Label4.Caption = rs.Fields!konter
tipedeh = 2
kntr = Label4.Caption
ang = rs.Fields!antrian
DoEvents
jreng = PlaySound(App.Path & "\antrian.wav", SND_FILENAME Or SND_ASYNC, SND_ASYNC)
Delay 3

DoEvents
SayNex (ang)
End If
Set rs1 = con.Execute("update h_call set status=1 where id=" & rs.Fields(0))

Else
Timer2.Enabled = True
End If
End Sub
