VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   Caption         =   "TR TICKETING"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3510
   LinkTopic       =   "Form3"
   ScaleHeight     =   4155
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2280
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ULANGI PANGGILAN"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PANGGIL ANTRIAN BERIKUTNYA"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form3.frx":0000
      Left            =   240
      List            =   "Form3.frx":000A
      TabIndex        =   3
      Text            =   "TELLER"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form3.frx":0028
      Left            =   240
      List            =   "Form3.frx":0032
      TabIndex        =   1
      Text            =   "TELLER"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA KONTER"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS KONTER"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hitu As Integer
Private Sub Command1_Click()
Dim ykn As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim hitung As Integer, hitung1 As Integer
ykn = MsgBox("Apakah anda yakin akan memanggil antrian berikutnya ?", vbYesNo)
If ykn = vbYes Then
If Combo1.Text = "TELLER" Then
Set rs = con.Execute("select id,konter2,konter from konter_teler where datevalue(tgl)=datevalue(now) and jenis=1")
If rs.EOF Then
MsgBox "Data Antrian Tidak ada"


ElseIf rs.Fields(1) = 0 Then
hitu = 1
Set rs1 = con.Execute("update konter_teler set konter2=konter2+1 where id=" & rs.Fields(0))
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now,1,'" & Combo2.Text & "',1)")
Set rs1 = con.Execute("select id,kount from hit_konter where konter='" & Combo2.Text & "' and jenis=1 and datevalue(tgl)=datevalue(now)")
If rs1.EOF Then
Set rs2 = con.Execute("insert into hit_konter(tgl,konter,jenis,kount) values(now,'" & Combo2.Text & _
                    "',1,1)")
                    
Else
Set rs2 = con.Execute("Update hit_konter set kount=kount+1 where id=" & rs1.Fields(0))

End If

Else
hitu = rs.Fields(1)
hitung = rs.Fields(1) + 1

hitung1 = rs.Fields(2)
If hitung > hitung1 Then
MsgBox "Data Antrian Tidak ada"
Else
hitu = hitung
Set rs1 = con.Execute("update konter_teler set konter2=konter2+1 where id=" & rs.Fields(0))
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now," & hitung & ",'" & Combo2.Text & "',1)")
Set rs1 = con.Execute("select id,kount from hit_konter where konter='" & Combo2.Text & "' and jenis=1 and datevalue(tgl)=datevalue(now)")
If rs1.EOF Then
Set rs2 = con.Execute("insert into hit_konter(tgl,konter,jenis,kount) values(now,'" & Combo2.Text & _
                    "',1,1)")
                    
Else
Set rs2 = con.Execute("Update hit_konter set kount=kount+1 where id=" & rs1.Fields(0))

End If

End If
End If
Else
'COSTUMER SERVICE
Set rs = con.Execute("select id,konter2,konter from konter_teler where datevalue(tgl)=datevalue(now) and jenis=2")
If rs.EOF Then
MsgBox "Data Antrian Tidak ada"

ElseIf rs.Fields(1) = 0 Then
hitu = 1
Set rs1 = con.Execute("update konter_teler set konter2=konter2+1 where id=" & rs.Fields(0))
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now,1,'" & Combo2.Text & "',2)")
Set rs1 = con.Execute("select id,kount from hit_konter where konter='" & Combo2.Text & "' and jenis=2 and datevalue(tgl)=datevalue(now)")
If rs1.EOF Then
Set rs2 = con.Execute("insert into hit_konter(tgl,konter,jenis,kount) values(now,'" & Combo2.Text & _
                    "',2,1)")
                    
Else
Set rs2 = con.Execute("Update hit_konter set kount=kount+1 where id=" & rs1.Fields(0))

End If

Else
hitu = rs.Fields(1)
hitung = rs.Fields(1) + 1
hitung1 = rs.Fields(2)
If hitung > hitung1 Then
MsgBox "Data Antrian Tidak ada"
Else
hitu = hitung
Set rs1 = con.Execute("update konter_teler set konter2=konter2+1 where id=" & rs.Fields(0))
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now," & hitung & ",'" & Combo2.Text & "',2)")
Set rs1 = con.Execute("select id,kount from hit_konter where konter='" & Combo2.Text & "' and jenis=2 and datevalue(tgl)=datevalue(now)")
If rs1.EOF Then
Set rs2 = con.Execute("insert into hit_konter(tgl,konter,jenis,kount) values(now,'" & Combo2.Text & _
                    "',2,1)")
                    
Else
Set rs2 = con.Execute("Update hit_konter set kount=kount+1 where id=" & rs1.Fields(0))

End If

End If
End If


End If


End If
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Enabled = True
End Sub

Private Sub Command3_Click()
Dim ykn As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
ykn = MsgBox("Apakah anda yakin akan mengulangi panggilan?", vbYesNo)
If ykn = vbYes Then
If hitu = 0 Then
MsgBox "Anda tidak punya data"
Exit Sub
End If
If Combo1.Text = "TELLER" Then
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now," & hitu & ",'" & Combo2.Text & "',1)")

Else
Set rs2 = con.Execute("insert into h_call(tgl,antrian,konter,jenis) values(now," & hitu & ",'" & Combo2.Text & "',2)")

End If
End If

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim ki As Integer
Set rs = con.Execute("select nama,suara,jenis from konter where ip='" & Winsock1.LocalIP & "'")
If rs.EOF Then
MsgBox "Maaf anda belum terdaftar menjadi konter"
Command1.Enabled = False
Command3.Enabled = False
'Unload Form3
'MDIForm1.Enabled = True
Else
ki = rs.Fields(2)
If ki = 1 Then
Combo1.Text = "TELLER"

Else
Combo1.Text = "COSTUMER SERVICE"
End If
Combo2.Text = rs.Fields(0)
Label3.Caption = rs.Fields(1)

End If
hitu = 0
End Sub
