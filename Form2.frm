VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Ambil Karcis"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Langsung cetak tanpa preview cetak"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KELUAR"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AMBIL TIKET"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AMBIL TIKET"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "XXXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   2295
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "XXXX"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Caption         =   "No. Antrian CS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "No. Antrian Teller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim hitung As Integer
Dim ykn As String
Dim prw As New nomor
ykn = MsgBox("Apakah anda yakin akan mengambil antrian ?", vbYesNo)
If ykn = vbYes Then
prw.Field1.Text = Label3.Caption
prw.Label2.Caption = nomor.Label2.Caption & " TELLER"
prw.Label4.Caption = "Hari : " & hari(Format(Now, "dddd")) & ", " & Format(Now, " dd-MM-yyyy hh:mm:ss")
If Check1.Value = 0 Then
prw.Show
Else
prw.PrintReport False
End If
hitung = Val(Label3.Caption)
If hitung = 1 Then
Set rs = con.Execute("insert into konter_teler(tgl,konter,jenis) values(now,1,1)")
Else
Set rs = con.Execute("update konter_teler set konter=konter+1 where jenis=1 and datevalue(tgl)=datevalue(now)")
End If
hitung = hitung + 1
If hitung < 10 Then
Label3.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label3.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label3.Caption = hitung
End If
End If
End Sub
Function hari(inideh As String) As String
Select Case inideh

Case "Monday"
hari = "Senin"
Case "Sunday"
hari = "Minggu"

Case "Tuesday"
hari = "Selasa"

Case "Wednesday"
hari = "Rabu"

Case "Thursday"
hari = "Kamis"

Case "Friday"
hari = "Jum'at"

Case "Saturday"
hari = "Sabtu"
End Select

End Function
Private Sub Command2_Click()
Dim rs As New ADODB.Recordset
Dim hitung As Integer
Dim ykn As String
Dim prw As New nomor
ykn = MsgBox("Apakah anda yakin akan mengambil antrian ?", vbYesNo)
If ykn = vbYes Then
prw.Field1.Text = Label4.Caption
prw.Label2.Caption = nomor.Label2.Caption & " CS"
prw.Label4.Caption = "Hari : " & hari(Format(Now, "dddd")) & ", " & Format(Now, " dd-MM-yyyy hh:mm:ss")
If Check1.Value = 0 Then
prw.Show
Else
prw.PrintReport False
End If

hitung = Val(Label4.Caption)
If hitung = 1 Then
Set rs = con.Execute("insert into konter_teler(tgl,konter,jenis) values(now,1,2)")
Else
Set rs = con.Execute("update konter_teler set konter=konter+1 where jenis=2 and datevalue(tgl)=datevalue(now)")
End If
hitung = hitung + 1
If hitung < 10 Then
Label4.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label4.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label4.Caption = hitung
End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
MDIForm1.Enabled = True
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim hitung As Integer

Set rs = con.Execute("select id,konter from konter_teler where jenis=1 and datevalue(tgl)=datevalue(now)")
If rs.EOF Then
'Set rs = con.Execute("insert into konter_teler(tgl,konter,jenis) values(now,1,1)")
Label3.Caption = "001"
Else
hitung = rs.Fields(1) + 1
If hitung < 10 Then
Label3.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label3.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label3.Caption = hitung
End If

End If


Set rs = con.Execute("select id,konter from konter_teler where jenis=2 and datevalue(tgl)=datevalue(now)")
If rs.EOF Then
'Set rs = con.Execute("insert into konter_teler(tgl,konter,jenis) values(now,1,2)")
Label4.Caption = "001"
Else
hitung = rs.Fields(1) + 1
If hitung < 10 Then
Label4.Caption = "00" & hitung
ElseIf hitung > 9 And hitung < 100 Then
Label4.Caption = "0" & hitung
ElseIf hitung > 99 And hitung < 999 Then
Label4.Caption = hitung
End If


End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = True
End Sub

