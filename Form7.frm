VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Laporan Bulanan"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4845
   LinkTopic       =   "Form7"
   ScaleHeight     =   1650
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker tgl1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   16515075
      CurrentDate     =   40639
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
lpharian.DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\antrian.mdb;Persist Security Info=False"
lpharian.DataControl1.Source = "SELECT datevalue(tgl) as tgl,iif(jenis=1,'TELLER','COSTUMER SERVICE'),konter,kount FROM HIT_KONTER " & _
            " where month(tgl)=" & Month(tgl1.Value) & " and year(tgl)=" & _
            Year(tgl1.Value) & " order by tgl"
lpharian.Label1.Caption = "Laporan Bulanan"
lpharian.Label2.Caption = "Bulan " & bln(Month(tgl1.Value)) & " Tahun " & Format(tgl1.Value, "yyyy")
lpharian.Show

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Function bln(inta As Integer) As String
Select Case inta
Case 1
bln = "Januari"
Case 2
bln = "Februari"
Case 3
bln = "Maret"

Case 4
bln = "April"


Case 5
bln = "Mei"

Case 6
bln = "Juni"

Case 7
bln = "Juli"

Case 8
bln = "Agustus"
Case 9
bln = "September"
Case 10
bln = "Oktober"
Case 11
bln = "November"
Case 12
bln = "Desember"
End Select

End Function

Private Sub Form_Load()
tgl1.Value = Now
End Sub
