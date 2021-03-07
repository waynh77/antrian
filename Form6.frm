VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   Caption         =   "Laporan Harian"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form6"
   ScaleHeight     =   1590
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker tgl2 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   16580611
      CurrentDate     =   40639
   End
   Begin MSComCtl2.DTPicker tgl1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   16580611
      CurrentDate     =   40639
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "sampai"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
lpharian.DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\antrian.mdb;Persist Security Info=False"
lpharian.DataControl1.Source = "SELECT datevalue(tgl) as tgl,iif(jenis=1,'TELLER','COSTUMER SERVICE'),konter,kount FROM HIT_KONTER " & _
            " where datevalue(tgl) between #" & DateValue(tgl1.Value) & "# and #" & _
            DateValue(tgl2.Value) & "# " & " order by tgl"
lpharian.Label2.Caption = "Tanggal " & Format(tgl1.Value, "dd-MM-yyyy") & " sampai " & Format(tgl2.Value, "dd-MM-yyyy")
lpharian.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
tgl1.Value = Now
tgl2.Value = Now
End Sub

