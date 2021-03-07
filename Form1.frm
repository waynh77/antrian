VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form Form1 
   Caption         =   "Data Master Konter"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "                                                  DATA MASTER KONTER"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   1440
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":2831
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rs As New ADODB.Recordset
Private Sub Form_Activate()
'Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.CursorLocation = adUseClient
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\antrian.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from konter order by nama"
'rs.Open "select * from konter order by konter", con, adOpenDynamic, adLockBatchOptimistic
Set TDBGrid1.DataSource = Adodc1
TDBGrid1.Refresh
TDBGrid1.Refresh
'TDBGrid1.ReBind
End Sub

Private Sub Form_Resize()
TDBGrid1.Width = Abs(Me.Width - 700)
TDBGrid1.Height = Abs(Me.Height - TDBGrid1.Top - 700)

End Sub

