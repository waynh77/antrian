VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form Form5 
   Caption         =   "Master Data User"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9120
   LinkTopic       =   "Form5"
   ScaleHeight     =   5040
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Height          =   3255
      Left            =   240
      OleObjectBlob   =   "Form5.frx":0000
      TabIndex        =   1
      Top             =   1560
      Width           =   8655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   582
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
      RecordSource    =   "select * from tbluser"
      Caption         =   "                                                                MASTER DATA USER"
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
   Begin VB.Label Label1 
      Caption         =   $"Form5.frx":2831
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\antrian.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from tbluser "
Set TDBGrid1.DataSource = Adodc1
TDBGrid1.Refresh
TDBGrid1.ReBind
End Sub

Private Sub Form_Resize()
TDBGrid1.Width = Abs(Me.Width - 700)
TDBGrid1.Height = Abs(Me.Height - TDBGrid1.Top - 700)

End Sub


