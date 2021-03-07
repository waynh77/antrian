VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Program ANtRiaN"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   8700
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1323
      ButtonWidth     =   1984
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Master Konter"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ambil Karcis"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TR Konter"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mesin Hitung"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Laporan"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8496
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B64C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D770
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E802
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":EF7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1000E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":110A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicBackDrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   885
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8640
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   8700
      Begin VB.PictureBox PicOriginal 
         AutoSize        =   -1  'True
         Height          =   8160
         Left            =   720
         Picture         =   "MDIForm1.frx":12132
         ScaleHeight     =   8100
         ScaleWidth      =   10800
         TabIndex        =   3
         Top             =   240
         Width           =   10860
      End
      Begin VB.PictureBox PicStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   3000
         ScaleHeight     =   2625
         ScaleWidth      =   4230
         TabIndex        =   2
         Top             =   3360
         Width           =   4230
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4890
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "4/6/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:27 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu fil 
      Caption         =   "File"
      Begin VB.Menu mk 
         Caption         =   "Master Konter"
      End
      Begin VB.Menu Muser 
         Caption         =   "Master User"
      End
      Begin VB.Menu Keluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu tr 
      Caption         =   "Transaksi"
      Begin VB.Menu ambil 
         Caption         =   "Ambil Karcis"
      End
      Begin VB.Menu trkon 
         Caption         =   "Tr Konter"
      End
      Begin VB.Menu mhit 
         Caption         =   "Mesin Hitung"
      End
   End
   Begin VB.Menu lap 
      Caption         =   "Laporan"
      Begin VB.Menu lpharian 
         Caption         =   "Laporan Harian"
      End
      Begin VB.Menu lpbulanan 
         Caption         =   "Laporan Bulanan"
      End
   End
   Begin VB.Menu hel 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ambil_Click()
Form2.Show
Me.Enabled = False
End Sub

Private Sub Keluar_Click()
Dim ykn As String
ykn = MsgBox("Apakah anda yakin akan keluar dari program ?", vbYesNo, "keluar program")
If ykn = vbYes Then
End
End If
End Sub

Private Sub lpbulanan_Click()
Form7.Show
End Sub

Private Sub lpharian_Click()
Form6.Show
End Sub

Private Sub MDIForm_Load()
'Unload frmLogin
End Sub

Private Sub MDIForm_Resize()
Dim client_rect As RECT
Dim client_hwnd As Long

    PicStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    PicStretched.PaintPicture _
        PicOriginal.Picture, _
        0, 0, _
        PicStretched.ScaleWidth, _
        PicStretched.ScaleHeight, _
        0, 0, _
        PicOriginal.ScaleWidth, _
        PicOriginal.ScaleHeight

    ' Set the MDI form's picture.
    Picture = PicStretched.Image

    ' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mhit_Click()
Form4.Show
End Sub

Private Sub mk_Click()
Form1.Show
End Sub

Private Sub Muser_Click()
Form5.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
Form1.Show

Case 2
Form2.Show
Me.Enabled = False

Case 3
Form3.Show
Me.Enabled = False

Case 4
Form4.Show
'Me.Enabled = False

Case 5
PopupMenu lap
'Case 6



End Select
End Sub

Private Sub trkon_Click()
Form3.Show
Me.Enabled = False
End Sub
