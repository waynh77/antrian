VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Cetak_Transaksi 
   Caption         =   "Cetak Transaksi"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   15161
   SectionData     =   "ctk_trans.dsx":0000
End
Attribute VB_Name = "Cetak_Transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tot As Double
Dim h, i As Single

Private Sub ActiveReport_ReportStart()
Field1.DataField = "jenis_produk"
Field6.DataField = "nama_produk"
Field2.DataField = "harga"
Field3.DataField = "qty"
Field4.DataField = "jumlah"
tot = 0
i = Trans_frm.Data1.Recordset.RecordCount
h = 680 * (8 + i)
Cetak_Transaksi.PageSettings.PaperSize = 256
Me.PageSettings.PaperWidth = 4700
Cetak_Transaksi.Printer.PaperHeight = h
Label19.Caption = Format(Date, "dd mmm yyyy") & " - " & Format(Time, "hh:mm:ss")
End Sub

Private Sub ActiveReport_Terminate()
'    Trans_frm.simpan_Trans
'    Trans_frm.Show
'    Trans_frm.hapus_temp
End Sub

Private Sub Detail_Format()
Label6.Caption = Field1.Text & " : " & Field6.Text
tot = tot + Val(Format(Field4, "###.##"))
End Sub

Private Sub PageFooter_Format()
Field5 = Format(tot, "###,###.00")
Field9 = Format(tot - Val(Format(Field7, "###.00")) + Val(Format(Field8, "###.00")), "###,###.00")
End Sub
