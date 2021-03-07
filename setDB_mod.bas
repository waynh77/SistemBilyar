Attribute VB_Name = "setDB_mod"
Dim db As String

Public Sub data()
    db = App.Path & "\master penjualan.mdb"
End Sub

Public Sub open_db()
data
DBadmin_frm.Data1.DatabaseName = db
DBadmin_frm.Data1.RecordSource = "data_admin"
DBPool_frm.Data1.DatabaseName = db
DBPool_frm.Data1.RecordSource = "pool"
DBproduk_form.DataProduk.DatabaseName = db
DBproduk_form.DataProduk.RecordSource = "produk"
DbTransFrm.Data1.DatabaseName = db
DbTransFrm.Data1.RecordSource = "transaksi"
'frmlaporan.Data1.DatabaseName = db
'frmlaporan.Data2.DatabaseName = db
'frmlaporan.Data3.DatabaseName = db
'frmlaporan.Data4.DatabaseName = db
frmloginuser.Data1.DatabaseName = db
frmloginuser.Data1.RecordSource = "data_admin"
frmPool.dt1.DatabaseName = db
frmPool.dt1.RecordSource = "temp_pool"
frmPool.dt2.DatabaseName = db
frmPool.dt2.RecordSource = "pool"
frmseting.Data1.DatabaseName = db
frmseting.Data1.RecordSource = "setingharga"
'test_sql.Data1.DatabaseName = db
'test_sql.Data2.DatabaseName = db
Transaksi_form.Data1.DatabaseName = db
Transaksi_form.Data1.RecordSource = "transaksi"
Transaksi_form.Data2.DatabaseName = db
Transaksi_form.Data2.RecordSource = "temp_trans"
Transaksi_form.DB_Prod.DatabaseName = db
Transaksi_form.DB_Prod.RecordSource = "produk"
End Sub

