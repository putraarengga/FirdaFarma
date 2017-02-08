Public Class FormTransaksiKeuangan
    Dim day, month, year, vdate, vdate2 As String
    Dim simpan, nowTime, SelectedDatabase, hapussql As String
    Dim nama, jenis, pesan As String
    Dim dt1 As DateTime
    Dim tgl As String
    Dim amount As Integer
    Dim Flag As Integer = 0

    Private Sub FormTransaksiKeuangan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
        ComboBox1.Items.Add("PENGELUARAN")
        ComboBox1.Items.Add("PEMASUKAN")

    End Sub
    Sub enable()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        DateTimePicker1.Enabled = True
    End Sub
    Sub disable()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        DateTimePicker1.Enabled = False
    End Sub
    Sub clear()
        TextBox1.Text = ""
        TextBox2.Text = 0
        ComboBox1.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        enable()
        Button3.Enabled = True
        Button1.Enabled = False
        Flag = 1
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        year = DateTimePicker1.Value.Year
        month = DateTimePicker1.Value.Month
        day = DateTimePicker1.Value.Day
        vdate2 = year + "-" + month + "-" + day
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        If (TextBox1.Text = "" Or vdate2 = "" Or ComboBox1.Text = "" Or TextBox2.Text = "") Then
            MsgBox("Tidak bisa menyimpan data ke server data tidak lengkap")
        Else
            If Flag = 1 Then
                simpan = "INSERT INTO tkeuangan(NamaTransaksi,Tanggal,TglInput,JenisTransaksi,Jumlah)" +
                    "VALUES('" & TextBox1.Text & "', '" & vdate2 & "', '" & nowTime & "', '" & ComboBox1.Text & "', '" & TextBox2.Text & "')"
                jalankansql(simpan)
                IsiGridUmum1()
                clear()
                disable()
                Flag = 0
            ElseIf Flag = 2 Then
                simpan = "UPDATE tkeuangan SET NamaTransaksi ='" & TextBox1.Text & "',Tanggal='" & vdate2 & "',JenisTransaksi='" & ComboBox1.Text & "',Jumlah= '" & TextBox2.Text & "' " +
                    " WHERE tkeuangan.NamaTransaksi = '" & nama & "'  AND tkeuangan.TglInput ='" & nowTime & "' "
                jalankansql(simpan)
                DataGridView1.Refresh()
                IsiGridUmum1()
                clear()
                disable()
                Flag = 0
            End If
            Button1.Enabled = True
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        End If

    End Sub
    Sub IsiGridUmum1()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        SelectedDatabase = " SELECT tkeuangan.NamaTransaksi, tkeuangan.Tanggal, tkeuangan.JenisTransaksi, tkeuangan.Jumlah, tkeuangan.TglInput " +
            " FROM tkeuangan WHERE tkeuangan.TglInput ='" & nowTime & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(SelectedDatabase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tkeuangan")
        DataGridView1.DataSource = (DS.Tables("tkeuangan"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Transaksi"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Nominal"
            .Columns(4).HeaderCell.Value = "Tanggal Input"
        
        End With

    End Sub
    Private Sub jalankansql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            objcmd.ExecuteNonQuery()
            objcmd.Dispose()
            MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'Dim stringtemp As String
        'Dim nominal As Integer
        If IsNumeric(TextBox2.Text) Then
            'nominal = Val(TextBox2.Text)
            'stringtemp = FormatNumber(TextBox2.Text)
            'TextBox2.Text = stringtemp
        Else
            MsgBox("Bukan Angka", MsgBoxStyle.OkCancel)
            TextBox2.Text = 0
        End If
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            nama = DataGridView1.Rows(x).Cells(0).Value.ToString
            dt1 = DataGridView1.Rows(x).Cells(1).Value
            jenis = DataGridView1.Rows(x).Cells(2).Value.ToString
            amount = DataGridView1.Rows(x).Cells(3).Value
            tgl = DataGridView1.Rows(x).Cells(4).Value.ToString
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        Button2.Enabled = True
        Button4.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = nama
        DateTimePicker1.Value = dt1
        ComboBox1.Text = jenis
        TextBox2.Text = amount
        enable()
        Flag = 2
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If nama = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + nama, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            hapussql = "DELETE FROM tkeuangan WHERE tkeuangan.NamaTransaksi= '" & nama & "' "
            jalankansql(hapussql)
            DataGridView1.Refresh()
            IsiGridUmum1()
            nama = ""
            Button1.Enabled = True
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Flag = 0
        End If
    End Sub
End Class