Public Class FormDataResep
    Dim databaru As Boolean
    Dim idresep, nomorfaktur As Integer
    Dim namapasien, nameO, sb As String
    Dim JumlahObat, TotalHarga As Integer
    Dim SubTotal As Integer
    Dim vdate As String
    Dim selectedDatabase, simpan As String
    Dim Keterangan, ket, Jumlah, Jum As Integer
    Dim satuan, satuanlv1, satuanlv2, satuanlv3, satbeli, satbel As String
    Dim stokobatLv1, stokobatLv2, stokobatLv3 As Integer
    Dim sisaobatLv1, sisaobatLv2, sisaobatLv3 As Integer
    Dim sisobatLv1, sisobatLv2, sisobatLv3 As Integer
    Dim konversiLv2 As Integer
    Dim konversiLv3 As Integer
    Dim selectDataBase As String

    Shared Property IDobat As Integer
    Shared Property namaobat As String
    Shared Property hargaobat As Integer
    Shared Property tglKadaluarsa As String
    Shared Property fakturbeli As Integer
    Shared Property satuanlevel1 As String
    Shared Property jumlahlevel1 As Integer
    Shared Property satuanlevel2 As String
    Shared Property jumlahlevel2 As Integer
    Shared Property satuanlevel3 As String
    Shared Property jumlahlevel3 As Integer
    Shared Property Stoksatuan As Integer

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub FormDataSupplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        IsiGrid()
        IsiGridDetailResep()
        TextBox13.Enabled = False

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\1461484694_6.ico")
    End Sub
    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT * FROM tresep WHERE tresep.NomorFakturPenjualan='" & FormTransaksiPenjualan.TextBox1.Text & "'", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tresep")
        DataGridView1.DataSource = (DS.Tables("tresep"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Resep"
            .Columns(1).HeaderCell.Value = "Faktur Penjualan "
            .Columns(2).HeaderCell.Value = "Nama Pasien "
            .Columns(3).HeaderCell.Value = "Usia"
            .Columns(5).HeaderCell.Value = "No. Resep"
            .Columns(6).HeaderCell.Value = "Nama Dokter"
            .Columns(7).HeaderCell.Value = "Tanggal Resep"
            .Columns(8).HeaderCell.Value = "Copy Resep"
        End With
    End Sub
    Sub Bersih()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox13.Text = ""
    End Sub
    Private Sub isitextbox(ByVal x As Integer)
        Try
            idresep = DataGridView1.Rows(x).Cells(0).Value
            nomorfaktur = DataGridView1.Rows(x).Cells(1).Value
            namapasien = DataGridView1.Rows(x).Cells(2).Value.ToString
            TextBox13.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox2.Text = DataGridView1.Rows(x).Cells(2).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(3).Value
            TextBox4.Text = DataGridView1.Rows(x).Cells(4).Value
            TextBox5.Text = DataGridView1.Rows(x).Cells(5).Value
            TextBox6.Text = DataGridView1.Rows(x).Cells(6).Value
            TextBox8.Text = DataGridView1.Rows(x).Cells(8).Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox2.Focus()
        databaru = True
        Button3.Enabled = True
        Button2.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox8.Enabled = True
        DateTimePicker2.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String
        Dim vdate As String
        If TextBox2.Text = "" Then Exit Sub
        pesan = " "
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "yyyy-MM-dd"
        vdate = DateTimePicker2.Value

        If DataGridView1.Rows.Count > 0 Then
            If idresep = TextBox13.Text And namapasien = TextBox2.Text Then
                pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
                If pesan = MsgBoxResult.No Then
                    Exit Sub
                End If
                simpan = "UPDATE tresep SET NomorFakturPenjualan= '" & FormTransaksiPenjualan.TextBox1.Text & "', NamaPasien='" & TextBox2.Text & "',Usia='" & TextBox3.Text & "'," +
                    "Alamat='" & TextBox4.Text & "',NoResep= '" & TextBox5.Text & "', NamaDokter='" & TextBox6.Text & "',TglResep='" & vdate & "',CopyResep='" & TextBox8.Text & "' " +
                    "WHERE tresep.NomorFakturPenjualan= '" & nomorfaktur & "' AND tresep.IDResep='" & TextBox13.Text & "' AND tresep.NamaPasien='" & TextBox2.Text & "'"
                jalankansql(simpan)
                DataGridView1.Refresh()
                IsiGrid()
                Button2.Enabled = False
                Button3.Enabled = False
            Else
                pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
                If pesan = MsgBoxResult.No Then
                    Exit Sub
                End If
                simpan = "INSERT INTO tresep(NomorFakturPenjualan,NamaPasien, Usia, Alamat, NoResep, NamaDokter, TglResep, CopyResep) " +
                 "VALUES ('" & FormTransaksiPenjualan.TextBox1.Text & "', '" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & vdate & "','" & TextBox8.Text & "')"
                jalankansql(simpan)
                DataGridView1.Refresh()
                IsiGrid()
                Button2.Enabled = False
                Button3.Enabled = False
            End If
        
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tresep(NomorFakturPenjualan,NamaPasien, Usia, Alamat, NoResep, NamaDokter, TglResep, CopyResep) " +
             "VALUES ('" & FormTransaksiPenjualan.TextBox1.Text & "', '" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & vdate & "','" & TextBox8.Text & "')"
            jalankansql(simpan)
            DataGridView1.Refresh()
            IsiGrid()
            Button2.Enabled = False
            Button3.Enabled = False
        End If
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox8.Enabled = False
        DateTimePicker2.Enabled = False
        Bersih()
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
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        isitextbox(e.RowIndex)
        If DataGridView1.Rows.Count > 0 Then
            IsiGridDetailResep()
            Button5.Enabled = True
            ComboBox2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button2.Enabled = False
            databaru = False

            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox8.Enabled = True
            DateTimePicker2.Enabled = True

        Else
            Button5.Enabled = False
            ComboBox2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        If TextBox2.Text = "" Or TextBox13.Text = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Resep Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        ElseIf DataGridView2.Rows.Count > 0 Then
            pesan = MsgBox("Hapus Dulu data Obat ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server?" + TextBox2.Text, vbExclamation + vbYesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            hapussql = "DELETE FROM tresep WHERE IDResep ='" & TextBox13.Text & "'"
            jalankansql(hapussql)
            DataGridView1.Refresh()
            IsiGrid()
            Bersih()
            Button2.Enabled = True
            Button3.Enabled = False
            Button4.Enabled = False

            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox8.Enabled = False
            DateTimePicker2.Enabled = False

        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tsupplier WHERE NamaSupplier LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tsatuan")
            DataGridView1.DataSource = (DS.Tables("tsatuan"))
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button3.Enabled = False
        Button4.Enabled = False
        FormPencarianStokObat.IDPencariObat = 1
        FormPencarianStokObat.Show()
        FormPencarianStokObat.Focus()
        If TextBox2.Text = "" Then
            MsgBox("Tidak dapat mengetahui Nama Racikan Obat", MsgBoxStyle.OkCancel)
            FormPencarianStokObat.Close()
        End If
    End Sub

    Private Sub FormDataResep_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
    Sub IsiGridDetailResep()
        DataGridView2.Refresh()
        selectedDatabase = "SELECT tdetiltransaksi.NamaObat, tdetiltransaksi.Jumlah, tdetiltransaksi.Satuan, tdetiltransaksi.Harga, tdetiltransaksi.SubTotal," +
            "tdetiltransaksi.Kadaluarsa, tdetiltransaksi.FakturBeliObat FROM tdetiltransaksi WHERE tdetiltransaksi.IDResep = '" & idresep.ToString & "' AND tdetiltransaksi.NomorFakturPenjualan= '" & FormTransaksiPenjualan.TextBox1.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectedDatabase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdetiltransaksi")
        DataGridView2.DataSource = (DS.Tables("tdetiltransaksi"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Obat"
            .Columns(1).HeaderCell.Value = "Jumlah "
            .Columns(2).HeaderCell.Value = "Satuan Obat "
            .Columns(3).HeaderCell.Value = "Harga Jual"
            .Columns(5).HeaderCell.Value = "Sub Total"
            .Columns(5).HeaderCell.Value = "Kadaluarsa"
            .Columns(6).HeaderCell.Value = "Keterangan"
        End With
    End Sub
    Sub isitextbox2(ByVal x As Integer)
        Try
            namaobat = DataGridView2.Rows(x).Cells(0).Value.ToString
            Jumlah = DataGridView2.Rows(x).Cells(1).Value
            sb = DataGridView2.Rows(x).Cells(2).Value.ToString
            Keterangan = DataGridView2.Rows(x).Cells(6).Value

        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        isitextbox2(e.RowIndex)
        nameO = namaObat
        Jum = Jumlah
        satbel = sb
        ket = Keterangan
        Button6.Enabled = True
        Button7.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
    End Sub
    Sub tambahObat()
        Dim simpan, simpan2 As String
        Dim pesan As String
        Dim tmpString As String

        tmpString = Format(DateTime.Now, "yyyy-MM-dd")
        If TextBox2.Text = "" Then Exit Sub

        pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
        If pesan = MsgBoxResult.No Then
            Exit Sub
        End If
        simpan = "INSERT INTO tdetiltransaksi(IDResep, NomorFakturPenjualan, NamaPasien,IDObat, NamaObat, Jumlah, Satuan, Harga,SubTotal,Kadaluarsa, FakturBeliObat)" +
            "VALUES ('" & idresep & "','" & nomorfaktur & "','" & namapasien & "','" & IDobat & "','" & TextBox7.Text & "','" & TextBox9.Text & "','" & ComboBox2.Text & "','" & hargaobat & "', '" & TotalHarga & "','" & tglKadaluarsa & "', '" & fakturbeli & "')"
        jalankansql(simpan)
        HitungSisa1()
        simpan2 = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "', tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                "WHERE tpembelian.NamaObat = '" & TextBox7.Text & "' AND tpembelian.NomorFaktur = '" & fakturbeli & "'"
        jalankansql(simpan2)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        JumlahObat = Val(TextBox9.Text)
        TotalHarga = JumlahObat * hargaobat
        tambahObat()
        IsiGridDetailResep()
        TextBox7.Text = ""
        TextBox9.Text = ""
        ComboBox2.Text = ""
        Button8.Enabled = True
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        JumlahObat = Val(TextBox9.Text)
        If JumlahObat > Stoksatuan Then
            MsgBox("Stok Obat Tidak Mencukupi", MsgBoxStyle.OkCancel)
            TextBox9.Text = 0
        End If

    End Sub
    Function Ceiling(number As Double) As Long
        Ceiling = -Int(-number)
    End Function
    Sub Konversi()
        selectDataBase = "SELECT TP.KonversiLv2 AS KLV2, TP.KonversiLv3 AS KLV3  FROM tpembelian AS TP WHERE TP.NamaObat ='" & TextBox7.Text & "' AND TP.NomorFaktur='" & fakturbeli & "' "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                konversiLv2 = DT.Rows(i).Item("KLV2")
                konversiLv3 = DT.Rows(i).Item("KLV3")
            Next
        End If
    End Sub
    Sub Konversi2()
        selectDataBase = "SELECT TP.KonversiLv2 AS KLV2, TP.KonversiLv3 AS KLV3  FROM tpembelian AS TP WHERE TP.NamaObat ='" & nameO & "' AND TP.NomorFaktur='" & ket & "' "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                konversiLv2 = DT.Rows(i).Item("KLV2")
                konversiLv3 = DT.Rows(i).Item("KLV3")
            Next
        End If
    End Sub
 
    Sub HitungSisa1()
        sisaobatLv1 = 0
        sisaobatLv2 = 0
        sisaobatLv3 = 0
        Konversi()
        sisobatLv1 = 0
        sisobatLv2 = 0
        sisobatLv3 = 0
        JumlahObat = Val(TextBox9.Text)
        satbeli = ComboBox2.SelectedItem
        selectDataBase = "SELECT TP.SisaObatLv1 AS SOLV1, TP.SatuanLv1 AS SAT1, TP.SisaObatLv2 AS SOLV2, TP.SatuanLv2 AS SAT2," +
                    "TP.SisaObatLv3 AS SOLV3, TP.SatuanLv3 AS SAT3 FROM tpembelian AS TP WHERE TP.NamaObat = '" & TextBox7.Text & "' AND TP.NomorFaktur = '" & fakturbeli & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                'sisaobatLv1 = DT.Rows(i).Item("SOLV1") - sisobatLv1
                'sisaobatLv2 = DT.Rows(i).Item("SOLV2") - sisobatLv2
                'sisaobatLv3 = DT.Rows(i).Item("SOLV3") - sisobatLv3

                If Equals(DT.Rows(i).Item("SAT1"), satbeli) Then
                    sisaobatLv1 = DT.Rows(i).Item("SOLV1") - JumlahObat
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") - (konversiLv2 * JumlahObat)
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") - (konversiLv2 * konversiLv3 * JumlahObat)
                ElseIf Equals(DT.Rows(i).Item("SAT2"), satbeli) Then
                    'sisobatLv2 = JumlahObat
                    'sisobatLv3 = (konversiLv3 * JumlahObat)
                    'sisobatLv1 = Ceiling((JumlahObat / konversiLv2))
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") - JumlahObat
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") - (konversiLv3 * JumlahObat)
                    sisaobatLv1 = Int((sisaobatLv2 / konversiLv2))
                ElseIf Equals(DT.Rows(i).Item("SAT3"), satbeli) Then
                    'sisobatLv3 = JumlahObat
                    'sisobatLv1 = Ceiling((JumlahObat / (konversiLv2 * konversiLv3)))
                    'sisobatLv2 = Ceiling((JumlahObat / konversiLv3))
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") - JumlahObat
                    sisaobatLv1 = Int((sisaobatLv3 / (konversiLv2 * konversiLv3)))
                    sisaobatLv2 = Int((sisaobatLv3 / konversiLv3))

                End If
            Next
        End If
    End Sub

    Sub HitungSisa2()
        sisaobatLv1 = 0
        sisaobatLv2 = 0
        sisaobatLv3 = 0
        Konversi2()
        sisobatLv1 = 0
        sisobatLv2 = 0
        sisobatLv3 = 0
        selectDataBase = "SELECT TP.SisaObatLv1 AS SOLV1, TP.SatuanLv1 AS SAT1, TP.SisaObatLv2 AS SOLV2, TP.SatuanLv2 AS SAT2," +
                    "TP.SisaObatLv3 AS SOLV3, TP.SatuanLv3 AS SAT3 FROM tpembelian AS TP WHERE TP.NamaObat = '" & nameO & "' AND TP.NomorFaktur = '" & ket & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                'sisaobatLv1 = DT.Rows(i).Item("SOLV1") + sisobatLv1
                'sisaobatLv2 = DT.Rows(i).Item("SOLV2") + sisobatLv2
                'sisaobatLv3 = DT.Rows(i).Item("SOLV3") + sisobatLv3
                If Equals(DT.Rows(i).Item("SAT1"), satbel) Then
                    sisaobatLv1 = DT.Rows(i).Item("SOLV1") + Jum
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") + (konversiLv2 * Jum)
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + (konversiLv2 * konversiLv3 * Jum)
                ElseIf Equals(DT.Rows(i).Item("SAT2"), satbel) Then
                    'sisobatLv2 = Jum
                    'sisobatLv3 = (konversiLv3 * Jum)
                    'sisobatLv1 = Ceiling((Jum / konversiLv2))
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") + Jum
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + (konversiLv3 * Jum)
                    sisaobatLv1 = Int((sisaobatLv2 / konversiLv2))
                ElseIf Equals(DT.Rows(i).Item("SAT3"), satbel) Then
                    'sisobatLv3 = Jum
                    'sisobatLv1 = Ceiling((Jum / (konversiLv2 * konversiLv3)))
                    'sisobatLv2 = Ceiling((Jum / konversiLv3))
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + Jum
                    sisaobatLv1 = Int((sisaobatLv3 / (konversiLv2 * konversiLv3)))
                    sisaobatLv2 = Int((sisaobatLv3 / konversiLv3))
                End If

            Next
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        vdate = Format(DateTime.Now, "yyyy-MM-dd")
        If DataGridView2.RowCount > 0 Then
            For x As Integer = 0 To DataGridView2.RowCount - 1
                SubTotal += DataGridView2.Rows(x).Cells(4).Value
            Next
            simpan = "INSERT INTO ttransaksi(Urutan,FakturPenjualan,JenisTransaksi," +
                    "NamaObat,TotalHarga, TglTransaksi)" +
        "VALUES( '" & FormTransaksiPenjualan.Urutan_Faktur & "','" & FormTransaksiPenjualan.TextBox1.Text & "','Resep'," +
        "'" & namapasien & "','" & SubTotal & "','" & vdate & "')"
            jalankansql(simpan)
            FormTransaksiPenjualan.IsiGridUmum1()
            FormTransaksiPenjualan.Button11.Enabled = True
            'FormTransaksiPenjualan.HitungData2()
            'FormTransaksiPenjualan.RichTextBox1.Text = FormTransaksiPenjualan.total.ToString
            Me.Close()
        Else
            MsgBox("Tidak dapat mengetahui Nama Racikan Obat", MsgBoxStyle.OkCancel)
            Button8.Enabled = False
        End If
       
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim hapussql As String
        Dim pesan, simpan2 As String
        If nameO = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")

        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox3.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            HitungSisa2()
            hapussql = "DELETE FROM tdetiltransaksi WHERE  tdetiltransaksi.IDResep ='" & idresep & "' AND  tdetiltransaksi.NamaObat='" & nameO & "' AND tdetiltransaksi.FakturBeliObat='" & ket & "' "
            jalankansql(hapussql)
            simpan2 = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "', tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                "WHERE tpembelian.NamaObat = '" & nameO & "' AND tpembelian.NomorFaktur = '" & ket & "'"
            jalankansql(simpan2)
            DataGridView1.Refresh()
            DataGridView2.Refresh()
            IsiGrid()
            IsiGridDetailResep()
            Button6.Enabled = False
            Button8.Enabled = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "" Then
            TextBox9.Enabled = False
        Else
            TextBox9.Enabled = True
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = "" Then
            Button7.Enabled = False
        Else
            Button7.Enabled = True
        End If
    End Sub
End Class