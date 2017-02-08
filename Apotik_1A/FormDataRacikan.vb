Public Class FormDataRacikan
    Dim databaru As Boolean
    Dim selectDataBase As String
    Dim indexSatuan, indexKategori, indexLokasi, idRacikan, idObat As Integer
    Dim day, month, year, vdate, namaObat, nameO, sb As String
    Dim simpan, simpan2 As String
    Dim JumlahObat, TotalHarga, SubTotal, Keterangan, ket, Jumlah, Jum As Integer
    Dim satuan, satuanlv1, satuanlv2, satuanlv3, satbeli, satbel As String
    Dim stokobatLv1, stokobatLv2, stokobatLv3 As Integer
    Dim sisaobatLv1, sisaobatLv2, sisaobatLv3 As Integer
    Dim sisobatLv1, sisobatLv2, sisobatLv3 As Integer
    Dim temp As Integer

    Shared Property IndeksIDObat As Integer
    Shared Property Stoksatuan As Integer
    Shared Property HargaResep As Integer
    Shared Property Kadaluarsa As String
    Shared Property FakturJual As Integer
    Shared Property Flag As Integer
    Dim konversiLv2 As Integer
    Dim konversiLv3 As Integer

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormDataObats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        'If Flag = 1 Then
        IsiGrid()
        IsiGridDetailRacikan()
        'End If


        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Devcom-Medical-Pill.ico")
    End Sub
    Sub IsiGrid()
        selectDataBase = "SELECT * FROM tracikan WHERE tracikan.NomorFakturPenjualan= '" & FormTransaksiPenjualan.TextBox1.Text & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tracikan")
        DataGridView1.DataSource = (DS.Tables("tracikan"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Racikan"
            .Columns(1).HeaderCell.Value = "Faktur Racikan"
            .Columns(2).HeaderCell.Value = "Nama Racikan"
            .Columns(3).HeaderCell.Value = "Tanggal Pembuatan"
            .Columns(4).HeaderCell.Value = "Nama Pembeli"
            .Columns(5).HeaderCell.Value = "Nama Pasien"
        End With
    End Sub
    Sub Bersih()
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
    End Sub
    Sub ShowKategori()
        selectDataBase = "SELECT Kategori FROM tkategori"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            With ComboBox2
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("Kategori"))
                Next
                .SelectedIndex = -1
            End With
        End If
    End Sub
    
    Sub GetLokasi()
        indexLokasi = -1
        selectDataBase = "SELECT IDLokasi FROM tlokasi WHERE NamaLokasi ='" & ComboBox2.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexLokasi = DT.Rows(0).Item("IDLokasi")
        End If
    End Sub
    Sub GetKategori()
        indexKategori = -1
        selectDataBase = "SELECT IDKategori FROM tkategori WHERE Kategori ='" & ComboBox2.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexKategori = DT.Rows(0).Item("IDKategori")
        End If
    End Sub

    Sub tambahObat()
        Dim simpan As String
        Dim pesan As String
        Dim tmpString As String

        tmpString = Format(DateTime.Now, "yyyy-MM-dd")
        If TextBox2.Text = "" Then Exit Sub

        pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
        If pesan = MsgBoxResult.No Then
            Exit Sub
        End If
        simpan = "INSERT INTO tdetailracikan(IDRacikan, IDObat,IDSatuan,NomorFakturPenjualan, NamaObat,jumlahSatuan,SatuanDosis,HargaJualResep,SubTotal,Kadarluarsa, FakturBeliObat)" +
            "VALUES ('" & TextBox2.Text & "','" & IndeksIDObat & "','" & indexSatuan & "','" & FormTransaksiPenjualan.TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox2.Text & "', '" & HargaResep & "', '" & TotalHarga & "','" & Kadaluarsa & "', '" & FakturJual & "')"
        jalankansql(simpan)
        HitungSisa1()
        simpan2 = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "', tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                "WHERE tpembelian.NamaObat = '" & TextBox3.Text & "' AND tpembelian.NomorFaktur = '" & FakturJual & "'"
        jalankansql(simpan2)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Function Ceiling(number As Double) As Long
        Ceiling = -Int(-number)
    End Function
    Sub Konversi()
        selectDataBase = "SELECT TP.KonversiLv2 AS KLV2, TP.KonversiLv3 AS KLV3  FROM tpembelian AS TP WHERE TP.NamaObat ='" & TextBox3.Text & "' AND TP.NomorFaktur='" & FakturJual & "' "

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
        'SelectSatuanSisaObat()
        Konversi()
        satbeli = ComboBox2.SelectedItem.ToString
        sisaobatLv1 = 0
        sisaobatLv2 = 0
        sisaobatLv3 = 0
        selectDataBase = "SELECT TP.SisaObatLv1 AS SOLV1, TP.SatuanLv1 AS SAT1, TP.SisaObatLv2 AS SOLV2, TP.SatuanLv2 AS SAT2," +
                    "TP.SisaObatLv3 AS SOLV3, TP.SatuanLv3 AS SAT3  FROM tpembelian AS TP WHERE TP.NamaObat = '" & TextBox3.Text & "' AND TP.NomorFaktur = '" & FakturJual & "'"
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
        'SelectSatuanSisaObat2()
        Konversi2()
        sisaobatLv1 = 0
        sisaobatLv2 = 0
        sisaobatLv3 = 0
        selectDataBase = "SELECT TP.SisaObatLv1 AS SOLV1, TP.SatuanLv1 AS SAT1, TP.SisaObatLv2 AS SOLV2, TP.SatuanLv2 AS SAT2," +
                    "TP.SisaObatLv3 AS SOLV3, TP.SatuanLv3 AS SAT3  FROM tpembelian AS TP WHERE TP.NamaObat = '" & nameO & "' AND TP.NomorFaktur = '" & ket & "'"
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        JumlahObat = Val(TextBox4.Text)
        TotalHarga = JumlahObat * HargaResep
        GetSatuan()
        tambahObat()
        Bersih()
        IsiGridDetailRacikan()
        TextBox3.Focus()
        databaru = True
        TextBox3.Enabled = True
        ComboBox2.Enabled = True

    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        'Dim hit As String

        'temp = Val(TextBox5.Text)
        'hit = FormatNumber(temp)
        ''TextBox5.Text = ""
        'TextBox5.Text = hit
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim simpan As String

        Button3.Enabled = True
        TextBox5.Enabled = False
        Button5.Enabled = False
        temp = Val(TextBox5.Text)

        simpan = "INSERT INTO tdetailracikan(NomorFakturPenjualan, NamaObat, SubTotal)" +
            "VALUES ('" & FormTransaksiPenjualan.TextBox1.Text & "',' Biaya Apoteker, '" & temp & "')"
        jalankansql(simpan)


    End Sub
    
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim simpan As String
        Dim pesan As String

        If TextBox3.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        ComboBox2.Enabled = False
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
        vdate = ""
        Bersih()
        isitextbox(e.RowIndex)
        IsiGridDetailRacikan()
        TextBox3.Enabled = True
        ComboBox2.Enabled = True
        Button10.Enabled = True
        Button7.Enabled = True
        Button9.Enabled = False
        TextBox5.Enabled = True
        Button5.Enabled = True
        databaru = False
    End Sub
    Sub IsiGridDetailRacikan()
        DataGridView2.Refresh()
        selectDataBase = "SELECT tracikan.namaRacikan,tobats.namaObat,tdetailracikan.jumlahSatuan,tsatuan.Satuan,tdetailracikan.HargaJualResep,tdetailracikan.SubTotal, tdetailracikan.FakturBeliObat FROM tdetailracikan " +
            "join tracikan " +
                                "on tracikan.IDRacikan= tdetailracikan.IDRacikan " +
            "join tobats " +
                                "on tobats.IDObat= tdetailracikan.IDObat " +
            "join tsatuan " +
                                "on tsatuan.IDSatuan = tdetailracikan.IDSatuan WHERE tdetailracikan.IDRacikan = '" & idRacikan.ToString & "' AND tdetailracikan.NomorFakturPenjualan= '" & FormTransaksiPenjualan.TextBox1.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tracikan")
        DataGridView2.DataSource = (DS.Tables("tracikan"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Racikan"
            .Columns(1).HeaderCell.Value = "Nama Obat "
            .Columns(2).HeaderCell.Value = "Jumlah "
            .Columns(3).HeaderCell.Value = "Satuan"
            .Columns(5).HeaderCell.Value = "Harga Jual"
            .Columns(5).HeaderCell.Value = "SubTotal"
            .Columns(6).HeaderCell.Value = "Keterangan"
        End With
    End Sub
    Sub isitextbox(ByVal x As Integer)
        Try
            TextBox2.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox6.Text = DataGridView1.Rows(x).Cells(2).Value
            idRacikan = DataGridView1.Rows(x).Cells(0).Value
        Catch ex As Exception
        End Try
    End Sub
    Sub isitextbox2(ByVal x As Integer)
        Try
            namaObat = DataGridView2.Rows(x).Cells(1).Value.ToString
            Jumlah = DataGridView2.Rows(x).Cells(2).Value
            sb = DataGridView2.Rows(x).Cells(3).Value.ToString
            Keterangan = DataGridView2.Rows(x).Cells(6).Value

        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        If nameO = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")

        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox3.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            HitungSisa2()
            hapussql = "DELETE FROM tdetailracikan  WHERE  tdetailracikan.IDRacikan ='" & TextBox2.Text & "' AND  tdetailracikan.NamaObat='" & nameO & "' AND tdetailracikan.FakturBeliObat='" & ket & "' "
            jalankansql(hapussql)
            simpan2 = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "', tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                "WHERE tpembelian.NamaObat = '" & nameO & "' AND tpembelian.NomorFaktur = '" & ket & "'"
            jalankansql(simpan2)
            DataGridView1.Refresh()
            DataGridView2.Refresh()
            IsiGrid()
            IsiGridDetailRacikan()
            'Button2.Enabled = True
            Button4.Enabled = False
            Button3.Enabled = True
        End If

    End Sub
    Sub GetSatuan()
        indexSatuan = -1
        selectDataBase = "SELECT IDSatuan FROM tsatuan WHERE Satuan ='" & ComboBox2.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexSatuan = DT.Rows(0).Item("IDSatuan")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tobat WHERE NamaObat LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tsatuan")
            DataGridView1.DataSource = (DS.Tables("tsatuan"))
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        FormDataSatuan.Show()
        FormDataSatuan.Focus()
    End Sub

    Sub ShowSatuan()
        selectDataBase = "SELECT Satuan FROM tsatuan "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            With ComboBox2
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("Satuan"))
                Next
                .SelectedIndex = -1
            End With
            ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
        End If
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs)

        vdate = year + "-" + month + "-" + day
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs)
        GetLokasi()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        FormDataKategori.Show()
        FormDataKategori.Focus()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        FormDataLokasi.Show()
        FormDataLokasi.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        FormDaftarRacikan.Show()
        FormDaftarRacikan.Focus()
        Button9.Enabled = False
    End Sub

    Private Sub FormDataRacikan_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Dim hapussql As String
        Dim pesan As String
        If TextBox2.Text = "" Or TextBox6.Text = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        ElseIf DataGridView2.Rows.Count > 0 Then
            pesan = MsgBox("Hapus Dulu data Obat ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server?" + TextBox6.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If

            hapussql = "DELETE FROM tracikan WHERE IDRacikan='" & TextBox2.Text & "'"
            jalankansql(hapussql)
            DataGridView1.Refresh()
            IsiGrid()
            TextBox2.Text = ""
            TextBox6.Text = ""
            If DataGridView1.CurrentCell Is Nothing Then
                Button7.Enabled = False
                Button10.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                Button2.Enabled = False
                ComboBox2.Enabled = False
            Else
                Button7.Enabled = True
                Button10.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                Button2.Enabled = True
                ComboBox2.Enabled = True
            End If
        End If
        Button9.Enabled = True
       
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
        Button4.Enabled = False
        Button7.Enabled = False
        FormPencarianStokObat.IDPencariObat = 2
        FormPencarianStokObat.Show()
        FormPencarianStokObat.Focus()
        If TextBox2.Text = "" Then
            MsgBox("Tidak dapat mengetahui Nama Racikan Obat", MsgBoxStyle.OkCancel)
            FormPencarianStokObat.Close()
        End If
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        vdate = Format(DateTime.Now, "yyyy-MM-dd")
        If DataGridView2.RowCount > 0 Then
            For x As Integer = 0 To DataGridView2.RowCount - 1
                SubTotal += DataGridView2.Rows(x).Cells(5).Value
            Next
            simpan = "INSERT INTO ttransaksi(Urutan,FakturPenjualan,JenisTransaksi," +
                    "NamaObat,TotalHarga, TglTransaksi)" +
        "VALUES( '" & FormTransaksiPenjualan.Urutan_Faktur & "','" & FormTransaksiPenjualan.TextBox1.Text & "','Racikan'," +
        "'" & TextBox6.Text & "','" & SubTotal & "','" & vdate & "')"
            jalankansql(simpan)
            FormTransaksiPenjualan.IsiGridUmum1()
            FormTransaksiPenjualan.Button11.Enabled = True
            Me.Close()
        Else
            MsgBox("Tidak dapat mengetahui Nama Racikan Obat", MsgBoxStyle.OkCancel)
            Button3.Enabled = False
        End If
        
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        isitextbox2(e.RowIndex)
        nameO = namaObat
        Jum = Jumlah
        satbel = sb
        ket = Keterangan
        Button4.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        JumlahObat = Val(TextBox4.Text)
        TotalHarga = JumlahObat * HargaResep
        If JumlahObat > Stoksatuan Then
            MsgBox("Stok Obat Tidak Mencukupi", MsgBoxStyle.OkCancel)
            TextBox4.Text = 0
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "" Then
            TextBox4.Enabled = False
        Else
            TextBox4.Enabled = True
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

   
    
End Class