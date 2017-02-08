Public Class FormPencarianStokObat
    Dim databaru As Boolean
    Dim indeksObat As String
    Dim indeksSupplier As String
    Dim dataSelected, dataSelected2 As Boolean
    Dim namaObat, nO As String
    Dim tglKadaluarsa As String
    Dim idSatuan As Integer
    Dim namaSatuan As String
    Dim selectDataBase As String
    Dim HargaJualUmum1, HargaJualUmum2, HargaJualUmum3 As Integer
    Dim HargaJualResep As Integer
    Dim DiskonJualSatuan1, DiskonJualSatuan2, DiskonJualSatuan3 As Integer
    Dim DiskonJualUmum1, DiskonJualUmum2, DiskonJualUmum3 As Integer
    Dim DiskonJualResep As Integer
    Dim SisaObatLv1, SisaObatLv2, SisaObatLv3 As Integer
    Dim SatLv1, SatLv2, SatLv3 As String
    Dim nowTime As String
    Dim kadaluarsa As Date
    Dim FakturBeli, Fb As Integer
    Dim Satuanracikan, HargaResep As Integer
    Dim pesan As String
    Shared Property IDPencariObat As Integer


    Private Sub FormPencarianStokObat_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        dataSelected2 = False

        IsiGrid1()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
    End Sub

    Sub IsiGrid1()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase = "SELECT tpembelian.IDObat, tpembelian.NamaObat, tpembelian.HargaJualUmum1, tpembelian.SatDisLv1, tpembelian.HargaDisLv1," +
            " tpembelian.HargaJualUmum2, tpembelian.SatDisLv2, tpembelian.HargaDisLv2," +
            " tpembelian.HargaJualUmum3, tpembelian.SatDisLv3, tpembelian.HargaDisLv3, tpembelian.HargaJualResep" +
            " FROM tpembelian WHERE (tpembelian.TglKadaluarsa > '" & nowTime & "') AND ((tpembelian.SisaObatLv1 >'" & 0 & "') OR (tpembelian.SisaObatLv2 >'" & 0 & "') OR (tpembelian.SisaObatLv3 >'" & 0 & "')) GROUP BY tpembelian.NamaObat ASC LIMIT 0,50 "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Harga Jual Umum Lv1"
            .Columns(3).HeaderCell.Value = "Minimum Pembelian Lv1"
            .Columns(4).HeaderCell.Value = "Diskon Pembelian Lv1"
            .Columns(5).HeaderCell.Value = "Harga Jual Umum Lv2"
            .Columns(6).HeaderCell.Value = "Minimum Pembelian Lv2"
            .Columns(7).HeaderCell.Value = "Diskon Pembelian Lv2"
            .Columns(8).HeaderCell.Value = "Harga Jual Umum Lv3"
            .Columns(9).HeaderCell.Value = "Minimum Pembelian Lv3"
            .Columns(10).HeaderCell.Value = "Diskon Pembelian Lv3"
            .Columns(11).HeaderCell.Value = "Harga Jual Resep"

        End With
    End Sub

    Sub IsiGrid2()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase = "SELECT tpembelian.TglKadaluarsa, tpembelian.SisaObatLv1, tpembelian.SatuanLv1 , tpembelian.SisaObatLv2,tpembelian.SatuanLv2," +
            "tpembelian.SisaObatLv3, tpembelian.SatuanLv3, " +
            "tpembelian.NomorFaktur FROM tpembelian " +
            "WHERE (tpembelian.NamaObat ='" & nO & "') AND (tpembelian.TglKadaluarsa > '" & nowTime & "') AND ((tpembelian.SisaObatLv1 >'" & 0 & "') OR (tpembelian.SisaObatLv2 >'" & 0 & "') OR (tpembelian.SisaObatLv3 >'" & 0 & "')) ORDER BY tpembelian.TglKadaluarsa ASC LIMIT 0,50 "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView2.DataSource = (DS.Tables("tpembelian"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Tanggal Kadaluarsa"
            .Columns(1).HeaderCell.Value = "Sisa Obat LV 1"
            .Columns(2).HeaderCell.Value = "Satuan Obat LV 1"
            .Columns(3).HeaderCell.Value = "Sisa Obat LV 2"
            .Columns(4).HeaderCell.Value = "Satuan Obat LV 2"
            .Columns(5).HeaderCell.Value = "Sisa Obat LV 3"
            .Columns(6).HeaderCell.Value = "Satuan Obat LV 3"
            .Columns(7).HeaderCell.Value = "Faktur Pembelian"
        End With
    End Sub

    Sub SelectSatuan()
        selectDataBase = "SELECT t1.Satuan AS Level1, t2.Satuan AS Level2, t3.Satuan AS Level3 FROM tobats AS T " +
                            " inner join tsatuan AS t1 " +
                                "on t1.IDSatuan = T.IDSatuan" +
                            " inner join tsatuan AS t2 " +
                                "on t2.IDSatuan = T.IDSatuanLv2" +
                            " inner join tsatuan AS t3 " +
                                "on t3.IDSatuan = T.IDSatuanLv3 WHERE T.NamaObat ='" & nO & "' "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            If IDPencariObat = 1 Then
                With FormDataResep.ComboBox2
                    .Items.Clear()
                    For i As Integer = 0 To DT.Rows.Count - 1
                        If Equals(DT.Rows(i).Item("Level3"), "Tidak Terdeteksi") Then
                            If Equals(DT.Rows(i).Item("Level2"), "Tidak Terdeteksi") Then
                                FormDataRacikan.ComboBox2.Refresh()
                                .Items.Clear()
                                .Items.Add(DT.Rows(i).Item("Level1"))
                            Else
                                .Items.Clear()
                                FormDataRacikan.ComboBox2.Refresh()
                                .Items.Add(DT.Rows(i).Item("Level2"))
                            End If
                        Else
                            .Items.Clear()
                            FormDataRacikan.ComboBox2.Refresh()
                            .Items.Add(DT.Rows(i).Item("Level3"))
                        End If
                    Next
                    .SelectedIndex = -1
                End With
            End If
            If IDPencariObat = 2 Then
                With FormDataRacikan.ComboBox2
                    .Items.Clear()
                    For i As Integer = 0 To DT.Rows.Count - 1
                        If Equals(DT.Rows(i).Item("Level3"), "Tidak Terdeteksi") Then
                            If Equals(DT.Rows(i).Item("Level2"), "Tidak Terdeteksi") Then
                                FormDataRacikan.ComboBox2.Refresh()
                                .Items.Clear()
                                .Items.Add(DT.Rows(i).Item("Level1"))
                            Else
                                .Items.Clear()
                                FormDataRacikan.ComboBox2.Refresh()
                                .Items.Add(DT.Rows(i).Item("Level2"))
                            End If
                        Else
                            .Items.Clear()
                            FormDataRacikan.ComboBox2.Refresh()
                            .Items.Add(DT.Rows(i).Item("Level3"))
                        End If
                    Next
                    .SelectedIndex = -1
                End With

            ElseIf IDPencariObat = 3 Then

                With FormTransaksiPenjualan.ComboBox1
                    .Items.Clear()
                    For i As Integer = 0 To DT.Rows.Count - 1
                        If Equals(DT.Rows(i).Item("Level1"), "Tidak Terdeteksi") Then
                            If Equals(DT.Rows(i).Item("Level2"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level3"))
                            ElseIf Equals(DT.Rows(i).Item("Level3"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level2"))
                            Else
                                .Items.Add(DT.Rows(i).Item("Level2"))
                                .Items.Add(DT.Rows(i).Item("Level3"))
                            End If

                        ElseIf Equals(DT.Rows(i).Item("Level2"), "Tidak Terdeteksi") Then
                            If Equals(DT.Rows(i).Item("Level1"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level3"))
                            ElseIf Equals(DT.Rows(i).Item("Level3"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level1"))
                            Else
                                .Items.Add(DT.Rows(i).Item("Level1"))
                                .Items.Add(DT.Rows(i).Item("Level3"))
                            End If

                        ElseIf Equals(DT.Rows(i).Item("Level3"), "Tidak Terdeteksi") Then
                            If Equals(DT.Rows(i).Item("Level2"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level1"))
                            ElseIf Equals(DT.Rows(i).Item("Level1"), "Tidak Terdeteksi") Then
                                .Items.Add(DT.Rows(i).Item("Level2"))
                            Else
                                .Items.Add(DT.Rows(i).Item("Level1"))
                                .Items.Add(DT.Rows(i).Item("Level2"))
                            End If

                        Else
                            .Items.Add(DT.Rows(i).Item("Level1"))
                            .Items.Add(DT.Rows(i).Item("Level2"))
                            .Items.Add(DT.Rows(i).Item("Level3"))
                        End If
                    Next
                    .SelectedIndex = -1
                End With
            End If
        End If
    End Sub

    Sub Konversi()
        selectDataBase = "SELECT TP.KonversiLv2 AS KLV2, TP.KonversiLv3 AS KLV3  FROM tpembelian AS TP WHERE TP.NamaObat ='" & nO & "' AND TP.NomorFaktur='" & Fb & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
                For i As Integer = 0 To DT.Rows.Count - 1
                    FormTransaksiPenjualan.konversiLv2 = DT.Rows(i).Item("KLV2")
                    FormTransaksiPenjualan.konversiLv3 = DT.Rows(i).Item("KLV3")
                Next

        End If
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            indeksObat = DataGridView1.Rows(x).Cells(0).Value.ToString
            namaObat = DataGridView1.Rows(x).Cells(1).Value.ToString
            HargaJualUmum1 = DataGridView1.Rows(x).Cells(2).Value
            DiskonJualSatuan1 = DataGridView1.Rows(x).Cells(3).Value
            DiskonJualUmum1 = DataGridView1.Rows(x).Cells(4).Value
            HargaJualUmum2 = DataGridView1.Rows(x).Cells(5).Value
            DiskonJualSatuan2 = DataGridView1.Rows(x).Cells(6).Value
            DiskonJualUmum2 = DataGridView1.Rows(x).Cells(7).Value
            HargaJualUmum3 = DataGridView1.Rows(x).Cells(8).Value
            DiskonJualSatuan3 = DataGridView1.Rows(x).Cells(9).Value
            DiskonJualUmum3 = DataGridView1.Rows(x).Cells(10).Value
            HargaJualResep = DataGridView1.Rows(x).Cells(11).Value
        Catch ex As Exception
        End Try
    End Sub

    Sub GetIndeks2(ByVal x As Integer)
        Try
            tglKadaluarsa = DataGridView2.Rows(x).Cells(0).Value.ToString
            SisaObatLv1 = DataGridView2.Rows(x).Cells(1).Value
            SatLv1 = DataGridView2.Rows(x).Cells(2).Value.ToString
            SisaObatLv2 = DataGridView2.Rows(x).Cells(3).Value
            SatLv2 = DataGridView2.Rows(x).Cells(4).Value.ToString
            SisaObatLv3 = DataGridView2.Rows(x).Cells(5).Value
            SatLv3 = DataGridView2.Rows(x).Cells(6).Value.ToString
            FakturBeli = DataGridView2.Rows(x).Cells(7).Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
        nO = namaObat
        IsiGrid2()
        SelectSatuan()
    End Sub

    Private Sub DataGridView2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseClick
        GetIndeks2(e.RowIndex)
        Fb = FakturBeli
        dataSelected = True
        dataSelected2 = True
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Fb = 0 Then
            pesan = MsgBox("Pilih Dulu Data Obat ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            If dataSelected = True And dataSelected2 = True Then
                If IDPencariObat = 1 Then
                    FormDataResep.IDobat = indeksObat
                    FormDataResep.TextBox7.Text = namaObat
                    FormDataResep.hargaobat = HargaJualResep
                    FormDataResep.tglKadaluarsa = tglKadaluarsa
                    FormDataResep.fakturbeli = FakturBeli

                    If Equals(SatLv3, "Tidak Terdeteksi") Then
                        If Equals(SatLv2, "Tidak Terdeteksi") Then
                            FormDataResep.Stoksatuan = SisaObatLv1
                        Else
                            FormDataResep.Stoksatuan = SisaObatLv2
                        End If
                    Else
                        FormDataResep.Stoksatuan = SisaObatLv3
                    End If

                ElseIf IDPencariObat = 2 Then
                    FormDataRacikan.TextBox3.Text = namaObat
                    FormDataRacikan.HargaResep = HargaJualResep
                    FormDataRacikan.IndeksIDObat = indeksObat
                    FormDataRacikan.FakturJual = FakturBeli
                    FormDataRacikan.Kadaluarsa = tglKadaluarsa

                    If Equals(SatLv3, "Tidak Terdeteksi") Then
                        If Equals(SatLv2, "Tidak Terdeteksi") Then
                            FormDataRacikan.Stoksatuan = SisaObatLv1
                        Else
                            FormDataRacikan.Stoksatuan = SisaObatLv2
                        End If
                    Else
                        FormDataRacikan.Stoksatuan = SisaObatLv3
                    End If
                ElseIf IDPencariObat = 3 Then
                    FormTransaksiPenjualan.indexObat = indeksObat
                    FormTransaksiPenjualan.TextBox22.Text = indeksObat
                    FormTransaksiPenjualan.indexSupplier = indeksSupplier
                    FormTransaksiPenjualan.TextBox2.Text = namaObat
                    FormTransaksiPenjualan.HJU1 = HargaJualUmum1
                    FormTransaksiPenjualan.MinimumPembelianlv1 = DiskonJualSatuan1
                    FormTransaksiPenjualan.DiskonPembelianlv1 = DiskonJualUmum1
                    FormTransaksiPenjualan.HJU2 = HargaJualUmum2
                    FormTransaksiPenjualan.MinimumPembelianlv2 = DiskonJualSatuan2
                    FormTransaksiPenjualan.DiskonPembelianlv2 = DiskonJualUmum2
                    FormTransaksiPenjualan.HJU3 = HargaJualUmum3
                    FormTransaksiPenjualan.MinimumPembelianlv3 = DiskonJualSatuan3
                    FormTransaksiPenjualan.DiskonPembelianlv3 = DiskonJualUmum3
                    FormTransaksiPenjualan.TextBox9.Text = SisaObatLv1
                    FormTransaksiPenjualan.TextBox14.Text = SatLv1
                    FormTransaksiPenjualan.TextBox16.Text = SisaObatLv2
                    FormTransaksiPenjualan.TextBox15.Text = SatLv2
                    FormTransaksiPenjualan.TextBox21.Text = SisaObatLv3
                    FormTransaksiPenjualan.TextBox17.Text = SatLv3
                    FormTransaksiPenjualan.FakturJual = FakturBeli
                    FormTransaksiPenjualan.Kadaluarsa = tglKadaluarsa
                    Dim tgl As Date
                    tgl = Convert.ToDateTime(tglKadaluarsa)
                    FormTransaksiPenjualan.DateTimePicker2.Format = DateTimePickerFormat.Custom
                    FormTransaksiPenjualan.DateTimePicker2.CustomFormat = tgl.Day.ToString + "/" + tgl.Month.ToString + "/" + tgl.Year.ToString
                    GetNamaSatuan()
                    Konversi()
                End If
            End If
            Me.Close()
        End If
    End Sub

    Private Sub DataGridView2_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseDoubleClick
        If Fb = 0 Then
            pesan = MsgBox("Pilih Dulu Data Obat ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            If dataSelected = True And dataSelected2 = True Then
                If IDPencariObat = 1 Then
                    FormDataResep.IDobat = indeksObat
                    FormDataResep.TextBox7.Text = namaObat
                    FormDataResep.hargaobat = HargaJualResep
                    FormDataResep.tglKadaluarsa = tglKadaluarsa
                    FormDataResep.fakturbeli = FakturBeli

                    If Equals(SatLv3, "Tidak Terdeteksi") Then
                        If Equals(SatLv2, "Tidak Terdeteksi") Then
                            FormDataResep.Stoksatuan = SisaObatLv1
                        Else
                            FormDataResep.Stoksatuan = SisaObatLv2
                        End If
                    Else
                        FormDataResep.Stoksatuan = SisaObatLv3
                    End If

                ElseIf IDPencariObat = 2 Then
                    FormDataRacikan.TextBox3.Text = namaObat
                    'SelectSatuan()
                    FormDataRacikan.HargaResep = HargaJualResep
                    FormDataRacikan.IndeksIDObat = indeksObat
                    FormDataRacikan.FakturJual = FakturBeli
                    FormDataRacikan.Kadaluarsa = tglKadaluarsa
                    If Equals(SatLv3, "Tidak Terdeteksi") Then
                        If Equals(SatLv2, "Tidak Terdeteksi") Then
                            FormDataRacikan.Stoksatuan = SisaObatLv1
                        Else
                            FormDataRacikan.Stoksatuan = SisaObatLv2
                        End If
                    Else
                        FormDataRacikan.Stoksatuan = SisaObatLv3
                    End If
                ElseIf IDPencariObat = 3 Then
                    FormTransaksiPenjualan.indexObat = indeksObat
                    FormTransaksiPenjualan.TextBox22.Text = indeksObat
                    FormTransaksiPenjualan.indexSupplier = indeksSupplier
                    FormTransaksiPenjualan.TextBox2.Text = namaObat
                    FormTransaksiPenjualan.HJU1 = HargaJualUmum1
                    FormTransaksiPenjualan.MinimumPembelianlv1 = DiskonJualSatuan1
                    FormTransaksiPenjualan.DiskonPembelianlv1 = DiskonJualUmum1
                    FormTransaksiPenjualan.HJU2 = HargaJualUmum2
                    FormTransaksiPenjualan.MinimumPembelianlv2 = DiskonJualSatuan2
                    FormTransaksiPenjualan.DiskonPembelianlv2 = DiskonJualUmum2
                    FormTransaksiPenjualan.HJU3 = HargaJualUmum3
                    FormTransaksiPenjualan.MinimumPembelianlv3 = DiskonJualSatuan3
                    FormTransaksiPenjualan.DiskonPembelianlv3 = DiskonJualUmum3
                    FormTransaksiPenjualan.TextBox9.Text = SisaObatLv1
                    FormTransaksiPenjualan.TextBox14.Text = SatLv1
                    FormTransaksiPenjualan.TextBox16.Text = SisaObatLv2
                    FormTransaksiPenjualan.TextBox15.Text = SatLv2
                    FormTransaksiPenjualan.TextBox21.Text = SisaObatLv3
                    FormTransaksiPenjualan.TextBox17.Text = SatLv3
                    FormTransaksiPenjualan.FakturJual = FakturBeli
                    FormTransaksiPenjualan.Kadaluarsa = tglKadaluarsa
                    Dim tgl As Date
                    tgl = Convert.ToDateTime(tglKadaluarsa)
                    FormTransaksiPenjualan.DateTimePicker2.Format = DateTimePickerFormat.Custom
                    FormTransaksiPenjualan.DateTimePicker2.CustomFormat = tgl.Day.ToString + "/" + tgl.Month.ToString + "/" + tgl.Year.ToString
                    GetNamaSatuan()
                    Konversi()
                End If
            End If
            Me.Close()
        End If
    End Sub
    Sub GetNamaSatuan()
        namaSatuan = ""
        selectDataBase = "SELECT Satuan FROM tsatuan WHERE IDSatuan ='" & idSatuan & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            namaSatuan = DT.Rows(0).Item("Satuan")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid1()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT tpembelian.IDObat, tpembelian.NamaObat, tpembelian.HargaJualUmum1, tpembelian.SatDisLv1, tpembelian.HargaDisLv1," +
            " tpembelian.HargaJualUmum2, tpembelian.SatDisLv2, tpembelian.HargaDisLv2," +
            " tpembelian.HargaJualUmum3, tpembelian.SatDisLv3, tpembelian.HargaDisLv3, tpembelian.HargaJualResep" +
            " FROM tpembelian WHERE (tpembelian.NamaObat LIKE '%" & TextBox1.Text & "%') AND (tpembelian.TglKadaluarsa > '" & nowTime & "' ) AND ((tpembelian.SisaObatLv1 >'" & 0 & "') OR (tpembelian.SisaObatLv2 >'" & 0 & "') OR (tpembelian.SisaObatLv3 >'" & 0 & "')) GROUP BY tpembelian.NamaObat ASC LIMIT 0,50 ", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tpembelian")
            DataGridView1.DataSource = (DS.Tables("tpembelian"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID Obat"
                .Columns(1).HeaderCell.Value = "Nama Obat"
                .Columns(2).HeaderCell.Value = "Harga Jual Umum Lv1"
                .Columns(3).HeaderCell.Value = "Minimum Pembelian Lv1"
                .Columns(4).HeaderCell.Value = "Diskon Pembelian Lv1"
                .Columns(5).HeaderCell.Value = "Harga Jual Umum Lv2"
                .Columns(6).HeaderCell.Value = "Minimum Pembelian Lv2"
                .Columns(7).HeaderCell.Value = "Diskon Pembelian Lv2"
                .Columns(8).HeaderCell.Value = "Harga Jual Umum Lv3"
                .Columns(9).HeaderCell.Value = "Minimum Pembelian Lv3"
                .Columns(10).HeaderCell.Value = "Diskon Pembelian Lv3"
                .Columns(11).HeaderCell.Value = "Harga Jual Resep"

            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click

    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown

    End Sub

    Private Sub DataGridView1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
        GetIndeks(e.RowIndex)
        dataSelected = True
        nO = namaObat
        IsiGrid2()
        SelectSatuan()
    End Sub
End Class