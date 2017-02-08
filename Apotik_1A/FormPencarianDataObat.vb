Public Class FormPencarianDataObat
    Dim databaru As Boolean
    Dim indeksObat As String
    Dim dataSelected As Boolean
    Dim namaObat As String
    Dim tglKadaluarsa As Date
    Dim idSatuan, pesan As String
    Dim namaSatuan As String
    Dim selectDataBase As String
    Dim HargaJualUmum As Integer
    Dim HargaJualResep As Integer
    Dim stokObat As Integer
    Dim Satuanlv1, Satuanlv2, Satuanlv3 As String
    Shared Property IDPencariObat As Integer
    'Dim StockRemain As Integer


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormPencarianDataObat_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        IsiGrid()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img

        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
    End Sub

    Sub IsiGrid()
        bukaDB()
        selectDataBase = "SELECT T.IDObat, T.NamaObat, t1.Satuan AS Level1, t2.Satuan AS Level2, t3.Satuan AS Level3, tk.Kategori AS Kategori, tl.NamaLokasi AS Lokasi , T.StokMinimal FROM tobats AS T " +
                            " inner join tsatuan AS t1 " +
                                "on t1.IDSatuan = T.IDSatuan" +
                            " inner join tsatuan AS t2 " +
                                "on t2.IDSatuan = T.IDSatuanLv2" +
                            " inner join tsatuan AS t3 " +
                                "on t3.IDSatuan = T.IDSatuanLv3" +
                            " join tkategori AS tk " +
                                "on tk.IDKategori = T.IDKategori" +
                            " join tlokasi AS tl " +
                                "on tl.IDLokasi = T.IDLokasi ORDER BY IDObat ASC LIMIT 0,30"

        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tobats")
        DataGridView1.DataSource = (DS.Tables("tobats"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Satuan Lv 1"
            .Columns(3).HeaderCell.Value = "Satuan Lv2"
            .Columns(4).HeaderCell.Value = "Satuan Lv 3"
            .Columns(5).HeaderCell.Value = "Kategori Obat"
            .Columns(6).HeaderCell.Value = "Lokasi Obat"
            .Columns(7).HeaderCell.Value = "Stok Minimal"
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormDataObats.Show()
        FormDataObats.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If namaObat = "" Then
            pesan = MsgBox("Pilih Dulu Data Obat ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")

        Else
            selectObat()
        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        
    End Sub
    Sub SelectSatuan()
        selectDataBase = "SELECT t1.Satuan AS Level1, t2.Satuan AS Level2, t3.Satuan AS Level3 FROM tobats AS T " +
                            " inner join tsatuan AS t1 " +
                                "on t1.IDSatuan = T.IDSatuan" +
                            " inner join tsatuan AS t2 " +
                                "on t2.IDSatuan = T.IDSatuanLv2" +
                            " inner join tsatuan AS t3 " +
                                "on t3.IDSatuan = T.IDSatuanLv3 WHERE T.NamaObat ='" & namaObat & "' "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            With FormDataRacikan.ComboBox2
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                   
                    .Items.Add(DT.Rows(i).Item("Level1"))
                    .Items.Add(DT.Rows(i).Item("Level2"))
                    .Items.Add(DT.Rows(i).Item("Level3"))
                Next
                .SelectedIndex = DT.Rows.Count - 1
            End With
        End If
    End Sub
    Sub selectObat()
        If dataSelected = True Then

            'If IDPencariObat = 2 Then

            '    FormDataRacikan.TextBox3.Text = namaObat
            '    FormTransaksiPembelian.indexObat = indeksObat
            '    SelectSatuan()
            'Else
            FormTransaksiPembelian.TextBox2.Text = namaObat
            FormTransaksiPembelian.TextBox3.Text = indeksObat
            FormTransaksiPembelian.DateTimePicker2.CustomFormat = "MM/dd/yyyy"
            FormTransaksiPembelian.indexObat = indeksObat
            FormTransaksiPembelian.stokObat = stokObat
            FormTransaksiPembelian.NamaObat = namaObat
            FormTransaksiPembelian.TextBox10.Text = Satuanlv1
            FormTransaksiPembelian.TextBox9.Text = Satuanlv2
            FormTransaksiPembelian.TextBox11.Text = Satuanlv3

            FormTransaksiPembelian.TextBox42.Text = Satuanlv1
            FormTransaksiPembelian.TextBox44.Text = Satuanlv2
            FormTransaksiPembelian.TextBox45.Text = Satuanlv3

            'FormTransaksiPembelian.RemainStock = StockRemain
            'FormTransaksiPembelian.TextBox17.Text = StockRemain

            GetNamaSatuan()
            If Equals(Satuanlv3, "Tidak Terdeteksi") Then

                FormTransaksiPembelian.TextBox26.Enabled = False
                FormTransaksiPembelian.TextBox39.Enabled = False
                FormTransaksiPembelian.TextBox40.Enabled = False
                FormTransaksiPembelian.TextBox43.Enabled = False
                FormTransaksiPembelian.FLagdata = 3

                If Equals(Satuanlv2, "Tidak Terdeteksi") Then

                    FormTransaksiPembelian.TextBox25.Enabled = False
                    FormTransaksiPembelian.TextBox33.Enabled = False
                    FormTransaksiPembelian.TextBox37.Enabled = False
                    FormTransaksiPembelian.TextBox38.Enabled = False

                    FormTransaksiPembelian.TextBox26.Enabled = False
                    FormTransaksiPembelian.TextBox39.Enabled = False
                    FormTransaksiPembelian.TextBox40.Enabled = False
                    FormTransaksiPembelian.TextBox43.Enabled = False
                    FormTransaksiPembelian.FLagdata = 2

                    If Equals(Satuanlv1, "Tidak Terdeteksi") Then
                        FormTransaksiPembelian.TextBox5.Enabled = False
                        FormTransaksiPembelian.TextBox31.Enabled = False
                        FormTransaksiPembelian.TextBox35.Enabled = False
                        FormTransaksiPembelian.TextBox36.Enabled = False

                        FormTransaksiPembelian.TextBox25.Enabled = False
                        FormTransaksiPembelian.TextBox33.Enabled = False
                        FormTransaksiPembelian.TextBox37.Enabled = False
                        FormTransaksiPembelian.TextBox38.Enabled = False

                        FormTransaksiPembelian.TextBox26.Enabled = False
                        FormTransaksiPembelian.TextBox39.Enabled = False
                        FormTransaksiPembelian.TextBox40.Enabled = False
                        FormTransaksiPembelian.TextBox43.Enabled = False

                        FormTransaksiPembelian.TextBox14.Enabled = False
                        FormTransaksiPembelian.TextBox6.Enabled = False
                        FormTransaksiPembelian.TextBox7.Enabled = False
                        FormTransaksiPembelian.Button7.Enabled = False
                        FormTransaksiPembelian.FLagdata = 1
                    Else

                        FormTransaksiPembelian.TextBox25.Enabled = False
                        FormTransaksiPembelian.TextBox33.Enabled = False
                        FormTransaksiPembelian.TextBox37.Enabled = False
                        FormTransaksiPembelian.TextBox38.Enabled = False

                        FormTransaksiPembelian.TextBox26.Enabled = False
                        FormTransaksiPembelian.TextBox39.Enabled = False
                        FormTransaksiPembelian.TextBox40.Enabled = False
                        FormTransaksiPembelian.TextBox43.Enabled = False
                    End If
                Else
                    FormTransaksiPembelian.TextBox26.Enabled = False
                    FormTransaksiPembelian.TextBox39.Enabled = False
                    FormTransaksiPembelian.TextBox40.Enabled = False
                    FormTransaksiPembelian.TextBox43.Enabled = False
                End If
                'TextBox1.SelectedIndex = FormTransaksiPembelian.ComboBox1.FindStringExact(namaSatuan)
            Else
                FormTransaksiPembelian.FLagdata = 4
            End If

            Me.Close()
        End If

    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            indeksObat = DataGridView1.Rows(x).Cells(0).Value.ToString
            namaObat = DataGridView1.Rows(x).Cells(1).Value
            Satuanlv1 = DataGridView1.Rows(x).Cells(2).Value
            Satuanlv2 = DataGridView1.Rows(x).Cells(3).Value
            Satuanlv3 = DataGridView1.Rows(x).Cells(4).Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        selectObat()
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
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
    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick

        FormTransaksiPembelian.TextBox2.Text = namaObat
        FormTransaksiPembelian.TextBox3.Text = indeksObat
        FormTransaksiPembelian.indexObat = indeksObat
        FormTransaksiPembelian.stokObat = stokObat
        FormTransaksiPembelian.NamaObat = namaObat
        FormTransaksiPembelian.TextBox10.Text = Satuanlv1
        FormTransaksiPembelian.TextBox9.Text = Satuanlv2
        FormTransaksiPembelian.TextBox11.Text = Satuanlv3

        FormTransaksiPembelian.TextBox42.Text = Satuanlv1
        FormTransaksiPembelian.TextBox44.Text = Satuanlv2
        FormTransaksiPembelian.TextBox45.Text = Satuanlv3

        'FormTransaksiPembelian.RemainStock = StockRemain
        'FormTransaksiPembelian.TextBox17.Text = StockRemain

        GetNamaSatuan()
        If Equals(Satuanlv3, "Tidak Terdeteksi") Then

            FormTransaksiPembelian.TextBox26.Enabled = False
            FormTransaksiPembelian.TextBox39.Enabled = False
            FormTransaksiPembelian.TextBox40.Enabled = False
            FormTransaksiPembelian.TextBox43.Enabled = False
            FormTransaksiPembelian.FLagdata = 3


            If Equals(Satuanlv2, "Tidak Terdeteksi") Then

                FormTransaksiPembelian.TextBox25.Enabled = False
                FormTransaksiPembelian.TextBox33.Enabled = False
                FormTransaksiPembelian.TextBox37.Enabled = False
                FormTransaksiPembelian.TextBox38.Enabled = False

                FormTransaksiPembelian.TextBox26.Enabled = False
                FormTransaksiPembelian.TextBox39.Enabled = False
                FormTransaksiPembelian.TextBox40.Enabled = False
                FormTransaksiPembelian.TextBox43.Enabled = False
                FormTransaksiPembelian.FLagdata = 2

                If Equals(Satuanlv1, "Tidak Terdeteksi") Then
                    FormTransaksiPembelian.TextBox5.Enabled = False
                    FormTransaksiPembelian.TextBox31.Enabled = False
                    FormTransaksiPembelian.TextBox35.Enabled = False
                    FormTransaksiPembelian.TextBox36.Enabled = False

                    FormTransaksiPembelian.TextBox25.Enabled = False
                    FormTransaksiPembelian.TextBox33.Enabled = False
                    FormTransaksiPembelian.TextBox37.Enabled = False
                    FormTransaksiPembelian.TextBox38.Enabled = False

                    FormTransaksiPembelian.TextBox26.Enabled = False
                    FormTransaksiPembelian.TextBox39.Enabled = False
                    FormTransaksiPembelian.TextBox40.Enabled = False
                    FormTransaksiPembelian.TextBox43.Enabled = False

                    FormTransaksiPembelian.TextBox14.Enabled = False
                    FormTransaksiPembelian.TextBox6.Enabled = False
                    FormTransaksiPembelian.TextBox7.Enabled = False
                    FormTransaksiPembelian.Button7.Enabled = False
                    FormTransaksiPembelian.FLagdata = 1

                Else

                    FormTransaksiPembelian.TextBox25.Enabled = False
                    FormTransaksiPembelian.TextBox33.Enabled = False
                    FormTransaksiPembelian.TextBox37.Enabled = False
                    FormTransaksiPembelian.TextBox38.Enabled = False

                    FormTransaksiPembelian.TextBox26.Enabled = False
                    FormTransaksiPembelian.TextBox39.Enabled = False
                    FormTransaksiPembelian.TextBox40.Enabled = False
                    FormTransaksiPembelian.TextBox43.Enabled = False
                End If
            Else
                FormTransaksiPembelian.TextBox26.Enabled = False
                FormTransaksiPembelian.TextBox39.Enabled = False
                FormTransaksiPembelian.TextBox40.Enabled = False
                FormTransaksiPembelian.TextBox43.Enabled = False
            End If
            'TextBox1.SelectedIndex = FormTransaksiPembelian.ComboBox1.FindStringExact(namaSatuan)
        Else
            FormTransaksiPembelian.FLagdata = 4
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
       
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            selectDataBase = "SELECT T.IDObat, T.NamaObat, t1.Satuan AS Level1, t2.Satuan AS Level2, t3.Satuan AS Level3, tk.Kategori AS Kategori, tl.NamaLokasi AS Lokasi , T.StokMinimal FROM tobats AS T " +
                            " inner join tsatuan AS t1 " +
                                "on t1.IDSatuan = T.IDSatuan" +
                            " inner join tsatuan AS t2 " +
                                "on t2.IDSatuan = T.IDSatuanLv2" +
                            " inner join tsatuan AS t3 " +
                                "on t3.IDSatuan = T.IDSatuanLv3" +
                            " join tkategori AS tk " +
                                "on tk.IDKategori = T.IDKategori" +
                            " join tlokasi AS tl " +
                                "on tl.IDLokasi = T.IDLokasi  WHERE NamaObat LIKE '%" & TextBox1.Text & "%' ORDER BY IDObat ASC"

            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tsatuan")
            DataGridView1.DataSource = (DS.Tables("tsatuan"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID Obat"
                .Columns(1).HeaderCell.Value = "Nama Obat"
                .Columns(2).HeaderCell.Value = "Satuan Lv 1"
                .Columns(3).HeaderCell.Value = "Satuan Lv2"
                .Columns(4).HeaderCell.Value = "Satuan Lv 3"
                .Columns(5).HeaderCell.Value = "Kategori Obat"
                .Columns(6).HeaderCell.Value = "Lokasi Obat"
                .Columns(7).HeaderCell.Value = "Stok Minimal"
            End With
        End If
    End Sub

    Private Sub FormPencarianDataObat_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class