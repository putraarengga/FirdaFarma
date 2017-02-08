Public Class FormDataObats
    Dim databaru As Boolean
    Dim selectDataBase As String
    Dim indexSatuan, indexSatuanLv2, indexSatuanlv3, indexKategori, indexLokasi As Integer
    Dim tmpSatuan, tmpKategori, tmpLokasi As Integer
    Dim day, month, year, vdate As String
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormDataObats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        IsiGrid()
        TextBox2.Enabled = False

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Devcom-Medical-Pill.ico")
    End Sub
    Sub IsiGrid()
        'selectDataBase = "SELECT tobats.IDObat, tobats.NamaObat, tsatuan.Satuan, tkategori.Kategori, tlokasi.NamaLokasi, tobats.StokMinimal FROM tobats As S " +
        '                    "join tsatuan AS R1 " +
        '                        "on R1.IDSatuan = S.IDSatuan  LIMIT 0,30"
        'selectDataBase = "SELECT * FROM tobats LIMIT 0,30"
        selectDataBase = "SELECT T.IDObat, T.NamaObat, t1.Satuan AS Level1, t2.Satuan AS Level2, t3.Satuan AS Level3, tk.Kategori AS Kategori, tl.NamaLokasi AS Lokasi, T.StokMinimal  FROM tobats AS T " +
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

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable()
        DA.Fill(DT)
        DataGridView1.DataSource = DT
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
    Sub Bersih()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox10.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
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
            With ComboBox1
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("Satuan"))

                Next
                .SelectedIndex = -1
            End With
            With ComboBox4
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("Satuan"))

                Next
                .SelectedIndex = -1
            End With
            With ComboBox5
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("Satuan"))

                Next
                .SelectedIndex = -1
            End With
        End If
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
    Sub ShowLokasi()
        selectDataBase = "SELECT NamaLokasi FROM tlokasi"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            With ComboBox3
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("NamaLokasi"))
                Next
                .SelectedIndex = -1
            End With
        End If
    End Sub
    Sub GetSatuan1()
        indexSatuan = -1
        selectDataBase = "SELECT IDSatuan FROM tsatuan WHERE Satuan ='" & ComboBox1.SelectedItem & "'"
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
    Sub GetSatuan2()
        indexSatuanLv2 = -1
        selectDataBase = "SELECT IDSatuan FROM tsatuan WHERE Satuan ='" & ComboBox4.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexSatuanLv2 = DT.Rows(0).Item("IDSatuan")
        End If
    End Sub
    Sub GetSatuan3()
        indexSatuanlv3 = -1
        selectDataBase = "SELECT IDSatuan FROM tsatuan WHERE Satuan ='" & ComboBox5.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexSatuanlv3 = DT.Rows(0).Item("IDSatuan")
        End If
    End Sub
    Sub GetLokasi()
        indexLokasi = -1
        selectDataBase = "SELECT IDLokasi FROM tlokasi WHERE NamaLokasi ='" & ComboBox3.SelectedItem & "'"
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

    Private Sub isitextbox(ByVal x As Integer)
        Try
            ComboBox1.SelectedIndex = ComboBox1.FindStringExact(DataGridView1.Rows(x).Cells(2).Value)
            ComboBox2.SelectedIndex = ComboBox2.FindStringExact(DataGridView1.Rows(x).Cells(5).Value)
            ComboBox3.SelectedIndex = ComboBox3.FindStringExact(DataGridView1.Rows(x).Cells(6).Value)
            ComboBox4.SelectedIndex = ComboBox4.FindStringExact(DataGridView1.Rows(x).Cells(3).Value)
            ComboBox5.SelectedIndex = ComboBox5.FindStringExact(DataGridView1.Rows(x).Cells(4).Value)
            TextBox2.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox10.Text = DataGridView1.Rows(x).Cells(7).Value
            'If DataGridView1.Rows(x).Cells(7).Value.ToString = "" Then
            'DateTimePicker1.CustomFormat = " "  'An empty SPACE
            'DateTimePicker1.Format = DateTimePickerFormat.Custom
            'Else
            'DateTimePicker1.CustomFormat = "dd/MM/yyyy"
            'DateTimePicker1.Value = DataGridView1.Rows(x).Cells(7).Value
            'DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.CustomFormat = "yyyy-MM-dd"
            'vdate = DateTimePicker1.Value.Year + "-" + DateTimePicker1.Value.Month + "-" + DateTimePicker1.Value.Day
            'End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        ShowSatuan()
        ShowLokasi()
        ShowKategori()
        TextBox3.Focus()
        DateTimePicker1.Enabled = True
        databaru = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox10.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String

        If TextBox3.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tobats(IDObat,NamaObat,IDSatuan,IDSatuanLv2,IDSatuanLv3,IDKategori,IDLokasi,StokMinimal) " +
                     "VALUES ('" & TextBox2.Text & "','" & TextBox3.Text & "','" & indexSatuan & "','" & indexSatuanLv2 & "','" & indexSatuanlv3 & "','" & indexKategori & "','" & indexLokasi & "','" & TextBox10.Text & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tobats SET NamaObat= '" & TextBox3.Text & "', IDSatuan = '" & indexSatuan & "', IDSatuanLv2 = '" & indexSatuanLv2 & "', IDSatuanLv3 = '" & indexSatuanlv3 & "',IDKategori= '" & indexKategori & "',IDLokasi= '" & indexLokasi & "',StokMinimal = '" & TextBox10.Text & "' WHERE IDObat= '" & TextBox2.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox10.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
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
        ShowSatuan()
        ShowLokasi()
        ShowKategori()
        isitextbox(e.RowIndex)
        TextBox3.Enabled = True
        TextBox10.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        DateTimePicker1.Enabled = True
        databaru = False
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox3.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub
        hapussql = "DELETE FROM tobats WHERE tobats.IDObat ='" & TextBox2.Text & "'"
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FormDataSatuan.Show()
        FormDataSatuan.Focus()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        year = DateTimePicker1.Value.Year
        month = DateTimePicker1.Value.Month
        day = DateTimePicker1.Value.Day
        vdate = year + "-" + month + "-" + day
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        GetSatuan1()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        GetKategori()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        GetLokasi()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        GetSatuan2()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        GetSatuan3()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        FormDataKategori.Show()
        FormDataKategori.Focus()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        FormDataLokasi.Show()
        FormDataLokasi.Focus()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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

    Private Sub FormDataObats_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        FormDataSatuan.Show()
        FormDataSatuan.Focus()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        FormDataSatuan.Show()
        FormDataSatuan.Focus()
    End Sub


End Class