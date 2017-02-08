Public Class FormLaporanAbsensi
    Dim databaru As Boolean
    Dim indeksSupplier As Integer
    Dim dataSelected As Boolean
    Dim namaSupplier As String
    Dim tanggalAwal, tanggalAkhir As String

    Private Sub FormPencarianDataPelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        IsiGrid()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\user-male-female.ico")

        ComboBox1.SelectedItem = "Semua Data"

        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        Label4.Visible = False
        Label5.Visible = False



    End Sub

    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tabsensi")
        DataGridView1.DataSource = (DS.Tables("tabsensi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama User"
            .Columns(1).HeaderCell.Value = "Tanggal Masuk"
            .Columns(2).HeaderCell.Value = "Waktu Masuk"
            .Columns(3).HeaderCell.Value = "Waktu Keluar"
        End With
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        bukaDB()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        If ComboBox1.SelectedItem = "Semua Data" Then
            DisableForm()
            DA = New Odbc.OdbcDataAdapter("SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser", konek)
        ElseIf ComboBox1.SelectedItem = "Tanggal" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "dd/MMM/yyyy"
            DateTimePicker2.CustomFormat = "dd/MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-dd")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-dd")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Bulan" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "MMM/yyyy"
            DateTimePicker2.CustomFormat = "MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-00")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Tahun" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "yyyy"
            DateTimePicker2.CustomFormat = "yyyy"
            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-01-01")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-12-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If
        End If
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tabsensi")
        DataGridView1.DataSource = (DS.Tables("tabsensi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama User"
            .Columns(1).HeaderCell.Value = "Tanggal Masuk"
            .Columns(2).HeaderCell.Value = "Waktu Masuk"
            .Columns(3).HeaderCell.Value = "Waktu Keluar"
        End With
    End Sub

    Private Sub EnableForm()
        DateTimePicker1.Visible = True
        DateTimePicker2.Visible = True
        Label4.Visible = True
        Label5.Visible = True
    End Sub
    Private Sub DisableForm()
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        Label4.Visible = False
        Label5.Visible = False
    End Sub

    Private Sub ComboBox1_SystemColorsChanged(sender As Object, e As EventArgs) Handles ComboBox1.SystemColorsChanged

    End Sub

    Private Sub FormLaporanAbsensi_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        bukaDB()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        If ComboBox1.SelectedItem = "Semua Data" Then
            DisableForm()
            DA = New Odbc.OdbcDataAdapter("SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser", konek)
        ElseIf ComboBox1.SelectedItem = "Tanggal" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "dd/MMM/yyyy"
            DateTimePicker2.CustomFormat = "dd/MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-dd")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-dd")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Bulan" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "MMM/yyyy"
            DateTimePicker2.CustomFormat = "MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-00")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Tahun" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "yyyy"
            DateTimePicker2.CustomFormat = "yyyy"
            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-01-01")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-12-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If
        End If
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tabsensi")
        DataGridView1.DataSource = (DS.Tables("tabsensi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama User"
            .Columns(1).HeaderCell.Value = "Tanggal Masuk"
            .Columns(2).HeaderCell.Value = "Waktu Masuk"
            .Columns(3).HeaderCell.Value = "Waktu Keluar"
        End With
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        bukaDB()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        If ComboBox1.SelectedItem = "Semua Data" Then
            DisableForm()
            DA = New Odbc.OdbcDataAdapter("SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser", konek)
        ElseIf ComboBox1.SelectedItem = "Tanggal" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "dd/MMM/yyyy"
            DateTimePicker2.CustomFormat = "dd/MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-dd")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-dd")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Bulan" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "MMM/yyyy"
            DateTimePicker2.CustomFormat = "MMM/yyyy"

            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-MM-00")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-MM-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If

        ElseIf ComboBox1.SelectedItem = "Tahun" Then
            EnableForm()
            DateTimePicker1.CustomFormat = "yyyy"
            DateTimePicker2.CustomFormat = "yyyy"
            tanggalAwal = Format(DateTimePicker1.Value.Date, "yyyy-01-01")
            tanggalAkhir = Format(DateTimePicker2.Value.Date, "yyyy-12-31")
            If tanggalAwal = tanggalAkhir Then
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser  WHERE tabsensi.tglMasuk LIKE '%" & tanggalAkhir & "%'", konek)
            Else
                DA = New Odbc.OdbcDataAdapter(
                "SELECT tuser.NamaUser, tabsensi.tglMasuk, tabsensi.wktMasuk, tabsensi.wktKeluar " +
                "FROM tabsensi join tuser on tuser.IDUser = tabsensi.IDUser WHERE tabsensi.tglMasuk BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'", konek)

            End If
        End If
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tabsensi")
        DataGridView1.DataSource = (DS.Tables("tabsensi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama User"
            .Columns(1).HeaderCell.Value = "Tanggal Masuk"
            .Columns(2).HeaderCell.Value = "Waktu Masuk"
            .Columns(3).HeaderCell.Value = "Waktu Keluar"
        End With
    End Sub
End Class