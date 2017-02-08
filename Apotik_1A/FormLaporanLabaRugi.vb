Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormLaporanLabaRugi
    Dim dt1, nowTime, now As Date
    Dim day1, month1, year1, vdate1, tgl1 As String
    Dim day2, month2, year2, vdate2, tgl2 As String
    Dim selectDataBase As String
    Dim Subtotal1, Subtotal2, Selisih As Integer
    Dim Countkeluar1, Countkeluar2, Countmasuk1, Countmasuk2 As Integer
    Dim Stringsub1, Stringsub2, StringSelisih As String
    Dim lastData As Integer = 0
    Dim lastData2 As Integer = 0
    Dim Strings, Strings2, Stringmasuk1, Stringmasuk2, Stringkeluar1, Stringkeluar2 As String


    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
    Private Sub FormLaporanLabaRugi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Devcom-Medical-Pill.ico")
        PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"
        IsiGrid1()
        IsiGrid2()
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        year1 = DateTimePicker1.Value.Year
        month1 = DateTimePicker1.Value.Month
        day1 = DateTimePicker1.Value.Day
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd-MMM-yyyy"
        vdate1 = year1 + "-" + month1 + "-" + day1
        tgl1 = day1 + "-" + month1 + "-" + year1
        IsiGrid1()
        IsiGrid2()
        CountingPengeluaran1()
        CountingPemasukan1()
        CountingPengeluaran2()
        CountingPemasukan2()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        year2 = DateTimePicker2.Value.Year
        month2 = DateTimePicker2.Value.Month
        day2 = DateTimePicker2.Value.Day
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "dd-MMM-yyyy"
        vdate2 = year2 + "-" + month2 + "-" + day2
        tgl2 = day2 + "-" + month2 + "-" + year2
        IsiGrid1()
        IsiGrid2()
        CountingPengeluaran1()
        CountingPemasukan1()
        CountingPengeluaran2()
        CountingPemasukan2()
    End Sub
    Sub IsiGrid1()
        Subtotal1 = 0
        selectDataBase = "SELECT ttransaksi.NamaObat, ttransaksi.TglTransaksi, ttransaksi.JenisTransaksi," +
            " ttransaksi.TotalHarga, ttransaksi.FakturPenjualan " +
            "FROM ttransaksi WHERE (ttransaksi.TglTransaksi BETWEEN '" & vdate1 & "' AND '" & vdate2 & "') UNION " +
        "SELECT tkeuangan.NamaTransaksi, tkeuangan.Tanggal, tkeuangan.JenisTransaksi, tkeuangan.Jumlah, tkeuangan.TglInput " +
        "FROM tkeuangan WHERE tkeuangan.JenisTransaksi= 'PEMASUKAN' AND (tkeuangan.Tanggal BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "ttransaksi,tkeuangan")
        DataGridView1.DataSource = (DS.Tables("ttransaksi,tkeuangan"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Jumlah"
            'DataGridView1.Columns(3).DefaultCellStyle.Format = "N2"
            .Columns(4).HeaderCell.Value = "Keterangan"

        End With
        If DataGridView1.RowCount > 0 Then
            For x As Integer = 0 To DataGridView1.RowCount - 1
                Subtotal1 += DataGridView1.Rows(x).Cells(3).Value
            Next
            Stringsub1 = FormatNumber(Subtotal1, 2, , , TriState.True)
            TextBox1.Text = Stringsub1
        End If
    End Sub

    Sub CountingPemasukan1()

        Countmasuk1 = 0
        selectDataBase = "SELECT ttransaksi.NamaObat, ttransaksi.TglTransaksi, ttransaksi.JenisTransaksi," +
            " ttransaksi.TotalHarga, ttransaksi.FakturPenjualan " +
            "FROM ttransaksi WHERE (ttransaksi.TglTransaksi BETWEEN '" & vdate1 & "' AND '" & vdate2 & "') "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                Countmasuk1 += DT.Rows(i).Item("TotalHarga")
            Next
            'TextBox4.Text = Countmasuk1.ToString()
        End If
    End Sub
    Sub CountingPemasukan2()

        Countmasuk2 = 0
        selectDataBase = "SELECT tkeuangan.NamaTransaksi, tkeuangan.Tanggal, tkeuangan.JenisTransaksi, tkeuangan.Jumlah, tkeuangan.TglInput " +
        "FROM tkeuangan WHERE tkeuangan.JenisTransaksi= 'PEMASUKAN' AND (tkeuangan.Tanggal BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                Countmasuk2 += DT.Rows(i).Item("Jumlah")
            Next
            'TextBox4.Text = Countmasuk2.ToString()
        End If
    End Sub


    Sub IsiGrid2()
        Subtotal2 = 0
        selectDataBase = "SELECT tpembelian.NamaObat, tpembelian.TglPembelianObat,tpembelian.JenisTransaksi, tpembelian.SubTotal, tpembelian.NomorFaktur " +
            "FROM tpembelian WHERE (tpembelian.TglPembelianObat BETWEEN '" & vdate1 & "' AND '" & vdate2 & "') UNION " +
        "SELECT tkeuangan.NamaTransaksi, tkeuangan.Tanggal,tkeuangan.JenisTransaksi, tkeuangan.Jumlah, tkeuangan.TglInput " +
        "FROM tkeuangan WHERE tkeuangan.JenisTransaksi= 'PENGELUARAN' AND (tkeuangan.Tanggal BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian,tkeuangan")
        DataGridView2.DataSource = (DS.Tables("tpembelian,tkeuangan"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Jumlah"
            'DataGridView2.Columns(3).DefaultCellStyle.Format = "N2"
            .Columns(4).HeaderCell.Value = "Keterangan"
        End With
        If DataGridView2.RowCount > 0 Then
            For x As Integer = 0 To DataGridView2.RowCount - 1
                Subtotal2 += DataGridView2.Rows(x).Cells(3).Value
            Next
            Stringsub2 = FormatNumber(Subtotal2, 2, , , TriState.True)
            TextBox2.Text = Stringsub2
            Selisih = Subtotal1 - Subtotal2
            StringSelisih = FormatNumber(Selisih, 2, , , TriState.True)
            TextBox3.Text = StringSelisih
        End If

    End Sub

    Sub CountingPengeluaran1()

        Countkeluar1 = 0
        selectDataBase = "SELECT tpembelian.NamaObat, tpembelian.TglPembelianObat, tpembelian.SubTotal, tpembelian.NomorFaktur " +
            "FROM tpembelian WHERE (tpembelian.TglPembelianObat BETWEEN '" & vdate1 & "' AND '" & vdate2 & "') "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                Countkeluar1 += DT.Rows(i).Item("SubTotal")
            Next
            'TextBox4.Text = Countkeluar1.ToString()
        End If
    End Sub
    Sub CountingPengeluaran2()

        Countkeluar2 = 0
        selectDataBase = "SELECT tkeuangan.NamaTransaksi, tkeuangan.Tanggal, tkeuangan.Jumlah, tkeuangan.TglInput " +
        "FROM tkeuangan WHERE tkeuangan.JenisTransaksi= 'PENGELUARAN' AND (tkeuangan.Tanggal BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                Countkeluar2 += DT.Rows(i).Item("Jumlah")
            Next
            'TextBox4.Text = Countkeluar2.ToString()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (DataGridView1.RowCount > 0) And (DataGridView2.RowCount > 0) Then
            PrintDocument1.DefaultPageSettings.Landscape = True
            PrintPreviewDialog1.ShowDialog()
            Me.Close()
        Else
            MsgBox("Tidak dapat Data Keuangan", MsgBoxStyle.OkCancel)
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim height As Integer = 0
        Dim height2 As Integer = 0
        Dim width As Integer = 0
        Dim i, j As Integer
        Dim BlackPen As New Pen(Brushes.Black, 2.5F)
        nowTime = Format(DateTime.Now, "dd-MM-yyyy")
        Dim rect1 As New Rectangle(250, 10, 500, 140)
        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center
        Dim stringFormat2 As New StringFormat()
        stringFormat2.Alignment = StringAlignment.Far
        stringFormat2.LineAlignment = StringAlignment.Far
        Dim text1 As String = "LAPORAN LABA/RUGI APOTEK FIRDA FARMA"
        Dim font1 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point)
        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logo.png")
        Dim tanggal As String = tgl1 + "  " + "Sampai" + "  " + tgl2
        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(185, 20)

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)
        e.HasMorePages = False
        width = 189
        width += DataGridView1.Rows(0).Height


        e.Graphics.DrawString(text1, font1, Brushes.Black, rect1, stringFormat)

        PrintDocument1.PrinterSettings.DefaultPageSettings.Margins.Bottom = 205


        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(200, 100, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(200, 100, 110, 27))
        e.Graphics.DrawString("Tanggal", DataGridView1.Font, Brushes.Black, New Rectangle(200, 100, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(310, 100, 490, 27))
        e.Graphics.DrawString(tanggal, DataGridView1.Font, Brushes.Black, New Rectangle(310, 100, 490, 27))


        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(50, 140, 550, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, 140, 550, 27))
        e.Graphics.DrawString("PEMASUKAN", DataGridView1.Font, Brushes.Black, New Rectangle(50, 140, 550, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, 167, 200, 27))
        e.Graphics.DrawString("Nama Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(50, 167, 200, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, 167, 150, 27))
        e.Graphics.DrawString("Tgl Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(250, 167, 150, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, 167, 200, 27))
        e.Graphics.DrawString("Jumlah Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(400, 167, 200, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(600, 140, 550, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, 140, 550, 27))
        e.Graphics.DrawString("PENGELUARAN", DataGridView1.Font, Brushes.Black, New Rectangle(600, 140, 550, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, 167, 200, 27))
        e.Graphics.DrawString("Nama Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(600, 167, 200, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(800, 167, 150, 27))
        e.Graphics.DrawString("Tgl Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(800, 167, 150, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, 167, 200, 27))
        e.Graphics.DrawString("Jumlah Transaksi", DataGridView1.Font, Brushes.Black, New Rectangle(950, 167, 200, 27))

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, 194, 350, 20))
        e.Graphics.DrawString("Penjualan Obat", DataGridView1.Font, Brushes.Black, New Rectangle(50, 194, 350, 20))

        Stringmasuk1 = FormatNumber(Countmasuk1, 2, , , TriState.True)

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, 194, 200, 20))
        e.Graphics.DrawString("RP " & Stringmasuk1, DataGridView1.Font, Brushes.Black, New Rectangle(400, 194, 200, 20), stringFormat2)

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, 194, 350, 20))
        e.Graphics.DrawString("Pembelian Obat", DataGridView1.Font, Brushes.Black, New Rectangle(600, 194, 350, 20))

        Stringkeluar1 = FormatNumber(Countkeluar1, 2, , , TriState.True)

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, 194, 200, 20))
        e.Graphics.DrawString("RP " & Stringkeluar1, DataGridView1.Font, Brushes.Black, New Rectangle(950, 194, 200, 20), stringFormat2)


        ' '';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        height = 214
        height2 = 214
        i = lastData
        j = lastData2

        While (j < DataGridView1.Rows.Count)
            If (DataGridView1.Rows(j).Cells(2).Value = "PEMASUKAN") Then
                If (height > e.MarginBounds.Height) Then
                    'height = 214
                    e.HasMorePages = True
                    Return
                Else
                    e.HasMorePages = False
                End If


                Strings2 = FormatNumber(DataGridView1.Rows(j).Cells(3).FormattedValue.ToString(), 2, , , TriState.True)

                e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, height, 200, 20))
                e.Graphics.DrawString(DataGridView1.Rows(j).Cells(0).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(50, height, 200, 20))

                e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, height, 150, 20))
                e.Graphics.DrawString(DataGridView1.Rows(j).Cells(1).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(250, height, 150, 20))

                e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, height, 200, 20))
                e.Graphics.DrawString("RP " & Strings2, DataGridView1.Font, Brushes.Black, New Rectangle(400, height, 200, 20), stringFormat2)

                height += 20

            End If

            j += 1
            lastData2 = j

        End While


        While (i < DataGridView2.Rows.Count)
            If (DataGridView2.Rows(i).Cells(2).Value = "PENGELUARAN") Then
                If (height > e.MarginBounds.Height) Then
                    'height = 214
                    e.HasMorePages = True
                    Return
                Else
                    e.HasMorePages = False
                End If



                Strings = FormatNumber(DataGridView2.Rows(i).Cells(3).FormattedValue.ToString(), 2, , , TriState.True)

                e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, height2, 200, 20))
                e.Graphics.DrawString(DataGridView2.Rows(i).Cells(0).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(600, height2, 200, 20))

                e.Graphics.DrawRectangle(BlackPen, New Rectangle(800, height2, 150, 20))
                e.Graphics.DrawString(DataGridView2.Rows(i).Cells(1).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(800, height2, 150, 20))


                e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, height2, 200, 20))
                e.Graphics.DrawString("RP " & Strings, DataGridView2.Font, Brushes.Black, New Rectangle(950, height2, 200, 20), stringFormat2)

                height2 += 20

            End If
            i += 1
            lastData = i

        End While

        If height <= height2 Then
            Dim ting As Integer = height2 + 1

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(50, ting, 350, 20))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, ting, 350, 20))
            e.Graphics.DrawString("Total Pemasukan", DataGridView2.Font, Brushes.Black, New Rectangle(50, ting, 350, 20))

            Stringmasuk2 = FormatNumber(Countmasuk2, 2, , , TriState.True)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, ting, 200, 20))
            e.Graphics.DrawString("RP " & Stringsub1, DataGridView2.Font, Brushes.Black, New Rectangle(400, ting, 200, 20), stringFormat2)

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(600, ting, 350, 20))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, ting, 350, 20))
            e.Graphics.DrawString("Total Pengeluaran", DataGridView2.Font, Brushes.Black, New Rectangle(600, ting, 350, 20))

            Stringkeluar2 = FormatNumber(Countkeluar2, 2, , , TriState.True)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, ting, 200, 20))
            e.Graphics.DrawString("RP " & Stringsub2, DataGridView2.Font, Brushes.Black, New Rectangle(950, ting, 200, 20), stringFormat2)

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(250, ting + 21, 350, 27))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, ting + 21, 350, 27))
            e.Graphics.DrawString("LABA/RUGI", DataGridView2.Font, Brushes.Black, New Rectangle(250, ting + 21, 350, 27))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, ting + 21, 350, 27))
            e.Graphics.DrawString("RP " & StringSelisih, DataGridView2.Font, Brushes.Black, New Rectangle(600, ting + 21, 350, 27), stringFormat2)

        ElseIf height > height2 Then
            Dim ting As Integer = height + 1

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(50, ting, 350, 20))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, ting, 350, 20))
            e.Graphics.DrawString("Total Pemasukan", DataGridView2.Font, Brushes.Black, New Rectangle(50, ting, 350, 20))

            Stringmasuk2 = FormatNumber(Countmasuk2, 2, , , TriState.True)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, ting, 200, 20))
            e.Graphics.DrawString("RP " & Stringsub1, DataGridView2.Font, Brushes.Black, New Rectangle(400, ting, 200, 20), stringFormat2)

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(600, ting, 350, 20))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, ting, 350, 20))
            e.Graphics.DrawString("Total Pengeluaran", DataGridView2.Font, Brushes.Black, New Rectangle(600, ting, 350, 20))

            Stringkeluar2 = FormatNumber(Countkeluar2, 2, , , TriState.True)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, ting, 200, 20))
            e.Graphics.DrawString("RP " & Stringsub2, DataGridView2.Font, Brushes.Black, New Rectangle(950, ting, 200, 20), stringFormat2)

            e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(250, ting + 21, 350, 27))
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, ting + 21, 350, 27))
            e.Graphics.DrawString("LABA/RUGI", DataGridView2.Font, Brushes.Black, New Rectangle(250, ting + 21, 350, 27))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, ting + 21, 350, 27))
            e.Graphics.DrawString("RP " & StringSelisih, DataGridView2.Font, Brushes.Black, New Rectangle(600, ting + 21, 350, 27), stringFormat2)

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim nows As String = Format(DateTime.Now, "dd-MMM-yyy")
        Stringmasuk1 = FormatNumber(Countmasuk1, 2, , , TriState.True)
        Stringkeluar2 = FormatNumber(Countkeluar2, 2, , , TriState.True)

        'verfying the datagridview having data or not
        If ((DataGridView1.Columns.Count = 0) Or (DataGridView1.Rows.Count = 0)) Then
            MsgBox("Tidak Dapat mengetahui data")
        Else

            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()
            'add column to that table
            For i As Integer = 0 To DataGridView1.ColumnCount - 1
                If DataGridView1.Columns(i).Visible = True Then
                    dset.Tables(0).Columns.Add(DataGridView1.Columns(i).HeaderText)
                End If
            Next
            Dim count As Integer = -1
            'add rows to the table
            Dim dr1 As DataRow

            For i As Integer = 0 To DataGridView1.RowCount - 1
                If Equals(DataGridView1.Rows(i).Cells(2).Value(), "PEMASUKAN") Then
                    dr1 = dset.Tables(0).NewRow

                    For j As Integer = 0 To DataGridView1.Columns.Count - 1
                        If DataGridView1.Columns(j).Visible = True Then
                            count = count + 1
                            dr1(count) = DataGridView1.Rows(i).Cells(j).Value
                        End If
                    Next

                    count = -1
                    dset.Tables(0).Rows.Add(dr1)
                End If
            Next
            ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


            'Creating dataset to export
            Dim dset2 As New DataSet
            'add table to dataset
            dset2.Tables.Add()
            'add column to that table
            For i As Integer = 0 To DataGridView2.ColumnCount - 1
                If DataGridView2.Columns(i).Visible = True Then
                    dset2.Tables(0).Columns.Add(DataGridView2.Columns(i).HeaderText)
                End If
            Next
            Dim count2 As Integer = -1
            'add rows to the table
            Dim dr2 As DataRow
            'Dim sis1 As DataRow
            For i As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Rows(i).Cells(2).Value = "PENGELUARAN" Then
                    dr2 = dset2.Tables(0).NewRow
                    'sis1 = dset2.Tables(0).NewRow
                    For j As Integer = 0 To DataGridView2.Columns.Count - 1
                        If DataGridView2.Columns(j).Visible = True Then
                            count2 = count2 + 1
                            'DataGridView2.Columns(3).ValueType = GetType(Double)
                            dr2(count2) = DataGridView2.Rows(i).Cells(j).Value
                            'Dim sis As String = FormatNumber(DataGridView2.Rows(i).Cells(3).Value(), TriState.True)
                            'sis1(count2) = sis
                End If
                    Next

                    count2 = -1
                    dset2.Tables(0).Rows.Add(dr2)
                    'dset2.Tables(0).Rows.Add(sis1)
                End If
            Next
            ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

            Dim excel As New Excel.Application
            Dim wBook As Excel.Workbook
            Dim wSheet As Excel.Worksheet

            wBook = excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()


            Dim dt As System.Data.DataTable = dset.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 1

            wSheet.Cells(1, 1) = "PEMASUKAN"

            wSheet.Cells(3, 1) = "Penjualan Obat"
            wSheet.Cells(3, 4) = Stringmasuk1


            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(2, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 2, colIndex) = dr(dc.ColumnName)

                Next
            Next

            Dim dt2 As System.Data.DataTable = dset2.Tables(0)
            Dim dc2 As System.Data.DataColumn
            Dim dr_2 As System.Data.DataRow
            Dim colIndex2 As Integer = 5
            Dim rowIndex2 As Integer = 1


            wSheet.Cells(1, 6) = "PENGELUARAN"

            wSheet.Cells(3, 6) = "Pembelian Obat"
            wSheet.Cells(3, 9) = Stringkeluar2

            For Each dc2 In dt2.Columns
                colIndex2 = colIndex2 + 1
                excel.Cells(2, colIndex2) = dc2.ColumnName
            Next

            For Each dr_2 In dt2.Rows
                rowIndex2 = rowIndex2 + 1
                colIndex2 = 5
                For Each dc2 In dt2.Columns
                    colIndex2 = colIndex2 + 1
                    excel.Cells(rowIndex2 + 2, colIndex2) = dr_2(dc2.ColumnName)

                Next
            Next


            wSheet.Columns.AutoFit()

            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx"
            saveFileDialog1.Title = "Save Excel File"
            saveFileDialog1.FileName = "File LAPORAN LABA-RUGI APOTEK FIRDA FARMA -" & nows & ".xls"
            saveFileDialog1.ShowDialog()

            saveFileDialog1.InitialDirectory = "D:\"
            If saveFileDialog1.FileName <> "" Then

                Dim fs As System.IO.FileStream = CType(saveFileDialog1.OpenFile(), System.IO.FileStream)
                fs.Close()
            End If


            Dim strFileName As String = saveFileDialog1.FileName
            Dim blnFileOpen As Boolean = False


            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
                Exit Sub
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            wBook.SaveAs(strFileName)
            excel.Workbooks.Open(strFileName)
            excel.Visible = True
            Exit Sub
errorhandler:
            MsgBox(Err.Description)
        End If
    End Sub
End Class