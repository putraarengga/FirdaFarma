Imports Excel = Microsoft.Office.Interop.Excel


Public Class FormLaporanPenjualan
    Dim selectDataBase As String
    Dim selectDataBase2, JenisTransaksi, JT As String
    Dim nowTime, TglTransaksi, TT As String
    Dim FakturPenjualan, FP, GrandTotal, GT As String
    Dim dataSelected As Boolean
    Dim lastData As Integer = 0
    Dim dt1, now As Date
    Dim day1, month1, year1, vdate1, tgl1 As String
    Dim day2, month2, year2, vdate2, tgl2 As String

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormLaporanPenjualan_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dataSelected = False

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
        PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        'FormPencarianDataObat.Show()
        'FormPencarianDataObat.Focus()
        PrintPreviewDialog1.ShowDialog()

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub FormLaporanPenjualan_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Sub IsiGridUmum1()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT ttransaksi.FakturPenjualan, ttransaksi.TglTransaksi, ttransaksi.JenisTransaksi, ttransaksi.GrandTotal FROM ttransaksi WHERE ( ttransaksi.TglTransaksi BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) GROUP BY FakturPenjualan", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "ttransaksi")
        DataGridView1.DataSource = (DS.Tables("ttransaksi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nomor Faktur"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Grand Total"

        End With

    End Sub

    Sub IsiGridUmum2()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT ttransaksi.NamaObat, ttransaksi.TglTransaksi, ttransaksi.HargaJualUmum, ttransaksi.Jumlah, ttransaksi.hargadiskon, ttransaksi.TotalHarga FROM ttransaksi WHERE ttransaksi.FakturPenjualan ='" & FP & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "ttransaksi")
        DataGridView2.DataSource = (DS.Tables("ttransaksi"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Obat"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "HJU"
            .Columns(3).HeaderCell.Value = "Jumlah"
            .Columns(4).HeaderCell.Value = "Diskon"
            .Columns(5).HeaderCell.Value = "Sub Total"
        End With
    End Sub
    Sub GetIndeks(ByVal x As Integer)
        Try
            FakturPenjualan = DataGridView1.Rows(x).Cells(0).Value
            TglTransaksi = DataGridView1.Rows(x).Cells(1).Value.ToString
            JenisTransaksi = DataGridView1.Rows(x).Cells(2).Value
            GrandTotal = DataGridView1.Rows(x).Cells(3).Value
        Catch ex As Exception
        End Try

    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
        FP = FakturPenjualan
        JT = JenisTransaksi
        TT = TglTransaksi
        Dim grand As String = FormatNumber(GrandTotal, 2, , , TriState.True)
        RichTextBox1.Text = grand
        IsiGridUmum2()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        year1 = DateTimePicker1.Value.Year
        month1 = DateTimePicker1.Value.Month
        day1 = DateTimePicker1.Value.Day
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd-MMM-yyyy"
        vdate1 = year1 + "-" + month1 + "-" + day1
        tgl1 = day1 + "-" + month1 + "-" + year1
        IsiGridUmum1()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        year2 = DateTimePicker2.Value.Year
        month2 = DateTimePicker2.Value.Month
        day2 = DateTimePicker2.Value.Day
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "dd-MMM-yyyy"
        vdate2 = year2 + "-" + month2 + "-" + day2
        tgl2 = day2 + "-" + month2 + "-" + year2
        IsiGridUmum1()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT ttransaksi.FakturPenjualan, ttransaksi.TglTransaksi, ttransaksi.JenisTransaksi, ttransaksi.GrandTotal FROM ttransaksi WHERE ttransaksi.FakturPenjualan = '" & TextBox2.Text & "' AND( ttransaksi.TglTransaksi BETWEEN '" & vdate1 & "' AND '" & vdate2 & "' ) GROUP BY FakturPenjualan", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "ttransaksi")
        DataGridView1.DataSource = (DS.Tables("ttransaksi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nomor Faktur"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Grand Total"

        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'verfying the datagridview having data or not
        If ((DataGridView1.Columns.Count = 0) Or (DataGridView1.Rows.Count = 0) Or (DataGridView2.Columns.Count = 0) Or (DataGridView2.Rows.Count = 0)) Then
            MsgBox("Tidak Dapat mengetahui data")
        Else

            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()
            'add column to that table
            For i As Integer = 0 To DataGridView2.ColumnCount - 1
                If DataGridView2.Columns(i).Visible = True Then
                    dset.Tables(0).Columns.Add(DataGridView2.Columns(i).HeaderText)
                End If
            Next
            Dim count As Integer = -1
            'add rows to the table
            Dim dr1 As DataRow
            For i As Integer = 0 To DataGridView2.RowCount - 1
                dr1 = dset.Tables(0).NewRow


                For j As Integer = 0 To DataGridView2.Columns.Count - 1
                    If DataGridView2.Columns(j).Visible = True Then
                        count = count + 1

                        dr1(count) = DataGridView2.Rows(i).Cells(j).Value
                    End If
                Next

                count = -1
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim excel As New Excel.Application
            Dim wBook As Excel.Workbook
            Dim wSheet As Excel.Worksheet

            wBook = excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()


            Dim dt As System.Data.DataTable = dset.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 4

            wSheet.Cells(1, 1) = "Nomor Faktur"
            wSheet.Cells(2, 1) = FakturPenjualan

            wSheet.Cells(1, 2) = "Tanggal Transaksi"
            wSheet.Cells(2, 2) = TglTransaksi

            wSheet.Cells(1, 3) = "Jenis Transaksi"
            wSheet.Cells(2, 3) = JenisTransaksi

            wSheet.Cells(1, 4) = "Grand Total"
            wSheet.Cells(2, 4) = GrandTotal

            Dim last As Integer

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(4, colIndex) = dc.ColumnName
            Next
            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
                last = rowIndex
            Next

            wSheet.Columns.AutoFit()

            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx"
            saveFileDialog1.Title = "Save Excel File"
            saveFileDialog1.FileName = "File Data Penjualan " & FakturPenjualan & ".xls"
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