Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormLaporanHistoryHarga
    Dim selectDataBase2 As String
    Dim nowTime As String
    Dim dataSelected As Boolean
    Dim sum As Integer
    Dim IDSupplierObat, IDObat, NamaObat As String
    Dim ISO, IO, NO As String
    Dim lastData As Integer = 0

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormLaporanHistoryHarga_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        IsiGrid()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\1461484694_6.ico")
        PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub FormLaporanHistoryHarga_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub

    Sub IsiGrid()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase2 = "SELECT thistoryharga.IDObat, thistoryharga.NamaObat FROM thistoryharga GROUP BY NamaObat ORDER BY NamaObat ASC "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "thistoryharga")
        DataGridView1.DataSource = (DS.Tables("thistoryharga"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
        End With
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            IDObat = DataGridView1.Rows(x).Cells(0).Value
            NamaObat = DataGridView1.Rows(x).Cells(1).Value

        Catch ex As Exception
        End Try
    End Sub
    Private Sub periksasql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            sum = Convert.ToInt32(objcmd.ExecuteScalar())
            objcmd.Dispose()
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
        IO = IDObat
        NO = NamaObat
        IsiGrid2()

    End Sub

    Sub IsiGrid2()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase2 = "SELECT thistoryharga.NomorFaktur, tsupplier.NamaSupplier, thistoryharga.HargaBeliObat, thistoryharga.TglPembelian," +
            "thistoryharga.HJU1, thistoryharga.DiskonJual1, thistoryharga.Jumlah1 , thistoryharga.HJU2, thistoryharga.DiskonJual2," +
            "thistoryharga.Jumlah2 , thistoryharga.HJU3, thistoryharga.DiskonJual3, thistoryharga.Jumlah3, thistoryharga.HJR, thistoryharga.PajakPPN" +
            " FROM thistoryharga join tsupplier on tsupplier.IDSupplier = thistoryharga.IDSupplier WHERE thistoryharga.IDObat='" & IO & "' AND thistoryharga.NamaObat='" & NO & "' ORDER BY HargaBeliObat ASC "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "thistoryharga")
        DataGridView2.DataSource = (DS.Tables("thistoryharga"))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nomor Faktur"
            .Columns(1).HeaderCell.Value = "Nama Supplier"
            .Columns(2).HeaderCell.Value = "Harga Beli Obat"
            .Columns(3).HeaderCell.Value = "Tanggal Pembelian"
            .Columns(4).HeaderCell.Value = "Harga Jual Umum 1"
            .Columns(5).HeaderCell.Value = "Diskon Jual Umum 1"
            .Columns(6).HeaderCell.Value = "Jumlah Level 1"
            .Columns(7).HeaderCell.Value = "Harga Jual Umum 2"
            .Columns(8).HeaderCell.Value = "Diskon Jual Umum 2"
            .Columns(9).HeaderCell.Value = "Jumlah Level 2"
            .Columns(10).HeaderCell.Value = "Harga Jual Umum 3"
            .Columns(11).HeaderCell.Value = "Diskon Jual Umum 3"
            .Columns(12).HeaderCell.Value = "Jumlah Level 3"
            .Columns(13).HeaderCell.Value = "Harga Jual Resep"
            .Columns(14).HeaderCell.Value = "Pajak PPN %"

        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView2.RowCount > 0 Then
            PrintDocument1.DefaultPageSettings.Landscape = True
            PrintPreviewDialog1.ShowDialog()
            Me.Close()
        Else
            MsgBox("Tidak dapat Data Obat", MsgBoxStyle.OkCancel)
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim height As Integer = 0
        Dim height2 As Integer = 0
        Dim width As Integer = 0
        Dim i As Integer
        Dim BlackPen As New Pen(Brushes.Black, 2.5F)
        nowTime = Format(DateTime.Now, "dd-MM-yyyy")
        Dim rect1 As New Rectangle(250, 10, 500, 140)
        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center
        Dim stringFormat2 As New StringFormat()
        stringFormat2.Alignment = StringAlignment.Far
        stringFormat2.LineAlignment = StringAlignment.Far
        Dim text1 As String = "LAPORAN HISTORY HARGA OBAT APOTEK FIRDA FARMA"
        Dim font1 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point)

        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logo.png")

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
        e.Graphics.DrawString("ID Obat", DataGridView1.Font, Brushes.Black, New Rectangle(200, 100, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(310, 100, 490, 27))
        e.Graphics.DrawString(IO, DataGridView1.Font, Brushes.Black, New Rectangle(310, 100, 490, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(200, 127, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(200, 127, 110, 27))
        e.Graphics.DrawString("Nama Obat", DataGridView1.Font, Brushes.Black, New Rectangle(200, 127, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(310, 127, 490, 27))
        e.Graphics.DrawString(NO, DataGridView1.Font, Brushes.Black, New Rectangle(310, 127, 490, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(200, 154, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(200, 154, 110, 27))
        e.Graphics.DrawString("Tanggal Cetak", DataGridView1.Font, Brushes.Black, New Rectangle(200, 154, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(310, 154, 490, 27))
        e.Graphics.DrawString(nowTime, DataGridView1.Font, Brushes.Black, New Rectangle(310, 154, 490, 27))
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
        '.Columns(0).HeaderCell.Value = "Nomor Faktur"
        '.Columns(1).HeaderCell.Value = "Nama Supplier"
        '.Columns(2).HeaderCell.Value = "Harga Beli Obat"
        '.Columns(3).HeaderCell.Value = "Tanggal Pembelian"
        '.Columns(4).HeaderCell.Value = "Harga Jual Umum 1"
        '.Columns(5).HeaderCell.Value = "Diskon Jual Umum 1"
        '.Columns(6).HeaderCell.Value = "Jumlah Level 1"
        '.Columns(7).HeaderCell.Value = "Harga Jual Umum 2"
        '.Columns(8).HeaderCell.Value = "Diskon Jual Umum 2"
        '.Columns(9).HeaderCell.Value = "Jumlah Level 2"
        '.Columns(10).HeaderCell.Value = "Harga Jual Umum 3"
        '.Columns(11).HeaderCell.Value = "Diskon Jual Umum 3"
        '.Columns(12).HeaderCell.Value = "Jumlah Level 3"
        '.Columns(13).HeaderCell.Value = "Harga Jual Resep"
        '.Columns(14).HeaderCell.Value = "Pajak PPN"

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(50, 200, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, 200, 100, 40))
        e.Graphics.DrawString("Nomor Faktur", DataGridView2.Font, Brushes.Black, New Rectangle(50, 200, 100, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(150, 200, 150, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(150, 200, 150, 40))
        e.Graphics.DrawString("Nama Supplier", DataGridView2.Font, Brushes.Black, New Rectangle(150, 200, 150, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(300, 200, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, 200, 100, 40))
        e.Graphics.DrawString("Harga Beli Obat", DataGridView2.Font, Brushes.Black, New Rectangle(300, 200, 100, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(400, 200, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, 200, 100, 40))
        e.Graphics.DrawString("Tgl. Pembelian", DataGridView2.Font, Brushes.Black, New Rectangle(400, 200, 100, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(500, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, 200, 50, 40))
        e.Graphics.DrawString("HJU Lv 1", DataGridView2.Font, Brushes.Black, New Rectangle(500, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(550, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, 200, 50, 40))
        e.Graphics.DrawString("Diskon Lv1", DataGridView2.Font, Brushes.Black, New Rectangle(550, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(600, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, 200, 50, 40))
        e.Graphics.DrawString("Jumlah Lv1", DataGridView2.Font, Brushes.Black, New Rectangle(600, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(650, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(650, 200, 50, 40))
        e.Graphics.DrawString("HJU Lv 2", DataGridView2.Font, Brushes.Black, New Rectangle(650, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(700, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(700, 200, 50, 40))
        e.Graphics.DrawString("Diskon Lv2", DataGridView2.Font, Brushes.Black, New Rectangle(700, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(750, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(750, 200, 50, 40))
        e.Graphics.DrawString("Jumlah Lv2", DataGridView2.Font, Brushes.Black, New Rectangle(750, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(800, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(800, 200, 50, 40))
        e.Graphics.DrawString("HJU Lv 3", DataGridView2.Font, Brushes.Black, New Rectangle(800, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(850, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(850, 200, 50, 40))
        e.Graphics.DrawString("Diskon Lv3", DataGridView2.Font, Brushes.Black, New Rectangle(850, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(900, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(900, 200, 50, 40))
        e.Graphics.DrawString("Jumlah Lv3", DataGridView2.Font, Brushes.Black, New Rectangle(900, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(950, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, 200, 50, 40))
        e.Graphics.DrawString("Harga Jual Resep", DataGridView2.Font, Brushes.Black, New Rectangle(950, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(1000, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(1000, 200, 50, 40))
        e.Graphics.DrawString("Pajak PPN %", DataGridView2.Font, Brushes.Black, New Rectangle(1000, 200, 50, 40), stringFormat)


        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        height = 218
        i = lastData
        While (i < DataGridView2.Rows.Count)

            If (height > e.MarginBounds.Height) Then
                height = 218
                e.HasMorePages = True
                Return
            Else
                e.HasMorePages = False
            End If

            height += 30

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, height, DataGridView2.Columns(0).Width, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(0).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(50, height, DataGridView2.Columns(0).Width, 30))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(150, height, 150, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(1).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(150, height, 150, 30))

            Dim h_beli As String
            h_beli = FormatNumber(DataGridView2.Rows(i).Cells(2).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, height, 100, 30))
            e.Graphics.DrawString(h_beli, DataGridView2.Font, Brushes.Black, New Rectangle(300, height, 100, 30), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, height, 100, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(3).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(400, height, 100, 30))

            Dim h_HJU1 As String
            h_HJU1 = FormatNumber(DataGridView2.Rows(i).Cells(4).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, height, 50, 30))
            e.Graphics.DrawString(h_HJU1, DataGridView2.Font, Brushes.Black, New Rectangle(500, height, 50, 30), stringFormat2)

            Dim Dis1 As String
            Dis1 = FormatNumber(DataGridView2.Rows(i).Cells(5).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, height, 50, 30))
            e.Graphics.DrawString(Dis1, DataGridView2.Font, Brushes.Black, New Rectangle(550, height, 50, 30), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, height, 50, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(6).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(600, height, 50, 30), stringFormat2)

            Dim h_HJU2 As String
            h_HJU2 = FormatNumber(DataGridView2.Rows(i).Cells(7).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(650, height, 50, 30))
            e.Graphics.DrawString(h_HJU2, DataGridView2.Font, Brushes.Black, New Rectangle(650, height, 50, 30), stringFormat2)

            Dim Dis2 As String
            Dis2 = FormatNumber(DataGridView2.Rows(i).Cells(8).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(700, height, 50, 30))
            e.Graphics.DrawString(Dis2, DataGridView2.Font, Brushes.Black, New Rectangle(700, height, 50, 30), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(750, height, 50, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(9).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(750, height, 50, 30), stringFormat2)


            Dim h_HJU3 As String
            h_HJU3 = FormatNumber(DataGridView2.Rows(i).Cells(10).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(800, height, 50, 30))
            e.Graphics.DrawString(h_HJU3, DataGridView2.Font, Brushes.Black, New Rectangle(800, height, 50, 30), stringFormat2)

            Dim Dis3 As String
            Dis3 = FormatNumber(DataGridView2.Rows(i).Cells(11).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(850, height, 50, 30))
            e.Graphics.DrawString(Dis3, DataGridView2.Font, Brushes.Black, New Rectangle(850, height, 50, 30), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(900, height, 50, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(12).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(900, height, 50, 30), stringFormat2)

            Dim h_HJR As String
            h_HJR = FormatNumber(DataGridView2.Rows(i).Cells(13).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(950, height, 50, 30))
            e.Graphics.DrawString(h_HJR, DataGridView2.Font, Brushes.Black, New Rectangle(950, height, 50, 30), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(1000, height, 50, 30))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(14).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(1000, height, 50, 30), stringFormat2)


            i += 1
            lastData = i
        End While

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub PrintPreviewDialog1_Load(sender As Object, e As EventArgs) Handles PrintPreviewDialog1.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase2 = "SELECT thistoryharga.IDObat, thistoryharga.NamaObat FROM thistoryharga WHERE thistoryharga.NamaObat= '" & TextBox1.Text & "' GROUP BY NamaObat ORDER BY NamaObat ASC "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "thistoryharga")
        DataGridView1.DataSource = (DS.Tables("thistoryharga"))
        DataGridView1.Enabled = True
        DataGridView1.Refresh()
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
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

            wSheet.Cells(1, 1) = "ID Obat"
            wSheet.Cells(2, 1) = IO

            wSheet.Cells(1, 2) = "Nama Obat"
            wSheet.Cells(2, 2) = NO


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
            saveFileDialog1.FileName = "File Data History Harga Obat APOTEK FIRDA FARMA " & NO & ".xls"
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