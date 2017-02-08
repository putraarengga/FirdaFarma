Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormLaporanStokObat

    Dim dataSelected As Boolean
    Dim selectedDataBase As String
    Dim nowTime As String
    Dim lastData As Integer = 0

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormLaporanStokObat_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        IsiGrid()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img

        PictureBox1.ImageLocation = appPath + ("\icons\Truck_supplier.ico")
        PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"

    End Sub

    Sub IsiGrid()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectedDataBase = "SELECT tpembelian.IDObat, tpembelian.NamaObat, tpembelian.TglKadaluarsa," +
            "tpembelian.HargaJualUmum1, tpembelian.SisaObatLv1, tpembelian.SatuanLv1," +
            " tpembelian.HargaJualUmum2,tpembelian.SisaObatLv2,tpembelian.SatuanLv2," +
            " tpembelian.HargaJualUmum3, tpembelian.SisaObatLv3, tpembelian.SatuanLv3, tpembelian.HargaJualResep" +
            " FROM tpembelian WHERE (tpembelian.TglKadaluarsa > '" & nowTime & "') AND ((tpembelian.SisaObatLv1 >'" & 0 & "') OR (tpembelian.SisaObatLv2 >'" & 0 & "') OR (tpembelian.SisaObatLv3 >'" & 0 & "')) ORDER BY tpembelian.NamaObat ASC LIMIT 0,50 "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectedDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Tanggal Kadaluarsa"
            .Columns(3).HeaderCell.Value = "Harga Jual Umum Lv1"
            .Columns(4).HeaderCell.Value = "Sisa Obat Lv1"
            .Columns(5).HeaderCell.Value = "Satuan Obat Lv1"
            .Columns(6).HeaderCell.Value = "Harga Jual Umum Lv2"
            .Columns(7).HeaderCell.Value = "Sisa Obat Lv2"
            .Columns(8).HeaderCell.Value = "Satuan Obat Lv2"
            .Columns(9).HeaderCell.Value = "Harga Jual Umum Lv3"
            .Columns(10).HeaderCell.Value = "Sisa Obat Lv3"
            .Columns(11).HeaderCell.Value = "Satuan Obat Lv3"
            .Columns(12).HeaderCell.Value = "Harga Jual Resep"
        End With
    End Sub
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs)

    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            selectedDataBase = "SELECT tpembelian.IDObat, tpembelian.NamaObat, tpembelian.TglKadaluarsa," +
            " tpembelian.HargaJualUmum1, tpembelian.SisaObatLv1,tpembelian.SatuanLv1 " +
            " tpembelian.HargaJualUmum2,tpembelian.SisaObatLv2, tpembelian.SatuanLv2" +
            " tpembelian.HargaJualUmum3, tpembelian.SisaObatLv3, tpembelian.SatuanLv3, tpembelian.HargaJualResep" +
            " FROM tpembelian WHERE tpembelian.NamaObat LIKE '%" & TextBox1.Text & "%' AND (tpembelian.TglKadaluarsa > '" & nowTime & "') AND ((tpembelian.SisaObatLv1 >'" & 0 & "') OR (tpembelian.SisaObatLv2 >'" & 0 & "') OR (tpembelian.SisaObatLv3 >'" & 0 & "')) ORDER BY tpembelian.NamaObat ASC LIMIT 0,50  "
            bukaDB()
            DA = New Odbc.OdbcDataAdapter(selectedDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tpembelian")
            DataGridView1.DataSource = (DS.Tables("tpembelian"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID Obat"
                .Columns(1).HeaderCell.Value = "Nama Obat"
                .Columns(2).HeaderCell.Value = "Tanggal Kadaluarsa"
                .Columns(3).HeaderCell.Value = "Harga Jual Umum Lv1"
                .Columns(4).HeaderCell.Value = "Sisa Obat Lv1"
                .Columns(5).HeaderCell.Value = "Satuan Obat Lv1"
                .Columns(6).HeaderCell.Value = "Harga Jual Umum Lv2"
                .Columns(7).HeaderCell.Value = "Sisa Obat Lv2"
                .Columns(8).HeaderCell.Value = "Satuan Obat Lv2"
                .Columns(9).HeaderCell.Value = "Harga Jual Umum Lv3"
                .Columns(10).HeaderCell.Value = "Sisa Obat Lv3"
                .Columns(11).HeaderCell.Value = "Satuan Obat Lv3"
                .Columns(12).HeaderCell.Value = "Harga Jual Resep"
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FormLaporanStokObat_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount > 0 Then
            PrintDocument1.DefaultPageSettings.Landscape = True
            PrintPreviewDialog1.ShowDialog()
            Me.Close()
        Else
            MsgBox("Tidak dapat Data Obat", MsgBoxStyle.OkCancel)
        End If
    End Sub

    Private Sub GroupBox2_Enter_1(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim height As Integer = 0
        Dim height2 As Integer = 0
        Dim width As Integer = 0
        Dim i As Integer
        Dim BlackPen As New Pen(Brushes.Black, 2.5F)

        Dim rect1 As New Rectangle(300, 10, 500, 140)
        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center

        Dim stringFormat2 As New StringFormat()
        stringFormat2.Alignment = StringAlignment.Far
        stringFormat2.LineAlignment = StringAlignment.Far

        Dim text1 As String = "LAPORAN STOK OBAT APOTEK FIRDA FARMA"
        Dim font1 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point)
        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logo.png")

        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(100, 45)

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)

        e.HasMorePages = False
        width = 189
        width += DataGridView1.Rows(0).Height


        e.Graphics.DrawString(text1, font1, Brushes.Black, rect1, stringFormat)

        PrintDocument1.PrinterSettings.DefaultPageSettings.Margins.Bottom = 205

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
       
        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(75, 150, 40, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(75, 150, 40, 40))
        e.Graphics.DrawString("No.", DataGridView1.Font, Brushes.Black, New Rectangle(75, 150, 40, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(115, 150, 200, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(115, 150, 200, 40))
        e.Graphics.DrawString("Nama Obat ", DataGridView1.Font, Brushes.Black, New Rectangle(115, 150, 200, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(315, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(315, 150, 75, 40))
        e.Graphics.DrawString("Tgl. Kadaluarsa", DataGridView1.Font, Brushes.Black, New Rectangle(315, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(390, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(390, 150, 75, 40))
        e.Graphics.DrawString("HJU Lv 1", DataGridView1.Font, Brushes.Black, New Rectangle(390, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(465, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(465, 150, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv 1", DataGridView1.Font, Brushes.Black, New Rectangle(465, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(540, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(540, 150, 75, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(540, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(615, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(615, 150, 75, 40))
        e.Graphics.DrawString("HJU Lv 2", DataGridView1.Font, Brushes.Black, New Rectangle(615, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(690, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(690, 150, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv 2", DataGridView1.Font, Brushes.Black, New Rectangle(690, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(765, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(765, 150, 75, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(765, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(840, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(840, 150, 75, 40))
        e.Graphics.DrawString("HJU Lv 3", DataGridView1.Font, Brushes.Black, New Rectangle(840, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(915, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(915, 150, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv 3", DataGridView1.Font, Brushes.Black, New Rectangle(915, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(990, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(990, 150, 75, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(990, 150, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(1065, 150, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(1065, 150, 75, 40))
        e.Graphics.DrawString("Harga Jual Resep", DataGridView1.Font, Brushes.Black, New Rectangle(1065, 150, 75, 40), stringFormat)

        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        height = 168
        i = lastData
        While (i < DataGridView1.Rows.Count)

            If (height > e.MarginBounds.Height) Then
                height = 168
                e.HasMorePages = True
                Return
            Else
                e.HasMorePages = False
            End If

            height += DataGridView1.Rows(0).Height

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(75, height, 40, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(i + 1, DataGridView1.Font, Brushes.Black, New Rectangle(75, height, 40, DataGridView1.Rows(0).Height), stringFormat)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(115, height, 200, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(1).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(115, height, 200, DataGridView1.Rows(0).Height))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(315, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(2).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(315, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            Dim h_HJU1 As String
            h_HJU1 = FormatNumber(DataGridView1.Rows(i).Cells(3).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(390, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(h_HJU1, DataGridView1.Font, Brushes.Black, New Rectangle(390, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(465, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(4).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(465, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(540, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(5).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(540, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            Dim h_HJU2 As String
            h_HJU2 = FormatNumber(DataGridView1.Rows(i).Cells(6).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(615, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(h_HJU2, DataGridView1.Font, Brushes.Black, New Rectangle(615, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(690, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(7).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(690, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(765, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(8).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(765, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            Dim h_HJU3 As String
            h_HJU3 = FormatNumber(DataGridView1.Rows(i).Cells(9).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(840, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(h_HJU3, DataGridView1.Font, Brushes.Black, New Rectangle(840, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(915, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(10).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(915, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(990, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(11).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(990, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            Dim h_HJR As String
            h_HJR = FormatNumber(DataGridView1.Rows(i).Cells(12).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(1065, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(h_HJR, DataGridView1.Font, Brushes.Black, New Rectangle(1065, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            i += 1
            lastData = i
        End While
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        nowTime = Format(DateTime.Now, "dd-MMM-yyy")

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
                dr1 = dset.Tables(0).NewRow


                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    If DataGridView1.Columns(j).Visible = True Then
                        count = count + 1

                        dr1(count) = DataGridView1.Rows(i).Cells(j).Value
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
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
            Next

            wSheet.Columns.AutoFit()

            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx"
            saveFileDialog1.Title = "Save Excel File"
            saveFileDialog1.FileName = "File LAPORAN STOK OBAT APOTEK FIRDA FARMA -" & nowTime & ".xls"
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