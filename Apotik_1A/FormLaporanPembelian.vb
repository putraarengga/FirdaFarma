Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormLaporanPembelian

    Dim dataSelected As Boolean
    Dim NamaSupplier As String
    Dim Supplier, selectedDatabase2 As String
    Dim TglPembelianObat As Date
    Dim TPO As Date
    Dim NamaObat, nowTime As String
    Dim selectDataBase As String
    Dim NomorFaktur, GrandTotal As Integer
    Dim Faktur, GT As Integer
    Dim sum As Integer
    Dim lastData As Integer = 0

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormLaporanPembelian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IsiGrid1()
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
    End Sub

    Sub IsiGrid1()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT tsupplier.NamaSupplier, tpembelian.NomorFaktur, tpembelian.TglPembelianObat, tpembelian.GrandTotal FROM tsupplier,tpembelian WHERE tsupplier.IDSupplier = tpembelian.IDSupplierObat GROUP BY NomorFaktur", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tsupplier,tpembelian")
        DataGridView1.DataSource = (DS.Tables("tsupplier,tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Supplier"
            .Columns(1).HeaderCell.Value = "Nomor Faktur"
            .Columns(2).HeaderCell.Value = "Tanggal Pembelian"
            .Columns(3).HeaderCell.Value = "Grand Total"
        End With
    End Sub

    Sub IsiGrid2()
        bukaDB()
        selectedDatabase2 = "SELECT tpembelian.NamaObat, tpembelian.TglKadaluarsa, tpembelian.JumlahPembelian, tpembelian.HargaBeli," +
          "tpembelian.DiskonPembelian, tpembelian.Pajak, tpembelian.SubTotal" +
          " FROM tpembelian WHERE tpembelian.NomorFaktur ='" & Faktur & "' "

        DA = New Odbc.OdbcDataAdapter(selectedDatabase2, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tpembelian.")
        DataGridView2.DataSource = (DS.Tables("tpembelian."))
        DataGridView2.Enabled = True
        With DataGridView2
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nama Obat"
            .Columns(1).HeaderCell.Value = "Tgl. Kadaluarsa"
            .Columns(2).HeaderCell.Value = "Jumlah Obat"
            .Columns(3).HeaderCell.Value = "Harga Beli Obat"
            .Columns(4).HeaderCell.Value = "Diskon Pembelian"
            .Columns(5).HeaderCell.Value = "Pajak PPN %"
            .Columns(6).HeaderCell.Value = "Sub Total"

        End With
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            NamaSupplier = DataGridView1.Rows(x).Cells(0).Value
            NomorFaktur = DataGridView1.Rows(x).Cells(1).Value
            TglPembelianObat = DataGridView1.Rows(x).Cells(2).Value.ToString
            GrandTotal = DataGridView1.Rows(x).Cells(3).Value

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
        Supplier = NamaSupplier
        Faktur = NomorFaktur
        TPO = TglPembelianObat
        RichTextBox2.Text = GrandTotal
        IsiGrid2()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        If DataGridView2.RowCount > 0 Then
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

        Dim rect1 As New Rectangle(189, 10, 500, 140)
        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center
        Dim stringFormat2 As New StringFormat()
        stringFormat2.Alignment = StringAlignment.Far
        stringFormat2.LineAlignment = StringAlignment.Far
        Dim text1 As String = "LAPORAN PEMBELIAN OBAT APOTEK FIRDA FARMA"
        Dim font1 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point)
        e.HasMorePages = False
        width = 189
        width += DataGridView1.Rows(0).Height


        e.Graphics.DrawString(text1, font1, Brushes.Black, rect1, stringFormat)

        PrintDocument1.PrinterSettings.DefaultPageSettings.Margins.Bottom = 205


        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 100, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 100, 110, 27))
        e.Graphics.DrawString("Nomor Faktur", DataGridView1.Font, Brushes.Black, New Rectangle(100, 100, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(210, 100, 490, 27))
        e.Graphics.DrawString(Faktur, DataGridView1.Font, Brushes.Black, New Rectangle(210, 100, 490, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 127, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 127, 110, 27))
        e.Graphics.DrawString("Nama Supplier", DataGridView1.Font, Brushes.Black, New Rectangle(100, 127, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(210, 127, 490, 27))
        e.Graphics.DrawString(Supplier, DataGridView1.Font, Brushes.Black, New Rectangle(210, 127, 490, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 154, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 154, 110, 27))
        e.Graphics.DrawString("Tanggal Pembelian Obat", DataGridView1.Font, Brushes.Black, New Rectangle(100, 154, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(210, 154, 490, 27))
        e.Graphics.DrawString(TPO, DataGridView1.Font, Brushes.Black, New Rectangle(210, 154, 490, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 181, 110, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 181, 110, 27))
        e.Graphics.DrawString("Tanggal Cetak", DataGridView1.Font, Brushes.Black, New Rectangle(100, 181, 110, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(210, 181, 490, 27))
        e.Graphics.DrawString(nowTime, DataGridView1.Font, Brushes.Black, New Rectangle(210, 181, 490, 27))
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
       
        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 250, 200, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 250, 200, 40))
        e.Graphics.DrawString("Nama Obat", DataGridView2.Font, Brushes.Black, New Rectangle(100, 250, 200, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(300, 250, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, 250, 100, 40))
        e.Graphics.DrawString("Tgl. Kadarluarsa", DataGridView2.Font, Brushes.Black, New Rectangle(300, 250, 100, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(400, 250, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, 250, 50, 40))
        e.Graphics.DrawString("Jumlah", DataGridView2.Font, Brushes.Black, New Rectangle(400, 250, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(450, 250, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(450, 250, 100, 40))
        e.Graphics.DrawString("Harga Obat", DataGridView2.Font, Brushes.Black, New Rectangle(450, 250, 100, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(550, 250, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, 250, 50, 40))
        e.Graphics.DrawString("Diskon Pembelian %", DataGridView2.Font, Brushes.Black, New Rectangle(550, 250, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(600, 250, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, 250, 50, 40))
        e.Graphics.DrawString("Pajak PPN %", DataGridView2.Font, Brushes.Black, New Rectangle(600, 250, 50, 40), stringFormat)

        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        height = 270
        i = lastData
        While (i < DataGridView2.Rows.Count)

            If (height > e.MarginBounds.Height) Then
                height = 270
                e.HasMorePages = True
                Return
            Else
                e.HasMorePages = False
            End If

            height += DataGridView2.Rows(0).Height
           
            e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, height, 200, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(0).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(100, height, 200, DataGridView2.Rows(0).Height))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, height, 100, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(1).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(300, height, 100, DataGridView2.Rows(0).Height))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(400, height, 50, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(2).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(400, height, 50, DataGridView2.Rows(0).Height), stringFormat2)

            Dim h_beli As String
            h_beli = FormatNumber(DataGridView2.Rows(i).Cells(3).FormattedValue.ToString(), TriState.False)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(450, height, 100, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(h_beli, DataGridView2.Font, Brushes.Black, New Rectangle(450, height, 100, DataGridView2.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, height, 50, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(4).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(550, height, 50, DataGridView2.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(600, height, 50, DataGridView2.Rows(0).Height))
            e.Graphics.DrawString(DataGridView2.Rows(i).Cells(5).FormattedValue.ToString(), DataGridView2.Font, Brushes.Black, New Rectangle(600, height, 50, DataGridView2.Rows(0).Height), stringFormat2)

            i += 1
            lastData = i
        End While
        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, height + 25, 400, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, height + 25, 400, 27))
        e.Graphics.DrawString("Grand Total Pembelian", RichTextBox2.Font, Brushes.Black, New Rectangle(100, height + 25, 400, 27))

        Dim h_total As String
        h_total = FormatNumber(RichTextBox2.Text, TriState.False)

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, height + 25, 150, RichTextBox2.Height))
        e.Graphics.DrawString("RP " & h_total, RichTextBox2.Font, Brushes.Black, New Rectangle(500, height + 25, 150, 27), stringFormat2)
        '';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, height + 75, 200, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, height + 75, 200, 27))
        e.Graphics.DrawString("Distributor", TextBox2.Font, Brushes.Black, New Rectangle(100, height + 75, 435, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, height + 102, 200, 100))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(500, height + 75, 200, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, height + 75, 200, 27))
        e.Graphics.DrawString("Apotek Firda Farma", TextBox2.Font, Brushes.Black, New Rectangle(500, height + 75, 435, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, height + 102, 200, 100))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub FormLaporanPembelian_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

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
            wSheet.Cells(2, 1) = Faktur

            wSheet.Cells(1, 2) = "Nama Supplier"
            wSheet.Cells(2, 2) = Supplier

            wSheet.Cells(1, 3) = "Tanggal Pembelian Obat"
            wSheet.Cells(2, 3) = TPO

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

            wSheet.Cells(last + 2, 1) = "GRAND TOTAL"
            wSheet.Cells(last + 2, 7) = RichTextBox2.Text

            wSheet.Columns.AutoFit()



            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx"
            saveFileDialog1.Title = "Save Excel File"
            saveFileDialog1.FileName = "File " & Faktur & " - " & Supplier & ".xls"
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