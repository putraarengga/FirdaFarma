Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormLaporanObatExpiredHabis

    Dim dataSelected As Boolean
    Dim selectDataBase As String
    Dim selectDataBase2 As String
    Dim selectDate As String
    Dim nowTime As String
    Dim lastData As Integer = 0
    Dim JenLa As String


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormLaporanObatExpiredHabis_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Devcom-Medical-Pill.ico")
        PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"
        'PrintDocument1.PrinterSettings.PaperSizes = 9
        IsiGrid()


    End Sub
    Sub IsiGrid()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        If ComboBox1.SelectedItem = "Semua" Then
            selectDataBase = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1, tpembelian.SisaObatLv2, tpembelian.SatuanLv2, tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                        "FROM tpembelian WHERE (tpembelian.TglKadaluarsa < '" & nowTime & "') OR ((tpembelian.SisaObatLv1 ='" & 0 & "') AND(tpembelian.SisaObatLv2 ='" & 0 & "') AND (tpembelian.SisaObatLv3 ='" & 0 & "'))  ORDER BY tpembelian.NamaObat LIMIT 0,50"
        ElseIf ComboBox1.SelectedItem = "Obat Habis" Then
            selectDataBase = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1, tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                        "FROM tpembelian WHERE ((tpembelian.SisaObatLv1 ='" & 0 & "') AND(tpembelian.SisaObatLv2 ='" & 0 & "') AND (tpembelian.SisaObatLv3 ='" & 0 & "')) ORDER BY tpembelian.NamaObat LIMIT 0,50"
        ElseIf ComboBox1.SelectedItem = "Obat Kadaluarsa" Then
            selectDataBase = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1,tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                        "FROM tpembelian WHERE tpembelian.TglKadaluarsa < '" & nowTime & "' ORDER BY tpembelian.NamaObat LIMIT 0,50"

        End If
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Sisa Obat LV1"
            .Columns(3).HeaderCell.Value = "Satuan Obat LV1"
            .Columns(4).HeaderCell.Value = "Sisa Obat LV2"
            .Columns(5).HeaderCell.Value = "Satuan Obat LV2"
            .Columns(6).HeaderCell.Value = "Sisa Obat LV3"
            .Columns(7).HeaderCell.Value = "Satuan Obat LV3"
            .Columns(8).HeaderCell.Value = "Tanggal Kadaluarsa"
        End With

    End Sub

    Sub IsiGridExperied()
        Dim nowTime As String

        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1, tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                          "FROM tpembelian WHERE tpembelian.TglKadaluarsa < '" & nowTime & "' ORDER BY tpembelian.NamaObat LIMIT 0,50"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Sisa Obat LV1"
            .Columns(3).HeaderCell.Value = "Satuan Obat LV1"
            .Columns(4).HeaderCell.Value = "Sisa Obat LV2"
            .Columns(5).HeaderCell.Value = "Satuan Obat LV2"
            .Columns(6).HeaderCell.Value = "Sisa Obat LV3"
            .Columns(7).HeaderCell.Value = "Satuan Obat LV3"
            .Columns(8).HeaderCell.Value = "Tanggal Kadaluarsa"
        End With

    End Sub

    Sub CariIsiGrid()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1,tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                        "FROM tpembelian WHERE (((tpembelian.SisaObatLv1 ='" & 0 & "') AND(tpembelian.SisaObatLv2 ='" & 0 & "') AND (tpembelian.SisaObatLv3 ='" & 0 & "')) OR tpembelian.TglKadaluarsa < '" & nowTime & "' ) AND tpembelian.NamaObat LIKE '%" & TextBox1.Text & "%' ORDER BY tpembelian.NamaObat LIMIT 0,50"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Sisa Obat LV1"
            .Columns(3).HeaderCell.Value = "Satuan Obat LV1"
            .Columns(4).HeaderCell.Value = "Sisa Obat LV2"
            .Columns(5).HeaderCell.Value = "Satuan Obat LV2"
            .Columns(6).HeaderCell.Value = "Sisa Obat LV3"
            .Columns(7).HeaderCell.Value = "Satuan Obat LV3"
            .Columns(8).HeaderCell.Value = "Tanggal Kadaluarsa"
        End With
    End Sub

    Sub CariObatHabis()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1,tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3,tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                          "FROM tpembelian WHERE ((tpembelian.SisaObatLv1 ='" & 0 & "') AND(tpembelian.SisaObatLv2 ='" & 0 & "') AND (tpembelian.SisaObatLv3 ='" & 0 & "')) AND tpembelian.NamaObat LIKE '%" & TextBox1.Text & "%' ORDER BY tpembelian.NamaObat LIMIT 0,50"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Sisa Obat LV1"
            .Columns(3).HeaderCell.Value = "Satuan Obat LV1"
            .Columns(4).HeaderCell.Value = "Sisa Obat LV2"
            .Columns(5).HeaderCell.Value = "Satuan Obat LV2"
            .Columns(6).HeaderCell.Value = "Sisa Obat LV3"
            .Columns(7).HeaderCell.Value = "Satuan Obat LV3"
            .Columns(8).HeaderCell.Value = "Tanggal Kadaluarsa"
        End With
    End Sub

    Sub CariObatExperied()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT tpembelian.IDObat,tpembelian.NamaObat,tpembelian.SisaObatLv1, tpembelian.SatuanLv1,tpembelian.SisaObatLv2, tpembelian.SatuanLv2,tpembelian.SisaObatLv3, tpembelian.SatuanLv3,tpembelian.TglKadaluarsa " +
                          "FROM tpembelian WHERE tpembelian.TglKadaluarsa < '" & nowTime & "' AND tpembelian.NamaObat LIKE '%" & TextBox1.Text & "%' ORDER BY tpembelian.NamaObat LIMIT 0,50"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase2, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tpembelian")
        DataGridView1.DataSource = (DS.Tables("tpembelian"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Obat"
            .Columns(1).HeaderCell.Value = "Nama Obat"
            .Columns(2).HeaderCell.Value = "Sisa Obat LV1"
            .Columns(3).HeaderCell.Value = "Satuan Obat LV1"
            .Columns(4).HeaderCell.Value = "Sisa Obat LV2"
            .Columns(5).HeaderCell.Value = "Satuan Obat LV2"
            .Columns(6).HeaderCell.Value = "Sisa Obat LV3"
            .Columns(7).HeaderCell.Value = "Satuan Obat LV3"
            .Columns(8).HeaderCell.Value = "Tanggal Kadaluarsa"
        End With
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            If ComboBox1.SelectedItem = "Obat Habis" Then
                DataGridView1.Refresh()
                CariObatHabis()
            ElseIf ComboBox1.SelectedItem = "Obat Kadaluarsa" Then
                DataGridView1.Refresh()
                CariObatExperied()
            Else
                DataGridView1.Refresh()
                CariIsiGrid()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pkCustomSize1 As New System.Drawing.Printing.PaperSize("A4", 800, 1100)

        If DataGridView1.RowCount > 0 Then
            PrintDocument1.DefaultPageSettings.PaperSize = pkCustomSize1
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

        Dim rect1 As New Rectangle(150, 10, 500, 140)
        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center
        Dim stringFormat2 As New StringFormat()
        stringFormat2.Alignment = StringAlignment.Far
        stringFormat2.LineAlignment = StringAlignment.Far
        Dim text1 As String = "LAPORAN OBAT HABIS/EXPERIED APOTEK FIRDA FARMA"
        Dim font1 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point)

        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logo.png")

        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(90, 20)

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)

        e.HasMorePages = False
        width = 189
        width += DataGridView1.Rows(0).Height


        e.Graphics.DrawString(text1, font1, Brushes.Black, rect1, stringFormat)

        PrintDocument1.PrinterSettings.DefaultPageSettings.Margins.Bottom = 205

        If ComboBox1.SelectedItem = "Semua" Then
            JenLa = "Obat Habis dan Experied"
        ElseIf ComboBox1.SelectedItem = "Obat Habis" Then
            JenLa = "Obat Habis"
        ElseIf ComboBox1.SelectedItem = "Obat Kadaluarsa" Then
            JenLa = "Obat Experied"
        End If

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 100, 150, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 100, 150, 27))
        e.Graphics.DrawString("Cetak Laporan", DataGridView1.Font, Brushes.Black, New Rectangle(100, 100, 150, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, 100, 490, 27))
        e.Graphics.DrawString(JenLa, DataGridView1.Font, Brushes.Black, New Rectangle(250, 100, 300, 27))

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 127, 150, 27))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 127, 150, 27))
        e.Graphics.DrawString("Tanggal Cetak", DataGridView1.Font, Brushes.Black, New Rectangle(100, 127, 150, 27))

        e.Graphics.DrawRectangle(BlackPen, New Rectangle(250, 127, 490, 27))
        e.Graphics.DrawString(nowTime, DataGridView1.Font, Brushes.Black, New Rectangle(250, 127, 300, 27))

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(50, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, 200, 50, 40))
        e.Graphics.DrawString("ID Obat", DataGridView1.Font, Brushes.Black, New Rectangle(50, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(100, 200, 200, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, 200, 200, 40))
        e.Graphics.DrawString("Nama Obat", DataGridView1.Font, Brushes.Black, New Rectangle(100, 200, 200, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(300, 200, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, 200, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv1", DataGridView1.Font, Brushes.Black, New Rectangle(300, 200, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(375, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(375, 200, 50, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(375, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(425, 200, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(425, 200, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv2", DataGridView1.Font, Brushes.Black, New Rectangle(425, 200, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(500, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, 200, 50, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(500, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(550, 200, 75, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, 200, 75, 40))
        e.Graphics.DrawString("Sisa Obat Lv3", DataGridView1.Font, Brushes.Black, New Rectangle(550, 200, 75, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(625, 200, 50, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(625, 200, 50, 40))
        e.Graphics.DrawString("Satuan", DataGridView1.Font, Brushes.Black, New Rectangle(625, 200, 50, 40), stringFormat)

        e.Graphics.FillRectangle(Brushes.DarkGray, New Rectangle(675, 200, 100, 40))
        e.Graphics.DrawRectangle(BlackPen, New Rectangle(675, 200, 100, 40))
        e.Graphics.DrawString("Tgl. Kadarluarsa", DataGridView1.Font, Brushes.Black, New Rectangle(675, 200, 100, 40), stringFormat)

        '';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

        height = 218
        i = lastData
        While (i < DataGridView1.Rows.Count)

            If (height > e.MarginBounds.Height) Then
                height = 218
                e.HasMorePages = True
                Return
            Else
                e.HasMorePages = False
            End If

            height += DataGridView1.Rows(0).Height

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(50, height, 50, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(0).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(50, height, 50, DataGridView1.Rows(0).Height), stringFormat)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(100, height, 200, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(1).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(100, height, 200, DataGridView1.Rows(0).Height))

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(300, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(2).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(300, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(375, height, 50, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(3).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(375, height, 50, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(425, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(4).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(425, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(500, height, 50, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(5).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(500, height, 50, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(550, height, 75, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(6).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(550, height, 75, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(625, height, 50, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(7).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(625, height, 50, DataGridView1.Rows(0).Height), stringFormat2)

            e.Graphics.DrawRectangle(BlackPen, New Rectangle(675, height, 100, DataGridView1.Rows(0).Height))
            e.Graphics.DrawString(DataGridView1.Rows(i).Cells(8).FormattedValue.ToString(), DataGridView1.Font, Brushes.Black, New Rectangle(675, height, 100, DataGridView1.Rows(0).Height), stringFormat2)

            i += 1
            lastData = i
        End While
        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        DataGridView1.Refresh()
        IsiGrid()
    End Sub

    Private Sub FormLaporanObatExpiredHabis_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
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
            saveFileDialog1.FileName = "File LAPORAN OBAT EXPERIED-HABIS APOTEK FIRDA FARMA -" & nowTime & ".xls"
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