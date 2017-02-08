Public Class FormPencarianDataSupplier
    Dim databaru As Boolean
    Dim indeksSupplier As Integer
    Dim dataSelected As Boolean
    Dim namaSupplier As String

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormPencarianDataPelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        IsiGrid()

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img

        PictureBox1.ImageLocation = appPath + ("\icons\Truck_supplier.ico")
        TextBox1.Focus()

    End Sub

    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT * FROM tsupplier", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tsupplier")
        DataGridView1.DataSource = (DS.Tables("tsupplier"))
        DataGridView1.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormDataSupplier.Show()
        FormDataSupplier.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If dataSelected = True Then
            FormTransaksiPembelian.TextBox4.Text = namaSupplier
            FormTransaksiPembelian.indexSupplier = indeksSupplier
            Me.Close()
        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            indeksSupplier = DataGridView1.Rows(x).Cells(0).Value
            namaSupplier = DataGridView1.Rows(x).Cells(1).Value.ToString
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        FormTransaksiPembelian.TextBox4.Text = namaSupplier

        FormTransaksiPembelian.indexSupplier = indeksSupplier
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tsupplier WHERE NamaSupplier LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tsupplier")
            DataGridView1.DataSource = (DS.Tables("tsupplier"))
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub FormPencarianDataSupplier_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class