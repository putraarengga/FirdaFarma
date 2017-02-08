Public Class FormDataSupplier
    Dim databaru As Boolean
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub FormDataSupplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        IsiGrid()
        TextBox13.Enabled = False

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Truck_supplier.ico")
    End Sub
    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT * FROM tsupplier", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tsupplier")
        DataGridView1.DataSource = (DS.Tables("tsupplier"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Supplier"
            .Columns(1).HeaderCell.Value = "Nama Supplier"
            .Columns(2).HeaderCell.Value = "Alamat Supplier"
            .Columns(3).HeaderCell.Value = "Nomor HP"
            .Columns(4).HeaderCell.Value = "Nomor Rekening"
            .Columns(5).HeaderCell.Value = "Bank"
            .Columns(6).HeaderCell.Value = "Contact Person"
            .Columns(7).HeaderCell.Value = "Email"
            .Columns(8).HeaderCell.Value = "Website"
        End With
    End Sub
    Sub Bersih()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox13.Text = ""
    End Sub
    Sub Enable()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
    End Sub
    Sub Disable()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
    End Sub
    Private Sub isitextbox(ByVal x As Integer)
        Try
            TextBox13.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox2.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(2).Value
            TextBox4.Text = DataGridView1.Rows(x).Cells(3).Value
            TextBox5.Text = DataGridView1.Rows(x).Cells(4).Value
            TextBox6.Text = DataGridView1.Rows(x).Cells(5).Value
            TextBox7.Text = DataGridView1.Rows(x).Cells(6).Value
            TextBox8.Text = DataGridView1.Rows(x).Cells(7).Value
            TextBox9.Text = DataGridView1.Rows(x).Cells(8).Value
        Catch ex As Exception
        End Try
    End Sub
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox2.Focus()
        databaru = True
        Enable()
        Button3.Enabled = True
        Button2.Enabled = False
        Button4.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String
        If TextBox2.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO `tsupplier`(`NamaSupplier`, `AlamatSupplier`, `No.HP`, `No.Rekening`, `Bank`, `ContactPerson`, `Email`, `Website`) VALUES ('" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE `tsupplier` SET " _
                + "`NamaSupplier` = '" & TextBox2.Text & "'," _
                + "`AlamatSupplier` = '" & TextBox3.Text & "'," _
                + "`No.HP` = '" & TextBox4.Text & "'," _
                + "`No.Rekening` = '" & TextBox5.Text & "'," _
                + "`Bank` = '" & TextBox6.Text & "'," _
                + "`ContactPerson` = '" & TextBox7.Text & "'," _
                + "`Email` = '" & TextBox8.Text & "'," _
                + "`Website` = '" & TextBox9.Text & "' WHERE `IDSupplier` = '" & TextBox13.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Disable()
        Bersih()
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
        isitextbox(e.RowIndex)
        databaru = False
        Enable()
        Button3.Enabled = True
        Button4.Enabled = True
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server?" + TextBox2.Text, vbExclamation + vbYesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub
        hapussql = "DELETE FROM tsupplier WHERE IDSupplier ='" & TextBox13.Text & "'"
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
        Button3.Enabled = False
        Button4.Enabled = False
        Disable()
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
            DA.Fill(DS, "tsatuan")
            DataGridView1.DataSource = (DS.Tables("tsatuan"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID Supplier"
                .Columns(1).HeaderCell.Value = "Nama Supplier"
                .Columns(2).HeaderCell.Value = "Alamat Supplier"
                .Columns(3).HeaderCell.Value = "Nomor HP"
                .Columns(4).HeaderCell.Value = "Nomor Rekening"
                .Columns(5).HeaderCell.Value = "Bank"
                .Columns(6).HeaderCell.Value = "Contact Person"
                .Columns(7).HeaderCell.Value = "Email"
                .Columns(8).HeaderCell.Value = "Website"
            End With
        End If
    End Sub

    Private Sub FormDataSupplier_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class