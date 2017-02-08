Public Class FormDataUser
    Dim databaru As Boolean
    Dim selectDataBase, vJenisUser, tmpString As String
    Dim indexSatuan, indexKategori, indexLokasi As Integer
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormDataUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        IsiGrid()
        TextBox2.Enabled = False
        
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\user-male-female.ico")
    End Sub
    Sub IsiGrid()
        selectDataBase = "SELECT tuser.IDUser, tuser.NamaUser,tuser.Password,tuser.NamaLengkap,tjenisuser.JenisUser " +
                        " FROM tuser " +
                            "join tjenisuser " +
                                "on tuser.IDJenisUser = tjenisuser.IDJenisUser "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tuser")
        DataGridView1.DataSource = (DS.Tables("tuser"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID User"
            .Columns(1).HeaderCell.Value = "Nama User"
            .Columns(2).HeaderCell.Value = "Password"
            .Columns(3).HeaderCell.Value = "Nama Lengkap"
            .Columns(4).HeaderCell.Value = "Jenis User"
        End With
    End Sub
    Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.Text = ""
        ComboBox1.Items.Clear()
        
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox4.Focus()
        ShowJenisUser()
        databaru = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tuser WHERE NamaUser LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tuser")
            DataGridView1.DataSource = (DS.Tables("tuser"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID User"
                .Columns(1).HeaderCell.Value = "Nama User"
                .Columns(2).HeaderCell.Value = "Password"
                .Columns(3).HeaderCell.Value = "Nama Lengkap"
                .Columns(4).HeaderCell.Value = "Jenis User"
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String

        If TextBox4.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tuser(IDUser,NamaUser,Password,NamaLengkap,IDJenisUser) " +
                     "VALUES (LAST_INSERT_ID(),'" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox3.Text & "','" & vJenisUser & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tuser SET NamaUser= '" & TextBox4.Text & "', Password = '" & TextBox5.Text & "',NamaLengkap= '" & TextBox3.Text & "',IDJenisUser= '" & vJenisUser & "' WHERE IDUser= '" & TextBox2.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        ComboBox1.Enabled = False
        
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
    Sub GetJenisUser()
        vJenisUser = -1
        selectDataBase = "SELECT IDJenisUser FROM tjenisuser WHERE JenisUser='" & ComboBox1.SelectedItem & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            vJenisUser = DT.Rows(0).Item("IDJenisUser")
        End If
    End Sub

    Sub ShowJenisUser()
        selectDataBase = "SELECT JenisUser FROM tjenisuser "
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
                    .Items.Add(DT.Rows(i).Item("JenisUser"))
                Next
                .SelectedIndex = -1
            End With
        End If
    End Sub
    Sub isitextbox(ByVal x As Integer)

        Try
            ComboBox1.SelectedIndex = ComboBox1.FindStringExact(DataGridView1.Rows(x).Cells(4).Value.ToString)
            TextBox2.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(3).Value
            TextBox4.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox5.Text = DataGridView1.Rows(x).Cells(2).Value

        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Bersih()
        ShowJenisUser()
        isitextbox(e.RowIndex)
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        ComboBox1.Enabled = True
        databaru = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        GetJenisUser()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox4.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub

        hapussql = "DELETE FROM tuser WHERE tuser.IDUser ='" & TextBox2.Text & "'"
        If TextBox2.Text = "" Then Exit Sub
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
    End Sub

    Private Sub FormDataUser_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class