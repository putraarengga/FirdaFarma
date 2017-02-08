Public Class FormDaftarRacikan
    Dim databaru As Boolean
    Private Sub FormDataSatuan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = True
        IsiGrid()
        TextBox3.Enabled = False

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Pills.ico")

    End Sub
    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT * FROM tracikan", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tracikan")
    End Sub
    Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Bersih()
        TextBox2.Focus()
        databaru = True
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String
        Dim tmpString As String

        tmpString = Format(DateTime.Now, "yyyy-MM-dd")
        If TextBox2.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tRacikan(NomorFakturPenjualan,namaRacikan,namaPembeli,namaPasien,tglPembuatan)" +
                "VALUES ('" & FormTransaksiPenjualan.TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox1.Text & "','" & TextBox4.Text & "','" & tmpString & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", vbYesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tracikan SET namaRacikan = '" & TextBox2.Text & "',namaPembeli = '" & TextBox1.Text & "',namaPasien = '" & TextBox4.Text & "',tglPembuatan = '" & tmpString & "' WHERE IDRacikan = '" & TextBox3.Text & "' "
        End If
        jalankansql(simpan)
        FormDataRacikan.IsiGrid()

        Me.Close()

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
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        databaru = False
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server?" + TextBox2.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub
        hapussql = "DELETE FROM tracikan WHERE IDRacikan='" & TextBox3.Text & "'"
        jalankansql(hapussql)
        IsiGrid()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
       
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FormDaftarRacikan_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class