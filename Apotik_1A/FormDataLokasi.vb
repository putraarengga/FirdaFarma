﻿Public Class FormDataLokasi
    Dim databaru As Boolean
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormDataLokasi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img

        databaru = False
        IsiGrid()
        TextBox3.Enabled = False
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
    End Sub
    Sub IsiGrid()
        bukaDB()
        DA = New Odbc.OdbcDataAdapter("SELECT * FROM tlokasi", konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tlokasi")
        DataGridView1.DataSource = (DS.Tables("tlokasi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID Lokasi"
            .Columns(1).HeaderCell.Value = "Lokasi"
        End With
    End Sub
    Sub Bersih()
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Private Sub isitextbox(ByVal x As Integer)
        Try
            TextBox2.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(0).Value
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox2.Focus()
        databaru = True
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
            simpan = "INSERT INTO tlokasi(NamaLokasi) VALUES ('" & TextBox2.Text & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tlokasi SET NamaLokasi = '" & TextBox2.Text & "' WHERE IDLokasi = '" & TextBox3.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
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
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server?" + TextBox2.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub
        hapussql = "DELETE FROM tlokasi WHERE IDLokasi ='" & TextBox3.Text & "'"
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tlokasi WHERE NamaLokasi LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tlokasi")
            DataGridView1.DataSource = (DS.Tables("tlokasi"))
            DataGridView1.Enabled = True
            With DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ID Lokasi"
                .Columns(1).HeaderCell.Value = "Lokasi"
            End With
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub FormDataLokasi_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class