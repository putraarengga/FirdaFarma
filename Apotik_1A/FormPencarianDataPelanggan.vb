Public Class FormPencarianDataPelanggan

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormPencarianDataPelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormDataPelanggan.Show()
        FormDataPelanggan.Focus()
    End Sub
End Class