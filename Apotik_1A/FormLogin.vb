Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Public Class FormLogin
    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim simpan As String
    Dim tmpAbsensi, tmpDate, tmpTime As String
    Dim countAbsensi As Integer
    Dim tmpCount As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        
        connect = New MySqlConnection
        'connect.ConnectionString = "server=192.168.1.1;userid=root;database=giandi"
        'connect.ConnectionString = "server=localhost;userid=root;password=123Siapasaja;database=giandi2"
        connect.ConnectionString = "server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=apotek"
        Dim reader As MySqlDataReader
        Dim userFound As Boolean = False
        Dim FullName As String = ""
        Dim idJenisUser As Integer


        Try
            connect.Open()
            Dim Query As String
            Query = String.Format("SELECT * FROM tuser WHERE NamaUser = '{0}' AND Password = '{1}'", Me.TB_Nama.Text.Trim(), Me.TB_Password.Text.Trim())
            command = New MySqlCommand(Query, connect)
            reader = command.ExecuteReader

            Dim count As Integer
            FormMenu.statusLog = 0
            count = 0
            While reader.Read
                count = count + 1
                userFound = True
                FullName = reader("NamaUser").ToString
                idJenisUser = reader("IDJenisUser")
                FormMenu.idUser = reader("IDUser")
            End While
            connect.Close()
            'If count = 1 Then
            'MessageBox.Show("Username and password are correct")
            'ElseIf count > 1 Then
            'MessageBox.Show("Username and password are duplicate")
            If count < 1 Then
                MsgBox("Sorry, username or password not found", MsgBoxStyle.OkOnly, "Invalid Login")
            End If
            If userFound = True Then
                TB_Nama.Clear()
                TB_Password.Clear()
                Hide()
                FormMenu.Enabled = True
                FormMenu.Show()
                FormMenu.MenuStrip1.Enabled = True
                FormMenu.LOGINToolStripMenuItem.Visible = False
                FormMenu.LOGINToolStripMenuItem.Enabled = False
                FormMenu.Label1.Text = "User : " & FullName
                FormMenu.MASTERToolStripMenuItem.Enabled = True
                FormMenu.TRANSAKSIToolStripMenuItem.Enabled = True
                FormMenu.LAPORANToolStripMenuItem.Enabled = True
                FormMenu.fullName = FullName

                FormMenu.DATAToolStripMenuItem.Enabled = False
                If idJenisUser = 2 Then
                    FormMenu.MASTERToolStripMenuItem.Enabled = False
                    FormMenu.LAPORANToolStripMenuItem.Enabled = False
                    FormMenu.MAINTENANCEToolStripMenuItem.Enabled = False
                    FormMenu.TRANSAKSIToolStripMenuItem.Enabled = True
                ElseIf idJenisUser = 1 Then
                    FormMenu.DATAToolStripMenuItem.Enabled = True
                End If
                FormMenu.KELUARToolStripMenuItem.Visible = True
                tmpDate = Format(Date.Now, "yyyy-MM-dd")
                tmpTime = Format(DateTime.Now, "HH:mm:ss")


                'simpan = "SELECT COUNT(*) FROM tabsensi WHERE IDAbsensi = '" & tmpAbsensi & "'"
                'periksasql(simpan)

                If countAbsensi = 0 And Not (FormMenu.statusLog Mod 2 = 1) Then

                    Try
                        connect.Open()
                        Query = String.Format("SELECT * FROM tabsensi ORDER BY status DESC LIMIT 1")
                        command = New MySqlCommand(Query, connect)
                        reader = command.ExecuteReader

                        While reader.Read
                            FormMenu.statusLog = reader("status")
                        End While
                        connect.Close()
                    Catch ex As MySqlException
                        MessageBox.Show(ex.Message)
                    Finally
                        connect.Dispose()
                    End Try

                    FormMenu.statusLog = FormMenu.statusLog + 1
                    tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & FormMenu.idUser & FormMenu.statusLog

                    simpan = "INSERT INTO tabsensi(IDAbsensi, IDUser, tglMasuk, wktMasuk) " +
                               "VALUES ('" & tmpAbsensi & "','" & FormMenu.idUser & "','" & tmpDate & "','" & tmpTime & "')"

                    jalankansql(simpan)
                End If

            End If



        Catch ex As MySqlException
            MessageBox.Show(ex.Message)
        Finally
            connect.Dispose()
        End Try
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
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub periksasql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            countAbsensi = Convert.ToInt16(objcmd.ExecuteScalar())
            objcmd.Dispose()
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormMenu.Show()
        TB_Nama.Focus()
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\wallpaper.jpg")

    End Sub

    Private Sub TB_Password_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_Password.KeyDown
        If e.KeyCode = Keys.Enter Then
            btn_login.Focus()
        End If
    End Sub

    Private Sub TB_Nama_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_Nama.KeyDown
        If e.KeyCode = Keys.Enter Then
            TB_Password.Focus()
        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub OpenFileDialog3_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog3.FileOk

    End Sub

    Private Sub OpenFileDialog4_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog4.FileOk

    End Sub
End Class
