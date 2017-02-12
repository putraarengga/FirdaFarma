Public Class FormMenu
    Dim simpan As String
    Dim tmpAbsensi, tmpDate, tmpTime As String
    Dim idJenisUser As Integer
    Dim countAbsensi As Integer
    Shared Property statusLog As Integer
    Shared Property idUser As Integer
    Shared Property fullName As String
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        FormTransaksiPenjualan.Show()
    End Sub

    Private Sub btn_pembelian_Click(sender As Object, e As EventArgs)
        FormTransaksiPembelian.Show()

    End Sub


    Private Sub btn_profile_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub PENJUALANToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PEMBELIANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MASTERToolStripMenuItem.Click

    End Sub

    Private Sub REKAPANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSAKSIToolStripMenuItem.Click

    End Sub

    Private Sub BANTUANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LAPORANToolStripMenuItem.Click

    End Sub

    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Me.Width = screenWidth
        Me.Height = screenHeight
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\backgrounds.jpg")
        Me.BackgroundImageLayout = ImageLayout.Stretch

        'Me.Enabl.ed = False
        MenuStrip1.Enabled = True
        MASTERToolStripMenuItem.Enabled = False
        TRANSAKSIToolStripMenuItem.Enabled = False
        LAPORANToolStripMenuItem.Enabled = False
        LOGINToolStripMenuItem.Visible = True
        LOGINToolStripMenuItem.Enabled = True
        countAbsensi = 0
        Timer1.Start()
    End Sub

    Private Sub TRANSAKSIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSAKSIPEMBELIANToolStripMenuItem.Click
        FormTransaksiPembelian.MdiParent = Me
        FormTransaksiPembelian.Show()
        FormTransaksiPembelian.Focus()
    End Sub

    Private Sub RETURPEMBELIANToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TRANSAKSIPENJUALANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSAKSIPENJUALANToolStripMenuItem.Click
        FormTransaksiPenjualan.MdiParent = Me
        FormTransaksiPenjualan.Show()
        FormTransaksiPenjualan.Focus()
    End Sub

    Private Sub USERToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormDataUser.MdiParent = Me
        FormDataUser.Show()
        FormDataUser.Focus()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Lbl_Date.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub DATAAPOTEKERToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormDataApoteker.MdiParent = Me
        FormDataApoteker.Show()
        FormDataApoteker.Focus()

    End Sub

    Private Sub TRANSAKSIPENJUALANToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TRANSAKSIPENJUALANToolStripMenuItem1.Click
        FormLaporanPenjualan.MdiParent = Me
        FormLaporanPenjualan.Show()
        FormLaporanPenjualan.Focus()
    End Sub

    Private Sub DATADOKTERToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormDataAsistenApoteker.MdiParent = Me
        FormDataAsistenApoteker.Show()
        FormDataAsistenApoteker.Focus()
    End Sub

    Private Sub DATAToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        FormDataKasir.MdiParent = Me
        FormDataKasir.Show()
        FormDataKasir.Focus()
    End Sub

    Private Sub DATASUPPLIERToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATASUPPLIERToolStripMenuItem.Click
        FormDataSupplier.MdiParent = Me
        FormDataSupplier.Show()
        FormDataSupplier.Focus()
    End Sub

    Private Sub DATASATUANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATASATUANToolStripMenuItem.Click
        FormDataSatuan.MdiParent = Me
        FormDataSatuan.Show()
        FormDataSatuan.Focus()
    End Sub

    Private Sub DATAKATEGORIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATAKATEGORIToolStripMenuItem.Click
        FormDataKategori.MdiParent = Me
        FormDataKategori.Show()
        FormDataKategori.Focus()
    End Sub

    Private Sub DATAOBATToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DATAOBATToolStripMenuItem2.Click
        FormDataObats.MdiParent = Me
        FormDataObats.Show()
        FormDataObats.Focus()
    End Sub

    Private Sub DATALOKASIOBATToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATALOKASIOBATToolStripMenuItem.Click
        FormDataLokasi.MdiParent = Me
        FormDataLokasi.Show()
        FormDataLokasi.Focus()
    End Sub

    Private Sub UPDATEHARGAOBATToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEHARGAOBATToolStripMenuItem.Click
        FormMaintenanceDataObats.MdiParent = Me
        FormMaintenanceDataObats.Show()
        FormMaintenanceDataObats.Focus()

    End Sub

    Private Sub DATAOBATToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DATAOBATToolStripMenuItem1.Click
        FormLaporanStokObat.MdiParent = Me
        FormLaporanStokObat.Show()
        FormLaporanStokObat.Focus()
    End Sub

    Private Sub OBATEXPIREDHILANGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OBATEXPIREDHILANGToolStripMenuItem.Click
        FormLaporanObatExpiredHabis.MdiParent = Me
        FormLaporanObatExpiredHabis.Show()
        FormLaporanObatExpiredHabis.Focus()
    End Sub

    Private Sub HISTORYHARGAOBATToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HISTORYHARGAOBATToolStripMenuItem.Click
        FormLaporanHistoryHarga.MdiParent = Me
        FormLaporanHistoryHarga.Show()
        FormLaporanHistoryHarga.Focus()

    End Sub

    Private Sub TRANSAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSAToolStripMenuItem.Click
        FormLaporanPembelian.MdiParent = Me
        FormLaporanPembelian.Show()
        FormLaporanPembelian.Focus()
    End Sub

    Private Sub DATAADMINToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormDataAdmin.MdiParent = Me
        FormDataAdmin.Show()
        FormDataAdmin.Focus()
    End Sub

    Private Sub DATAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATAToolStripMenuItem.Click
        FormDataUser.MdiParent = Me
        FormDataUser.Show()
        FormDataUser.Focus()
    End Sub

    Private Sub ABSENSIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ABSENSIToolStripMenuItem.Click
        FormLaporanAbsensi.MdiParent = Me
        FormLaporanAbsensi.Show()
        FormLaporanAbsensi.Focus()
    End Sub

    Private Sub LAPORANLABARUGIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LAPORANLABARUGIToolStripMenuItem.Click
        FormLaporanLabaRugi.MdiParent = Me
        FormLaporanLabaRugi.Show()
        FormLaporanLabaRugi.Focus()
    End Sub
    Private Sub TRANSAKSIKEUANGANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSAKSIKEUANGANToolStripMenuItem.Click
        FormTransaksiKeuangan.MdiParent = Me
        FormTransaksiKeuangan.Show()
        FormTransaksiKeuangan.Focus()
    End Sub
    Private Sub FormMenu_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        tmpDate = Format(Date.Now, "yyyy-MM-dd")
        tmpTime = Format(DateTime.Now, "HH:mm:ss")
        tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & idUser & FormMenu.statusLog
        simpan = "SELECT COUNT(*) FROM tabsensi WHERE IDAbsensi = '" & tmpAbsensi & "'"
        periksasql(simpan)
        If Not (countAbsensi = 0) And Not (FormMenu.statusLog = 0) Then

            tmpDate = Format(Date.Now, "yyyy-MM-dd")
            tmpTime = Format(DateTime.Now, "HH:mm:ss")
            tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & idUser & FormMenu.statusLog


            'simpan = "INSERT INTO tabsensi(IDAbsensi, IDUser, tglMasuk, wktMasuk) " +
            '                       "VALUES ('" & tmpAbsensi & "','" & idUser & "','" & tmpDate & "','" & tmpTime & "')"

            'simpan = "UPDATE 'tabsensi' SET('wktKeluar') " +
            '                      "VALUES ('" & tmpTime & "') WHERE `IDAbsensi` = '" & tmpAbsensi & "' "
            simpan = "UPDATE tabsensi SET wktKeluar= '" & tmpTime & "' WHERE IDAbsensi= '" & tmpAbsensi & "' "
            jalankansql(simpan)
        End If
        
        For Each prog As Process In Process.GetProcesses
            If prog.ProcessName = "Apotik_1A" Then
                prog.Kill()
            End If
            If prog.ProcessName = "Apotik_1A.vshost" Then
                prog.Kill()
            End If
        Next
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

    Private Sub LOGINToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGINToolStripMenuItem.Click
        FormLogin.Show()
        FormLogin.BringToFront()

    End Sub

    Private Sub KELUARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KELUARToolStripMenuItem.Click
        tmpDate = Format(Date.Now, "yyyy-MM-dd")
        tmpTime = Format(DateTime.Now, "HH:mm:ss")
        tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & idUser & FormMenu.statusLog
        simpan = "SELECT COUNT(*) FROM tabsensi WHERE IDAbsensi = '" & tmpAbsensi & "'"
        periksasql(simpan)
        If Not (countAbsensi = 0) And Not (FormMenu.statusLog = 0) Then

            tmpDate = Format(Date.Now, "yyyy-MM-dd")
            tmpTime = Format(DateTime.Now, "HH:mm:ss")
            tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & idUser & FormMenu.statusLog


            'simpan = "INSERT INTO tabsensi(IDAbsensi, IDUser, tglMasuk, wktMasuk) " +
            '                       "VALUES ('" & tmpAbsensi & "','" & idUser & "','" & tmpDate & "','" & tmpTime & "')"

            'simpan = "UPDATE 'tabsensi' SET('wktKeluar') " +
            '                      "VALUES ('" & tmpTime & "') WHERE `IDAbsensi` = '" & tmpAbsensi & "' "
            simpan = "UPDATE tabsensi SET wktKeluar= '" & tmpTime & "' WHERE IDAbsensi= '" & tmpAbsensi & "' "
            jalankansql(simpan)
        End If
        FormLogin.Show()
        KELUARToolStripMenuItem.Visible = False
        LOGINToolStripMenuItem.Visible = True
        LOGINToolStripMenuItem.Enabled = True
        MASTERToolStripMenuItem.Enabled = False
        TRANSAKSIToolStripMenuItem.Enabled = False
        LAPORANToolStripMenuItem.Enabled = False
        FormDataApoteker.Close()
        FormLaporanPenjualan.Close()
        FormTransaksiPembelian.Close()
        FormTransaksiPenjualan.Close()
        FormDataUser.Close()
        FormDataAsistenApoteker.Close()
        FormDataKasir.Close()
        FormDataSupplier.Close()
        FormDataSatuan.Close()
        FormDataKategori.Close()
        FormDataObats.Close()
        FormDataLokasi.Close()
        FormMaintenanceDataObats.Close()
        FormLaporanStokObat.Close()
        FormLaporanObatExpiredHabis.Close()
        FormLaporanHistoryHarga.Close()
        FormLaporanPembelian.Close()
        FormDataAdmin.Close()
        FormDataUser.Close()
        FormLaporanAbsensi.Close()
        FormLaporanLabaRugi.Close()
        FormTransaksiKeuangan.Close()
    End Sub

    Private Sub FormMenu_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        
    End Sub

    Private Sub FormMenu_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F1 Then

        End If
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    
    Private Sub DATAOPNAMEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DATAOPNAMEToolStripMenuItem.Click
        FormDataOpname.MdiParent = Me
        FormDataOpname.Show()
        FormDataOpname.Focus()
    End Sub
End Class
