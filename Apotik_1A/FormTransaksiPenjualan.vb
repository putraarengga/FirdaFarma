Imports MySql.Data.MySqlClient
Imports System.Data.OleDb

Public Class FormTransaksiPenjualan
    Dim databaru As Boolean
    Dim selectDataBase As String
    Dim simpan, simpan2 As String
    Dim persenDiskon As Integer
    Dim day, month, year, vdate As String
    Dim selectDataBase2, JenisTransaksi As String
    Dim nowTime As String
    Dim GetIDSupplier, GetIDObat As String
    Dim reader As MySqlDataReader
    Dim jumlah, Sisa, persen_diskon, harga_diskon, hargajual As Integer
    Dim tgltranskasi As Date
    Dim Urutan As Integer
    Dim hitung, bayar, kembalian, diskon, Total_harga As Integer
    Dim dt1, dt2 As DateTime
    Dim drugname, jenistrans As String
    Dim amount, price, totalprice, discprice As Integer
    Dim hargadiskon
    Dim panjang As Integer = 100
    Dim i As Integer
    Dim satuan, satuanlv1, satuanlv2, satuanlv3, satbeli As String
    Dim stokobatLv1, stokobatLv2, stokobatLv3 As Integer
    Dim sisaobatLv1, sisaobatLv2, sisaobatLv3 As Integer
    Dim sisobatLv1, sisobatLv2, sisobatLv3 As Integer
    Dim hasilmodulus1, hasilmodulus2 As Integer
    Dim dataSelected As Boolean
    Dim grandTotal, kuranggrand1, kuranggrand2 As Integer
    Dim Factur, Pkonversi2, Pkonversi3 As Integer
    Dim BuyDate As String
    Dim StringHitung, Stringtotal, Stringkurang As String
    Shared Property indexObat As String
    Shared Property indexSupplier As String
    Shared Property stok As Integer
    'Shared Property Satuans As String
    Shared Property HJU1 As Integer
    Shared Property MinimumPembelianlv1 As Integer
    Shared Property DiskonPembelianlv1 As Integer
    Shared Property HJU2 As Integer
    Shared Property MinimumPembelianlv2 As Integer
    Shared Property DiskonPembelianlv2 As Integer
    Shared Property HJU3 As Integer
    Shared Property MinimumPembelianlv3 As Integer
    Shared Property DiskonPembelianlv3 As Integer
    Shared Property HargaJualResep As Integer
    Shared Property Kadaluarsa As String
    Shared Property FakturJual As Integer
    Shared Property konversiLv2 As Integer
    Shared Property konversiLv3 As Integer
    Shared Property Urutan_Faktur As Integer
    Shared Property total As Integer
    Dim TextToPrint As String = ""


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormTransaksiPenjualan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vdate = Format(DateTime.Now, "dd/MM/yyyy")
        FormDisabeld()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        dt1 = DateTimePicker1.Value
        TextBox4.Text = FormMenu.fullName

        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")

        PrintDocument1.PrinterSettings.PrinterName = "POS-58"
        'PrintDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer"

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FormPencarianStokObat.IDPencariObat = 3
        FormPencarianStokObat.Show()
        FormPencarianStokObat.Focus()
        Button8.Enabled = True
        FormEnabeld4()

    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs)
        FormPencarianDataPelanggan.Show()
        FormPencarianDataPelanggan.Focus()
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        FormDataRacikan.Show()
        FormDataRacikan.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        FormDataPelanggan.Show()
        FormDataPelanggan.Focus()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
    End Sub

    Sub FormEnabeld()
        TextBox18.Enabled = True
        TextBox19.Enabled = True
        TextBox20.Enabled = True
        TextBox12.Enabled = True

        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox5.Enabled = True

        TextBox6.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True

        Button5.Enabled = True

        Button8.Enabled = True
        Button9.Enabled = True

    End Sub
    Sub FormEnabeld2()
        TextBox18.Enabled = True
        TextBox19.Enabled = True
        TextBox20.Enabled = True
        TextBox12.Enabled = True
        Button9.Enabled = True
        Button6.Enabled = True
    End Sub

    Sub FormEnabeld3()

        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox5.Enabled = True

        TextBox6.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True

        Button5.Enabled = True
        Button8.Enabled = True
    End Sub

    Sub FormEnabeld4()

        Button5.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Sub FormDisabeld()
        TextBox18.Enabled = False
        TextBox19.Enabled = False
        TextBox20.Enabled = False
        TextBox12.Enabled = False

        '        TextBox11.Enabled = False

        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox5.Enabled = False

        TextBox6.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False

        Button5.Enabled = False
        Button6.Enabled = False
        '        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False

        RadioButton1.Enabled = False
        RadioButton1.Checked = False
        RadioButton2.Enabled = False

    End Sub

    Sub FormDisabeld2()
        TextBox18.Enabled = False
        TextBox19.Enabled = False
        TextBox20.Enabled = False
        TextBox12.Enabled = False
        Button9.Enabled = False
        Button6.Enabled = False
    End Sub

    Sub FormDisable3()
        '        TextBox11.Enabled = False

        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox5.Enabled = False

        TextBox6.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False

        Button5.Enabled = False
        '        Button7.Enabled = False
        Button8.Enabled = False
    End Sub

    Sub FormDisabeld4()

        TextBox2.Enabled = False
        TextBox5.Enabled = False
        Button5.Enabled = False
        ComboBox1.Enabled = False
    End Sub

    Sub FormEnabeld5()
        TextBox18.Enabled = True
        TextBox19.Enabled = True
        TextBox20.Enabled = True
        TextBox12.Enabled = True

    End Sub

    Sub FormDisabeld5()
        TextBox18.Enabled = False
        TextBox19.Enabled = False
        TextBox20.Enabled = False
        TextBox12.Enabled = False

    End Sub

    Sub Bersih()
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox12.Text = ""

        '        TextBox11.Text = ""

        TextBox1.Text = ""

        TextBox2.Text = ""
        TextBox3.Text = 0
        TextBox5.Text = 0

        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        RichTextBox1.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub
    Sub Bersih2()
        TextBox2.Text = ""
        TextBox3.Text = 0
        TextBox5.Text = 0
        TextBox22.Text = ""
        ComboBox1.Text = ""
        TextBox9.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox21.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox13.Text = ""
    End Sub
    Sub IsiGridUmum1()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase = "SELECT ttransaksi.FakturPenjualan, ttransaksi.TglTransaksi, ttransaksi.JenisTransaksi," +
            "ttransaksi.NamaObat, ttransaksi.HargaJualUmum, ttransaksi.Jumlah, ttransaksi.Satuan, " +
            "ttransaksi.hargadiskon, ttransaksi.TotalHarga, ttransaksi.Kadarluarsa, ttransaksi.FakturBeliObat  " +
            "FROM ttransaksi WHERE ttransaksi.TglTransaksi ='" & nowTime & "' and ttransaksi.FakturPenjualan='" & TextBox1.Text & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "ttransaksi")
        DataGridView1.DataSource = (DS.Tables("ttransaksi"))
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nomor Faktur"
            .Columns(1).HeaderCell.Value = "Tanggal Transaksi"
            .Columns(2).HeaderCell.Value = "Jenis Transaksi"
            .Columns(3).HeaderCell.Value = "Nama Obat"
            .Columns(4).HeaderCell.Value = "Harga Jual"
            .Columns(5).HeaderCell.Value = "Jumlah Obat"
            .Columns(6).HeaderCell.Value = "Satuan Beli Obat"
            .Columns(7).HeaderCell.Value = "Diskon Pembelian"
            .Columns(8).HeaderCell.Value = "Sub Total"
            .Columns(9).HeaderCell.Value = "Tanggal Kadaluarsa"
            .Columns(10).HeaderCell.Value = "Keterangan"
        End With
        
    End Sub

    Sub NomorFakturPembelian()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")

        selectDataBase = "SELECT ttransaksi.TglTransaksi FROM ttransaksi ORDER BY TglTransaksi DESC LIMIT 1 "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            tgltranskasi = DT.Rows(0).Item("TglTransaksi")
            If tgltranskasi <> nowTime Then
                Urutan_Faktur = 1
                TextBox1.Text = Format(DateTime.Now, "yyyyMMdd") & Urutan_Faktur
            ElseIf tgltranskasi = nowTime Then
                indexSupplier = -1
                selectDataBase = "SELECT ttransaksi.Urutan FROM ttransaksi ORDER BY IDTransaksi DESC LIMIT 1"
                bukaDB()
                DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
                DS = New DataSet
                DT = New DataTable
                DS.Clear()
                DA.Fill(DT)
                If DT.Rows.Count > 0 Then
                    Urutan = DT.Rows(0).Item("Urutan")
                    Urutan_Faktur = Urutan + 1
                    TextBox1.Text = Format(DateTime.Now, "yyyyMMdd") & Urutan_Faktur
                End If
            End If
        ElseIf DT.Rows.Count = 0 Then
            Urutan_Faktur = 1
            TextBox1.Text = Format(DateTime.Now, "yyyyMMdd") & Urutan_Faktur
        End If
    End Sub

    Function Ceiling(number As Double) As Long
        Ceiling = -Int(-number)
    End Function

    Sub HitungJumlah()
        satuan = ComboBox1.SelectedItem
        satuanlv1 = TextBox14.Text.ToString
        satuanlv2 = TextBox15.Text.ToString
        satuanlv3 = TextBox17.Text.ToString
        jumlah = Val(TextBox5.Text)
        stokobatLv1 = Val(TextBox9.Text)
        stokobatLv2 = Val(TextBox16.Text)
        stokobatLv3 = Val(TextBox21.Text)

        If satuan = "" Then
            TextBox5.Text = 0
        ElseIf String.Equals(satuan, satuanlv1) Then
            If (jumlah > stokobatLv1) Then
                MsgBox("Stok Obat Tidak Mencukupi", MsgBoxStyle.OkCancel)
                jumlah = 0
                TextBox5.Text = 0
            ElseIf (jumlah <= stokobatLv1) Then
                jumlah = Val(TextBox5.Text)
            End If
        ElseIf String.Equals(satuan, satuanlv2) Then
            If (jumlah > stokobatLv2) Then
                MsgBox("Stok Obat Tidak Mencukupi", MsgBoxStyle.OkCancel)
                jumlah = 0
                TextBox5.Text = 0
            ElseIf (jumlah <= stokobatLv2) Then
                jumlah = Val(TextBox5.Text)
            End If
        ElseIf String.Equals(satuan, satuanlv3) Then
            If (jumlah > stokobatLv3) Then
                MsgBox("Stok Obat Tidak Mencukupi", MsgBoxStyle.OkCancel)
                jumlah = 0
                TextBox5.Text = 0
            ElseIf (jumlah <= stokobatLv3) Then
                jumlah = Val(TextBox5.Text)
            End If
        End If
    End Sub

    Sub HitungData()
        satuan = ""
        satuan = ComboBox1.SelectedItem
        satuanlv1 = TextBox14.Text.ToString
        satuanlv2 = TextBox15.Text.ToString
        satuanlv3 = TextBox17.Text.ToString
        jumlah = Val(TextBox5.Text)

        If satuan = "" Then
            TextBox6.Text = 0
            TextBox8.Text = 0
        ElseIf String.Equals(satuan, satuanlv1) Then
            TextBox6.Text = HJU1
            If jumlah >= MinimumPembelianlv1 Then
                TextBox8.Text = DiskonPembelianlv1.ToString
                Total_harga = (HJU1 * jumlah) - DiskonPembelianlv1
                TextBox13.Text = Total_harga.ToString
                hargadiskon = DiskonPembelianlv1
            Else
                TextBox8.Text = 0
                Total_harga = (HJU1 * jumlah)
                TextBox13.Text = Total_harga.ToString
                hargadiskon = 0
            End If
        ElseIf String.Equals(satuan, satuanlv2) Then
            TextBox6.Text = HJU2
            If jumlah >= MinimumPembelianlv2 Then
                TextBox8.Text = DiskonPembelianlv2.ToString
                Total_harga = (HJU2 * jumlah) - DiskonPembelianlv2
                TextBox13.Text = Total_harga.ToString
                hargadiskon = DiskonPembelianlv2
            Else
                TextBox8.Text = 0
                Total_harga = (HJU2 * jumlah)
                TextBox13.Text = Total_harga.ToString
                hargadiskon = 0
            End If
        ElseIf String.Equals(satuan, satuanlv3) Then
            TextBox6.Text = HJU3
            If jumlah >= MinimumPembelianlv3 Then
                TextBox8.Text = DiskonPembelianlv3.ToString
                Total_harga = (HJU3 * jumlah) - DiskonPembelianlv3
                TextBox13.Text = Total_harga.ToString
                hargadiskon = DiskonPembelianlv3
            Else
                TextBox8.Text = 0
                Total_harga = (HJU3 * jumlah)
                TextBox13.Text = Total_harga.ToString
                hargadiskon = 0
            End If
        End If

    End Sub

    Sub HitungData2()
       
        bayar = Val(TextBox3.Text)
        hitung = Val(TextBox7.Text)
        If bayar > 0 Then
            kembalian = bayar - hitung
            StringHitung = FormatNumber(kembalian)
            TextBox10.Text = StringHitung
            If kembalian >= 0 Then
                Button10.Enabled = True
            Else
                Button10.Enabled = False
            End If
        End If
    End Sub

    Sub HitungSisa()
        hasilmodulus1 = 0
        hasilmodulus2 = 0
        satuan = ComboBox1.SelectedItem
        satuanlv1 = TextBox14.Text.ToString
        satuanlv2 = TextBox15.Text.ToString
        satuanlv3 = TextBox17.Text.ToString
        jumlah = Val(TextBox5.Text)
        stokobatLv1 = Val(TextBox9.Text)
        stokobatLv2 = Val(TextBox16.Text)
        stokobatLv3 = Val(TextBox21.Text)

        If satuan = "" Then
            TextBox5.Text = 0
        ElseIf String.Equals(satuan, satuanlv1) Then
            If (jumlah <= stokobatLv1) Then
                sisaobatLv1 = stokobatLv1 - jumlah
                sisaobatLv2 = stokobatLv2 - (konversiLv2 * jumlah)
                sisaobatLv3 = stokobatLv3 - (konversiLv2 * konversiLv3 * jumlah)
            End If
        ElseIf String.Equals(satuan, satuanlv2) Then
            If (jumlah <= stokobatLv2) Then
                'sisaobatLv2 = stokobatLv2 - jumlah
                'sisaobatLv3 = stokobatLv3 - (konversiLv3 * jumlah)
                'sisaobatLv1 = stokobatLv1 - Ceiling((jumlah / konversiLv2))
                'TextBox3.Text = jumlah / konversiLv2
                'TextBox10.Text = Ceiling((jumlah / konversiLv2))

                sisaobatLv2 = stokobatLv2 - jumlah
                sisaobatLv3 = stokobatLv3 - (konversiLv3 * jumlah)
                sisaobatLv1 = Int((sisaobatLv2 / konversiLv2))

            End If
        ElseIf String.Equals(satuan, satuanlv3) Then
            If (jumlah <= stokobatLv3) Then
                'hasilmodulus1 = stokobatLv3 Mod konversiLv3
                'sisaobatLv3 = stokobatLv3 - jumlah
                'hasilmodulus2 = sisaobatLv3 Mod konversiLv3
                'If (hasilmodulus1 = 0 And hasilmodulus2 = 0) Then
                '    sisaobatLv3 = sisaobatLv3
                '    sisaobatLv1 = stokobatLv1 - Ceiling((jumlah / (konversiLv2 * konversiLv3)))
                '    sisaobatLv2 = stokobatLv2 - Ceiling((jumlah / konversiLv3))
                'Else
                '    sisaobatLv3 = sisaobatLv3
                '    sisaobatLv1 = stokobatLv1
                '    sisaobatLv2 = stokobatLv2
                'End If

                'sisaobatLv3 = stokobatLv3 - jumlah
                'sisaobatLv1 = stokobatLv1 - Ceiling((jumlah / (konversiLv2 * konversiLv3)))
                'sisaobatLv2 = stokobatLv2 - Ceiling((jumlah / konversiLv3))
                sisaobatLv3 = stokobatLv3 - jumlah
                sisaobatLv1 = Int((sisaobatLv3 / (konversiLv2 * konversiLv3)))
                sisaobatLv2 = Int((sisaobatLv3 / konversiLv3))
            End If
        End If
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        vdate = Format(DateTime.Now, "yyyy-MM-dd")

        If (TextBox5.Text = "" Or TextBox5.Text = 0 Or ComboBox1.SelectedItem = "") Then
            MsgBox("Isi Dulu Jumlah Barang Yang Dibeli", MsgBoxStyle.OkCancel)
        Else
            HitungData()
            HitungSisa()

            simpan = "INSERT INTO ttransaksi(Urutan,FakturPenjualan,JenisTransaksi," +
                "NamaObat, HargaJualUmum, Jumlah, Satuan, " +
                "TotalHarga, TglTransaksi,hargadiskon," +
                "Kadarluarsa, FakturBeliObat)" +
    "VALUES( '" & Urutan_Faktur & "','" & TextBox1.Text & "','" & JenisTransaksi & "'," +
    "'" & TextBox2.Text & "','" & TextBox6.Text & "','" & TextBox5.Text & "', '" & ComboBox1.Text & "'," +
    "'" & TextBox13.Text & "','" & vdate & "','" & hargadiskon & "'," +
    "'" & Kadaluarsa & "', '" & FakturJual & "')"
            jalankansql(simpan)
            simpan2 = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "', tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                "WHERE tpembelian.NamaObat = '" & TextBox2.Text & "' AND tpembelian.NomorFaktur = '" & FakturJual & "'"
            jalankansql(simpan2)
            IsiGridUmum1()
            Bersih2()
            Button11.Enabled = True
            'TextBox3.Enabled = True
            Button8.Enabled = False
            'Button2.Enabled = True
            TextBox5.Enabled = False
            ComboBox1.Enabled = False

        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        total = 0
        For x As Integer = 0 To DataGridView1.RowCount - 1
            total += DataGridView1.Rows(x).Cells(8).Value
        Next
        Stringtotal = FormatNumber(total)
        RichTextBox1.Text = Stringtotal
        TextBox7.Text = total.ToString
        Button11.Enabled = False
        TextBox3.Enabled = True
        TextBox3.Focus()
    End Sub
    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim pesan As String

        If (TextBox3.Text = "" Or TextBox3.Text = 0) Then
            MsgBox("Masukkan Pembayaran Uang", MsgBoxStyle.OkCancel)
            FormPencarianDataObat.Close()
        Else
            If kembalian < 0 Then
                MsgBox("Pembayaran Uang Kurang", MsgBoxStyle.OkCancel)
            Else
                pesan = MsgBox("Apakah Ingin Cetak Struk?" + TextBox1.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
                If pesan = MsgBoxResult.Yes Then
                    simpan2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & hitung & "', ttransaksi.Bayar= '" & bayar & "', ttransaksi.Kembalian= '" & kembalian & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(simpan2)
                    PrintHeader()
                    ItemsToBePrinted1()
                    ItemsToBePrinted2()
                    printFooter()
                    PrintFooter2()
                    Dim printControl = New Printing.StandardPrintController
                    PrintDocument1.PrintController = printControl
                    Try
                        PrintDocument1.Print()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                    Bersih()
                    FormDisabeld()
                    Button2.Enabled = True
                    RichTextBox1.Text = ""
                    RichTextBox1.Text = 0
                    TextBox1.Text = ""
                    Bersih2()
                    'Me.Close()

                ElseIf pesan = MsgBoxResult.No Then
                    simpan2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & hitung & "', ttransaksi.Bayar= '" & bayar & "', ttransaksi.Kembalian= '" & kembalian & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(simpan2)
                    Bersih()
                    FormDisabeld()
                    Button2.Enabled = True
                    RichTextBox1.Text = ""
                    RichTextBox1.Text = 0
                    TextBox1.Text = ""
                    Bersih2()
                    'Me.Close()
                    Me.Refresh()
                    Exit Sub
                End If
            End If
        End If

    End Sub
    Sub GetIndeks(ByVal x As Integer)
        Try
            drugname = DataGridView1.Rows(x).Cells(3).Value
            amount = DataGridView1.Rows(x).Cells(5).Value
            price = DataGridView1.Rows(x).Cells(4).Value
            totalprice = DataGridView1.Rows(x).Cells(8).Value
            satbeli = DataGridView1.Rows(x).Cells(6).Value
            BuyDate = DataGridView1.Rows(x).Cells(9).Value.ToString
            Factur = DataGridView1.Rows(x).Cells(10).Value
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        Konversi()
        HitungSisa2()
        dataSelected = True
        Button4.Enabled = True

    End Sub
    Sub Konversi()
        selectDataBase = "SELECT TP.KonversiLv2 AS KLV2, TP.KonversiLv3 AS KLV3  FROM tpembelian AS TP WHERE TP.NamaObat ='" & drugname & "' AND TP.NomorFaktur='" & Factur & "' "

        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
                Pkonversi2 = DT.Rows(i).Item("KLV2")
                Pkonversi3 = DT.Rows(i).Item("KLV3")
            Next
        End If
    End Sub
   
    Sub HitungSisa2()
        Konversi()

        selectDataBase = "SELECT TP.SisaObatLv1 AS SOLV1, TP.SatuanLv1 AS SAT1, TP.SisaObatLv2 AS SOLV2, TP.SatuanLv2 AS SAT2," +
                    "TP.SisaObatLv3 AS SOLV3, TP.SatuanLv3 AS SAT3 FROM tpembelian AS TP WHERE TP.NamaObat = '" & drugname & "' AND TP.NomorFaktur = '" & Factur & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            For i As Integer = 0 To DT.Rows.Count - 1
               
                If Equals(DT.Rows(i).Item("SAT1"), satbeli) Then
                    sisaobatLv1 = DT.Rows(i).Item("SOLV1") + amount
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") + (Pkonversi2 * amount)
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + (Pkonversi2 * Pkonversi3 * amount)
                ElseIf Equals(DT.Rows(i).Item("SAT2"), satbeli) Then
                   
                    sisaobatLv2 = DT.Rows(i).Item("SOLV2") + amount
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + (Pkonversi3 * amount)
                    sisaobatLv1 = Int((sisaobatLv2 / Pkonversi2))
                ElseIf Equals(DT.Rows(i).Item("SAT3"), satbeli) Then
                   
                    sisaobatLv3 = DT.Rows(i).Item("SOLV3") + amount
                    sisaobatLv1 = Int((sisaobatLv3 / (Pkonversi2 * Pkonversi3)))
                    sisaobatLv2 = Int((sisaobatLv3 / Pkonversi3))
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql, updatedata As String
        Dim pesan As String
        If drugname = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + drugname, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            If Equals(DataGridView1.Rows(i).Cells(2).FormattedValue.ToString(), "Umum") Then
                If Button11.Enabled = True Then
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    jalankansql(hapussql)

                    DataGridView1.Refresh()

                    HitungSisa2()
                    updatedata = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "',tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                                 " WHERE (tpembelian.NamaObat ='" & drugname & "') AND tpembelian.NomorFaktur='" & Factur & "'"
                    jalankansql(updatedata)
                    IsiGridUmum1()
                Else
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    kuranggrand1 = Val(TextBox7.Text)
                    kuranggrand2 = kuranggrand1 - totalprice
                    RichTextBox1.Refresh()
                    Stringkurang = FormatNumber(kuranggrand2)
                    RichTextBox1.Text = Stringkurang
                    'RichTextBox1.Text = kuranggrand2.ToString
                    jalankansql(hapussql)

                    DataGridView1.Refresh()
                    grandTotal = Val(TextBox7.Text)
                    selectDataBase2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & grandTotal & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(selectDataBase2)

                    HitungSisa2()
                    updatedata = " UPDATE tpembelian SET tpembelian.SisaObatLv1 = '" & sisaobatLv1 & "', tpembelian.SisaObatLv2 = '" & sisaobatLv2 & "',tpembelian.SisaObatLv3 = '" & sisaobatLv3 & "'" +
                                 " WHERE (tpembelian.NamaObat ='" & drugname & "') AND tpembelian.NomorFaktur='" & Factur & "'"
                    jalankansql(updatedata)

                    IsiGridUmum1()
                End If

            ElseIf Equals(DataGridView1.Rows(i).Cells(2).FormattedValue.ToString(), "Racikan") Then
                If Button11.Enabled = True Then
                    FormDataRacikan.Show()
                    FormDataRacikan.Focus()
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    jalankansql(hapussql)

                    DataGridView1.Refresh()
                    selectDataBase2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & grandTotal & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(selectDataBase2)
                    IsiGridUmum1()
                Else
                    FormDataRacikan.Show()
                    FormDataRacikan.Focus()
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    kuranggrand1 = Val(TextBox7.Text)
                    kuranggrand2 = kuranggrand1 - totalprice
                    RichTextBox1.Refresh()
                    Stringkurang = FormatNumber(kuranggrand2, 2)
                    RichTextBox1.Text = Stringkurang
                    'RichTextBox1.Text = kuranggrand2.ToString
                    jalankansql(hapussql)

                    DataGridView1.Refresh()
                    grandTotal = Val(TextBox7.Text)
                    selectDataBase2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & grandTotal & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(selectDataBase2)
                    IsiGridUmum1()

                    'FormDataRacikan.Flag = 1
                End If
            ElseIf Equals(DataGridView1.Rows(i).Cells(2).FormattedValue.ToString(), "Resep") Then
                If Button11.Enabled = True Then
                    FormDataResep.Show()
                    FormDataResep.Focus()
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    jalankansql(hapussql)

                    DataGridView1.Refresh()
                    selectDataBase2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & grandTotal & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(selectDataBase2)
                    IsiGridUmum1()
                Else
                    FormDataResep.Show()
                    FormDataResep.Focus()
                    hapussql = "DELETE FROM ttransaksi WHERE ttransaksi.FakturPenjualan= '" & TextBox1.Text & "' AND ttransaksi.NamaObat ='" & drugname & "'"
                    kuranggrand1 = Val(TextBox7.Text)
                    kuranggrand2 = kuranggrand1 - totalprice
                    RichTextBox1.Refresh()
                    Stringkurang = FormatNumber(kuranggrand2, 2)
                    RichTextBox1.Text = Stringkurang
                    'RichTextBox1.Text = kuranggrand2.ToString
                    jalankansql(hapussql)

                    DataGridView1.Refresh()
                    grandTotal = Val(TextBox7.Text)
                    selectDataBase2 = " UPDATE ttransaksi SET ttransaksi.GrandTotal = '" & grandTotal & "' WHERE ttransaksi.FakturPenjualan = '" & TextBox1.Text & "' "
                    jalankansql(selectDataBase2)
                    IsiGridUmum1()

                End If
            End If
        End If
        Button4.Enabled = False

    End Sub
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        HitungJumlah()
        HitungData()
        HitungSisa()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If ComboBox1.SelectedItem = "" Then
            TextBox5.Enabled = False
        Else
            TextBox5.Enabled = True
            If IsNumeric(TextBox5.Text) Then
                HitungJumlah()
                HitungData()
                HitungSisa()
            Else
                MsgBox("Bukan Angka", MsgBoxStyle.OkCancel)
                TextBox5.Text = 0
            End If
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True And RadioButton2.Checked = False Then
            FormEnabeld4()
            Button6.Enabled = False
            '            CheckBox1.Enabled = False
            '            CheckBox1.Checked = False
            'NomorFakturPembelian()
            JenisTransaksi = "Umum"
            'RadioButton2.Checked = False
            'RadioButton2.Enabled = False
            'ElseIf RadioButton2.Checked = True Then
            '    Button6.Enabled = True
            '    RadioButton1.Checked = False
            '    RadioButton1.Enabled = False
            '    'NomorFakturPembelian()
            '    JenisTransaksi = "Resep"
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True And RadioButton1.Checked = False Then
            FormDisabeld4()
            'FormEnabeld5()
            Button6.Enabled = True
        End If
            
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        NomorFakturPembelian()
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs)
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox3_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) Then
            HitungData()
            HitungData2()
        Else
            MsgBox("Bukan Angka", MsgBoxStyle.OkCancel)
            TextBox3.Text = 0
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        HitungData()
        HitungData2()
        'HitungJumlah()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        HitungJumlah()
        HitungData()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "" Then
            TextBox5.Enabled = False
        Else
            TextBox5.Enabled = True
            HitungJumlah()
            HitungData()
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        HitungJumlah()
        HitungData()
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        HitungJumlah()
        HitungData()
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


    Public Sub PrintHeader()

        TextToPrint = ""

        TextToPrint &= Environment.NewLine
        Dim StringToPrint As String = "Apotek Firda Farma"
        Dim LineLen As Integer = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((33 - LineLen) / 2)) 'This line is used to center text in the middle of the receipt
        TextToPrint &= spcLen1 & StringToPrint & Environment.NewLine


        StringToPrint = "Jalan Tambak Rejo,"
        LineLen = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen2 & StringToPrint & Environment.NewLine

        StringToPrint = "Waru-Sidoarjo,"
        LineLen = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen3 & StringToPrint & Environment.NewLine

        StringToPrint = "085109129300"
        LineLen = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen4 & StringToPrint & Environment.NewLine


        StringToPrint = "Jawa Timur"
        LineLen = StringToPrint.Length
        Dim spcLen4b As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen4b & StringToPrint & Environment.NewLine

        StringToPrint = "=================================="
        LineLen = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= StringToPrint & Environment.NewLine

    End Sub

    Public Sub ItemsToBePrinted1()

        'TextToPrint &= " "
        Dim globalLengt As Integer = 0

        Dim NomorFakturPenjualan As String = DataGridView1.Rows(0).Cells(0).FormattedValue.ToString()
        Dim TglPenjualanObat As Date = DataGridView1.Rows(0).Cells(1).FormattedValue.ToString()
        'Dim JenisPenjualanObat As String = DataGridView1.Rows(i).Cells(2).FormattedValue.ToString()


        Dim StringToPrint As String = "Nomor Faktur:"
        Dim StringToPrint2 As String = NomorFakturPenjualan
        Dim LineLen As String
        Dim LineLen2 As String
        globalLengt = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((0)))
        Dim spcLen5b As New String(" "c, Math.Round((3)))
        TextToPrint &= spcLen5 & StringToPrint & spcLen5b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Tgl:"
        StringToPrint2 = TglPenjualanObat
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen6 As New String(" "c, Math.Round(1))
        Dim spcLen6b As New String(" "c, Math.Round(12))
        TextToPrint &= StringToPrint & spcLen6b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Kasir:"
        StringToPrint2 = TextBox4.Text
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen7 As New String(" "c, Math.Round((1)))
        Dim spcLen7b As New String(" "c, Math.Round((10)))
        TextToPrint &= StringToPrint & spcLen7b & StringToPrint2 & Environment.NewLine

        StringToPrint = "=================================="
        LineLen = StringToPrint.Length
        Dim spcLen8 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= StringToPrint & Environment.NewLine

    End Sub

    Public Sub ItemsToBePrinted2()
        '.Columns(0).HeaderCell.Value = "Nomor Faktur"
        '.Columns(1).HeaderCell.Value = "Tanggal Transaksi"
        '.Columns(2).HeaderCell.Value = "Jenis Transaksi"
        '.Columns(3).HeaderCell.Value = "Nama Obat"
        '.Columns(4).HeaderCell.Value = "Harga Jual"
        '.Columns(5).HeaderCell.Value = "Jumlah Obat"
        '.Columns(6).HeaderCell.Value = "Satuan Beli Obat"
        '.Columns(7).HeaderCell.Value = "Diskon Pembelian"
        '.Columns(8).HeaderCell.Value = "Sub Total"
        Dim spcLen0 As New String(" "c, Math.Round((1)))
        Dim spcLeny As New String(" "c, Math.Round((4)))
        i = 0

        While (i < DataGridView1.Rows.Count)

            panjang += DataGridView1.Rows(0).Height

            jenistrans = DataGridView1.Rows(i).Cells(2).FormattedValue.ToString()
            drugname = DataGridView1.Rows(i).Cells(3).FormattedValue.ToString()
            amount = DataGridView1.Rows(i).Cells(5).FormattedValue.ToString()
            price = DataGridView1.Rows(i).Cells(4).FormattedValue.ToString()
            totalprice = DataGridView1.Rows(i).Cells(8).FormattedValue.ToString()
            discprice = DataGridView1.Rows(i).Cells(7).FormattedValue.ToString()


            If Equals(jenistrans, "Umum") Then
                Dim StringToPrint As String = "[" & amount & " x " & price & "]-" & discprice
                Dim PrintCurrency As String = FormatNumber(totalprice, 2, True, True, True)
                'PrintCurrency.Alignment = StringAlignment.Far
                'PrintCurrency.LineAlignment = StringAlignment.Far
                Dim LineLen As String = PrintCurrency.Length
                Dim panjang1 As String = StringToPrint.Length
                Dim spcLen5 As New String(" "c, Math.Round(18 - (spcLen0 + panjang1)))
                TextToPrint &= drugname & Environment.NewLine
                TextToPrint &= StringToPrint & spcLen5 & PrintCurrency & Environment.NewLine
            End If

            If Equals(jenistrans, "Racikan") Then
                Dim StringToPrint As String = "        "
                Dim PrintCurrency As String = FormatNumber(totalprice, 2, True, True, True)
                Dim LineLen As String = PrintCurrency.Length
                Dim spcLen6 As New String(" "c, Math.Round((18)))

                TextToPrint &= "Nama Racikan: " & drugname & Environment.NewLine
                TextToPrint &= spcLen6 & PrintCurrency & Environment.NewLine
            End If

            If Equals(jenistrans, "Resep") Then
                Dim StringToPrint As String = "        "
                Dim PrintCurrency As String = FormatNumber(totalprice, 2, True, True, True)
                Dim LineLen As String = PrintCurrency.Length
                Dim spcLen7 As New String(" "c, Math.Round((18)))

                TextToPrint &= "Resep- Attn: " & drugname & Environment.NewLine
                TextToPrint &= spcLen7 & PrintCurrency & Environment.NewLine
            End If
            i += 1
        End While
    End Sub

    Public Sub printFooter()

        TextToPrint &= Environment.NewLine & Environment.NewLine
        Dim globalLengt As Integer = 0
        Dim StringToPrint2 As String
        Dim PrintCurrency As String = FormatNumber(totalprice, 2, True, True, True)

        Dim StringToPrint As String = "=================================="
        Dim LineLen As String = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= StringToPrint & Environment.NewLine

        StringToPrint = "Total    RP."
        StringToPrint2 = RichTextBox1.Text
        PrintCurrency = FormatNumber(RichTextBox1.Text, 2, True, True, True)
        Dim LineLen2 = PrintCurrency.Length
        globalLengt = PrintCurrency.Length
        Dim spcLen6 As New String(" "c, Math.Round((2)))
        Dim spcLen1 As New String(" "c, Math.Round((13 - LineLen2)))
        TextToPrint &= Environment.NewLine & StringToPrint & spcLen1 & PrintCurrency & Environment.NewLine

        StringToPrint = "CASH     RP."
        StringToPrint2 = TextBox3.Text
        PrintCurrency = FormatNumber(TextBox3.Text, 2, True, True, True)
        LineLen = globalLengt
        Dim spcLen7 As New String(" "c, Math.Round((2)))
        TextToPrint &= Environment.NewLine & StringToPrint & spcLen1 & PrintCurrency & Environment.NewLine

        If (TextBox10.Text.Length < RichTextBox1.Text.Length) Then
            Dim beda As Integer = RichTextBox1.Text.Length - TextBox10.Text.Length
            Dim spcLen2 As New String(" "c, Math.Round(((13 + beda) - LineLen2)))
            StringToPrint = "KEMBALI  RP."
            StringToPrint2 = TextBox10.Text
            PrintCurrency = FormatNumber(TextBox10.Text, 2, True, True, True)
            LineLen = globalLengt
            Dim spcLen8 As New String(" "c, Math.Round((2)))
            TextToPrint &= Environment.NewLine & StringToPrint & spcLen2 & PrintCurrency & Environment.NewLine
        Else
            Dim spcLen2 As New String(" "c, Math.Round((13 - LineLen2)))
            StringToPrint = "KEMBALI  RP."
            StringToPrint2 = TextBox10.Text
            PrintCurrency = FormatNumber(TextBox10.Text, 2, True, True, True)
            LineLen = globalLengt
            Dim spcLen8 As New String(" "c, Math.Round((2)))
            TextToPrint &= Environment.NewLine & StringToPrint & spcLen2 & PrintCurrency & Environment.NewLine

        End If
    End Sub
    Public Sub PrintFooter2()

        TextToPrint &= Environment.NewLine
        Dim StringToPrint As String = "Barang Yang Sudah dibeli"
        Dim LineLen As Integer = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((30 - LineLen) / 2)) 'This line is used to center text in the middle of the receipt
        TextToPrint &= spcLen1 & StringToPrint & Environment.NewLine


        StringToPrint = "Tidak dapat Dikembalikan"
        LineLen = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round((30 - LineLen) / 2))
        TextToPrint &= spcLen2 & StringToPrint & Environment.NewLine

        StringToPrint = "Semoga Lekas Sembuh"
        LineLen = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((30 - LineLen) / 2))
        TextToPrint &= spcLen3 & StringToPrint & Environment.NewLine

    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static currentChar As Integer
        Dim textfont As Font = New Font("Courier New", 8, FontStyle.Bold)

        Dim h, w As Integer
        Dim left, top As Integer
        With PrintDocument1.DefaultPageSettings
            h = 0
            w = 0
            left = 0
            top = 0
        End With


        Dim lines As Integer = CInt(Math.Round(h / 1))
        Dim b As New Rectangle(left, top, w, h)
        Dim format As StringFormat
        format = New StringFormat(StringFormatFlags.LineLimit)
        Dim line, chars As Integer


        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\tes logo.png")

        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(1, 20)

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)

        e.Graphics.MeasureString(Mid(TextToPrint, currentChar + 1), textfont, New SizeF(w, h), format, chars, line)
        e.Graphics.DrawString(TextToPrint.Substring(currentChar, chars), New Font("Courier New", 8, FontStyle.Bold), Brushes.Black, b, format)


        currentChar = currentChar + chars
        If currentChar < TextToPrint.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            currentChar = 0
        End If
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'NomorFakturPembelian()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
End Class

