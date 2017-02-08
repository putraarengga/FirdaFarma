Imports MySql.Data.MySqlClient
Imports System.Math
Public Class FormTransaksiPembelian

    Dim databaru As Boolean
    Dim selectDataBase As String
    Dim hBeli, total, totalsatuan, grandTotal As Double
    Dim hDiskon, hitung As Double
    Dim hPajak As Double
    Dim hHJULv1, hHJULv2, hHJULv3 As Double
    Dim hHJR As Double
    Dim simpan As String
    Dim persenHJULv1, persenHJULv2, persenHJULv3 As Double
    Dim persenHJR As Double
    Dim persenPajak As Double
    Dim persenDiskon As Double
    Dim day, month, year, vdate, vdate2 As String
    Dim dt1, dt2 As DateTime
    Dim selectDataBase2, selectDataBase3 As String
    Dim nowTime As String
    Dim GetIDSupplier, NoFaktur As String
    Dim reader As MySqlDataReader
    Dim anInteger, hargabelilama, stokSimpan As Integer
    Dim StokBaru, konversilv1, konversilv2, konversilv3 As Integer
    Dim hbelilv2, hbelilv3 As Double
    Dim pdislv1, pdislv2, pdislv3 As Double
    Dim dislv1, dislv2, dislv3 As Integer
    Dim StoksimpanLv1, StoksimpanLv2, StoksimpanLv3 As Integer
    Dim dataSelected As Boolean
    Dim DrugName As String
    Dim Amount, Price, Discount, PDiscount, Tax, PTax As Integer
    Dim kurangsatuan, kuranggrand1, kuranggrand2, totalprice As Integer
    Dim Stringtotal As String
    Dim ceks1, ceks2, ceks3, ceks4 As Boolean
    Private _right As String

    Shared Property indexObat As String
    Shared Property hargabeli As String
    Shared Property indexSupplier As Integer
    Shared Property Satuans As String
    Shared Property NamaObat As String
    Shared Property stokObat As Integer
    Shared Property satuanlv1 As String
    Shared Property FLagdata As Integer

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormTransaksiPembelian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataSelected = False
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        dt1 = DateTimePicker1.Value
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "yyyy-MM-dd"
        dt2 = DateTimePicker2.Value
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Medical-Drug.ico")
        total = 0
        IsiGrid()
        FormDisabeld()
        Button8.Enabled = False
        Button9.Enabled = False

    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        year = DateTimePicker2.Value.Year
        month = DateTimePicker2.Value.Month
        day = DateTimePicker2.Value.Day
        vdate2 = year + "-" + month + "-" + day
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "yyyy-MM-dd"
        dt2 = DateTimePicker2.Value
    End Sub
    Sub IsiGrid()
        nowTime = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase2 = "SELECT tpembelian.NamaObat, tpembelian.JumlahPembelian, tpembelian.DiskonPembelian, tpembelian.Pajak, tpembelian.HargaBeli," +
            "tpembelian.SubTotal, tpembelian.KonversiLv2, tpembelian.KonversiLv3 " +
            "FROM tpembelian WHERE tpembelian.TglPembelianObat = '" & nowTime & "' and tpembelian.NomorFaktur= '" & TextBox1.Text & "'"
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
            .Columns(0).HeaderCell.Value = "Nama Obat"
            .Columns(1).HeaderCell.Value = "Jumlah Pembelian Obat"
            .Columns(2).HeaderCell.Value = "Diskon Pembelian"
            .Columns(3).HeaderCell.Value = "Pajak PPN"
            .Columns(4).HeaderCell.Value = "Harga Pembelian"
            .Columns(5).HeaderCell.Value = "Sub Total"
            .Columns(6).HeaderCell.Value = "Konversi Obat Lv2"
            .Columns(7).HeaderCell.Value = "Konversi Obat Lv3"
        End With
    End Sub

    Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox5.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""

        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""

        TextBox13.Text = ""
        TextBox14.Text = ""

        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""

        TextBox12.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""

        TextBox31.Text = ""
        TextBox33.Text = ""
        TextBox39.Text = ""

        TextBox32.Text = ""
        TextBox34.Text = ""
        TextBox41.Text = ""

        TextBox42.Text = ""
        TextBox44.Text = ""
        TextBox45.Text = ""

        TextBox35.Text = ""
        TextBox37.Text = ""
        TextBox40.Text = ""

        TextBox36.Text = ""
        TextBox38.Text = ""
        TextBox43.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Sub Bersih2()
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox5.Text = 0
        TextBox25.Text = 0
        TextBox26.Text = 0
        TextBox17.Text = 0

        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""


        TextBox13.Text = ""
        TextBox14.Text = ""

        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""

        TextBox12.Text = 0
        TextBox15.Text = 0
        TextBox16.Text = 0

        TextBox31.Text = ""
        TextBox33.Text = ""
        TextBox39.Text = ""

        TextBox32.Text = 0
        TextBox34.Text = 0
        TextBox41.Text = 0

        TextBox42.Text = ""
        TextBox44.Text = ""
        TextBox45.Text = ""

        TextBox35.Text = ""
        TextBox37.Text = ""
        TextBox40.Text = ""

        TextBox36.Text = ""
        TextBox38.Text = ""
        TextBox43.Text = ""

    End Sub

    Sub FormEnabeld()
        TextBox2.Enabled = True
        TextBox3.Enabled = True

        TextBox5.Enabled = True
        TextBox25.Enabled = True
        TextBox26.Enabled = True

        TextBox6.Enabled = True
        TextBox7.Enabled = True

        TextBox31.Enabled = True
        TextBox33.Enabled = True
        TextBox39.Enabled = True
        TextBox14.Enabled = True

        TextBox35.Enabled = True
        TextBox37.Enabled = True
        TextBox40.Enabled = True

        TextBox36.Enabled = True
        TextBox38.Enabled = True
        TextBox43.Enabled = True

        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        DateTimePicker2.Enabled = True
        CheckBox1.Enabled = True


    End Sub
    Sub FormDisabeld()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False

        TextBox5.Enabled = False
        TextBox25.Enabled = False
        TextBox26.Enabled = False

        TextBox6.Enabled = False
        TextBox7.Enabled = False

        TextBox31.Enabled = False
        TextBox33.Enabled = False
        TextBox39.Enabled = False

        TextBox14.Enabled = False
        TextBox16.Enabled = False

        TextBox35.Enabled = False
        TextBox37.Enabled = False
        TextBox40.Enabled = False

        TextBox36.Enabled = False
        TextBox38.Enabled = False
        TextBox43.Enabled = False


        Button5.Enabled = False
        Button6.Enabled = False
        DateTimePicker2.Enabled = False
        CheckBox1.Enabled = False

        TextBox12.Enabled = False
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Function bulat(angka As Integer) As Integer
        Dim strs, pembulatan As String
        Dim tex1 As Integer
        strs = Str(angka)
        pembulatan = Microsoft.VisualBasic.Right(strs, 2)
        If Val(pembulatan) < 50 Then
            tex1 = angka - Val(pembulatan)
        ElseIf Val(pembulatan) > 50 Then
            tex1 = angka - Val(pembulatan)
            tex1 = tex1 + 100
        ElseIf Val(pembulatan) = 50 Then
            tex1 = angka
        End If
        bulat = tex1
    End Function

    Private Sub HitungData()

        If CheckBox1.Checked = True Then
            persenPajak = 10
        ElseIf CheckBox1.Checked = False Then
            persenPajak = 0
        End If

        hBeli = Val(TextBox6.Text)
        persenDiskon = Val(TextBox7.Text)
        StokBaru = Val(TextBox5.Text)
        pdislv1 = Val(TextBox36.Text)
        pdislv2 = Val(TextBox38.Text)
        pdislv3 = Val(TextBox43.Text)


        hPajak = Convert.ToInt32((hBeli * StokBaru) * (persenPajak / 100))
        hDiskon = Convert.ToInt32((hBeli * StokBaru) * (persenDiskon / 100))

        konversilv2 = Val(TextBox25.Text)
        konversilv3 = Val(TextBox26.Text)

        StoksimpanLv1 = StokBaru
        StoksimpanLv2 = StokBaru * konversilv2
        StoksimpanLv3 = StoksimpanLv2 * konversilv3

        TextBox8.Text = hDiskon.ToString

        totalsatuan = ((hBeli * StokBaru) + hPajak) - hDiskon
        TextBox17.Text = totalsatuan.ToString
        'hHJULv1 = totalsatuan + (totalsatuan * persenHJULv1 / 100)
        'hbelilv2 = totalsatuan / (StokBaru * konversilv2)
        'hHJULv2 = hbelilv2 + (hbelilv2 * persenHJULv2 / 100)
        'hbelilv3 = hbelilv2 / (StokBaru * konversilv3)

        persenHJULv1 = Val(TextBox31.Text)
        hHJULv1 = Convert.ToInt32(hBeli + (hBeli * (persenHJULv1 / 100) + (hBeli * (persenPajak / 100))))
        Dim h_HJULv1 As Integer
        h_HJULv1 = bulat(hHJULv1)
        TextBox32.Text = h_HJULv1.ToString
        'TextBox32.Text = hHJULv1.ToString
        If konversilv2 <= 0 Then
            hbelilv2 = 0
        Else
            hbelilv2 = Convert.ToInt32((hBeli + (hBeli * (persenPajak / 100))) / konversilv2)

        End If
        dislv1 = Convert.ToInt32(hHJULv1 * (pdislv1 / 100))
        Dim ddislv1 As Integer
        ddislv1 = bulat(dislv1)
        TextBox12.Text = (ddislv1.ToString)

        persenHJULv2 = Val(TextBox33.Text)
        hHJULv2 = Convert.ToInt32(hbelilv2 + (hbelilv2 * (persenHJULv2 / 100)))
        If hHJULv2 > 0 Then
            Dim h_HJULv2 As Integer
            h_HJULv2 = bulat(hHJULv2)
            TextBox34.Text = h_HJULv2.ToString
            If konversilv3 <= 0 Then
                hbelilv3 = 0
            Else
                hbelilv3 = Convert.ToInt32(hbelilv2 / konversilv3)
            End If
            dislv2 = Convert.ToInt32(hHJULv2 * (pdislv2 / 100))
            Dim ddislv2 As Integer
            ddislv2 = bulat(dislv2)
            TextBox15.Text = ddislv2.ToString
        Else
            TextBox34.Text = 0
            hbelilv3 = 0
            dislv2 = 0
            TextBox15.Text = dislv2.ToString
        End If


        persenHJULv3 = Val(TextBox39.Text)
        hHJULv3 = Convert.ToInt32(hbelilv3 + (hbelilv3 * persenHJULv3 / 100))
        If hHJULv3 > 0 Then
            Dim h_HJULv3 As Integer
            h_HJULv3 = bulat(hHJULv3)
            TextBox41.Text = h_HJULv3.ToString
            dislv3 = Convert.ToInt32(hHJULv3 * (pdislv3 / 100))
            Dim ddislv3 As Integer = bulat(dislv3)
            TextBox16.Text = ddislv3.ToString
        Else
            TextBox41.Text = 0
            TextBox16.Text = 0
        End If

        If FLagdata = 2 Then
            persenHJR = Val(TextBox14.Text)
            hHJR = Convert.ToInt32(hBeli + (hBeli * persenHJR / 100))
            Dim h_HJR As Integer = bulat(hHJR)
            TextBox13.Text = h_HJR.ToString
        ElseIf FLagdata = 3 Then
            persenHJR = Val(TextBox14.Text)
            hHJR = Convert.ToInt32(hbelilv2 + (hbelilv2 * persenHJR / 100))
            Dim h_HJR As Integer = bulat(hHJR)
            TextBox13.Text = h_HJR.ToString
        Else
            persenHJR = Val(TextBox14.Text)
            hHJR = Convert.ToInt32(hbelilv3 + (hbelilv3 * persenHJR / 100))
            Dim h_HJR As Integer = bulat(hHJR)
            TextBox13.Text = h_HJR.ToString
        End If
        '    persenHJR = Val(TextBox14.Text)
        'hHJR = Convert.ToInt32(hBeli + (hBeli * persenHJR / 100))
        '    TextBox13.Text = hHJR.ToString

    End Sub

    Sub GetSupplierAndNoFaktur()
        indexSupplier = -1
        selectDataBase = "SELECT IDSupplier FROM tsupplier WHERE NamaSupplier ='" & TextBox4.Text & "'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            indexSupplier = DT.Rows(0).Item("IDSupplier")
            GetIDSupplier = indexSupplier
            TextBox1.Text = Format(DateTime.Now, "yyyyMMdd") & indexSupplier
        End If
    End Sub

    Sub cek()
        Dim tex5 As Integer = Val(TextBox5.Text)
        Dim tex31 As Double = Val(TextBox31.Text)
        Dim tex25 As Integer = Val(TextBox25.Text)
        Dim tex33 As Double = Val(TextBox33.Text)
        Dim tex26 As Integer = Val(TextBox26.Text)
        Dim tex39 As Double = Val(TextBox39.Text)
        Dim tex14 As Double = Val(TextBox14.Text)

        If TextBox5.Enabled = True Then
            If TextBox5.Text = "" Or tex5 = 0 Or TextBox31.Text = "" Or tex31 = 0.0 Then
                ceks1 = True
            Else
                ceks1 = False
            End If
        End If
        If TextBox25.Enabled = True Then
            If TextBox25.Text = "" Or tex25 = 0 Or TextBox33.Text = "" Or tex33 = 0.0 Then
                ceks2 = True
            Else
                ceks2 = False
            End If
        End If
        If TextBox26.Enabled = True Then
            If TextBox26.Text = "" Or tex26 = 0 Or TextBox39.Text = "" Or tex39 = 0.0 Then
                ceks3 = True
            Else
                ceks3 = False
            End If
        End If
        If TextBox14.Enabled = True Then
            If TextBox14.Text = "" Or tex14 = 0.0 Then
                ceks4 = True
            Else
                ceks4 = False
            End If
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        cek()
        If ceks1 = False And ceks2 = False And ceks3 = False And ceks4 = False Then
            Try
                vdate = Format(DateTime.Now, "yyyy-MM-dd")
                If (DateTime.Compare(vdate, vdate2) = 0) Then
                    MsgBox("Isi Dulu Tanggal Kadaluarsa", MsgBoxStyle.OkCancel)
                ElseIf (DateTime.Compare(vdate2, vdate) = -1) Then
                    MsgBox("Tanggal Kadaluarsa Salah", MsgBoxStyle.OkCancel)
                Else

                    simpan = "INSERT INTO tpembelian(NomorFaktur, IDObat, NamaObat," +
                        "JumlahPembelian, KonversiLv2, KonversiLv3, " +
                        "HargaBeli, HargaJualUmum1, HargaJualUmum2, HargaJualUmum3, " +
                        "HargaJualResep, Pajak, DiskonPembelian," +
                        "KeuntunganLv1, KeuntunganLv2, KeuntunganLv3, " +
                        "KeuntunganResep, TglKadaluarsa, TglPembelianObat, IDSupplierObat," +
                        "SisaObatLv1, SisaObatLv2, SisaObatLv3, " +
                        "SatuanLv1,SatuanLv2,SatuanLv3," +
                        "SatDisLv1, DisLv1, HargaDisLv1," +
                        "SatDisLv2, DisLv2, HargaDisLv2," +
                        "SatDisLv3, DisLv3, HargaDisLv3, " +
                        "SubTotal)" +
                                   "VALUES ('" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox2.Text & "'," +
                                   "'" & TextBox5.Text & "','" & TextBox25.Text & "','" & TextBox26.Text & "'," +
                                   "'" & TextBox6.Text & "','" & TextBox32.Text & "','" & TextBox34.Text & "','" & TextBox41.Text & "'," +
                                   "'" & TextBox13.Text & "','" & persenPajak & "','" & TextBox7.Text & "'," +
                                   "'" & TextBox31.Text & "','" & TextBox33.Text & "','" & TextBox39.Text & "'," +
                                   "'" & TextBox14.Text & "','" & vdate2 & "','" & vdate & "','" & GetIDSupplier & "'," +
                                   "'" & StoksimpanLv1 & "', '" & StoksimpanLv2 & "', '" & StoksimpanLv3 & "'," +
                                   "'" & TextBox10.Text & "','" & TextBox9.Text & "','" & TextBox11.Text & "'," +
                                   "'" & TextBox35.Text & "','" & TextBox36.Text & "', '" & TextBox12.Text & "'," +
                                   "'" & TextBox37.Text & "','" & TextBox38.Text & "', '" & TextBox15.Text & "'," +
                                   "'" & TextBox40.Text & "','" & TextBox43.Text & "', '" & TextBox16.Text & "'," +
                                   "'" & TextBox17.Text & "')"

                    selectDataBase = "INSERT INTO thistoryharga(NomorFaktur, IDObat, NamaObat, HargaBeliObat, TglPembelian," +
                        "HJR, HJU1, HJU2, HJU3," +
                        "PajakPPN, DiskonPembelian, " +
                        "DiskonJual1, DiskonJual2, DiskonJual3," +
                        "Jumlah1, Jumlah2, Jumlah3," +
                        "IDSupplier, Kadarluarsa ) " +
                                   "VALUES ('" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox2.Text & "','" & TextBox6.Text & "','" & vdate & "'," +
                                   "'" & TextBox13.Text & "','" & TextBox32.Text & "','" & TextBox34.Text & "','" & TextBox41.Text & "'," +
                                   "'" & persenPajak & "','" & TextBox7.Text & "'," +
                                   "'" & TextBox12.Text & "', '" & TextBox15.Text & "', '" & TextBox16.Text & "', " +
                                   "'" & TextBox5.Text & "','" & TextBox25.Text & "','" & TextBox26.Text & "'," +
                                   "'" & GetIDSupplier & "','" & vdate2 & "')"

                    selectDataBase2 = "UPDATE tpembelian SET HargaJualUmum1= '" & TextBox32.Text & "', HargaJualUmum2= '" & TextBox34.Text & "', HargaJualUmum3= '" & TextBox41.Text & "'," +
                        "KeuntunganLv1= '" & TextBox31.Text & "',KeuntunganLv2= '" & TextBox33.Text & "', KeuntunganLv3= '" & TextBox39.Text & "'," +
                        "HargaJualResep= '" & TextBox13.Text & "', KeuntunganResep= '" & TextBox14.Text & "'," +
                        "SatDisLv1='" & TextBox35.Text & "', DisLv1='" & TextBox36.Text & "', HargaDisLv1='" & TextBox12.Text & "'," +
                        "SatDisLv2='" & TextBox37.Text & "', DisLv2='" & TextBox38.Text & "', HargaDisLv2='" & TextBox15.Text & "'," +
                        "SatDisLv3='" & TextBox40.Text & "', DisLv3='" & TextBox43.Text & "', HargaDisLv3='" & TextBox16.Text & "'" +
                        "WHERE tpembelian.NamaObat = '" & TextBox2.Text & "'"


                    jalankansql(simpan)
                    jalankansql(selectDataBase)
                    jalankansql(selectDataBase2)

                    IsiGrid()
                    FormDisabeld()
                    Button8.Enabled = True
                    Button9.Enabled = True
                    Button7.Enabled = False
                    Button6.Enabled = False

                    total = 0
                    For x As Integer = 0 To DataGridView1.RowCount - 1
                        total += DataGridView1.Rows(x).Cells(5).Value
                    Next
                    Stringtotal = FormatNumber(total, 2, , , TriState.True)
                    RichTextBox1.Text = Stringtotal
                End If
            Catch ex As Exception
                MsgBox("Isi Dulu Tanggal Kadaluarsa" & ex.Message)
            End Try
            Button2.Focus()
            DateTimePicker2.CustomFormat = "dd/MM/yyyy"
            DateTimePicker2.Value = Now()
        Else
            MsgBox("Data Belum Lengkap", MsgBoxStyle.OkCancel)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If DataGridView1.RowCount > 0 Then
            grandTotal = total
            selectDataBase2 = " UPDATE tpembelian SET tpembelian.GrandTotal = '" & grandTotal & "' WHERE tpembelian.NomorFaktur = '" & TextBox1.Text & "' "
            jalankansql(selectDataBase2)
            FormDisabeld()
            RichTextBox1.Text = 0.0
            Button2.Enabled = True
            TextBox4.Text = ""
            Bersih()
            Me.Close()
        Else
            MsgBox("Tidak Ada Data", MsgBoxStyle.OkCancel)
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql, hapussql2, hapussql3 As String
        Dim pesan As String

        If DrugName = "" Then
            pesan = MsgBox("Tidak Dapat Mengetahui Jenis Obat Yang diHapus ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")

        Else
            pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + DrugName, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            hapussql = "DELETE FROM tpembelian WHERE tpembelian.NomorFaktur = '" & TextBox1.Text & "' AND tpembelian.NamaObat ='" & DrugName & "'"
            hapussql3 = "DELETE FROM thistoryharga WHERE thistoryharga.NomorFaktur = '" & TextBox1.Text & "' AND thistoryharga.NamaObat ='" & DrugName & "'"

            jalankansql(hapussql)
            jalankansql(hapussql3)

            DataGridView1.Refresh()
            IsiGrid()
            total = 0

            If DataGridView1.RowCount > 0 Then
                For x As Integer = 0 To DataGridView1.RowCount - 1
                    total += DataGridView1.Rows(x).Cells(5).Value
                Next
                Stringtotal = FormatNumber(total, 2, , , TriState.True)
                RichTextBox1.Text = Stringtotal
            Else
                RichTextBox1.Text = 0
            End If
            grandTotal = total
            selectDataBase2 = " UPDATE tpembelian SET tpembelian.GrandTotal = '" & grandTotal & "' WHERE tpembelian.NomorFaktur = '" & TextBox1.Text & "' "
            jalankansql(selectDataBase2)
        End If
        Button4.Enabled = False
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Bersih2()
        FormEnabeld()
        FormPencarianDataObat.Show()
        FormPencarianDataObat.Focus()
        If TextBox4.Text = "" Then
            MsgBox("Tidak dapat mengetahui Supplier", MsgBoxStyle.OkCancel)
            FormPencarianDataObat.Close()
        Else
            GetSupplierAndNoFaktur()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FormPencarianDataSupplier.Show()
        FormPencarianDataSupplier.Focus()
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
    Private Sub TextBox7_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox7.MouseClick
        TextBox7.SelectAll()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        HitungData()

    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged
        HitungData()
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        hitung = Val(TextBox6.Text)
        HitungData()
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        HitungData()
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        HitungData()
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles TextBox31.TextChanged
        HitungData()

    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox32_TextChanged(sender As Object, e As EventArgs) Handles TextBox32.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox34_TextChanged(sender As Object, e As EventArgs) Handles TextBox34.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox35_TextChanged(sender As Object, e As EventArgs) Handles TextBox35.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox37_TextChanged(sender As Object, e As EventArgs) Handles TextBox37.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles TextBox36.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged
        HitungData()
    End Sub
    Private Sub TextBox43_TextChanged(sender As Object, e As EventArgs) Handles TextBox43.TextChanged
        HitungData()
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormEnabeld()
        TextBox4.Focus()
        Button2.Enabled = False
        TextBox4.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False

    End Sub

    Private Sub FormTransaksiPembelian_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Bersih2()
        Button2.Enabled = True
        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
    End Sub
    Sub GetIndeks(ByVal x As Integer)

        '.Columns(0).HeaderCell.Value = "Nama Obat"
        '.Columns(1).HeaderCell.Value = "Jumlah Pembelian Obat"
        '.Columns(2).HeaderCell.Value = "Diskon Pembelian"
        '.Columns(3).HeaderCell.Value = "Pajak PPN"
        '.Columns(4).HeaderCell.Value = "Harga Pembelian"
        '.Columns(5).HeaderCell.Value = "Sub Total"
        Try
            DrugName = DataGridView1.Rows(x).Cells(0).Value
            Amount = DataGridView1.Rows(x).Cells(1).Value
            Price = DataGridView1.Rows(x).Cells(4).Value
            PDiscount = DataGridView1.Rows(x).Cells(2).Value
            PTax = DataGridView1.Rows(x).Cells(3).Value
            totalprice = DataGridView1.Rows(x).Cells(5).Value
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        GetIndeks(e.RowIndex)
        dataSelected = True
        Button4.Enabled = True
    End Sub
    Sub HitungUlang()
        Tax = Convert.ToInt32((Price * Amount) * (PTax / 100))
        Discount = Convert.ToInt32((Price * Amount) * (PDiscount / 100))
        kurangsatuan = ((Price * Amount) + Tax) - Discount
        kuranggrand1 = Val(RichTextBox1.Text)
        kuranggrand2 = kuranggrand1 - kurangsatuan
        RichTextBox1.Refresh()
        RichTextBox1.Text = kuranggrand2.ToString
    End Sub

End Class

