Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Module MdlKoneksi
    Public konek As OdbcConnection
    Public DA As OdbcDataAdapter
    Public DS As DataSet
    Public DT As DataTable
    Public DR As OdbcDataReader
    Public cmd As OdbcCommand
    Public konek2 As OdbcConnection
    Public DA2 As OdbcDataAdapter
    Public DS2 As DataSet
    Public DT2 As DataTable
    Public DR2 As OdbcDataReader
    Public cmd2 As OdbcCommand
    Sub bukaDB()
        Try
            'konek = New OdbcConnection("Dsn=apotek;server=192.168.1.1;userid=root;database=giandi;port=3306")
            'konek = New OdbcConnection("Dsn=apotek2;server=localhost;userid=root;password=123Siapasaja;database=giandi2;port=3306")
            konek = New OdbcConnection("Dsn=apotek2;server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=giandi2;port=3306")
            'konek = New OdbcConnection("Dsn=apoxsy;server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=giandi;port=3306")
            If konek.State = ConnectionState.Closed Then
                konek.Open()
            End If
        Catch ex As Exception
            MsgBox("Koneksi DataBase Bermasalah, Silahkan Periksa Koneksi Anda!")
        End Try
    End Sub

End Module
