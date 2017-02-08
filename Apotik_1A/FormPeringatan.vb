Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Public Class FormPeringatan
    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        Close()


    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
      
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        Me.Icon = img
        Dim posisi As Point
        posisi.X = 50
        posisi.Y = 500
        Location = posisi

        'Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\wallpaper.jpg")

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub FormPeringatan_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
End Class
