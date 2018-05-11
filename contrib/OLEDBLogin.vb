Imports System.Data.OleDb
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\data\Database2.accdb;Persist Security Info=False;")
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM [login] WHERE [username] = '" & TextBox1.Text & "' AND [password] = '" & TextBox2.Text & "'", con)
        Dim user As String = ""
        Dim pass As String = ""

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Isian masih kosong!", MsgBoxStyle.Information, "Informasi")
        Else
        con.Open()
        Dim sdr As OleDbDataReader = cmd.ExecuteReader
        If sdr.Read = True Then
            user = "username"
            pass = "password"
            MessageBox.Show("Login Berhasil!")
            Me.Visible = False
            Form2.Show()
        Else
            MessageBox.Show("Login Gagal!")
            End If
        End If
    End Sub
End Class
