Imports System.Data.OleDb
Public Class Form1
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If txtKullaniciAdi.Text = "admin" AndAlso txtSifre.Text = "123456" Then
            MessageBox.Show("Giriş Başarılı !")
            Dim menu As Menu = New Menu()
            menu.Show()
            Me.Hide()
        Else
            MessageBox.Show("Şifre ve Kullanıcı Adı uyuşmuyor.")
        End If
    End Sub
End Class
