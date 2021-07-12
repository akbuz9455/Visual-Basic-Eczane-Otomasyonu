Imports System.IO
Imports System.Net
Public Class Menu
    Private Sub Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim jsonVerileri, bugunkiKoronaCozumle As String()

        Try

            Using wc As WebClient = New WebClient()
                ServicePointManager.Expect100Continue = True
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                Dim json = wc.DownloadString("https://raw.githubusercontent.com/ozanerturk/covid19-turkey-api/master/dataset/timeline.json")
                jsonVerileri = json.ToString().Split("{"c)
            End Using

            bugunkiKoronaCozumle = jsonVerileri(jsonVerileri.Length - 1).Split(""""c)
            label6.Text = bugunkiKoronaCozumle(3)
            label7.Text = bugunkiKoronaCozumle(31)
            label8.Text = bugunkiKoronaCozumle(35)
            label9.Text = bugunkiKoronaCozumle(55)
            label10.Text = bugunkiKoronaCozumle(51)
            label13.Text = bugunkiKoronaCozumle(39)
        Catch hata As Exception
            MessageBox.Show(hata.Message)
        End Try
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Dim ilacForm As ilacBilgileriForm = New ilacBilgileriForm()
        ilacForm.Show()
        Me.Hide()
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        Dim personel As personel_islemleri = New personel_islemleri()
        personel.Show()
        Me.Hide()
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        Dim hasta As hastaKayit = New hastaKayit()
        hasta.Show()
        Me.Hide()
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        Dim asiTakip As AsiTakipFormu = New AsiTakipFormu()
        asiTakip.Show()
        Me.Hide()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Dim asiTakip As hastaTakip = New hastaTakip()
        asiTakip.Show()
        Me.Hide()
    End Sub

    Private Sub button6_Click(sender As Object, e As EventArgs) Handles button6.Click
        Dim _cari As cari = New cari()
        _cari.Show()
        Me.Hide()
    End Sub
End Class