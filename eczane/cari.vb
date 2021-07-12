Imports System.Data
Imports System.Data.OleDb

Public Class cari
    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")


    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()

    Public Sub satilanIlacSayisi()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT COUNT(*)from hasta"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label6.Text = oku(0).ToString() & " Adet"
        End If

        oku.Dispose()

        baglan.Close()
    End Sub

    Public Sub toplamPersonelSayisi()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT COUNT(tc_kimlik) from personel"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label10.Text = oku(0).ToString() & " Personel"
        End If

        oku.Dispose()
        baglan.Close()
    End Sub

    Public Sub hastaSayisiToplam()
        Try

            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
            kmt.Connection = baglan
            kmt.CommandText = "SELECT COUNT(tc_kimlik) from hasta"
            Dim oku As OleDbDataReader
            oku = kmt.ExecuteReader()

            If oku.Read() Then
                label9.Text = oku(0).ToString() & " Hasta Kaydı"
            End If

            oku.Dispose()
            baglan.Close()
        Catch ex As Exception

        End Try


    End Sub

    Public Sub toplamvurulanAsi()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT COUNT(*)from asiTablosu"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label8.Text = oku(0).ToString() & " Adet"
        End If

        oku.Dispose()
        baglan.Close()
    End Sub

    Public Sub kazanilanUcret()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT sum(ilac.fiyati) from hasta,ilac where hasta.ilac_barkod=ilac.barkod_no"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label7.Text = oku(0).ToString() & " TL"
        End If

        oku.Dispose()
        baglan.Close()
    End Sub

    Private Sub cari_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        satilanIlacSayisi()
        kazanilanUcret()
        toplamvurulanAsi()
        hastaSayisiToplam()
        toplamPersonelSayisi()
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        Dim menu As Menu = New Menu()
        menu.Show()
        Me.Hide()
    End Sub
End Class