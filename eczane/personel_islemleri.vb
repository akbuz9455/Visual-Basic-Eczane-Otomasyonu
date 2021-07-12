Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class personel_islemleri
    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")

    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()


    Public Sub sigortaDoldur()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "Select * from sigortaDurumTablosu"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox1.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        baglan.Close()

    End Sub

    Public Sub personelListele()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From personel", baglan)
        adtr.Fill(dtst, "personel")
        dataGridView1.DataMember = "personel"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Yaka Kart No"
        dataGridView1.Columns(2).HeaderText = "Adı Soyadı"
        dataGridView1.Columns(3).HeaderText = "TC Kimlik No"
        dataGridView1.Columns(4).HeaderText = "Doğum Tarihi"
        dataGridView1.Columns(5).HeaderText = "Adresi"
        dataGridView1.Columns(6).HeaderText = "Telefonu"
        dataGridView1.Columns(7).HeaderText = "Email"
        dataGridView1.Columns(8).HeaderText = "İşe Giriş Tarihi"
        dataGridView1.Columns(9).HeaderText = "Sigorta Girişi"
        textBox1.Text = ""
        textBox2.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        textBox6.Text = ""
        textBox7.Text = ""
        comboBox1.Text = ""
        baglan.Close()

    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Dim menu As Menu = New Menu()
        menu.Show()
        Me.Hide()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "INSERT INTO personel(yaka_kart,adi_soyadi,tc_kimlik,d_tarihi,adresi,telefonu,email,ise_giris,sigortasi) VALUES ('" & textBox1.Text & "' ,'" + textBox2.Text & "' ,'" + textBox3.Text & "' ,'" + textBox4.Text & "' ,'" + textBox5.Text & "' ,'" + textBox6.Text & "' ,'" + textBox7.Text & "' ,'" + textBox9.Text & "' ,'" + comboBox1.Text & "' )"
        kmt.ExecuteNonQuery()
        MessageBox.Show("Personel Ekleme işlemi tamamlandı ! ")
        kmt.Dispose()
        personelListele()
        baglan.Close()
    End Sub

    Private Sub personel_islemleri_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        personelListele()
        sigortaDoldur()
        textBox1.Focus()
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "DELETE from personel WHERE tc_kimlik = '" & dataGridView1.CurrentRow.Cells(3).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Personel Silme işlemi tamamlandı ! ")
        dtst.Clear()
        personelListele()
        baglan.Close()

    End Sub

    Private Sub dataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
        textBox1.Text = dataGridView1.CurrentRow.Cells(1).Value.ToString()
        textBox2.Text = dataGridView1.CurrentRow.Cells(2).Value.ToString()
        textBox3.Text = dataGridView1.CurrentRow.Cells(3).Value.ToString()
        textBox4.Text = dataGridView1.CurrentRow.Cells(4).Value.ToString()
        textBox5.Text = dataGridView1.CurrentRow.Cells(5).Value.ToString()
        textBox6.Text = dataGridView1.CurrentRow.Cells(6).Value.ToString()
        textBox7.Text = dataGridView1.CurrentRow.Cells(7).Value.ToString()
        textBox9.Text = dataGridView1.CurrentRow.Cells(8).Value.ToString()
        comboBox1.Text = dataGridView1.CurrentRow.Cells(9).Value.ToString()
    End Sub

    Private Sub dateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles dateTimePicker1.ValueChanged
        textBox9.Text = dateTimePicker1.Value.ToShortDateString()

    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        personelListele()
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "update personel set yaka_kart='" & textBox1.Text & "',adi_soyadi='" + textBox2.Text & "',tc_kimlik='" + textBox3.Text & "',d_tarihi='" + textBox4.Text & "',adresi='" + textBox5.Text & "',telefonu='" + textBox6.Text & "',email='" + textBox7.Text & "', ise_giris='" + textBox9.Text & "',sigortasi='" + comboBox1.Text & "' where  tc_kimlik = '" + dataGridView1.CurrentRow.Cells(3).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Personel Güncelleme işlemi tamamlandı ! ")
        dtst.Clear()
        personelListele()
        baglan.Close()
    End Sub

    Private Sub textBox8_TextChanged(sender As Object, e As EventArgs) Handles textBox8.TextChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From personel where tc_kimlik LIKE '%" & textBox8.Text & "%'", baglan)
        adtr.Fill(dtst, "personel")
        dataGridView1.DataMember = "personel"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Yaka Kart No"
        dataGridView1.Columns(2).HeaderText = "Adı Soyadı"
        dataGridView1.Columns(3).HeaderText = "TC Kimlik No"
        dataGridView1.Columns(4).HeaderText = "Doğum Tarihi"
        dataGridView1.Columns(5).HeaderText = "Adresi"
        dataGridView1.Columns(6).HeaderText = "Telefonu"
        dataGridView1.Columns(7).HeaderText = "Email"
        dataGridView1.Columns(8).HeaderText = "İşe Giriş Tarihi"
        dataGridView1.Columns(9).HeaderText = "Sigorta Girişi"
        textBox1.Text = ""
        textBox2.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        textBox6.Text = ""
        textBox7.Text = ""
        comboBox1.Text = ""
        baglan.Close()
    End Sub

    Private Sub button8_Click(sender As Object, e As EventArgs) Handles button8.Click
        Try
            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
            Dim cevap As DialogResult
            cevap = MessageBox.Show("Personel Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If cevap = DialogResult.Yes Then
                Dim tablobosalt As OleDbCommand = New OleDbCommand(" delete from personel", baglan)
                tablobosalt.ExecuteNonQuery()
            End If

            Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            excel.Visible = True
            Dim Missing As Object = Type.Missing
            Dim workbook As Workbook = excel.Workbooks.Add(Missing)
            Dim sheet1 As Worksheet = CType(workbook.Sheets(1), Worksheet)
            Dim StartCol As Integer = 1
            Dim StartRow As Integer = 1

            For j As Integer = 0 To dataGridView1.Columns.Count - 1
                Dim myRange As Range = CType(sheet1.Cells(StartRow, StartCol + j), Range)
                myRange.Value2 = dataGridView1.Columns(j).HeaderText
            Next

            StartRow += 1

            For i As Integer = 0 To dataGridView1.Rows.Count - 1

                For j As Integer = 0 To dataGridView1.Columns.Count - 1
                    Dim myRange As Range = CType(sheet1.Cells(StartRow + i, StartCol + j), Range)
                    myRange.Value2 = If(dataGridView1(j, i).Value Is Nothing, "", dataGridView1(j, i).Value)
                    myRange.[Select]()
                Next
            Next

            baglan.Close()


        Catch hata As Exception
            MessageBox.Show("Hata Aldınız" & hata.Message)
        End Try
    End Sub
End Class