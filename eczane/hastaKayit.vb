Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class hastaKayit


    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")


    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()



    Private Sub hastaKayit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''oluşturdumuz fonksiyonları form yüklenirken çağırdık
        hastaDoldur()
        guvence()
        ilacBarkodDoldur()
        kullanimAjDok()
        kullanimVakti()
    End Sub

    Public Sub hastaDoldur()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From hasta", baglan)
        adtr.Fill(dtst, "hasta")
        dataGridView1.DataMember = "hasta"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "TC Kimlik"
        dataGridView1.Columns(2).HeaderText = "Adı Soyadı"
        dataGridView1.Columns(3).HeaderText = "Sosyal Güvencesi"
        dataGridView1.Columns(4).HeaderText = "Adresi"
        dataGridView1.Columns(5).HeaderText = "Telefonu"
        dataGridView1.Columns(6).HeaderText = "İlaç Kullanımı"
        dataGridView1.Columns(7).HeaderText = "Kullanım Şekli"
        dataGridView1.Columns(8).HeaderText = "İlaç Barkod"
        textBox1.Text = ""
        textBox2.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        comboBox1.Text = ""
        comboBox2.Text = ""
        comboBox3.Text = ""
        comboBox4.Text = ""
        baglan.Close()
    End Sub



    Public Sub ilacBarkodDoldur()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        kmt.Connection = baglan
        kmt.CommandText = "Select * from ilac"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox4.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        comboBox4.Sorted = True

        baglan.Close()
    End Sub

    Public Sub kullanimAjDok()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        kmt.Connection = baglan
        kmt.CommandText = "Select * from kullanimActok"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox2.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        comboBox2.Sorted = True

        baglan.Close()
    End Sub

    Public Sub guvence()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        kmt.Connection = baglan
        kmt.CommandText = "Select * from guvence"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox1.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        comboBox1.Sorted = True

        baglan.Close()
    End Sub

    Public Sub kullanimVakti()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        kmt.Connection = baglan
        kmt.CommandText = "Select * from kullanimVakti"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox3.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        comboBox3.Sorted = True

        baglan.Close()
    End Sub



    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Dim menu As Menu = New Menu()
        menu.Show()
        Me.Hide()
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "update hasta set tc_kimlik='" & textBox1.Text & "',adi_soyadi='" + textBox2.Text & "',sosyal_güvencesi='" + comboBox1.Text & "',adresi='" + textBox3.Text & "',telefonu='" + textBox4.Text & "',ilac_kullanimi='" + comboBox2.Text & "',kullanim_sekli='" + comboBox3.Text & "',ilac_barkod='" + comboBox4.Text & "' where  tc_kimlik = '" + dataGridView1.CurrentRow.Cells(1).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Hasta Bilgileri Güncelleme işlemi tamamlandı ! ")
        dtst.Clear()
        hastaDoldur()


        baglan.Close()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "INSERT INTO hasta(tc_kimlik,adi_soyadi,sosyal_güvencesi,adresi,telefonu,ilac_kullanimi,kullanim_sekli,ilac_barkod) VALUES ('" & textBox1.Text & "' ,'" + textBox2.Text & "' ,'" + comboBox1.Text & "' ,'" + textBox3.Text & "' ,'" + textBox4.Text & "' ,'" + comboBox2.Text & "' ,'" + comboBox3.Text & "','" + comboBox4.Text & "')"
        kmt.ExecuteNonQuery()
        MessageBox.Show(" kayıt işleminiz tamamlanmıştır GEÇMİŞ OLSUN ! ")
        kmt.Dispose()
        hastaDoldur()
        baglan.Close()
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "DELETE from hasta WHERE tc_kimlik = '" & dataGridView1.CurrentRow.Cells(1).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Hasta Kaydı Silme işlemi tamamlandı ! ")
        dtst.Clear()
        hastaDoldur()
        baglan.Close()
    End Sub

    Private Sub dataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
        ''sql tablosundaki verileri form üzerindeki textbox ve comboboxlara aktardık.
        textBox1.Text = dataGridView1.CurrentRow.Cells(1).Value.ToString()
        textBox2.Text = dataGridView1.CurrentRow.Cells(2).Value.ToString()
        textBox3.Text = dataGridView1.CurrentRow.Cells(4).Value.ToString()
        comboBox1.Text = dataGridView1.CurrentRow.Cells(3).Value.ToString()
        textBox4.Text = dataGridView1.CurrentRow.Cells(5).Value.ToString()
        comboBox2.Text = dataGridView1.CurrentRow.Cells(6).Value.ToString()
        comboBox3.Text = dataGridView1.CurrentRow.Cells(7).Value.ToString()
        comboBox4.Text = dataGridView1.CurrentRow.Cells(8).Value.ToString()
    End Sub

    Private Sub textBox5_TextChanged(sender As Object, e As EventArgs) Handles textBox5.TextChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From hasta where tc_kimlik LIKE '%" & textBox5.Text & "%'", baglan)
        ''burada arama işlemi yaptık ancak  normal aramalardan farklı olarak listelerken LIKE Kullandık bu yazdıklarımız eğer tc içerisinde var ise sonuç verecektir
        ''bire bir de karşılaştırma yapmayacaktır.
        adtr.Fill(dtst, "hasta")
        dataGridView1.DataMember = "hasta"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        ''datagridviewi listeledimiz verilerle doldurduk
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        ''seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "TC Kimlik"
        dataGridView1.Columns(2).HeaderText = "Adı Soyadı"
        dataGridView1.Columns(3).HeaderText = "Sosyal Güvencesi"
        dataGridView1.Columns(4).HeaderText = "Adresi"
        dataGridView1.Columns(5).HeaderText = "Telefonu"
        dataGridView1.Columns(6).HeaderText = "İlaç Kullanımı"
        dataGridView1.Columns(7).HeaderText = "Kullanım Şekli"
        dataGridView1.Columns(8).HeaderText = "İlaç Barkod"
        baglan.Close()
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        ''hastadoldur fonksiyonunu çağırır.
        hastaDoldur()
    End Sub

    Private Sub button8_Click(sender As Object, e As EventArgs) Handles button8.Click
        Try
            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
            Dim cevap As DialogResult
            cevap = MessageBox.Show("Hasta Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If cevap = DialogResult.Yes Then
                Dim tablobosalt As OleDbCommand = New OleDbCommand(" delete from hasta", baglan)
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