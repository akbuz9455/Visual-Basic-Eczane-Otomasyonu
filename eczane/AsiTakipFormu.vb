Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class AsiTakipFormu

    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")


    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()

    Private Sub button7_Click(sender As Object, e As EventArgs) Handles button7.Click
        Dim menu As Menu = New Menu()
        menu.Show()
        Me.Hide()
    End Sub

    Public Sub hastaTC()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT DISTINCT * from hasta"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox1.Items.Add(oku(1).ToString())
        End While

        oku.Dispose()
        comboBox1.Sorted = True
        baglan.Close()

    End Sub

    Public Sub personelTC()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT DISTINCT * from personel"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        While oku.Read()
            comboBox2.Items.Add(oku(3).ToString())
        End While

        oku.Dispose()
        comboBox2.Sorted = True
        baglan.Close()

    End Sub

    Public Sub asiDoldur()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From asiTablosu", baglan)
        adtr.Fill(dtst, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Aşı Vurulan TC"
        dataGridView1.Columns(2).HeaderText = "Aşı Vuran TC"
        dataGridView1.Columns(3).HeaderText = "Aşı Adı"
        dataGridView1.Columns(4).HeaderText = "Etki Süresi"
        dataGridView1.Columns(5).HeaderText = "Etkisi"
        dataGridView1.Columns(6).HeaderText = "Aşı vurulma Tarihi"
        textBox1.Text = ""
        textBox2.Text = ""
        richTextBox1.Text = ""
        textBox9.Text = ""
        comboBox1.Text = ""
        comboBox2.Text = ""
        baglan.Close()

    End Sub

    Private Sub dateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles dateTimePicker1.ValueChanged
        textBox9.Text = dateTimePicker1.Value.ToShortDateString()
    End Sub

    Private Sub AsiTakipFormu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        asiDoldur()
        hastaTC()
        personelTC()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()

        kmt.Connection = baglan
        kmt.CommandText = "INSERT INTO asiTablosu(vurulanTC,vuranTC,asiAdi,etkiSuresi,etkisi,asiVurulmaTarihi) VALUES ('" & comboBox1.Text & "' ,'" + comboBox2.Text & "' ,'" + textBox1.Text & "' ,'" + textBox2.Text & "' ,'" + richTextBox1.Text & "' ,'" + textBox9.Text & "')"
        kmt.ExecuteNonQuery()
        MessageBox.Show(" Aşı Vurma işleminiz tamamlanmıştır GEÇMİŞ OLSUN ! ")
        kmt.Dispose()
        baglan.Close()
        asiDoldur()
    End Sub

    Private Sub comboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboBox1.SelectedIndexChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT DISTINCT * from hasta where tc_kimlik ='" & comboBox1.Text & "'"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label3.Text = oku(2).ToString()
        End If

        oku.Dispose()
        baglan.Close()
    End Sub

    Private Sub comboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboBox2.SelectedIndexChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "SELECT DISTINCT * from personel where tc_kimlik ='" & comboBox2.Text & "'"
        Dim oku As OleDbDataReader
        oku = kmt.ExecuteReader()

        If oku.Read() Then
            label7.Text = oku(2).ToString()
        End If

        oku.Dispose()
        baglan.Close()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From asiTablosu where vurulanTC='" & comboBox1.Text & "'", baglan)
        adtr.Fill(dtst, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Aşı Vurulan TC"
        dataGridView1.Columns(2).HeaderText = "Aşı Vuran TC"
        dataGridView1.Columns(3).HeaderText = "Aşı Adı"
        dataGridView1.Columns(4).HeaderText = "Etki Süresi"
        dataGridView1.Columns(5).HeaderText = "Etkisi"
        dataGridView1.Columns(6).HeaderText = "Aşı vurulma Tarihi"
        baglan.Close()
    End Sub

    Private Sub button6_Click(sender As Object, e As EventArgs) Handles button6.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From asiTablosu where vuranTC='" & comboBox2.Text & "'", baglan)
        adtr.Fill(dtst, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Aşı Vurulan TC"
        dataGridView1.Columns(2).HeaderText = "Aşı Vuran TC"
        dataGridView1.Columns(3).HeaderText = "Aşı Adı"
        dataGridView1.Columns(4).HeaderText = "Etki Süresi"
        dataGridView1.Columns(5).HeaderText = "Etkisi"
        dataGridView1.Columns(6).HeaderText = "Aşı vurulma Tarihi"
        baglan.Close()

    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        asiDoldur()
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "DELETE from asiTablosu WHERE vurulanTC = '" & dataGridView1.CurrentRow.Cells(1).Value.ToString() & "' and vuranTC='" & dataGridView1.CurrentRow.Cells(2).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Aşı Kaydı Silme işlemi tamamlandı ! ")
        dtst.Clear()
        baglan.Close()
        asiDoldur()
    End Sub

    Private Sub dataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
        comboBox1.Text = dataGridView1.CurrentRow.Cells(1).Value.ToString()
        comboBox2.Text = dataGridView1.CurrentRow.Cells(2).Value.ToString()
        textBox1.Text = dataGridView1.CurrentRow.Cells(3).Value.ToString()
        textBox2.Text = dataGridView1.CurrentRow.Cells(4).Value.ToString()
        richTextBox1.Text = dataGridView1.CurrentRow.Cells(5).Value.ToString()
        textBox9.Text = dataGridView1.CurrentRow.Cells(6).Value.ToString()
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "update asiTablosu set vurulanTC='" & comboBox1.Text & "',vuranTC='" + comboBox2.Text & "',asiAdi='" + textBox1.Text & "',etkiSuresi='" + textBox2.Text & "',etkisi='" + richTextBox1.Text & "',asiVurulmaTarihi='" + textBox9.Text & "' where  vurulanTC = '" & dataGridView1.CurrentRow.Cells(1).Value.ToString() & "' and vuranTC='" & dataGridView1.CurrentRow.Cells(2).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("Aşı Bilgileri Güncelleme işlemi tamamlandı ! ")
        dtst.Clear()
        baglan.Close()

        asiDoldur()
    End Sub

    Private Sub textBox3_TextChanged(sender As Object, e As EventArgs) Handles textBox3.TextChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From asiTablosu where vurulanTC LIKE '%" & textBox3.Text & "%' ", baglan)
        adtr.Fill(dtst, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Aşı Vurulan TC"
        dataGridView1.Columns(2).HeaderText = "Aşı Vuran TC"
        dataGridView1.Columns(3).HeaderText = "Aşı Adı"
        dataGridView1.Columns(4).HeaderText = "Etki Süresi"
        dataGridView1.Columns(5).HeaderText = "Etkisi"
        dataGridView1.Columns(6).HeaderText = "Aşı vurulma Tarihi"
        textBox1.Text = ""
        textBox2.Text = ""
        richTextBox1.Text = ""
        textBox9.Text = ""
        comboBox1.Text = ""
        comboBox2.Text = ""
        baglan.Close()

    End Sub

    Private Sub button8_Click(sender As Object, e As EventArgs) Handles button8.Click
        Try
            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
            Dim cevap As DialogResult
            cevap = MessageBox.Show("Aşı Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If cevap = DialogResult.Yes Then
                Dim tablobosalt As OleDbCommand = New OleDbCommand(" delete from asiTablosu", baglan)
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