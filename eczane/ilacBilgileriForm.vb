Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class ilacBilgileriForm
    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")


    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()

    Public Sub ilacListesiDoldur()
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From ilac", baglan)
        adtr.Fill(dtst, "ilac")
        dataGridView1.DataMember = "ilac"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Barkod No"
        dataGridView1.Columns(2).HeaderText = "İlacın Adı"
        dataGridView1.Columns(3).HeaderText = "Üretici Firma"
        dataGridView1.Columns(4).HeaderText = "Kutu Sayısı"
        dataGridView1.Columns(5).HeaderText = "Fiyatı"
        dataGridView1.Columns(6).HeaderText = "kullanım Amacı"
        dataGridView1.Columns(7).HeaderText = "Yan Etkileri"
        dataGridView1.Columns(8).HeaderText = "İlacı Teslim Alan Personel"
        textBox1.Text = ""
        textBox2.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        textBox6.Text = ""
        textBox8.Text = ""
        richTextBox1.Text = ""
        baglan.Close()
    End Sub

    Private Sub ilacBilgileriForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ilacListesiDoldur()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "INSERT INTO ilac(barkod_no,ilacin_adi,uretici_firma,kutu_sayisi,fiyati,kullanim_amaci,yan_etkileri,ilac_teslim_alan) VALUES ('" & textBox1.Text & "' ,'" + textBox2.Text & "' ,'" + textBox3.Text & "' ,'" + textBox4.Text & "' ,'" + textBox5.Text & "' ,'" + textBox6.Text & "' ,'" + richTextBox1.Text & "' ,'" + textBox8.Text & "'  )"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("İlaç kayıt işlemi tamamlandı ! ")
        dtst.Clear()
        baglan.Close()
        ilacListesiDoldur()


    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        kmt.CommandText = "DELETE from ilac WHERE barkod_no = '" & dataGridView1.CurrentRow.Cells(1).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("İlaç Silme işlemi tamamlandı ! ")
        dtst.Clear()
        baglan.Close()

        ilacListesiDoldur()
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        ilacListesiDoldur()
    End Sub

    Private Sub dataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
        textBox1.Text = dataGridView1.CurrentRow.Cells(1).Value.ToString()
        textBox2.Text = dataGridView1.CurrentRow.Cells(2).Value.ToString()
        textBox3.Text = dataGridView1.CurrentRow.Cells(3).Value.ToString()
        textBox4.Text = dataGridView1.CurrentRow.Cells(4).Value.ToString()
        textBox5.Text = dataGridView1.CurrentRow.Cells(5).Value.ToString()
        textBox6.Text = dataGridView1.CurrentRow.Cells(6).Value.ToString()
        richTextBox1.Text = dataGridView1.CurrentRow.Cells(7).Value.ToString()
        textBox8.Text = dataGridView1.CurrentRow.Cells(8).Value.ToString()
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
        kmt.CommandText = "update ilac set barkod_no='" & textBox1.Text & "',ilacin_adi='" + textBox2.Text & "',uretici_firma='" + textBox3.Text & "',kutu_sayisi='" + textBox4.Text & "',fiyati='" + textBox5.Text & "',kullanim_amaci='" + textBox6.Text & "',yan_etkileri='" + richTextBox1.Text & "', ilac_teslim_alan='" + textBox8.Text & "' where  barkod_no = '" + dataGridView1.CurrentRow.Cells(1).Value.ToString() & "'"
        kmt.ExecuteNonQuery()
        kmt.Dispose()
        MessageBox.Show("İlaç Güncelleme işlemi tamamlandı ! ")
        dtst.Clear()
        baglan.Close()
        ilacListesiDoldur()
    End Sub

    Private Sub textBox7_TextChanged(sender As Object, e As EventArgs) Handles textBox7.TextChanged
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        kmt.Connection = baglan
        dtst.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select * From ilac where barkod_no LIKE '%" & textBox7.Text & "%'", baglan)
        adtr.Fill(dtst, "ilac")
        dataGridView1.DataMember = "ilac"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).Visible = False
        dataGridView1.Columns(1).HeaderText = "Barkod No"
        dataGridView1.Columns(2).HeaderText = "İlacın Adı"
        dataGridView1.Columns(3).HeaderText = "Üretici Firma"
        dataGridView1.Columns(4).HeaderText = "Kutu Sayısı"
        dataGridView1.Columns(5).HeaderText = "Fiyatı"
        dataGridView1.Columns(6).HeaderText = "Kullanım Amacı"
        dataGridView1.Columns(7).HeaderText = "Yan Etkileri"
        dataGridView1.Columns(8).HeaderText = "İlacı Teslim Alan Personel"
        baglan.Close()
    End Sub

    Private Sub button8_Click(sender As Object, e As EventArgs) Handles button8.Click
        Try
            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
            kmt.Connection = baglan
            Dim cevap As DialogResult
            cevap = MessageBox.Show("İlaç Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If cevap = DialogResult.Yes Then
                Dim tablobosalt As OleDbCommand = New OleDbCommand(" delete from ilac", baglan)
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