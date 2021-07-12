Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class hastaTakip

    Dim baglan As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=eczane.accdb")
    Dim kmt As OleDbCommand = New OleDbCommand()
    Dim dtst As DataSet = New DataSet()
    Dim dtst2 As DataSet = New DataSet()
    Dim dtst3 As DataSet = New DataSet()

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

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        Dim menu As Menu = New Menu()
        menu.Show()
        Me.Hide()
    End Sub

    Private Sub hastaTakip_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        hastaTC()
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
            richTextBox1.Text = oku(4).ToString()
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
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select asiAdi,etkiSuresi,etkisi,asiVurulmaTarihi From asiTablosu where vurulanTC='" & comboBox1.Text & "'", baglan)
        adtr.Fill(dtst, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        dataGridView1.Columns(0).HeaderText = "Aşı Adı"
        dataGridView1.Columns(1).HeaderText = "Etki Süresi"
        dataGridView1.Columns(2).HeaderText = "Etkisi"
        dataGridView1.Columns(3).HeaderText = "Aşı vurulma Tarihi"
        baglan.Close()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst2.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select ilac.ilacin_adi From hasta,ilac where hasta.tc_kimlik='" & comboBox1.Text & "' and hasta.ilac_barkod=ilac.barkod_no", baglan)
        adtr.Fill(dtst2, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst2
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        Me.dataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dataGridView1.Columns(0).HeaderText = "Hastalığa Göre Verilen İlaç "
        baglan.Close()
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        If baglan.State = ConnectionState.Open Then
            baglan.Close()
        End If
        baglan.Open()
        dtst2.Clear()
        Dim adtr As OleDbDataAdapter = New OleDbDataAdapter("select ilac.ilacin_adi From hasta,ilac where hasta.tc_kimlik='" & comboBox1.Text & "' and hasta.ilac_barkod=ilac.barkod_no", baglan)
        adtr.Fill(dtst2, "asiTablosu")
        dataGridView1.DataMember = "asiTablosu"
        dataGridView1.DataSource = dtst2
        adtr.Dispose()
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView1.BackgroundColor = Color.White
        dataGridView1.RowHeadersVisible = False
        Me.dataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dataGridView1.Columns(0).HeaderText = "Hastalığa Göre Verilen İlaç "
        baglan.Close()
    End Sub

    Private Sub button8_Click(sender As Object, e As EventArgs) Handles button8.Click
        Try
            If baglan.State = ConnectionState.Open Then
                baglan.Close()
            End If
            baglan.Open()
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