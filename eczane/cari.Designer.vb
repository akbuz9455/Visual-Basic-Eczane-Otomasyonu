<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class cari
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.button3 = New System.Windows.Forms.Button()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.label10 = New System.Windows.Forms.Label()
        Me.label9 = New System.Windows.Forms.Label()
        Me.label8 = New System.Windows.Forms.Label()
        Me.label7 = New System.Windows.Forms.Label()
        Me.label6 = New System.Windows.Forms.Label()
        Me.label5 = New System.Windows.Forms.Label()
        Me.label4 = New System.Windows.Forms.Label()
        Me.label3 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.pictureBox2 = New System.Windows.Forms.PictureBox()
        Me.groupBox1.SuspendLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'button3
        '
        Me.button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(125, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.button3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.button3.Location = New System.Drawing.Point(794, 42)
        Me.button3.Name = "button3"
        Me.button3.Size = New System.Drawing.Size(77, 29)
        Me.button3.TabIndex = 32
        Me.button3.Text = "Anasayfa"
        Me.button3.UseVisualStyleBackColor = False
        '
        'groupBox1
        '
        Me.groupBox1.BackColor = System.Drawing.Color.Azure
        Me.groupBox1.Controls.Add(Me.label10)
        Me.groupBox1.Controls.Add(Me.label9)
        Me.groupBox1.Controls.Add(Me.label8)
        Me.groupBox1.Controls.Add(Me.label7)
        Me.groupBox1.Controls.Add(Me.label6)
        Me.groupBox1.Controls.Add(Me.label5)
        Me.groupBox1.Controls.Add(Me.label4)
        Me.groupBox1.Controls.Add(Me.label3)
        Me.groupBox1.Controls.Add(Me.label2)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Font = New System.Drawing.Font("Verdana", 11.25!)
        Me.groupBox1.Location = New System.Drawing.Point(41, 179)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(531, 230)
        Me.groupBox1.TabIndex = 31
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "İstatistik"
        '
        'label10
        '
        Me.label10.AutoSize = True
        Me.label10.Location = New System.Drawing.Point(256, 183)
        Me.label10.Name = "label10"
        Me.label10.Size = New System.Drawing.Size(18, 18)
        Me.label10.TabIndex = 9
        Me.label10.Text = "*"
        '
        'label9
        '
        Me.label9.AutoSize = True
        Me.label9.Location = New System.Drawing.Point(256, 146)
        Me.label9.Name = "label9"
        Me.label9.Size = New System.Drawing.Size(18, 18)
        Me.label9.TabIndex = 8
        Me.label9.Text = "*"
        '
        'label8
        '
        Me.label8.AutoSize = True
        Me.label8.Location = New System.Drawing.Point(256, 112)
        Me.label8.Name = "label8"
        Me.label8.Size = New System.Drawing.Size(18, 18)
        Me.label8.TabIndex = 7
        Me.label8.Text = "*"
        '
        'label7
        '
        Me.label7.AutoSize = True
        Me.label7.Location = New System.Drawing.Point(256, 75)
        Me.label7.Name = "label7"
        Me.label7.Size = New System.Drawing.Size(18, 18)
        Me.label7.TabIndex = 6
        Me.label7.Text = "*"
        '
        'label6
        '
        Me.label6.AutoSize = True
        Me.label6.Location = New System.Drawing.Point(256, 38)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(18, 18)
        Me.label6.TabIndex = 5
        Me.label6.Text = "*"
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(62, 183)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(188, 18)
        Me.label5.TabIndex = 4
        Me.label5.Text = "Toplam Personel Sayısı :"
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(82, 146)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(168, 18)
        Me.label4.TabIndex = 3
        Me.label4.Text = "Toplam Hasta Sayısı :"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(91, 112)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(159, 18)
        Me.label3.TabIndex = 2
        Me.label3.Text = "Toplam Vurulan Aşı :"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(82, 75)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(168, 18)
        Me.label2.TabIndex = 1
        Me.label2.Text = "Kazanılan İlaç Ücreti :"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(88, 38)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(162, 18)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Toplam Satılan İlaç : "
        '
        'pictureBox1
        '
        Me.pictureBox1.Image = Global.eczane.My.Resources.Resources.Cash_register_icon
        Me.pictureBox1.Location = New System.Drawing.Point(599, 186)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(210, 223)
        Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pictureBox1.TabIndex = 30
        Me.pictureBox1.TabStop = False
        '
        'pictureBox2
        '
        Me.pictureBox2.Image = Global.eczane.My.Resources.Resources.eczaneisim
        Me.pictureBox2.Location = New System.Drawing.Point(203, 53)
        Me.pictureBox2.Name = "pictureBox2"
        Me.pictureBox2.Size = New System.Drawing.Size(488, 81)
        Me.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pictureBox2.TabIndex = 29
        Me.pictureBox2.TabStop = False
        '
        'cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Azure
        Me.ClientSize = New System.Drawing.Size(912, 450)
        Me.Controls.Add(Me.button3)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.pictureBox1)
        Me.Controls.Add(Me.pictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "cari"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "İstatistik"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents button3 As Button
    Private WithEvents groupBox1 As GroupBox
    Private WithEvents label10 As Label
    Private WithEvents label9 As Label
    Private WithEvents label8 As Label
    Private WithEvents label7 As Label
    Private WithEvents label6 As Label
    Private WithEvents label5 As Label
    Private WithEvents label4 As Label
    Private WithEvents label3 As Label
    Private WithEvents label2 As Label
    Private WithEvents label1 As Label
    Private WithEvents pictureBox1 As PictureBox
    Private WithEvents pictureBox2 As PictureBox
End Class
