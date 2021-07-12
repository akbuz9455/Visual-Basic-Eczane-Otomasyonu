<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.button1 = New System.Windows.Forms.Button()
        Me.txtSifre = New System.Windows.Forms.TextBox()
        Me.txtKullaniciAdi = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.pictureBox2 = New System.Windows.Forms.PictureBox()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.groupBox1.SuspendLayout()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.groupBox1.Controls.Add(Me.button1)
        Me.groupBox1.Controls.Add(Me.txtSifre)
        Me.groupBox1.Controls.Add(Me.txtKullaniciAdi)
        Me.groupBox1.Controls.Add(Me.label2)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.groupBox1.Location = New System.Drawing.Point(325, 178)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(369, 203)
        Me.groupBox1.TabIndex = 4
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Giriş Bilgileri"
        '
        'button1
        '
        Me.button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(125, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button1.ForeColor = System.Drawing.Color.White
        Me.button1.Location = New System.Drawing.Point(200, 145)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(93, 32)
        Me.button1.TabIndex = 4
        Me.button1.Text = "Giriş Yap"
        Me.button1.UseVisualStyleBackColor = False
        '
        'txtSifre
        '
        Me.txtSifre.Location = New System.Drawing.Point(165, 96)
        Me.txtSifre.Name = "txtSifre"
        Me.txtSifre.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSifre.Size = New System.Drawing.Size(128, 22)
        Me.txtSifre.TabIndex = 3
        '
        'txtKullaniciAdi
        '
        Me.txtKullaniciAdi.Location = New System.Drawing.Point(165, 54)
        Me.txtKullaniciAdi.Name = "txtKullaniciAdi"
        Me.txtKullaniciAdi.Size = New System.Drawing.Size(128, 22)
        Me.txtKullaniciAdi.TabIndex = 2
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(105, 99)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(48, 14)
        Me.label2.TabIndex = 1
        Me.label2.Text = "Şifre :"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(63, 57)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(97, 14)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Kullanıcı Adı :"
        '
        'pictureBox2
        '
        Me.pictureBox2.Image = Global.eczane.My.Resources.Resources.eczaneisim
        Me.pictureBox2.Location = New System.Drawing.Point(75, 46)
        Me.pictureBox2.Name = "pictureBox2"
        Me.pictureBox2.Size = New System.Drawing.Size(601, 81)
        Me.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pictureBox2.TabIndex = 5
        Me.pictureBox2.TabStop = False
        '
        'pictureBox1
        '
        Me.pictureBox1.Image = Global.eczane.My.Resources.Resources.eczaneOtomasyonuGiris
        Me.pictureBox1.Location = New System.Drawing.Point(34, 149)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(253, 248)
        Me.pictureBox1.TabIndex = 3
        Me.pictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(744, 464)
        Me.Controls.Add(Me.pictureBox2)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.pictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Eczane Otomasyonu Giriş"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents pictureBox2 As PictureBox
    Private WithEvents groupBox1 As GroupBox
    Private WithEvents button1 As Button
    Private WithEvents txtSifre As TextBox
    Private WithEvents txtKullaniciAdi As TextBox
    Private WithEvents label2 As Label
    Private WithEvents label1 As Label
    Private WithEvents pictureBox1 As PictureBox
End Class
