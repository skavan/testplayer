<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.txtSend = New System.Windows.Forms.TextBox()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BackColor = System.Drawing.Color.DimGray
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.Black
        Me.SplitContainer1.Panel1.Controls.Add(Me.ListBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.Black
        Me.SplitContainer1.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1378, 746)
        Me.SplitContainer1.SplitterDistance = 409
        Me.SplitContainer1.TabIndex = 0
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 28
        Me.ListBox1.Location = New System.Drawing.Point(0, 232)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(409, 514)
        Me.ListBox1.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox1.Location = New System.Drawing.Point(0, 0)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(409, 232)
        Me.TextBox1.TabIndex = 0
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.RichTextBox1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(965, 746)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.Black
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Font = New System.Drawing.Font("Segoe UI", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.White
        Me.RichTextBox1.Location = New System.Drawing.Point(485, 3)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(477, 217)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Image = Global.testplayer.My.Resources.Resources.No_album
        Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.TableLayoutPanel1.SetRowSpan(Me.PictureBox1, 2)
        Me.PictureBox1.Size = New System.Drawing.Size(476, 440)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.ForeColor = System.Drawing.Color.Silver
        Me.Label1.Location = New System.Drawing.Point(485, 223)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(477, 223)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 7
        Me.TableLayoutPanel1.SetColumnSpan(Me.TableLayoutPanel2, 2)
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.ProgressBar1, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.btnConnect, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblStatus, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.txtIP, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtPort, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnDisconnect, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.btnSend, 5, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.txtSend, 4, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 449)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 4
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 37.79904!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 62.20096!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(959, 294)
        Me.TableLayoutPanel2.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.BackColor = System.Drawing.Color.DarkGray
        Me.TableLayoutPanel2.SetColumnSpan(Me.ProgressBar1, 5)
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ProgressBar1.Location = New System.Drawing.Point(98, 122)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(759, 28)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar1.TabIndex = 0
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(98, 37)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(147, 79)
        Me.btnConnect.TabIndex = 3
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.TableLayoutPanel2.SetColumnSpan(Me.lblStatus, 5)
        Me.lblStatus.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblStatus.ForeColor = System.Drawing.Color.White
        Me.lblStatus.Location = New System.Drawing.Point(98, 266)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(759, 28)
        Me.lblStatus.TabIndex = 5
        Me.lblStatus.Text = "Label2"
        '
        'txtIP
        '
        Me.txtIP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtIP.Location = New System.Drawing.Point(98, 3)
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(147, 34)
        Me.txtIP.TabIndex = 7
        Me.txtIP.Text = "192.168.1.126"
        '
        'txtPort
        '
        Me.txtPort.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtPort.Location = New System.Drawing.Point(251, 3)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(147, 34)
        Me.txtPort.TabIndex = 8
        Me.txtPort.Text = "9090"
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Location = New System.Drawing.Point(251, 37)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(147, 79)
        Me.btnDisconnect.TabIndex = 4
        Me.btnDisconnect.Text = "Disconnect"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(710, 37)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(147, 79)
        Me.btnSend.TabIndex = 6
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'txtSend
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.txtSend, 2)
        Me.txtSend.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSend.Location = New System.Drawing.Point(557, 3)
        Me.txtSend.Name = "txtSend"
        Me.txtSend.Size = New System.Drawing.Size(300, 34)
        Me.txtSend.TabIndex = 9
        Me.txtSend.Text = "status - 2 tags:A,l,j,C,N,u,t,y,L,c,K,o,r"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1378, 746)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMain"
        Me.Text = "Form1"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents txtIP As System.Windows.Forms.TextBox
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents txtSend As System.Windows.Forms.TextBox

End Class
