<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.TbExcFile = New System.Windows.Forms.TextBox()
        Me.TbPPfile = New System.Windows.Forms.TextBox()
        Me.lblExcelFile = New System.Windows.Forms.Label()
        Me.lblPPTemplate = New System.Windows.Forms.Label()
        Me.PbLogo = New System.Windows.Forms.PictureBox()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.pbCount = New System.Windows.Forms.ProgressBar()
        CType(Me.PbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.btnStart.ForeColor = System.Drawing.SystemColors.Highlight
        Me.btnStart.Location = New System.Drawing.Point(370, 5)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(84, 58)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = False
        '
        'TbExcFile
        '
        Me.TbExcFile.ForeColor = System.Drawing.Color.Green
        Me.TbExcFile.Location = New System.Drawing.Point(5, 69)
        Me.TbExcFile.Multiline = True
        Me.TbExcFile.Name = "TbExcFile"
        Me.TbExcFile.Size = New System.Drawing.Size(450, 40)
        Me.TbExcFile.TabIndex = 1
        '
        'TbPPfile
        '
        Me.TbPPfile.ForeColor = System.Drawing.Color.Green
        Me.TbPPfile.Location = New System.Drawing.Point(5, 129)
        Me.TbPPfile.Multiline = True
        Me.TbPPfile.Name = "TbPPfile"
        Me.TbPPfile.Size = New System.Drawing.Size(450, 40)
        Me.TbPPfile.TabIndex = 2
        '
        'lblExcelFile
        '
        Me.lblExcelFile.AutoSize = True
        Me.lblExcelFile.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblExcelFile.Location = New System.Drawing.Point(2, 58)
        Me.lblExcelFile.Name = "lblExcelFile"
        Me.lblExcelFile.Size = New System.Drawing.Size(52, 13)
        Me.lblExcelFile.TabIndex = 3
        Me.lblExcelFile.Text = "Excel File"
        '
        'lblPPTemplate
        '
        Me.lblPPTemplate.AutoSize = True
        Me.lblPPTemplate.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblPPTemplate.Location = New System.Drawing.Point(2, 118)
        Me.lblPPTemplate.Name = "lblPPTemplate"
        Me.lblPPTemplate.Size = New System.Drawing.Size(104, 13)
        Me.lblPPTemplate.TabIndex = 4
        Me.lblPPTemplate.Text = "PowerPoint template"
        '
        'PbLogo
        '
        Me.PbLogo.ImageLocation = "http://lcp.pl/images/logo.png"
        Me.PbLogo.Location = New System.Drawing.Point(5, 5)
        Me.PbLogo.Name = "PbLogo"
        Me.PbLogo.Size = New System.Drawing.Size(100, 50)
        Me.PbLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PbLogo.TabIndex = 5
        Me.PbLogo.TabStop = False
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.ForeColor = System.Drawing.Color.Blue
        Me.lblInfo.Location = New System.Drawing.Point(2, 186)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(35, 13)
        Me.lblInfo.TabIndex = 6
        Me.lblInfo.Text = "lblInfo"
        '
        'pbCount
        '
        Me.pbCount.Location = New System.Drawing.Point(5, 202)
        Me.pbCount.Name = "pbCount"
        Me.pbCount.Size = New System.Drawing.Size(450, 15)
        Me.pbCount.TabIndex = 7
        Me.pbCount.Visible = False
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.BackColor = System.Drawing.SystemColors.Info
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(460, 230)
        Me.Controls.Add(Me.pbCount)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.PbLogo)
        Me.Controls.Add(Me.lblPPTemplate)
        Me.Controls.Add(Me.lblExcelFile)
        Me.Controls.Add(Me.TbPPfile)
        Me.Controls.Add(Me.TbExcFile)
        Me.Controls.Add(Me.btnStart)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmMain"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CONV tool for LCP Properties Sp. z o. o."
        CType(Me.PbLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents TbExcFile As TextBox
    Friend WithEvents TbPPfile As TextBox
    Friend WithEvents lblExcelFile As Label
    Friend WithEvents lblPPTemplate As Label
    Friend WithEvents PbLogo As PictureBox
    Friend WithEvents lblInfo As Label
    Friend WithEvents pbCount As ProgressBar
End Class
