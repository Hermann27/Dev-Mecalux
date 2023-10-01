<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormVersionng
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormVersionng))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Table = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Mode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.codeTable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.version = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtVersion = New System.Windows.Forms.TextBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.BTupdate = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblinfos = New System.Windows.Forms.Label
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(565, 231)
        Me.SplitContainer1.SplitterDistance = 302
        Me.SplitContainer1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataListeIntegrer)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(302, 231)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Table, Me.Mode, Me.codeTable, Me.version})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.EnableHeadersVisualStyles = False
        Me.DataListeIntegrer.Location = New System.Drawing.Point(3, 16)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        Me.DataListeIntegrer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(296, 212)
        Me.DataListeIntegrer.TabIndex = 12
        '
        'Table
        '
        Me.Table.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Table.HeaderText = "Object"
        Me.Table.Name = "Table"
        Me.Table.ReadOnly = True
        '
        'Mode
        '
        Me.Mode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Mode.HeaderText = "Mode de Transfert"
        Me.Mode.Name = "Mode"
        '
        'codeTable
        '
        Me.codeTable.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.codeTable.DefaultCellStyle = DataGridViewCellStyle1
        Me.codeTable.HeaderText = "Code"
        Me.codeTable.Name = "codeTable"
        Me.codeTable.ReadOnly = True
        Me.codeTable.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.codeTable.Visible = False
        Me.codeTable.Width = 110
        '
        'version
        '
        Me.version.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.version.HeaderText = "Version"
        Me.version.Name = "version"
        Me.version.ReadOnly = True
        Me.version.Visible = False
        Me.version.Width = 150
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblinfos)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtVersion)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(259, 231)
        Me.SplitContainer2.SplitterDistance = 47
        Me.SplitContainer2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Version du fichier :"
        '
        'txtVersion
        '
        Me.txtVersion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVersion.Location = New System.Drawing.Point(98, 11)
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.Size = New System.Drawing.Size(41, 20)
        Me.txtVersion.TabIndex = 0
        Me.txtVersion.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.BTupdate)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.PictureBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(259, 180)
        Me.SplitContainer3.SplitterDistance = 41
        Me.SplitContainer3.TabIndex = 0
        '
        'BTupdate
        '
        Me.BTupdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BTupdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTupdate.Image = Global.Import_Planifier_IM.My.Resources.Resources.AnalyzeWizard11
        Me.BTupdate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTupdate.Location = New System.Drawing.Point(147, 6)
        Me.BTupdate.Name = "BTupdate"
        Me.BTupdate.Size = New System.Drawing.Size(100, 29)
        Me.BTupdate.TabIndex = 64
        Me.BTupdate.Text = "Modifier"
        Me.BTupdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTupdate.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        'Me.PictureBox1.Image = Global.Import_Planifier_IM.My.Resources.Resources.images__5_
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(259, 135)
        Me.PictureBox1.TabIndex = 65
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(144, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 65
        Me.Label2.Text = "Code  Table :"
        '
        'lblinfos
        '
        Me.lblinfos.AutoSize = True
        Me.lblinfos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblinfos.Location = New System.Drawing.Point(214, 14)
        Me.lblinfos.Name = "lblinfos"
        Me.lblinfos.Size = New System.Drawing.Size(27, 13)
        Me.lblinfos.TabIndex = 66
        Me.lblinfos.Text = "....."
        '
        'FormVersionng
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(565, 231)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormVersionng"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Formulaire de mise à jour de la version du fichier"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtVersion As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents BTupdate As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Table As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Mode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents codeTable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents version As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblinfos As System.Windows.Forms.Label
End Class
