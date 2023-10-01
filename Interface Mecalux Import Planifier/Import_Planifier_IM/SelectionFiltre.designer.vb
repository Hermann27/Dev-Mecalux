<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectionFiltre
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectionFiltre))
        Me.BT_Quit = New System.Windows.Forms.Button
        Me.BT_Save = New System.Windows.Forms.Button
        Me.FileSearched = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepListeFile = New System.Windows.Forms.FolderBrowserDialog
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Filtre = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Selection = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Txttable = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BT_Quit
        '
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quit.Location = New System.Drawing.Point(89, 7)
        Me.BT_Quit.Name = "BT_Quit"
        Me.BT_Quit.Size = New System.Drawing.Size(79, 23)
        Me.BT_Quit.TabIndex = 2
        Me.BT_Quit.Text = "&Quitter"
        Me.BT_Quit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Quit.UseVisualStyleBackColor = True
        '
        'BT_Save
        '
        Me.BT_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Save.Location = New System.Drawing.Point(246, 7)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(79, 23)
        Me.BT_Save.TabIndex = 1
        Me.BT_Save.Text = "&OK"
        Me.BT_Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Txttable)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Size = New System.Drawing.Size(431, 353)
        Me.SplitContainer1.SplitterDistance = 314
        Me.SplitContainer1.TabIndex = 3
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Filtre, Me.Selection})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataListeIntegrer.Size = New System.Drawing.Size(431, 314)
        Me.DataListeIntegrer.StandardTab = True
        Me.DataListeIntegrer.TabIndex = 0
        '
        'Filtre
        '
        Me.Filtre.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Filtre.DefaultCellStyle = DataGridViewCellStyle1
        Me.Filtre.HeaderText = "Nom de Colonne Filtre"
        Me.Filtre.Name = "Filtre"
        Me.Filtre.ReadOnly = True
        Me.Filtre.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Selection
        '
        Me.Selection.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle2.NullValue = False
        Me.Selection.DefaultCellStyle = DataGridViewCellStyle2
        Me.Selection.HeaderText = "Selection"
        Me.Selection.Name = "Selection"
        Me.Selection.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Selection.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Selection.Width = 60
        '
        'Txttable
        '
        Me.Txttable.Enabled = False
        Me.Txttable.Location = New System.Drawing.Point(140, 28)
        Me.Txttable.Name = "Txttable"
        Me.Txttable.ReadOnly = True
        Me.Txttable.Size = New System.Drawing.Size(216, 20)
        Me.Txttable.TabIndex = 1
        Me.Txttable.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.BT_Save)
        Me.Panel1.Controls.Add(Me.BT_Quit)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(431, 35)
        Me.Panel1.TabIndex = 0
        '
        'SelectionFiltre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ClientSize = New System.Drawing.Size(431, 353)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "SelectionFiltre"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sélection d'un Champ Filtre"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FileSearched As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BT_Quit As System.Windows.Forms.Button
    Friend WithEvents BT_Save As System.Windows.Forms.Button
    Friend WithEvents FolderRepListeFile As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Txttable As System.Windows.Forms.TextBox
    Friend WithEvents Filtre As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Selection As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
