<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_CorArticle
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_CorArticle))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.CodeFo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleFo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleDis = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.TextRecher = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_DelRow = New System.Windows.Forms.Button
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.CodeFo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleFo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleDis1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Modifier = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BT_SelAll = New System.Windows.Forms.Button
        Me.BT_DelAll = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BT_Delete = New System.Windows.Forms.Button
        Me.BT_Quit = New System.Windows.Forms.Button
        Me.BT_Save = New System.Windows.Forms.Button
        Me.FileSearched = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepListeFile = New System.Windows.Forms.FolderBrowserDialog
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_SelAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_DelAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Delete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Quit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Save)
        Me.SplitContainer1.Size = New System.Drawing.Size(770, 586)
        Me.SplitContainer1.SplitterDistance = 551
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DataListeSchema)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Size = New System.Drawing.Size(770, 551)
        Me.SplitContainer2.SplitterDistance = 142
        Me.SplitContainer2.TabIndex = 0
        '
        'DataListeSchema
        '
        Me.DataListeSchema.AllowUserToAddRows = False
        Me.DataListeSchema.AllowUserToDeleteRows = False
        Me.DataListeSchema.AllowUserToOrderColumns = True
        Me.DataListeSchema.AllowUserToResizeRows = False
        Me.DataListeSchema.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeSchema.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeFo, Me.ArticleFo, Me.ArticleDis})
        Me.DataListeSchema.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeSchema.Location = New System.Drawing.Point(0, 30)
        Me.DataListeSchema.MultiSelect = False
        Me.DataListeSchema.Name = "DataListeSchema"
        Me.DataListeSchema.RowHeadersVisible = False
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataListeSchema.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowTemplate.Height = 24
        Me.DataListeSchema.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataListeSchema.Size = New System.Drawing.Size(770, 112)
        Me.DataListeSchema.TabIndex = 44
        '
        'CodeFo
        '
        Me.CodeFo.HeaderText = "Tiers"
        Me.CodeFo.Name = "CodeFo"
        Me.CodeFo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ArticleFo
        '
        Me.ArticleFo.FillWeight = 90.90909!
        Me.ArticleFo.HeaderText = "Code  EAN"
        Me.ArticleFo.Name = "ArticleFo"
        '
        'ArticleDis
        '
        Me.ArticleDis.FillWeight = 90.90909!
        Me.ArticleDis.HeaderText = "Code Article Sage"
        Me.ArticleDis.Name = "ArticleDis"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(770, 30)
        Me.Panel2.TabIndex = 43
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.TextRecher)
        Me.GroupBox3.Controls.Add(Me.Button2)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(710, 30)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(216, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Code Article Sage"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(614, 7)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(44, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Tous"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextRecher
        '
        Me.TextRecher.Location = New System.Drawing.Point(6, 9)
        Me.TextRecher.Name = "TextRecher"
        Me.TextRecher.Size = New System.Drawing.Size(204, 20)
        Me.TextRecher.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(472, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(129, 23)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "Rechercher"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BT_ADD)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox2.Location = New System.Drawing.Point(710, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(29, 30)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'BT_ADD
        '
        Me.BT_ADD.Image = Global.Import_Planifier_IM.My.Resources.Resources.applications_161
        Me.BT_ADD.Location = New System.Drawing.Point(2, 7)
        Me.BT_ADD.Name = "BT_ADD"
        Me.BT_ADD.Size = New System.Drawing.Size(22, 20)
        Me.BT_ADD.TabIndex = 3
        Me.BT_ADD.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BT_DelRow)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox1.Location = New System.Drawing.Point(739, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(31, 30)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BT_DelRow
        '
        Me.BT_DelRow.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
        Me.BT_DelRow.Location = New System.Drawing.Point(3, 7)
        Me.BT_DelRow.Name = "BT_DelRow"
        Me.BT_DelRow.Size = New System.Drawing.Size(23, 20)
        Me.BT_DelRow.TabIndex = 2
        Me.BT_DelRow.UseVisualStyleBackColor = True
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeFo1, Me.ArticleFo1, Me.ArticleDis1, Me.Supprimer, Me.Modifier})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 15)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataListeIntegrer.Size = New System.Drawing.Size(770, 390)
        Me.DataListeIntegrer.TabIndex = 45
        '
        'CodeFo1
        '
        Me.CodeFo1.FillWeight = 60.80772!
        Me.CodeFo1.HeaderText = "Tiers"
        Me.CodeFo1.Name = "CodeFo1"
        Me.CodeFo1.ReadOnly = True
        Me.CodeFo1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ArticleFo1
        '
        Me.ArticleFo1.FillWeight = 55.27975!
        Me.ArticleFo1.HeaderText = "Code EAN"
        Me.ArticleFo1.Name = "ArticleFo1"
        Me.ArticleFo1.ReadOnly = True
        '
        'ArticleDis1
        '
        Me.ArticleDis1.FillWeight = 55.27975!
        Me.ArticleDis1.HeaderText = "Code Article Sage"
        Me.ArticleDis1.Name = "ArticleDis1"
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Supprimer"
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 70
        '
        'Modifier
        '
        Me.Modifier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Modifier.FillWeight = 210.451!
        Me.Modifier.HeaderText = "Modifier"
        Me.Modifier.Name = "Modifier"
        Me.Modifier.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Modifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Modifier.Text = "Ici"
        Me.Modifier.UseColumnTextForButtonValue = True
        Me.Modifier.Width = 50
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(770, 15)
        Me.Panel1.TabIndex = 9
        '
        'BT_SelAll
        '
        Me.BT_SelAll.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_SelAll.Location = New System.Drawing.Point(68, 4)
        Me.BT_SelAll.Name = "BT_SelAll"
        Me.BT_SelAll.Size = New System.Drawing.Size(101, 23)
        Me.BT_SelAll.TabIndex = 8
        Me.BT_SelAll.Text = "&Sélectionner Tous "
        Me.BT_SelAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_SelAll.UseVisualStyleBackColor = True
        '
        'BT_DelAll
        '
        Me.BT_DelAll.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_DelAll.Location = New System.Drawing.Point(176, 3)
        Me.BT_DelAll.Name = "BT_DelAll"
        Me.BT_DelAll.Size = New System.Drawing.Size(118, 23)
        Me.BT_DelAll.TabIndex = 7
        Me.BT_DelAll.Text = "&Désélectionner Tous"
        Me.BT_DelAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_DelAll.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(418, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(69, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Integrer "
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BT_Delete
        '
        Me.BT_Delete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BT_Delete.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.BT_Delete.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Delete.Location = New System.Drawing.Point(502, 4)
        Me.BT_Delete.Name = "BT_Delete"
        Me.BT_Delete.Size = New System.Drawing.Size(76, 23)
        Me.BT_Delete.TabIndex = 1
        Me.BT_Delete.Text = "Supprimer"
        Me.BT_Delete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Delete.UseVisualStyleBackColor = True
        '
        'BT_Quit
        '
        Me.BT_Quit.Image = Global.Import_Planifier_IM.My.Resources.Resources.image034
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quit.Location = New System.Drawing.Point(594, 4)
        Me.BT_Quit.Name = "BT_Quit"
        Me.BT_Quit.Size = New System.Drawing.Size(76, 23)
        Me.BT_Quit.TabIndex = 2
        Me.BT_Quit.Text = "&Quitter"
        Me.BT_Quit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Quit.UseVisualStyleBackColor = True
        '
        'BT_Save
        '
        Me.BT_Save.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_16
        Me.BT_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Save.Location = New System.Drawing.Point(309, 2)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(82, 23)
        Me.BT_Save.TabIndex = 1
        Me.BT_Save.Text = "&Enregistrer"
        Me.BT_Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'Frm_CorArticle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 586)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Frm_CorArticle"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Correspondance Articles"
        Me.TopMost = True
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents FileSearched As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BT_Quit As System.Windows.Forms.Button
    Friend WithEvents BT_Save As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BT_Delete As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents DataListeSchema As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents FolderRepListeFile As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BT_ADD As System.Windows.Forms.Button
    Friend WithEvents BT_DelRow As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextRecher As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BT_SelAll As System.Windows.Forms.Button
    Friend WithEvents BT_DelAll As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CodeFo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleFo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleDis As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeFo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleFo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleDis1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Modifier As System.Windows.Forms.DataGridViewButtonColumn
End Class
