<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Fr_ImportationMvtAchat
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Fr_ImportationMvtAchat))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Cible = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Comptable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Commercial = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NomFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FichierExport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Mode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FeuilleExcel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Deplace = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Valider = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TypeFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Creation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Modification = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Dossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CheminExport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.Datagridaffiche = New System.Windows.Forms.DataGridView
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.BT_SelAll = New System.Windows.Forms.Button
        Me.BT_DelAll = New System.Windows.Forms.Button
        Me.BT_integrer = New System.Windows.Forms.Button
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.BT_Quitter = New System.Windows.Forms.Button
        Me.BT_Apercue = New System.Windows.Forms.Button
        Me.FileSearched = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileXml = New System.Windows.Forms.SaveFileDialog
        Me.FileSearchedXml = New System.Windows.Forms.OpenFileDialog
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.Datagridaffiche, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(770, 576)
        Me.SplitContainer1.SplitterDistance = 216
        Me.SplitContainer1.TabIndex = 11
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.BackgroundColor = System.Drawing.Color.LightBlue
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cible, Me.Comptable, Me.Commercial, Me.NomFormat, Me.FichierExport, Me.Mode, Me.FeuilleExcel, Me.Deplace, Me.Valider, Me.TypeFormat, Me.Creation, Me.Modification, Me.Chemin, Me.Dossier, Me.CheminExport, Me.ID})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(770, 182)
        Me.DataListeIntegrer.TabIndex = 9
        '
        'Cible
        '
        Me.Cible.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Cible.HeaderText = "Cible"
        Me.Cible.Name = "Cible"
        Me.Cible.ReadOnly = True
        Me.Cible.Width = 70
        '
        'Comptable
        '
        Me.Comptable.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Comptable.HeaderText = "BaseCpta*"
        Me.Comptable.Name = "Comptable"
        Me.Comptable.Width = 90
        '
        'Commercial
        '
        Me.Commercial.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Commercial.HeaderText = "BaseCial*"
        Me.Commercial.Name = "Commercial"
        Me.Commercial.Width = 90
        '
        'NomFormat
        '
        Me.NomFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.NullValue = Nothing
        Me.NomFormat.DefaultCellStyle = DataGridViewCellStyle1
        Me.NomFormat.HeaderText = "Nom du Fichier Format*"
        Me.NomFormat.Name = "NomFormat"
        Me.NomFormat.ReadOnly = True
        Me.NomFormat.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NomFormat.Width = 150
        '
        'FichierExport
        '
        Me.FichierExport.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FichierExport.DefaultCellStyle = DataGridViewCellStyle2
        Me.FichierExport.HeaderText = "Nom du Fichier à Importer*"
        Me.FichierExport.Name = "FichierExport"
        Me.FichierExport.ReadOnly = True
        Me.FichierExport.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.FichierExport.Width = 170
        '
        'Mode
        '
        Me.Mode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Mode.HeaderText = "Mode"
        Me.Mode.Name = "Mode"
        Me.Mode.Width = 80
        '
        'FeuilleExcel
        '
        Me.FeuilleExcel.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FeuilleExcel.HeaderText = "Feuille Excel"
        Me.FeuilleExcel.Name = "FeuilleExcel"
        '
        'Deplace
        '
        Me.Deplace.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Deplace.HeaderText = "Deplace"
        Me.Deplace.Name = "Deplace"
        Me.Deplace.Width = 50
        '
        'Valider
        '
        Me.Valider.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Valider.HeaderText = "Valider"
        Me.Valider.Name = "Valider"
        Me.Valider.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Valider.ToolTipText = "Valider les Fichiers à exporter"
        Me.Valider.Width = 50
        '
        'TypeFormat
        '
        Me.TypeFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.TypeFormat.HeaderText = "Type*"
        Me.TypeFormat.Name = "TypeFormat"
        Me.TypeFormat.Width = 70
        '
        'Creation
        '
        Me.Creation.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Creation.HeaderText = "Creation"
        Me.Creation.Name = "Creation"
        Me.Creation.ReadOnly = True
        Me.Creation.Width = 80
        '
        'Modification
        '
        Me.Modification.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Modification.HeaderText = "Modification"
        Me.Modification.Name = "Modification"
        Me.Modification.ReadOnly = True
        Me.Modification.Width = 80
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.Format = "N0"
        Me.Chemin.DefaultCellStyle = DataGridViewCellStyle3
        Me.Chemin.HeaderText = "Chemin d'acces du Fichier Format"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Chemin.Visible = False
        Me.Chemin.Width = 170
        '
        'Dossier
        '
        Me.Dossier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Dossier.HeaderText = "Dossier"
        Me.Dossier.Name = "Dossier"
        Me.Dossier.Visible = False
        '
        'CheminExport
        '
        Me.CheminExport.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CheminExport.HeaderText = "Chemin d'acces du Fichier Export"
        Me.CheminExport.Name = "CheminExport"
        Me.CheminExport.ReadOnly = True
        Me.CheminExport.Visible = False
        Me.CheminExport.Width = 170
        '
        'ID
        '
        Me.ID.HeaderText = "ID"
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox1.Location = New System.Drawing.Point(0, 182)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(770, 34)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Fichier en Cours"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(498, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(0, 13)
        Me.Label5.TabIndex = 42
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(383, 12)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(0, 18)
        Me.Label8.TabIndex = 34
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ProgressBar1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.BT_SelAll)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BT_DelAll)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BT_integrer)
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BT_Quitter)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BT_Apercue)
        Me.SplitContainer2.Size = New System.Drawing.Size(770, 356)
        Me.SplitContainer2.SplitterDistance = 308
        Me.SplitContainer2.TabIndex = 0
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.IsSplitterFixed = True
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.Datagridaffiche)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.ListBox1)
        Me.SplitContainer4.Size = New System.Drawing.Size(770, 289)
        Me.SplitContainer4.SplitterDistance = 238
        Me.SplitContainer4.TabIndex = 0
        '
        'Datagridaffiche
        '
        Me.Datagridaffiche.AllowUserToAddRows = False
        Me.Datagridaffiche.AllowUserToDeleteRows = False
        Me.Datagridaffiche.AllowUserToOrderColumns = True
        Me.Datagridaffiche.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Datagridaffiche.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Datagridaffiche.Location = New System.Drawing.Point(0, 0)
        Me.Datagridaffiche.Name = "Datagridaffiche"
        Me.Datagridaffiche.RowHeadersVisible = False
        Me.Datagridaffiche.RowTemplate.Height = 24
        Me.Datagridaffiche.Size = New System.Drawing.Size(770, 238)
        Me.Datagridaffiche.TabIndex = 5
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(0, 0)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(770, 36)
        Me.ListBox1.Sorted = True
        Me.ListBox1.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 289)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(770, 19)
        Me.ProgressBar1.Step = 0
        Me.ProgressBar1.TabIndex = 3
        '
        'BT_SelAll
        '
        Me.BT_SelAll.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_SelAll.Location = New System.Drawing.Point(109, 15)
        Me.BT_SelAll.Name = "BT_SelAll"
        Me.BT_SelAll.Size = New System.Drawing.Size(101, 23)
        Me.BT_SelAll.TabIndex = 14
        Me.BT_SelAll.Text = "&Sélectionner Tous "
        Me.BT_SelAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_SelAll.UseVisualStyleBackColor = True
        '
        'BT_DelAll
        '
        Me.BT_DelAll.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_DelAll.Location = New System.Drawing.Point(238, 15)
        Me.BT_DelAll.Name = "BT_DelAll"
        Me.BT_DelAll.Size = New System.Drawing.Size(118, 23)
        Me.BT_DelAll.TabIndex = 13
        Me.BT_DelAll.Text = "&Désélectionner Tous"
        Me.BT_DelAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_DelAll.UseVisualStyleBackColor = True
        '
        'BT_integrer
        '
        Me.BT_integrer.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowforward_161
        Me.BT_integrer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_integrer.Location = New System.Drawing.Point(501, 14)
        Me.BT_integrer.Name = "BT_integrer"
        Me.BT_integrer.Size = New System.Drawing.Size(65, 23)
        Me.BT_integrer.TabIndex = 5
        Me.BT_integrer.Text = "&Integrer"
        Me.BT_integrer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_integrer.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox6.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(770, 9)
        Me.GroupBox6.TabIndex = 4
        Me.GroupBox6.TabStop = False
        '
        'BT_Quitter
        '
        Me.BT_Quitter.Image = Global.Import_Planifier_IM.My.Resources.Resources.image034
        Me.BT_Quitter.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quitter.Location = New System.Drawing.Point(599, 14)
        Me.BT_Quitter.Name = "BT_Quitter"
        Me.BT_Quitter.Size = New System.Drawing.Size(68, 23)
        Me.BT_Quitter.TabIndex = 2
        Me.BT_Quitter.Text = "&Quitter"
        Me.BT_Quitter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Quitter.UseVisualStyleBackColor = True
        '
        'BT_Apercue
        '
        Me.BT_Apercue.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowforward_161
        Me.BT_Apercue.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Apercue.Location = New System.Drawing.Point(412, 14)
        Me.BT_Apercue.Name = "BT_Apercue"
        Me.BT_Apercue.Size = New System.Drawing.Size(65, 23)
        Me.BT_Apercue.TabIndex = 0
        Me.BT_Apercue.Text = "&Aperçu"
        Me.BT_Apercue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Apercue.UseVisualStyleBackColor = True
        '
        'Fr_ImportationMvtAchat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 576)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Fr_ImportationMvtAchat"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Integration des docoments de Achat"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        CType(Me.Datagridaffiche, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents FileSearched As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BT_Quitter As System.Windows.Forms.Button
    Friend WithEvents BT_Apercue As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents SaveFileXml As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FileSearchedXml As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BT_integrer As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents Datagridaffiche As System.Windows.Forms.DataGridView
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents BT_SelAll As System.Windows.Forms.Button
    Friend WithEvents BT_DelAll As System.Windows.Forms.Button
    Friend WithEvents Cible As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Comptable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Commercial As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NomFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FichierExport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Mode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FeuilleExcel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Deplace As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Valider As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents TypeFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Creation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Modification As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Dossier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheminExport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ID As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
