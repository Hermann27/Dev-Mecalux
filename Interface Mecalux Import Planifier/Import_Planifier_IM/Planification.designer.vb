<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Planification
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Planification))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CbIntitule = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BTsup = New System.Windows.Forms.Button
        Me.BTupdate = New System.Windows.Forms.Button
        Me.SplitContain = New System.Windows.Forms.SplitContainer
        Me.dgvListeTrait = New System.Windows.Forms.DataGridView
        Me.intituleTrait = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.typeTrait = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Consulter = New System.Windows.Forms.DataGridViewButtonColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.btnSupprimer = New System.Windows.Forms.Button
        Me.btnEnregistrer = New System.Windows.Forms.Button
        Me.btnEnlever = New System.Windows.Forms.Button
        Me.btnSelect = New System.Windows.Forms.Button
        Me.dgvTraitementSelect = New System.Windows.Forms.DataGridView
        Me.intituleSelect = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossiers = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.rangSelect = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RefSelect = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvTraitementEnr = New System.Windows.Forms.DataGridView
        Me.Intitule = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Execution = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Heure1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Heure2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Heure = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Critere1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Critere2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateM = New System.Windows.Forms.DataGridViewButtonColumn
        Me.IDDossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Rang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.supEnr = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SplitContain.Panel1.SuspendLayout()
        Me.SplitContain.Panel2.SuspendLayout()
        Me.SplitContain.SuspendLayout()
        CType(Me.dgvListeTrait, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.dgvTraitementSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTraitementEnr, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContain)
        Me.SplitContainer1.Size = New System.Drawing.Size(883, 470)
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CbIntitule)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.BTsup)
        Me.GroupBox2.Controls.Add(Me.BTupdate)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(883, 50)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Tâche Planifiée en cours"
        '
        'CbIntitule
        '
        Me.CbIntitule.FormattingEnabled = True
        Me.CbIntitule.Location = New System.Drawing.Point(296, 18)
        Me.CbIntitule.Margin = New System.Windows.Forms.Padding(2)
        Me.CbIntitule.Name = "CbIntitule"
        Me.CbIntitule.Size = New System.Drawing.Size(172, 23)
        Me.CbIntitule.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(194, 20)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Intitulé Tâche :"
        '
        'BTsup
        '
        Me.BTsup.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTsup.Image = My.Resources.Resources.criticalind_status
        Me.BTsup.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTsup.Location = New System.Drawing.Point(594, 17)
        Me.BTsup.Name = "BTsup"
        Me.BTsup.Size = New System.Drawing.Size(83, 23)
        Me.BTsup.TabIndex = 64
        Me.BTsup.Text = "Supprimer"
        Me.BTsup.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTsup.UseVisualStyleBackColor = True
        '
        'BTupdate
        '
        Me.BTupdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.BTupdate.Image = My.Resources.Resources.btn_valider
        Me.BTupdate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTupdate.Location = New System.Drawing.Point(711, 17)
        Me.BTupdate.Name = "BTupdate"
        Me.BTupdate.Size = New System.Drawing.Size(79, 23)
        Me.BTupdate.TabIndex = 63
        Me.BTupdate.Text = "Modifier"
        Me.BTupdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTupdate.UseVisualStyleBackColor = True
        '
        'SplitContain
        '
        Me.SplitContain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContain.IsSplitterFixed = True
        Me.SplitContain.Location = New System.Drawing.Point(0, 0)
        Me.SplitContain.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContain.Name = "SplitContain"
        Me.SplitContain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContain.Panel1
        '
        Me.SplitContain.Panel1.Controls.Add(Me.dgvListeTrait)
        Me.SplitContain.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContain.Panel2
        '
        Me.SplitContain.Panel2.Controls.Add(Me.dgvTraitementEnr)
        Me.SplitContain.Size = New System.Drawing.Size(883, 417)
        Me.SplitContain.SplitterDistance = 243
        Me.SplitContain.SplitterWidth = 3
        Me.SplitContain.TabIndex = 2
        '
        'dgvListeTrait
        '
        Me.dgvListeTrait.AllowUserToAddRows = False
        Me.dgvListeTrait.AllowUserToDeleteRows = False
        Me.dgvListeTrait.AllowUserToOrderColumns = True
        Me.dgvListeTrait.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvListeTrait.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.intituleTrait, Me.typeTrait, Me.Consulter})
        Me.dgvListeTrait.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvListeTrait.Location = New System.Drawing.Point(0, 0)
        Me.dgvListeTrait.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvListeTrait.Name = "dgvListeTrait"
        Me.dgvListeTrait.RowHeadersVisible = False
        Me.dgvListeTrait.RowTemplate.Height = 24
        Me.dgvListeTrait.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvListeTrait.Size = New System.Drawing.Size(883, 243)
        Me.dgvListeTrait.TabIndex = 8
        '
        'intituleTrait
        '
        Me.intituleTrait.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.intituleTrait.HeaderText = "Intitulé du Traitement"
        Me.intituleTrait.Name = "intituleTrait"
        Me.intituleTrait.ReadOnly = True
        '
        'typeTrait
        '
        Me.typeTrait.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.typeTrait.HeaderText = "Type"
        Me.typeTrait.Name = "typeTrait"
        Me.typeTrait.Visible = False
        '
        'Consulter
        '
        Me.Consulter.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Consulter.DefaultCellStyle = DataGridViewCellStyle1
        Me.Consulter.HeaderText = "Visualiser"
        Me.Consulter.Name = "Consulter"
        Me.Consulter.Text = "..."
        Me.Consulter.UseColumnTextForButtonValue = True
        Me.Consulter.Width = 70
        '
        'SplitContainer2
        '
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(246, 226)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(344, 16)
        Me.SplitContainer2.SplitterDistance = 299
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 1
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnSupprimer)
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnEnregistrer)
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnEnlever)
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnSelect)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.dgvTraitementSelect)
        Me.SplitContainer3.Size = New System.Drawing.Size(42, 16)
        Me.SplitContainer3.SplitterDistance = 25
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 0
        '
        'btnSupprimer
        '
        Me.btnSupprimer.Image = My.Resources.Resources.delete_161
        Me.btnSupprimer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSupprimer.Location = New System.Drawing.Point(5, 214)
        Me.btnSupprimer.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSupprimer.Name = "btnSupprimer"
        Me.btnSupprimer.Size = New System.Drawing.Size(75, 22)
        Me.btnSupprimer.TabIndex = 15
        Me.btnSupprimer.Text = "Supprimer"
        Me.btnSupprimer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSupprimer.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Image = My.Resources.Resources.save_16
        Me.btnEnregistrer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnEnregistrer.Location = New System.Drawing.Point(5, 176)
        Me.btnEnregistrer.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(75, 23)
        Me.btnEnregistrer.TabIndex = 14
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnEnlever
        '
        Me.btnEnlever.Image = My.Resources.Resources.arrowback_16
        Me.btnEnlever.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEnlever.Location = New System.Drawing.Point(5, 40)
        Me.btnEnlever.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEnlever.Name = "btnEnlever"
        Me.btnEnlever.Size = New System.Drawing.Size(75, 22)
        Me.btnEnlever.TabIndex = 13
        Me.btnEnlever.Text = "Enlever"
        Me.btnEnlever.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnEnlever.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Image = My.Resources.Resources.arrowforward_16
        Me.btnSelect.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSelect.Location = New System.Drawing.Point(5, 2)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 25)
        Me.btnSelect.TabIndex = 12
        Me.btnSelect.Text = "Ajouter"
        Me.btnSelect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'dgvTraitementSelect
        '
        Me.dgvTraitementSelect.AllowUserToAddRows = False
        Me.dgvTraitementSelect.AllowUserToDeleteRows = False
        Me.dgvTraitementSelect.AllowUserToOrderColumns = True
        Me.dgvTraitementSelect.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTraitementSelect.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.intituleSelect, Me.IDDossiers, Me.rangSelect, Me.RefSelect})
        Me.dgvTraitementSelect.Location = New System.Drawing.Point(282, 0)
        Me.dgvTraitementSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvTraitementSelect.MultiSelect = False
        Me.dgvTraitementSelect.Name = "dgvTraitementSelect"
        Me.dgvTraitementSelect.RowHeadersVisible = False
        Me.dgvTraitementSelect.RowTemplate.Height = 24
        Me.dgvTraitementSelect.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTraitementSelect.Size = New System.Drawing.Size(166, 243)
        Me.dgvTraitementSelect.TabIndex = 4
        '
        'intituleSelect
        '
        Me.intituleSelect.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.intituleSelect.HeaderText = "Intitule"
        Me.intituleSelect.Name = "intituleSelect"
        Me.intituleSelect.ReadOnly = True
        '
        'IDDossiers
        '
        Me.IDDossiers.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossiers.HeaderText = "ID dossier"
        Me.IDDossiers.Name = "IDDossiers"
        Me.IDDossiers.ReadOnly = True
        Me.IDDossiers.Width = 80
        '
        'rangSelect
        '
        Me.rangSelect.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.rangSelect.HeaderText = "Rang"
        Me.rangSelect.Name = "rangSelect"
        Me.rangSelect.Width = 50
        '
        'RefSelect
        '
        Me.RefSelect.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RefSelect.HeaderText = "Critère"
        Me.RefSelect.Name = "RefSelect"
        Me.RefSelect.ReadOnly = True
        '
        'dgvTraitementEnr
        '
        Me.dgvTraitementEnr.AllowUserToAddRows = False
        Me.dgvTraitementEnr.AllowUserToDeleteRows = False
        Me.dgvTraitementEnr.AllowUserToOrderColumns = True
        Me.dgvTraitementEnr.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTraitementEnr.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Intitule, Me.Execution, Me.Heure1, Me.Heure2, Me.Heure, Me.Critere1, Me.Critere2, Me.DateM, Me.IDDossier, Me.Rang, Me.supEnr})
        Me.dgvTraitementEnr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvTraitementEnr.Location = New System.Drawing.Point(0, 0)
        Me.dgvTraitementEnr.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvTraitementEnr.Name = "dgvTraitementEnr"
        Me.dgvTraitementEnr.RowHeadersVisible = False
        Me.dgvTraitementEnr.RowTemplate.Height = 24
        Me.dgvTraitementEnr.Size = New System.Drawing.Size(883, 171)
        Me.dgvTraitementEnr.TabIndex = 1
        '
        'Intitule
        '
        Me.Intitule.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Intitule.HeaderText = "Intitule Traitement"
        Me.Intitule.Name = "Intitule"
        '
        'Execution
        '
        Me.Execution.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Execution.HeaderText = "Der. Exécution"
        Me.Execution.Name = "Execution"
        Me.Execution.ReadOnly = True
        Me.Execution.Width = 130
        '
        'Heure1
        '
        Me.Heure1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Heure1.HeaderText = "Heure 1"
        Me.Heure1.Name = "Heure1"
        Me.Heure1.Width = 70
        '
        'Heure2
        '
        Me.Heure2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Heure2.HeaderText = "Heure 2"
        Me.Heure2.Name = "Heure2"
        Me.Heure2.Width = 70
        '
        'Heure
        '
        Me.Heure.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Heure.DefaultCellStyle = DataGridViewCellStyle2
        Me.Heure.HeaderText = "Heure"
        Me.Heure.Name = "Heure"
        Me.Heure.Text = "..."
        Me.Heure.UseColumnTextForButtonValue = True
        Me.Heure.Width = 50
        '
        'Critere1
        '
        Me.Critere1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Critere1.HeaderText = "Critere 1"
        Me.Critere1.Name = "Critere1"
        '
        'Critere2
        '
        Me.Critere2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Critere2.HeaderText = "Critere 2"
        Me.Critere2.Name = "Critere2"
        Me.Critere2.Width = 70
        '
        'DateM
        '
        Me.DateM.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateM.DefaultCellStyle = DataGridViewCellStyle3
        Me.DateM.HeaderText = "Date"
        Me.DateM.Name = "DateM"
        Me.DateM.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DateM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DateM.Text = "..."
        Me.DateM.UseColumnTextForButtonValue = True
        Me.DateM.Width = 55
        '
        'IDDossier
        '
        Me.IDDossier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier.HeaderText = "ID dossier"
        Me.IDDossier.Name = "IDDossier"
        Me.IDDossier.ReadOnly = True
        Me.IDDossier.Width = 80
        '
        'Rang
        '
        Me.Rang.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Rang.HeaderText = "Rang"
        Me.Rang.Name = "Rang"
        Me.Rang.Width = 55
        '
        'supEnr
        '
        Me.supEnr.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.supEnr.HeaderText = "Supprimer"
        Me.supEnr.Name = "supEnr"
        Me.supEnr.Width = 60
        '
        'Planification
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(883, 470)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Planification"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Planification"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.SplitContain.Panel1.ResumeLayout(False)
        Me.SplitContain.Panel2.ResumeLayout(False)
        Me.SplitContain.ResumeLayout(False)
        CType(Me.dgvListeTrait, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.dgvTraitementSelect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTraitementEnr, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContain As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvTraitementEnr As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents CbIntitule As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvListeTrait As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnSupprimer As System.Windows.Forms.Button
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnEnlever As System.Windows.Forms.Button
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents dgvTraitementSelect As System.Windows.Forms.DataGridView
    Friend WithEvents intituleSelect As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossiers As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents rangSelect As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefSelect As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents intituleTrait As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents typeTrait As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Consulter As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents BTsup As System.Windows.Forms.Button
    Friend WithEvents BTupdate As System.Windows.Forms.Button
    Friend WithEvents Intitule As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Execution As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Heure1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Heure2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Heure As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Critere1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Critere2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateM As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents IDDossier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Rang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents supEnr As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
