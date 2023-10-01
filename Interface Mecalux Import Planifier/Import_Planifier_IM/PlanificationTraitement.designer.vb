<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlanificationTraitement
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PlanificationTraitement))
        Me.GrvTraitement = New System.Windows.Forms.DataGridView
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BtValider = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.lblInfos = New System.Windows.Forms.Label
        Me.Societe1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.charge = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Heure1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Heure2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Heure = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Critere1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Critere2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateM = New System.Windows.Forms.DataGridViewButtonColumn
        Me.IDDossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Rang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.supEnr = New System.Windows.Forms.DataGridViewCheckBoxColumn
        CType(Me.GrvTraitement, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GrvTraitement
        '
        Me.GrvTraitement.AllowUserToAddRows = False
        Me.GrvTraitement.AllowUserToDeleteRows = False
        Me.GrvTraitement.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GrvTraitement.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GrvTraitement.Location = New System.Drawing.Point(0, 0)
        Me.GrvTraitement.Name = "GrvTraitement"
        Me.GrvTraitement.RowHeadersVisible = False
        Me.GrvTraitement.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GrvTraitement.Size = New System.Drawing.Size(824, 296)
        Me.GrvTraitement.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lblInfos)
        Me.Panel1.Controls.Add(Me.BtValider)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 296)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(824, 35)
        Me.Panel1.TabIndex = 1
        '
        'BtValider
        '
        Me.BtValider.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtValider.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtValider.Location = New System.Drawing.Point(623, 5)
        Me.BtValider.Name = "BtValider"
        Me.BtValider.Size = New System.Drawing.Size(82, 26)
        Me.BtValider.TabIndex = 1
        Me.BtValider.Text = "Valider"
        Me.BtValider.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(728, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Quitter"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Societe1, Me.charge, Me.Heure1, Me.Heure2, Me.Heure, Me.Critere1, Me.Critere2, Me.DateM, Me.IDDossier, Me.Rang, Me.supEnr})
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.Margin = New System.Windows.Forms.Padding(2)
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(824, 296)
        Me.DataListeIntegrer.TabIndex = 2
        '
        'lblInfos
        '
        Me.lblInfos.AutoSize = True
        Me.lblInfos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfos.ForeColor = System.Drawing.Color.Red
        Me.lblInfos.Location = New System.Drawing.Point(3, 10)
        Me.lblInfos.Name = "lblInfos"
        Me.lblInfos.Size = New System.Drawing.Size(32, 16)
        Me.lblInfos.TabIndex = 2
        Me.lblInfos.Text = "......"
        '
        'Societe1
        '
        Me.Societe1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Societe1.HeaderText = "Société"
        Me.Societe1.Name = "Societe1"
        '
        'charge
        '
        Me.charge.HeaderText = "Ajouter"
        Me.charge.Name = "charge"
        '
        'Heure1
        '
        Me.Heure1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Heure1.HeaderText = "Heure debut"
        Me.Heure1.Name = "Heure1"
        Me.Heure1.Width = 70
        '
        'Heure2
        '
        Me.Heure2.HeaderText = "Heure fin"
        Me.Heure2.Name = "Heure2"
        '
        'Heure
        '
        Me.Heure.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Heure.DefaultCellStyle = DataGridViewCellStyle1
        Me.Heure.HeaderText = "Heure"
        Me.Heure.Name = "Heure"
        Me.Heure.Text = "..."
        Me.Heure.UseColumnTextForButtonValue = True
        Me.Heure.Width = 50
        '
        'Critere1
        '
        Me.Critere1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Critere1.HeaderText = "Date debut"
        Me.Critere1.Name = "Critere1"
        '
        'Critere2
        '
        Me.Critere2.HeaderText = "Date fin"
        Me.Critere2.Name = "Critere2"
        '
        'DateM
        '
        Me.DateM.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateM.DefaultCellStyle = DataGridViewCellStyle2
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
        Me.IDDossier.Visible = False
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
        Me.supEnr.Visible = False
        Me.supEnr.Width = 60
        '
        'PlanificationTraitement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(824, 331)
        Me.Controls.Add(Me.DataListeIntegrer)
        Me.Controls.Add(Me.GrvTraitement)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PlanificationTraitement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Traitement à Planifier"
        CType(Me.GrvTraitement, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GrvTraitement As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BtValider As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblInfos As System.Windows.Forms.Label
    Friend WithEvents Societe1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents charge As System.Windows.Forms.DataGridViewCheckBoxColumn
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
