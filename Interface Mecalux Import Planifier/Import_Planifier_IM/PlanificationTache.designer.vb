<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlanificationTache
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PlanificationTache))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BTsup = New System.Windows.Forms.Button
        Me.BTupdate = New System.Windows.Forms.Button
        Me.ChkLancer = New System.Windows.Forms.CheckBox
        Me.Bt_New = New System.Windows.Forms.Button
        Me.Bt_Enregistrer = New System.Windows.Forms.Button
        Me.TxtIntitule = New System.Windows.Forms.TextBox
        Me.TxtIDTache = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Intitule = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Tache = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Traitement = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Lancer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer1.Size = New System.Drawing.Size(833, 337)
        Me.SplitContainer1.SplitterDistance = 114
        Me.SplitContainer1.TabIndex = 60
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BTsup)
        Me.GroupBox1.Controls.Add(Me.BTupdate)
        Me.GroupBox1.Controls.Add(Me.ChkLancer)
        Me.GroupBox1.Controls.Add(Me.Bt_New)
        Me.GroupBox1.Controls.Add(Me.Bt_Enregistrer)
        Me.GroupBox1.Controls.Add(Me.TxtIntitule)
        Me.GroupBox1.Controls.Add(Me.TxtIDTache)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(833, 114)
        Me.GroupBox1.TabIndex = 60
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Gestion de la planification des tâches"
        '
        'BTsup
        '
        Me.BTsup.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTsup.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTsup.Location = New System.Drawing.Point(591, 63)
        Me.BTsup.Name = "BTsup"
        Me.BTsup.Size = New System.Drawing.Size(85, 23)
        Me.BTsup.TabIndex = 62
        Me.BTsup.Text = "Supprimer"
        Me.BTsup.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTsup.UseVisualStyleBackColor = True
        '
        'BTupdate
        '
        Me.BTupdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTupdate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTupdate.Location = New System.Drawing.Point(717, 63)
        Me.BTupdate.Name = "BTupdate"
        Me.BTupdate.Size = New System.Drawing.Size(81, 23)
        Me.BTupdate.TabIndex = 61
        Me.BTupdate.Text = "Modifier"
        Me.BTupdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTupdate.UseVisualStyleBackColor = True
        '
        'ChkLancer
        '
        Me.ChkLancer.AutoSize = True
        Me.ChkLancer.Location = New System.Drawing.Point(138, 88)
        Me.ChkLancer.Name = "ChkLancer"
        Me.ChkLancer.Size = New System.Drawing.Size(137, 20)
        Me.ChkLancer.TabIndex = 2
        Me.ChkLancer.Text = "Activer/Désactiver"
        Me.ChkLancer.UseVisualStyleBackColor = True
        '
        'Bt_New
        '
        Me.Bt_New.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bt_New.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Bt_New.Location = New System.Drawing.Point(253, 30)
        Me.Bt_New.Name = "Bt_New"
        Me.Bt_New.Size = New System.Drawing.Size(77, 23)
        Me.Bt_New.TabIndex = 3
        Me.Bt_New.Text = "Nouveau"
        Me.Bt_New.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Bt_New.UseVisualStyleBackColor = True
        '
        'Bt_Enregistrer
        '
        Me.Bt_Enregistrer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bt_Enregistrer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Bt_Enregistrer.Location = New System.Drawing.Point(350, 30)
        Me.Bt_Enregistrer.Name = "Bt_Enregistrer"
        Me.Bt_Enregistrer.Size = New System.Drawing.Size(88, 23)
        Me.Bt_Enregistrer.TabIndex = 4
        Me.Bt_Enregistrer.Text = "Enregistrer"
        Me.Bt_Enregistrer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Bt_Enregistrer.UseVisualStyleBackColor = True
        '
        'TxtIntitule
        '
        Me.TxtIntitule.Location = New System.Drawing.Point(138, 61)
        Me.TxtIntitule.Name = "TxtIntitule"
        Me.TxtIntitule.Size = New System.Drawing.Size(300, 22)
        Me.TxtIntitule.TabIndex = 0
        '
        'TxtIDTache
        '
        Me.TxtIDTache.Location = New System.Drawing.Point(138, 31)
        Me.TxtIDTache.Name = "TxtIDTache"
        Me.TxtIDTache.Size = New System.Drawing.Size(91, 22)
        Me.TxtIDTache.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 16)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "Intitulé Tâche *"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 16)
        Me.Label1.TabIndex = 59
        Me.Label1.Text = "Rang Tâche"
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Intitule, Me.Tache, Me.Traitement, Me.Lancer, Me.Supprimer})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataListeIntegrer.Size = New System.Drawing.Size(833, 219)
        Me.DataListeIntegrer.TabIndex = 46
        '
        'Intitule
        '
        Me.Intitule.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Intitule.FillWeight = 55.27975!
        Me.Intitule.HeaderText = "Intitulé Tache"
        Me.Intitule.Name = "Intitule"
        Me.Intitule.ReadOnly = True
        '
        'Tache
        '
        Me.Tache.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Tache.FillWeight = 55.27975!
        Me.Tache.HeaderText = "Rang Tâche"
        Me.Tache.Name = "Tache"
        '
        'Traitement
        '
        Me.Traitement.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Traitement.DefaultCellStyle = DataGridViewCellStyle1
        Me.Traitement.HeaderText = "Traitement Rattaché"
        Me.Traitement.Name = "Traitement"
        Me.Traitement.Text = "..."
        Me.Traitement.UseColumnTextForButtonValue = True
        Me.Traitement.Width = 110
        '
        'Lancer
        '
        Me.Lancer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Lancer.HeaderText = "Lancer"
        Me.Lancer.Name = "Lancer"
        Me.Lancer.Width = 55
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Supprimer"
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 70
        '
        'PlanificationTache
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(833, 337)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "PlanificationTache"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gestion de la planification des tâches"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtIntitule As System.Windows.Forms.TextBox
    Friend WithEvents TxtIDTache As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Bt_Enregistrer As System.Windows.Forms.Button
    Friend WithEvents Bt_New As System.Windows.Forms.Button
    Friend WithEvents ChkLancer As System.Windows.Forms.CheckBox
    Friend WithEvents BTupdate As System.Windows.Forms.Button
    Friend WithEvents BTsup As System.Windows.Forms.Button
    Friend WithEvents Intitule As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Tache As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Traitement As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Lancer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
