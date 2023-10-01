<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AjouterUnFormatMvt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AjouterUnFormatMvt))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DataListeFormat = New System.Windows.Forms.DataGridView
        Me.TypeFormat = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Format = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Fichier = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ima = New System.Windows.Forms.DataGridViewImageColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_Del = New System.Windows.Forms.Button
        Me.BT_Quit = New System.Windows.Forms.Button
        Me.BT_Save = New System.Windows.Forms.Button
        Me.OpenFileFormat = New System.Windows.Forms.OpenFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.DataGridViewImageColumn1 = New System.Windows.Forms.DataGridViewImageColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataListeFormat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataListeFormat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Quit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Save)
        Me.SplitContainer1.Size = New System.Drawing.Size(688, 394)
        Me.SplitContainer1.SplitterDistance = 362
        Me.SplitContainer1.TabIndex = 0
        '
        'DataListeFormat
        '
        Me.DataListeFormat.AllowUserToAddRows = False
        Me.DataListeFormat.AllowUserToDeleteRows = False
        Me.DataListeFormat.AllowUserToOrderColumns = True
        Me.DataListeFormat.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeFormat.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.TypeFormat, Me.Format, Me.Chemin, Me.Fichier, Me.ima})
        Me.DataListeFormat.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeFormat.Location = New System.Drawing.Point(0, 33)
        Me.DataListeFormat.Name = "DataListeFormat"
        Me.DataListeFormat.RowHeadersVisible = False
        Me.DataListeFormat.Size = New System.Drawing.Size(688, 329)
        Me.DataListeFormat.TabIndex = 4
        '
        'TypeFormat
        '
        Me.TypeFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.TypeFormat.HeaderText = "Type de Format"
        Me.TypeFormat.Items.AddRange(New Object() {"Longueur Fixe", "Point virgule", "Tabulation", "Excel", "Table", "Pipe"})
        Me.TypeFormat.Name = "TypeFormat"
        Me.TypeFormat.Width = 120
        '
        'Format
        '
        Me.Format.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Format.HeaderText = "Nom du Format"
        Me.Format.Name = "Format"
        Me.Format.ReadOnly = True
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Chemin.HeaderText = "Chemin du Format"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.ReadOnly = True
        '
        'Fichier
        '
        Me.Fichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Fichier.HeaderText = "Fichier"
        Me.Fichier.Name = "Fichier"
        Me.Fichier.Text = "Cliquer"
        Me.Fichier.UseColumnTextForButtonValue = True
        Me.Fichier.Width = 60
        '
        'ima
        '
        Me.ima.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ima.HeaderText = "?"
        Me.ima.Image = Global.Import_Planifier_IM.My.Resources.Resources.documents_16
        Me.ima.Name = "ima"
        Me.ima.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ima.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ima.Width = 30
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(688, 33)
        Me.Panel1.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(620, 33)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BT_ADD)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox2.Location = New System.Drawing.Point(620, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(35, 33)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        '
        'BT_ADD
        '
        Me.BT_ADD.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.BT_ADD.Location = New System.Drawing.Point(5, 10)
        Me.BT_ADD.Name = "BT_ADD"
        Me.BT_ADD.Size = New System.Drawing.Size(28, 20)
        Me.BT_ADD.TabIndex = 1
        Me.BT_ADD.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BT_Del)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox1.Location = New System.Drawing.Point(655, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(33, 33)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'BT_Del
        '
        Me.BT_Del.Image = Global.Import_Planifier_IM.My.Resources.Resources.accepter
        Me.BT_Del.Location = New System.Drawing.Point(3, 10)
        Me.BT_Del.Name = "BT_Del"
        Me.BT_Del.Size = New System.Drawing.Size(27, 20)
        Me.BT_Del.TabIndex = 1
        Me.BT_Del.UseVisualStyleBackColor = True
        '
        'BT_Quit
        '
        Me.BT_Quit.Image = Global.Import_Planifier_IM.My.Resources.Resources.image034
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quit.Location = New System.Drawing.Point(458, 0)
        Me.BT_Quit.Name = "BT_Quit"
        Me.BT_Quit.Size = New System.Drawing.Size(72, 25)
        Me.BT_Quit.TabIndex = 1
        Me.BT_Quit.Text = "&Quitter"
        Me.BT_Quit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Quit.UseVisualStyleBackColor = True
        '
        'BT_Save
        '
        Me.BT_Save.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_161
        Me.BT_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Save.Location = New System.Drawing.Point(343, 0)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(78, 25)
        Me.BT_Save.TabIndex = 0
        Me.BT_Save.Text = "&Enregistrer"
        Me.BT_Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'DataGridViewImageColumn1
        '
        Me.DataGridViewImageColumn1.HeaderText = "?"
        Me.DataGridViewImageColumn1.Image = Global.Import_Planifier_IM.My.Resources.Resources.documents_16
        Me.DataGridViewImageColumn1.Name = "DataGridViewImageColumn1"
        Me.DataGridViewImageColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewImageColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'AjouterUnFormatMvt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 394)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AjouterUnFormatMvt"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ajouter les formats d'import <des documents> Existant"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataListeFormat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BT_Quit As System.Windows.Forms.Button
    Friend WithEvents BT_Save As System.Windows.Forms.Button
    Friend WithEvents OpenFileFormat As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BT_ADD As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents DataListeFormat As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BT_Del As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridViewImageColumn1 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents TypeFormat As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Format As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fichier As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ima As System.Windows.Forms.DataGridViewImageColumn
End Class
