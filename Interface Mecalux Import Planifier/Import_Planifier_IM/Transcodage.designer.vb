<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Transcodage
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Transcodage))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataCompte = New System.Windows.Forms.DataGridView
        Me.Fichier = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.CompteFichier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompteImport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Menu2 = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Ligne = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Entete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Datacompte1 = New System.Windows.Forms.DataGridView
        Me.Fichier1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDTraitement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompteFichier1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompteImport1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Menu1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Categorie1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Ligne1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Entete1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Supprime1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_DelRow = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Cbtraitement = New System.Windows.Forms.ComboBox
        Me.CbCat = New System.Windows.Forms.ComboBox
        Me.BT_Del = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BT_Creer = New System.Windows.Forms.Button
        Me.BT_Update = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DataCompte, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Datacompte1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Del)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Creer)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Update)
        Me.SplitContainer1.Size = New System.Drawing.Size(1028, 515)
        Me.SplitContainer1.SplitterDistance = 482
        Me.SplitContainer1.TabIndex = 29
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 44)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DataCompte)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Datacompte1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1028, 438)
        Me.SplitContainer2.SplitterDistance = 121
        Me.SplitContainer2.TabIndex = 45
        '
        'DataCompte
        '
        Me.DataCompte.AllowUserToAddRows = False
        Me.DataCompte.AllowUserToDeleteRows = False
        Me.DataCompte.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataCompte.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Fichier, Me.CompteFichier, Me.CompteImport, Me.Menu2, Me.Ligne, Me.Entete})
        Me.DataCompte.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataCompte.Location = New System.Drawing.Point(0, 0)
        Me.DataCompte.Name = "DataCompte"
        Me.DataCompte.RowHeadersVisible = False
        Me.DataCompte.Size = New System.Drawing.Size(1028, 121)
        Me.DataCompte.TabIndex = 1
        '
        'Fichier
        '
        Me.Fichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Fichier.HeaderText = "Fichier Sage"
        Me.Fichier.Name = "Fichier"
        Me.Fichier.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Fichier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'CompteFichier
        '
        Me.CompteFichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CompteFichier.HeaderText = "Valeur Lue"
        Me.CompteFichier.Name = "CompteFichier"
        Me.CompteFichier.Width = 128
        '
        'CompteImport
        '
        Me.CompteImport.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CompteImport.HeaderText = "Valeur Sage"
        Me.CompteImport.Name = "CompteImport"
        Me.CompteImport.Width = 128
        '
        'Menu2
        '
        Me.Menu2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Menu2.HeaderText = "Menu"
        Me.Menu2.Items.AddRange(New Object() {"Importation", "Exportation"})
        Me.Menu2.Name = "Menu2"
        '
        'Ligne
        '
        Me.Ligne.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Ligne.FalseValue = "False"
        Me.Ligne.HeaderText = "Ligne"
        Me.Ligne.Name = "Ligne"
        Me.Ligne.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Ligne.TrueValue = "True"
        Me.Ligne.Width = 40
        '
        'Entete
        '
        Me.Entete.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Entete.FalseValue = "False"
        Me.Entete.HeaderText = "Entête"
        Me.Entete.Name = "Entete"
        Me.Entete.TrueValue = "True"
        Me.Entete.Width = 40
        '
        'Datacompte1
        '
        Me.Datacompte1.AllowUserToAddRows = False
        Me.Datacompte1.AllowUserToDeleteRows = False
        Me.Datacompte1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Datacompte1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Fichier1, Me.IDTraitement, Me.CompteFichier1, Me.CompteImport1, Me.Menu1, Me.Categorie1, Me.Ligne1, Me.Entete1, Me.Supprime1})
        Me.Datacompte1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Datacompte1.Location = New System.Drawing.Point(0, 0)
        Me.Datacompte1.Name = "Datacompte1"
        Me.Datacompte1.RowHeadersVisible = False
        Me.Datacompte1.Size = New System.Drawing.Size(1028, 313)
        Me.Datacompte1.TabIndex = 2
        '
        'Fichier1
        '
        Me.Fichier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Fichier1.HeaderText = "Fichier Sage"
        Me.Fichier1.Name = "Fichier1"
        Me.Fichier1.ReadOnly = True
        '
        'IDTraitement
        '
        Me.IDTraitement.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDTraitement.HeaderText = "ID du Traitement"
        Me.IDTraitement.Name = "IDTraitement"
        Me.IDTraitement.Width = 120
        '
        'CompteFichier1
        '
        Me.CompteFichier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CompteFichier1.HeaderText = "Valeur Lue"
        Me.CompteFichier1.Name = "CompteFichier1"
        Me.CompteFichier1.ReadOnly = True
        Me.CompteFichier1.Width = 128
        '
        'CompteImport1
        '
        Me.CompteImport1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CompteImport1.HeaderText = "Valeur Sage"
        Me.CompteImport1.Name = "CompteImport1"
        Me.CompteImport1.Width = 128
        '
        'Menu1
        '
        Me.Menu1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Menu1.HeaderText = "Menu"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.ReadOnly = True
        Me.Menu1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Menu1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Categorie1
        '
        Me.Categorie1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Categorie1.HeaderText = "Categorie"
        Me.Categorie1.Name = "Categorie1"
        Me.Categorie1.ReadOnly = True
        Me.Categorie1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Categorie1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Categorie1.Width = 120
        '
        'Ligne1
        '
        Me.Ligne1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Ligne1.HeaderText = "Ligne"
        Me.Ligne1.Name = "Ligne1"
        Me.Ligne1.ReadOnly = True
        Me.Ligne1.Width = 40
        '
        'Entete1
        '
        Me.Entete1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Entete1.HeaderText = "Entête"
        Me.Entete1.Name = "Entete1"
        Me.Entete1.ReadOnly = True
        Me.Entete1.Width = 40
        '
        'Supprime1
        '
        Me.Supprime1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprime1.HeaderText = "Supprimer"
        Me.Supprime1.Name = "Supprime1"
        Me.Supprime1.Width = 60
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1028, 44)
        Me.Panel2.TabIndex = 44
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BT_ADD)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox2.Location = New System.Drawing.Point(968, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(29, 44)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'BT_ADD
        '
        Me.BT_ADD.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
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
        Me.GroupBox1.Location = New System.Drawing.Point(997, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(31, 44)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BT_DelRow
        '
        Me.BT_DelRow.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
        Me.BT_DelRow.Location = New System.Drawing.Point(0, 7)
        Me.BT_DelRow.Name = "BT_DelRow"
        Me.BT_DelRow.Size = New System.Drawing.Size(23, 20)
        Me.BT_DelRow.TabIndex = 2
        Me.BT_DelRow.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Cbtraitement)
        Me.GroupBox3.Controls.Add(Me.CbCat)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1028, 44)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(482, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 15)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "ID Traitement"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(200, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(135, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Catégorie de traitement" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Cbtraitement
        '
        Me.Cbtraitement.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cbtraitement.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cbtraitement.FormattingEnabled = True
        Me.Cbtraitement.Location = New System.Drawing.Point(569, 11)
        Me.Cbtraitement.Name = "Cbtraitement"
        Me.Cbtraitement.Size = New System.Drawing.Size(121, 23)
        Me.Cbtraitement.TabIndex = 1
        '
        'CbCat
        '
        Me.CbCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CbCat.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CbCat.FormattingEnabled = True
        Me.CbCat.Items.AddRange(New Object() {"Document", "Ecritures", "Articles", "Tiers", "CompteA", "Document Stock", "Document Transfert"})
        Me.CbCat.Location = New System.Drawing.Point(341, 13)
        Me.CbCat.Name = "CbCat"
        Me.CbCat.Size = New System.Drawing.Size(121, 23)
        Me.CbCat.TabIndex = 0
        '
        'BT_Del
        '
        Me.BT_Del.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.BT_Del.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Del.Location = New System.Drawing.Point(477, 3)
        Me.BT_Del.Name = "BT_Del"
        Me.BT_Del.Size = New System.Drawing.Size(86, 23)
        Me.BT_Del.TabIndex = 32
        Me.BT_Del.Text = "&Supprimer"
        Me.BT_Del.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Del.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.Import_Planifier_IM.My.Resources.Resources.image033
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(742, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(86, 23)
        Me.Button1.TabIndex = 29
        Me.Button1.Text = "&Quitter"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BT_Creer
        '
        Me.BT_Creer.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_16
        Me.BT_Creer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Creer.Location = New System.Drawing.Point(607, 3)
        Me.BT_Creer.Name = "BT_Creer"
        Me.BT_Creer.Size = New System.Drawing.Size(86, 23)
        Me.BT_Creer.TabIndex = 30
        Me.BT_Creer.Text = "&Enregistrer"
        Me.BT_Creer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Creer.UseVisualStyleBackColor = True
        '
        'BT_Update
        '
        Me.BT_Update.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Update.Location = New System.Drawing.Point(345, 3)
        Me.BT_Update.Name = "BT_Update"
        Me.BT_Update.Size = New System.Drawing.Size(86, 23)
        Me.BT_Update.TabIndex = 31
        Me.BT_Update.Text = "&Modifier"
        Me.BT_Update.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Update.UseVisualStyleBackColor = True
        '
        'Transcodage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1028, 515)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Transcodage"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transcodage des informations des fichiers Sage"
        Me.TopMost = True
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataCompte, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Datacompte1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BT_Del As System.Windows.Forms.Button
    Friend WithEvents BT_Update As System.Windows.Forms.Button
    Friend WithEvents BT_Creer As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents BT_ADD As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BT_DelRow As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataCompte As System.Windows.Forms.DataGridView
    Friend WithEvents Datacompte1 As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Cbtraitement As System.Windows.Forms.ComboBox
    Friend WithEvents CbCat As System.Windows.Forms.ComboBox
    Friend WithEvents Fichier As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents CompteFichier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompteImport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Menu2 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Ligne As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Entete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Fichier1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDTraitement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompteFichier1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompteImport1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Menu1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Categorie1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Ligne1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Entete1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Supprime1 As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
