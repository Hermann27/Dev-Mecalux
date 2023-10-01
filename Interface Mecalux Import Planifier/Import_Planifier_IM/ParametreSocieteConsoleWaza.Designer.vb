<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ParametreSocieteConsoleWaza
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ParametreSocieteConsoleWaza))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.Societe = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UserSage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PasseSage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.find = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Serveur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.bdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Utilisateur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Passe = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.BT_DelRow = New System.Windows.Forms.Button
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Societe1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Chemin1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UserSage1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PasseSage1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Serveur1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.bdd1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NomUtil = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Mot = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BT_Quit = New System.Windows.Forms.Button
        Me.BtnTest = New System.Windows.Forms.Button
        Me.BT_Save = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.FindFile = New System.Windows.Forms.OpenFileDialog
        Me.FindFolder = New System.Windows.Forms.FolderBrowserDialog
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
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
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Quit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnTest)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Save)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1037, 489)
        Me.SplitContainer1.SplitterDistance = 452
        Me.SplitContainer1.TabIndex = 1
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
        Me.SplitContainer2.Size = New System.Drawing.Size(1037, 452)
        Me.SplitContainer2.SplitterDistance = 260
        Me.SplitContainer2.TabIndex = 0
        '
        'DataListeSchema
        '
        Me.DataListeSchema.AllowUserToAddRows = False
        Me.DataListeSchema.AllowUserToDeleteRows = False
        Me.DataListeSchema.AllowUserToOrderColumns = True
        Me.DataListeSchema.AllowUserToResizeRows = False
        Me.DataListeSchema.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeSchema.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DataListeSchema.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DataListeSchema.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Societe, Me.Type, Me.Chemin, Me.UserSage, Me.PasseSage, Me.find, Me.Serveur, Me.bdd, Me.Utilisateur, Me.Passe})
        Me.DataListeSchema.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeSchema.EnableHeadersVisualStyles = False
        Me.DataListeSchema.GridColor = System.Drawing.Color.SlateGray
        Me.DataListeSchema.Location = New System.Drawing.Point(0, 30)
        Me.DataListeSchema.MultiSelect = False
        Me.DataListeSchema.Name = "DataListeSchema"
        Me.DataListeSchema.RowHeadersVisible = False
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.DataListeSchema.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowTemplate.Height = 24
        Me.DataListeSchema.Size = New System.Drawing.Size(1037, 230)
        Me.DataListeSchema.TabIndex = 44
        '
        'Societe
        '
        Me.Societe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Format = "N0"
        Me.Societe.DefaultCellStyle = DataGridViewCellStyle1
        Me.Societe.HeaderText = " Société"
        Me.Societe.Name = "Societe"
        Me.Societe.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Societe.Width = 114
        '
        'Type
        '
        Me.Type.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Type.HeaderText = "Type Base"
        Me.Type.Name = "Type"
        Me.Type.Width = 110
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Chemin.HeaderText = "Fichier Sage"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.ReadOnly = True
        Me.Chemin.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Chemin.Width = 160
        '
        'UserSage
        '
        Me.UserSage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UserSage.HeaderText = "Nom Sage"
        Me.UserSage.Name = "UserSage"
        '
        'PasseSage
        '
        Me.PasseSage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PasseSage.HeaderText = "Mot de Passe Sage"
        Me.PasseSage.Name = "PasseSage"
        Me.PasseSage.Width = 125
        '
        'find
        '
        Me.find.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.find.DefaultCellStyle = DataGridViewCellStyle2
        Me.find.HeaderText = "Rep"
        Me.find.Name = "find"
        Me.find.Text = "..."
        Me.find.UseColumnTextForButtonValue = True
        Me.find.Width = 35
        '
        'Serveur
        '
        Me.Serveur.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Serveur.HeaderText = "Serveur SQL"
        Me.Serveur.Name = "Serveur"
        '
        'bdd
        '
        Me.bdd.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.bdd.HeaderText = "Base SQL"
        Me.bdd.Name = "bdd"
        '
        'Utilisateur
        '
        Me.Utilisateur.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Utilisateur.HeaderText = "Nom SQL"
        Me.Utilisateur.Name = "Utilisateur"
        Me.Utilisateur.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Utilisateur.Width = 80
        '
        'Passe
        '
        Me.Passe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Passe.HeaderText = "Mot de Passe"
        Me.Passe.Name = "Passe"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1037, 30)
        Me.Panel2.TabIndex = 43
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.DarkSlateGray
        Me.GroupBox3.Controls.Add(Me.BT_ADD)
        Me.GroupBox3.Controls.Add(Me.BT_DelRow)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1037, 30)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'BT_ADD
        '
        Me.BT_ADD.BackColor = System.Drawing.Color.White
        Me.BT_ADD.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_ADD.Image = Global.Import_Planifier_IM.My.Resources.Resources.create
        Me.BT_ADD.Location = New System.Drawing.Point(951, 9)
        Me.BT_ADD.Name = "BT_ADD"
        Me.BT_ADD.Size = New System.Drawing.Size(35, 18)
        Me.BT_ADD.TabIndex = 1
        Me.BT_ADD.UseVisualStyleBackColor = False
        '
        'BT_DelRow
        '
        Me.BT_DelRow.BackColor = System.Drawing.Color.White
        Me.BT_DelRow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_DelRow.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.BT_DelRow.Location = New System.Drawing.Point(992, 9)
        Me.BT_DelRow.Name = "BT_DelRow"
        Me.BT_DelRow.Size = New System.Drawing.Size(35, 18)
        Me.BT_DelRow.TabIndex = 0
        Me.BT_DelRow.UseVisualStyleBackColor = False
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
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Societe1, Me.Type1, Me.Chemin1, Me.UserSage1, Me.PasseSage1, Me.Serveur1, Me.bdd1, Me.NomUtil, Me.Mot, Me.Supprimer})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.EnableHeadersVisualStyles = False
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(1037, 188)
        Me.DataListeIntegrer.TabIndex = 10
        '
        'Societe1
        '
        Me.Societe1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Societe1.HeaderText = "Société"
        Me.Societe1.Name = "Societe1"
        Me.Societe1.ReadOnly = True
        Me.Societe1.Width = 114
        '
        'Type1
        '
        Me.Type1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Type1.DefaultCellStyle = DataGridViewCellStyle4
        Me.Type1.HeaderText = "Type Base"
        Me.Type1.Name = "Type1"
        Me.Type1.ReadOnly = True
        Me.Type1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Type1.Width = 110
        '
        'Chemin1
        '
        Me.Chemin1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Chemin1.HeaderText = "Fichier Sage"
        Me.Chemin1.Name = "Chemin1"
        Me.Chemin1.ReadOnly = True
        Me.Chemin1.Width = 150
        '
        'UserSage1
        '
        Me.UserSage1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UserSage1.HeaderText = "Nom Sage"
        Me.UserSage1.Name = "UserSage1"
        '
        'PasseSage1
        '
        Me.PasseSage1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PasseSage1.HeaderText = "Mot de Passe Sage"
        Me.PasseSage1.Name = "PasseSage1"
        Me.PasseSage1.Width = 125
        '
        'Serveur1
        '
        Me.Serveur1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Serveur1.HeaderText = "Serveur SQL"
        Me.Serveur1.Name = "Serveur1"
        '
        'bdd1
        '
        Me.bdd1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.bdd1.HeaderText = "Base SQL"
        Me.bdd1.Name = "bdd1"
        '
        'NomUtil
        '
        Me.NomUtil.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle5.Format = "N0"
        Me.NomUtil.DefaultCellStyle = DataGridViewCellStyle5
        Me.NomUtil.FillWeight = 40.0!
        Me.NomUtil.HeaderText = "Nom SQL"
        Me.NomUtil.Name = "NomUtil"
        Me.NomUtil.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.NomUtil.Width = 80
        '
        'Mot
        '
        Me.Mot.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Mot.HeaderText = "Mot de Passe"
        Me.Mot.Name = "Mot"
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Supprimer"
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 58
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1037, 188)
        Me.Panel1.TabIndex = 9
        '
        'BT_Quit
        '
        Me.BT_Quit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_Quit.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Quit.Location = New System.Drawing.Point(894, 4)
        Me.BT_Quit.Name = "BT_Quit"
        Me.BT_Quit.Size = New System.Drawing.Size(110, 26)
        Me.BT_Quit.TabIndex = 4
        Me.BT_Quit.Text = "&Quitter"
        Me.BT_Quit.UseVisualStyleBackColor = True
        '
        'BtnTest
        '
        Me.BtnTest.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        'Me.BtnTest.Image = Global.Import_Planifier_IM.My.Resources.Resources.btn_valider
        Me.BtnTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnTest.Location = New System.Drawing.Point(279, 4)
        Me.BtnTest.Name = "BtnTest"
        Me.BtnTest.Size = New System.Drawing.Size(163, 26)
        Me.BtnTest.TabIndex = 3
        Me.BtnTest.Text = "     Tester Les Connexions SQL"
        Me.BtnTest.UseVisualStyleBackColor = True
        '
        'BT_Save
        '
        Me.BT_Save.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        'Me.BT_Save.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_1611
        Me.BT_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.Location = New System.Drawing.Point(481, 4)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(110, 26)
        Me.BT_Save.TabIndex = 2
        Me.BT_Save.Text = "&Enregistrer"
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Image = Global.Import_Planifier_IM.My.Resources.Resources.AnalyzeWizard11
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(639, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(110, 26)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "&Modifier"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Image = Global.Import_Planifier_IM.My.Resources.Resources.criticalind_status
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(769, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(110, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "&Supprimer"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ParametreSocieteConsoleWaza
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1037, 489)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ParametreSocieteConsoleWaza"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parametrage des Societes"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataListeSchema As System.Windows.Forms.DataGridView
    Friend WithEvents Societe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UserSage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PasseSage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents find As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Serveur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents bdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Utilisateur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Passe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Societe1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Chemin1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UserSage1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PasseSage1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Serveur1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents bdd1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NomUtil As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Mot As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BT_Quit As System.Windows.Forms.Button
    Friend WithEvents BtnTest As System.Windows.Forms.Button
    Friend WithEvents BT_Save As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FindFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FindFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BT_ADD As System.Windows.Forms.Button
    Friend WithEvents BT_DelRow As System.Windows.Forms.Button
End Class
