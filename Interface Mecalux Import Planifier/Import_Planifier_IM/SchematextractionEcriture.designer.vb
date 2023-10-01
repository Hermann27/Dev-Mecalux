<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SchematextractionEcriture
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SchematextractionEcriture))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.Cible = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.BaseCpta = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Type = New System.Windows.Forms.DataGridViewButtonColumn
        Me.TypeFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FeuilleExcel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NomFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RepExtraction = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChoixExtraction = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Flag = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Valeur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EstEntete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_DelRow = New System.Windows.Forms.Button
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Cible1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BaseCpta1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TypeFormat1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FeuilleExcel1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NameFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossier1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RepExtraction1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChoixExtraction1 = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Flag1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Valeur1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EstEntete1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CheminForma = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BT_Update = New System.Windows.Forms.Button
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Update)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Delete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Quit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Save)
        Me.SplitContainer1.Size = New System.Drawing.Size(953, 586)
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
        Me.SplitContainer2.Size = New System.Drawing.Size(953, 551)
        Me.SplitContainer2.SplitterDistance = 321
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
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cible, Me.BaseCpta, Me.Type, Me.TypeFormat, Me.FeuilleExcel, Me.NomFormat, Me.IDDossier, Me.RepExtraction, Me.ChoixExtraction, Me.Flag, Me.Valeur, Me.EstEntete, Me.Chemin})
        Me.DataListeSchema.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeSchema.Location = New System.Drawing.Point(0, 30)
        Me.DataListeSchema.MultiSelect = False
        Me.DataListeSchema.Name = "DataListeSchema"
        Me.DataListeSchema.RowHeadersVisible = False
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.DataListeSchema.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowTemplate.Height = 24
        Me.DataListeSchema.Size = New System.Drawing.Size(953, 291)
        Me.DataListeSchema.TabIndex = 44
        '
        'Cible
        '
        Me.Cible.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Cible.HeaderText = "Cible"
        Me.Cible.Items.AddRange(New Object() {"FTP", "Repertoire"})
        Me.Cible.Name = "Cible"
        '
        'BaseCpta
        '
        Me.BaseCpta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BaseCpta.HeaderText = "Base Comptable*"
        Me.BaseCpta.Name = "BaseCpta"
        Me.BaseCpta.Width = 110
        '
        'Type
        '
        Me.Type.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Type.DefaultCellStyle = DataGridViewCellStyle1
        Me.Type.HeaderText = "Type"
        Me.Type.Name = "Type"
        Me.Type.Text = "..."
        Me.Type.UseColumnTextForButtonValue = True
        Me.Type.Width = 40
        '
        'TypeFormat
        '
        Me.TypeFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.TypeFormat.HeaderText = "Type de Format*"
        Me.TypeFormat.Name = "TypeFormat"
        Me.TypeFormat.ReadOnly = True
        Me.TypeFormat.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.TypeFormat.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'FeuilleExcel
        '
        Me.FeuilleExcel.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FeuilleExcel.HeaderText = "Feuille Excel"
        Me.FeuilleExcel.Name = "FeuilleExcel"
        '
        'NomFormat
        '
        Me.NomFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Format = "N0"
        Me.NomFormat.DefaultCellStyle = DataGridViewCellStyle2
        Me.NomFormat.HeaderText = "Fichier Format*"
        Me.NomFormat.Name = "NomFormat"
        Me.NomFormat.ReadOnly = True
        Me.NomFormat.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NomFormat.Width = 110
        '
        'IDDossier
        '
        Me.IDDossier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier.HeaderText = "IDDossier*"
        Me.IDDossier.Name = "IDDossier"
        Me.IDDossier.Width = 60
        '
        'RepExtraction
        '
        Me.RepExtraction.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RepExtraction.HeaderText = "Repertoire d'extraction/Ftp"
        Me.RepExtraction.Name = "RepExtraction"
        Me.RepExtraction.Width = 220
        '
        'ChoixExtraction
        '
        Me.ChoixExtraction.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChoixExtraction.DefaultCellStyle = DataGridViewCellStyle3
        Me.ChoixExtraction.HeaderText = "Rep"
        Me.ChoixExtraction.Name = "ChoixExtraction"
        Me.ChoixExtraction.Text = "..."
        Me.ChoixExtraction.UseColumnTextForButtonValue = True
        Me.ChoixExtraction.Width = 35
        '
        'Flag
        '
        Me.Flag.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Flag.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.Flag.HeaderText = "Flag entête"
        Me.Flag.Name = "Flag"
        Me.Flag.Width = 90
        '
        'Valeur
        '
        Me.Valeur.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Valeur.HeaderText = "Valeur entête"
        Me.Valeur.Name = "Valeur"
        Me.Valeur.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Valeur.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Valeur.Width = 90
        '
        'EstEntete
        '
        Me.EstEntete.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.EstEntete.FalseValue = "False"
        Me.EstEntete.HeaderText = "EstEntête"
        Me.EstEntete.Name = "EstEntete"
        Me.EstEntete.TrueValue = "True"
        Me.EstEntete.Width = 55
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle4.Format = "N0"
        Me.Chemin.DefaultCellStyle = DataGridViewCellStyle4
        Me.Chemin.HeaderText = "Repertoire du Format*"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.ReadOnly = True
        Me.Chemin.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Chemin.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(953, 30)
        Me.Panel2.TabIndex = 43
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(893, 30)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BT_ADD)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox2.Location = New System.Drawing.Point(893, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(29, 30)
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
        Me.GroupBox1.Location = New System.Drawing.Point(922, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(31, 30)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BT_DelRow
        '
        Me.BT_DelRow.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
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
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cible1, Me.BaseCpta1, Me.TypeFormat1, Me.FeuilleExcel1, Me.NameFormat, Me.IDDossier1, Me.RepExtraction1, Me.ChoixExtraction1, Me.Flag1, Me.Valeur1, Me.EstEntete1, Me.Supprimer, Me.CheminForma})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 15)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(953, 211)
        Me.DataListeIntegrer.TabIndex = 10
        '
        'Cible1
        '
        Me.Cible1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Cible1.HeaderText = "Cible"
        Me.Cible1.Name = "Cible1"
        Me.Cible1.ReadOnly = True
        '
        'BaseCpta1
        '
        Me.BaseCpta1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BaseCpta1.HeaderText = "Base Comptable*"
        Me.BaseCpta1.Name = "BaseCpta1"
        Me.BaseCpta1.Width = 120
        '
        'TypeFormat1
        '
        Me.TypeFormat1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.TypeFormat1.HeaderText = "Type de Format*"
        Me.TypeFormat1.Name = "TypeFormat1"
        Me.TypeFormat1.ReadOnly = True
        Me.TypeFormat1.Width = 110
        '
        'FeuilleExcel1
        '
        Me.FeuilleExcel1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FeuilleExcel1.HeaderText = "Feuille Excel"
        Me.FeuilleExcel1.Name = "FeuilleExcel1"
        '
        'NameFormat
        '
        Me.NameFormat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.NullValue = Nothing
        Me.NameFormat.DefaultCellStyle = DataGridViewCellStyle6
        Me.NameFormat.HeaderText = "Fichier Format*"
        Me.NameFormat.Name = "NameFormat"
        Me.NameFormat.ReadOnly = True
        Me.NameFormat.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NameFormat.Width = 110
        '
        'IDDossier1
        '
        Me.IDDossier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier1.HeaderText = "IDDossier*"
        Me.IDDossier1.Name = "IDDossier1"
        Me.IDDossier1.Width = 60
        '
        'RepExtraction1
        '
        Me.RepExtraction1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RepExtraction1.HeaderText = "Repertoire d'extraction/Ftp"
        Me.RepExtraction1.Name = "RepExtraction1"
        Me.RepExtraction1.Width = 200
        '
        'ChoixExtraction1
        '
        Me.ChoixExtraction1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChoixExtraction1.DefaultCellStyle = DataGridViewCellStyle7
        Me.ChoixExtraction1.HeaderText = "Rep"
        Me.ChoixExtraction1.Name = "ChoixExtraction1"
        Me.ChoixExtraction1.Text = "..."
        Me.ChoixExtraction1.UseColumnTextForButtonValue = True
        Me.ChoixExtraction1.Width = 35
        '
        'Flag1
        '
        Me.Flag1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Flag1.HeaderText = "Champ Flag"
        Me.Flag1.Name = "Flag1"
        Me.Flag1.ReadOnly = True
        Me.Flag1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Flag1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Flag1.Width = 90
        '
        'Valeur1
        '
        Me.Valeur1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Valeur1.HeaderText = "Valeur Flag"
        Me.Valeur1.Name = "Valeur1"
        Me.Valeur1.Width = 90
        '
        'EstEntete1
        '
        Me.EstEntete1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.EstEntete1.FalseValue = "False"
        Me.EstEntete1.HeaderText = "EstEntête"
        Me.EstEntete1.Name = "EstEntete1"
        Me.EstEntete1.TrueValue = "True"
        Me.EstEntete1.Width = 55
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Suppr."
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 40
        '
        'CheminForma
        '
        Me.CheminForma.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheminForma.DefaultCellStyle = DataGridViewCellStyle8
        Me.CheminForma.HeaderText = "Nom du repertoire*"
        Me.CheminForma.Name = "CheminForma"
        Me.CheminForma.ReadOnly = True
        Me.CheminForma.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CheminForma.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(953, 15)
        Me.Panel1.TabIndex = 9
        '
        'BT_Update
        '
        Me.BT_Update.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Update.Location = New System.Drawing.Point(292, 5)
        Me.BT_Update.Name = "BT_Update"
        Me.BT_Update.Size = New System.Drawing.Size(62, 23)
        Me.BT_Update.TabIndex = 32
        Me.BT_Update.Text = "&Modifier"
        Me.BT_Update.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Update.UseVisualStyleBackColor = True
        '
        'BT_Delete
        '
        Me.BT_Delete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BT_Delete.Image = Global.Import_Planifier_IM.My.Resources.Resources.criticalind_status
        Me.BT_Delete.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Delete.Location = New System.Drawing.Point(421, 6)
        Me.BT_Delete.Name = "BT_Delete"
        Me.BT_Delete.Size = New System.Drawing.Size(75, 22)
        Me.BT_Delete.TabIndex = 1
        Me.BT_Delete.Text = "&Supprimer"
        Me.BT_Delete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Delete.UseVisualStyleBackColor = True
        '
        'BT_Quit
        '
        Me.BT_Quit.Image = Global.Import_Planifier_IM.My.Resources.Resources.image034
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quit.Location = New System.Drawing.Point(669, 5)
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
        Me.BT_Save.Location = New System.Drawing.Point(540, 5)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(86, 23)
        Me.BT_Save.TabIndex = 1
        Me.BT_Save.Text = "&Enregistrer"
        Me.BT_Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'SchematextractionEcriture
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(953, 586)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SchematextractionEcriture"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parametrage d'extraction des <Ecritures Comptables>"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
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
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
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
    Friend WithEvents BT_Update As System.Windows.Forms.Button
    Friend WithEvents Cible As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents BaseCpta As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Type As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents TypeFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FeuilleExcel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NomFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RepExtraction As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChoixExtraction As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Flag As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Valeur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstEntete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Cible1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BaseCpta1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TypeFormat1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FeuilleExcel1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossier1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RepExtraction1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChoixExtraction1 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Flag1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Valeur1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstEntete1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CheminForma As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
