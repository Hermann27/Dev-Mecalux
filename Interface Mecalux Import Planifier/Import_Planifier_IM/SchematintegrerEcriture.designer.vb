<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SchematintegrerEcriture
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
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SchematintegrerEcriture))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.Cible = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.BaseCpta = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Type = New System.Windows.Forms.DataGridViewButtonColumn
        Me.TypeFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NomFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CheminExport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RechercheFichier = New System.Windows.Forms.DataGridViewButtonColumn
        Me.FeuilleExcel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Deplace = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CreationAuto = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.DossierExport = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.NameFormat = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CheminRepexpor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RechercheFichier1 = New System.Windows.Forms.DataGridViewButtonColumn
        Me.FeuilleExcel1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Deplace1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CreationAuto1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.NomRepxpor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CheminForma = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossier1 = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.SplitContainer1.Size = New System.Drawing.Size(1026, 586)
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
        Me.SplitContainer2.Size = New System.Drawing.Size(1026, 551)
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
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cible, Me.BaseCpta, Me.Type, Me.TypeFormat, Me.NomFormat, Me.CheminExport, Me.RechercheFichier, Me.FeuilleExcel, Me.Chemin, Me.IDDossier, Me.Deplace, Me.CreationAuto, Me.DossierExport})
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
        Me.DataListeSchema.Size = New System.Drawing.Size(1026, 291)
        Me.DataListeSchema.TabIndex = 44
        '
        'Cible
        '
        Me.Cible.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Cible.HeaderText = "Cible"
        Me.Cible.Items.AddRange(New Object() {"FTP", "Repertoire", "BaseSQL"})
        Me.Cible.Name = "Cible"
        '
        'BaseCpta
        '
        Me.BaseCpta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BaseCpta.HeaderText = "Base Comptable*"
        Me.BaseCpta.Name = "BaseCpta"
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
        'CheminExport
        '
        Me.CheminExport.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CheminExport.HeaderText = "Repertoire à importer/Ftp/BaseSQL*"
        Me.CheminExport.Name = "CheminExport"
        Me.CheminExport.Width = 250
        '
        'RechercheFichier
        '
        Me.RechercheFichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RechercheFichier.DefaultCellStyle = DataGridViewCellStyle3
        Me.RechercheFichier.HeaderText = "Rep"
        Me.RechercheFichier.Name = "RechercheFichier"
        Me.RechercheFichier.Text = "..."
        Me.RechercheFichier.ToolTipText = "Rechercher le Fichier à exporter"
        Me.RechercheFichier.UseColumnTextForButtonValue = True
        Me.RechercheFichier.Width = 40
        '
        'FeuilleExcel
        '
        Me.FeuilleExcel.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FeuilleExcel.HeaderText = "F.Excel/Filtre"
        Me.FeuilleExcel.Name = "FeuilleExcel"
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle4.Format = "N0"
        Me.Chemin.DefaultCellStyle = DataGridViewCellStyle4
        Me.Chemin.FillWeight = 40.0!
        Me.Chemin.HeaderText = "Repertoire du Format*"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.ReadOnly = True
        Me.Chemin.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Chemin.Width = 200
        '
        'IDDossier
        '
        Me.IDDossier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier.HeaderText = "IDDossier*"
        Me.IDDossier.Name = "IDDossier"
        Me.IDDossier.Width = 60
        '
        'Deplace
        '
        Me.Deplace.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Deplace.FalseValue = "False"
        Me.Deplace.HeaderText = "Deplace"
        Me.Deplace.Name = "Deplace"
        Me.Deplace.TrueValue = "True"
        Me.Deplace.Width = 50
        '
        'CreationAuto
        '
        Me.CreationAuto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CreationAuto.HeaderText = "Pas de création Auto"
        Me.CreationAuto.Name = "CreationAuto"
        Me.CreationAuto.Width = 120
        '
        'DossierExport
        '
        Me.DossierExport.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.DossierExport.HeaderText = "Repertoire Fichiers "
        Me.DossierExport.Name = "DossierExport"
        Me.DossierExport.ReadOnly = True
        Me.DossierExport.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DossierExport.Visible = False
        Me.DossierExport.Width = 120
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1026, 30)
        Me.Panel2.TabIndex = 43
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(966, 30)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BT_ADD)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox2.Location = New System.Drawing.Point(966, 0)
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
        Me.GroupBox1.Location = New System.Drawing.Point(995, 0)
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
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cible1, Me.BaseCpta1, Me.TypeFormat1, Me.NameFormat, Me.CheminRepexpor, Me.RechercheFichier1, Me.FeuilleExcel1, Me.Deplace1, Me.CreationAuto1, Me.Supprimer, Me.NomRepxpor, Me.CheminForma, Me.IDDossier1})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 15)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle10
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.DataListeIntegrer.Size = New System.Drawing.Size(1026, 211)
        Me.DataListeIntegrer.TabIndex = 10
        '
        'Cible1
        '
        Me.Cible1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Cible1.HeaderText = "Cible"
        Me.Cible1.Name = "Cible1"
        Me.Cible1.ReadOnly = True
        Me.Cible1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Cible1.Width = 70
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
        'CheminRepexpor
        '
        Me.CheminRepexpor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.CheminRepexpor.HeaderText = "Chemin du Repertoire/Ftp/BaseSQL*"
        Me.CheminRepexpor.Name = "CheminRepexpor"
        Me.CheminRepexpor.ReadOnly = True
        '
        'RechercheFichier1
        '
        Me.RechercheFichier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RechercheFichier1.DefaultCellStyle = DataGridViewCellStyle7
        Me.RechercheFichier1.HeaderText = "Rep"
        Me.RechercheFichier1.Name = "RechercheFichier1"
        Me.RechercheFichier1.Text = "..."
        Me.RechercheFichier1.UseColumnTextForButtonValue = True
        Me.RechercheFichier1.Width = 40
        '
        'FeuilleExcel1
        '
        Me.FeuilleExcel1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.FeuilleExcel1.HeaderText = "F.Excel/Filtre"
        Me.FeuilleExcel1.Name = "FeuilleExcel1"
        '
        'Deplace1
        '
        Me.Deplace1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Deplace1.FalseValue = "False"
        Me.Deplace1.HeaderText = "Deplace"
        Me.Deplace1.Name = "Deplace1"
        Me.Deplace1.TrueValue = "True"
        Me.Deplace1.Width = 50
        '
        'CreationAuto1
        '
        Me.CreationAuto1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CreationAuto1.HeaderText = "Pas de création Auto"
        Me.CreationAuto1.Name = "CreationAuto1"
        Me.CreationAuto1.Width = 120
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Supprimer"
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 60
        '
        'NomRepxpor
        '
        Me.NomRepxpor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NomRepxpor.DefaultCellStyle = DataGridViewCellStyle8
        Me.NomRepxpor.HeaderText = "Nom Repertoire*"
        Me.NomRepxpor.Name = "NomRepxpor"
        Me.NomRepxpor.ReadOnly = True
        Me.NomRepxpor.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NomRepxpor.Visible = False
        Me.NomRepxpor.Width = 110
        '
        'CheminForma
        '
        Me.CheminForma.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle9.Format = "N0"
        Me.CheminForma.DefaultCellStyle = DataGridViewCellStyle9
        Me.CheminForma.FillWeight = 40.0!
        Me.CheminForma.HeaderText = "Chemin d'acces du Fichier Format"
        Me.CheminForma.Name = "CheminForma"
        Me.CheminForma.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.CheminForma.Visible = False
        '
        'IDDossier1
        '
        Me.IDDossier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier1.HeaderText = "IDDossier*"
        Me.IDDossier1.Name = "IDDossier1"
        Me.IDDossier1.Width = 60
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1026, 15)
        Me.Panel1.TabIndex = 9
        '
        'BT_Update
        '
        Me.BT_Update.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Update.Location = New System.Drawing.Point(298, 5)
        Me.BT_Update.Name = "BT_Update"
        Me.BT_Update.Size = New System.Drawing.Size(68, 23)
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
        'SchematintegrerEcriture
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1026, 586)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SchematintegrerEcriture"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parametrage des integrations <Ecritures Comptables>"
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
    Friend WithEvents NomFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheminExport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RechercheFichier As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents FeuilleExcel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Deplace As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CreationAuto As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DossierExport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Cible1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BaseCpta1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TypeFormat1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameFormat As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheminRepexpor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RechercheFichier1 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents FeuilleExcel1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Deplace1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CreationAuto1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents NomRepxpor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CheminForma As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossier1 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
