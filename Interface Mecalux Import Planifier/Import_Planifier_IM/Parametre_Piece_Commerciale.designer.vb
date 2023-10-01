<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Parametre_Piece_Commerciale
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Parametre_Piece_Commerciale))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.basegescom = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.basecpta = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Piece = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Format = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.IDDossier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Fournisseur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dostraite = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.par3 = New System.Windows.Forms.DataGridViewButtonColumn
        Me.dosdest = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Par4 = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ServeurFtp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Ping = New System.Windows.Forms.DataGridViewButtonColumn
        Me.RepertoireFTP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UserFtp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PwdFtp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BT_ADD = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_DelRow = New System.Windows.Forms.Button
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.basegescom1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.basecpta1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Piece1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Format1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IDDossier1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Fournisseur1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dostraite1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dosdest1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.par2 = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ServeurFtp1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RepertoireFTP1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UserFtp1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PwdFtp1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Supprimer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BT_Delete = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.BT_Quit = New System.Windows.Forms.Button
        Me.BT_Save = New System.Windows.Forms.Button
        Me.FindFile = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepListeFile = New System.Windows.Forms.FolderBrowserDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Delete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Quit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Save)
        Me.SplitContainer1.Size = New System.Drawing.Size(770, 586)
        Me.SplitContainer1.SplitterDistance = 543
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
        Me.SplitContainer2.Size = New System.Drawing.Size(770, 543)
        Me.SplitContainer2.SplitterDistance = 314
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
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.basegescom, Me.basecpta, Me.Piece, Me.Format, Me.IDDossier, Me.Fournisseur, Me.dostraite, Me.par3, Me.dosdest, Me.Par4, Me.ServeurFtp, Me.Ping, Me.RepertoireFTP, Me.UserFtp, Me.PwdFtp})
        Me.DataListeSchema.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeSchema.Location = New System.Drawing.Point(0, 30)
        Me.DataListeSchema.MultiSelect = False
        Me.DataListeSchema.Name = "DataListeSchema"
        Me.DataListeSchema.RowHeadersVisible = False
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataListeSchema.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeSchema.RowTemplate.Height = 24
        Me.DataListeSchema.Size = New System.Drawing.Size(770, 284)
        Me.DataListeSchema.TabIndex = 44
        '
        'basegescom
        '
        Me.basegescom.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle1.Format = "N0"
        Me.basegescom.DefaultCellStyle = DataGridViewCellStyle1
        Me.basegescom.FillWeight = 67.61422!
        Me.basegescom.Frozen = True
        Me.basegescom.HeaderText = "Base Commerciale*"
        Me.basegescom.Name = "basegescom"
        Me.basegescom.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.basegescom.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.basegescom.Width = 160
        '
        'basecpta
        '
        Me.basecpta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.basecpta.FillWeight = 67.61422!
        Me.basecpta.Frozen = True
        Me.basecpta.HeaderText = "Base Comptable*"
        Me.basecpta.Name = "basecpta"
        Me.basecpta.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.basecpta.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.basecpta.Width = 150
        '
        'Piece
        '
        Me.Piece.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Piece.Frozen = True
        Me.Piece.HeaderText = "Type Piece*"
        Me.Piece.Items.AddRange(New Object() {"COMMANDE VENTE", "COMMANDE ACHAT", "B. R. VALORISE", "FACTURE VENTE", "FACTURE ACHAT", "B.L", "DEVIS"})
        Me.Piece.Name = "Piece"
        '
        'Format
        '
        Me.Format.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Format.Frozen = True
        Me.Format.HeaderText = "Format"
        Me.Format.Items.AddRange(New Object() {"EAN96", "EANCOM", "GESCOM100"})
        Me.Format.Name = "Format"
        Me.Format.Width = 70
        '
        'IDDossier
        '
        Me.IDDossier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier.HeaderText = "IDDossier*"
        Me.IDDossier.Name = "IDDossier"
        Me.IDDossier.Width = 60
        '
        'Fournisseur
        '
        Me.Fournisseur.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Fournisseur.HeaderText = "Tiers"
        Me.Fournisseur.Name = "Fournisseur"
        '
        'dostraite
        '
        Me.dostraite.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.dostraite.FillWeight = 67.61422!
        Me.dostraite.HeaderText = "Dossier Traitement*"
        Me.dostraite.Name = "dostraite"
        Me.dostraite.Width = 130
        '
        'par3
        '
        Me.par3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.par3.FillWeight = 67.61422!
        Me.par3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.par3.HeaderText = "Rep"
        Me.par3.Name = "par3"
        Me.par3.Width = 40
        '
        'dosdest
        '
        Me.dosdest.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.dosdest.FillWeight = 67.61422!
        Me.dosdest.HeaderText = "Dossier Destination"
        Me.dosdest.Name = "dosdest"
        Me.dosdest.Width = 130
        '
        'Par4
        '
        Me.Par4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Par4.FillWeight = 261.9289!
        Me.Par4.HeaderText = "Rep"
        Me.Par4.Name = "Par4"
        Me.Par4.Width = 40
        '
        'ServeurFtp
        '
        Me.ServeurFtp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ServeurFtp.HeaderText = "ServeurFtp"
        Me.ServeurFtp.Name = "ServeurFtp"
        '
        'Ping
        '
        Me.Ping.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Ping.HeaderText = "Ping du Serveur"
        Me.Ping.Name = "Ping"
        '
        'RepertoireFTP
        '
        Me.RepertoireFTP.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RepertoireFTP.HeaderText = "RepertoireFTP"
        Me.RepertoireFTP.Name = "RepertoireFTP"
        '
        'UserFtp
        '
        Me.UserFtp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UserFtp.HeaderText = "UserFtp"
        Me.UserFtp.Name = "UserFtp"
        '
        'PwdFtp
        '
        Me.PwdFtp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PwdFtp.HeaderText = "PwdFtp"
        Me.PwdFtp.Name = "PwdFtp"
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
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(710, 30)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
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
        Me.GroupBox1.Location = New System.Drawing.Point(739, 0)
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
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.basegescom1, Me.basecpta1, Me.Piece1, Me.Format1, Me.IDDossier1, Me.Fournisseur1, Me.dostraite1, Me.dosdest1, Me.par2, Me.ServeurFtp1, Me.RepertoireFTP1, Me.UserFtp1, Me.PwdFtp1, Me.Supprimer})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 15)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(770, 210)
        Me.DataListeIntegrer.TabIndex = 10
        '
        'basegescom1
        '
        Me.basegescom1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.NullValue = Nothing
        Me.basegescom1.DefaultCellStyle = DataGridViewCellStyle3
        Me.basegescom1.FillWeight = 229.7297!
        Me.basegescom1.HeaderText = "Base Commerciale"
        Me.basegescom1.Name = "basegescom1"
        Me.basegescom1.ReadOnly = True
        Me.basegescom1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.basegescom1.Width = 160
        '
        'basecpta1
        '
        Me.basecpta1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.basecpta1.DefaultCellStyle = DataGridViewCellStyle4
        Me.basecpta1.HeaderText = "Base Comptable"
        Me.basecpta1.MaxInputLength = 6
        Me.basecpta1.Name = "basecpta1"
        Me.basecpta1.ReadOnly = True
        Me.basecpta1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.basecpta1.Width = 150
        '
        'Piece1
        '
        Me.Piece1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Piece1.HeaderText = "Type Piece"
        Me.Piece1.Name = "Piece1"
        Me.Piece1.ReadOnly = True
        Me.Piece1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Piece1.Width = 110
        '
        'Format1
        '
        Me.Format1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Format1.HeaderText = "Format"
        Me.Format1.Name = "Format1"
        Me.Format1.ReadOnly = True
        Me.Format1.Width = 70
        '
        'IDDossier1
        '
        Me.IDDossier1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.IDDossier1.HeaderText = "IDDossier*"
        Me.IDDossier1.Name = "IDDossier1"
        Me.IDDossier1.Width = 60
        '
        'Fournisseur1
        '
        Me.Fournisseur1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Fournisseur1.HeaderText = "Tiers"
        Me.Fournisseur1.Name = "Fournisseur1"
        '
        'dostraite1
        '
        Me.dostraite1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.dostraite1.FillWeight = 67.56757!
        Me.dostraite1.HeaderText = "Dossier Traitement"
        Me.dostraite1.Name = "dostraite1"
        Me.dostraite1.ReadOnly = True
        Me.dostraite1.Width = 233
        '
        'dosdest1
        '
        Me.dosdest1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.dosdest1.FillWeight = 67.56757!
        Me.dosdest1.HeaderText = "Dossier Destination"
        Me.dosdest1.Name = "dosdest1"
        Me.dosdest1.Width = 160
        '
        'par2
        '
        Me.par2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.par2.FillWeight = 67.56757!
        Me.par2.HeaderText = "Rep"
        Me.par2.Name = "par2"
        Me.par2.Width = 40
        '
        'ServeurFtp1
        '
        Me.ServeurFtp1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ServeurFtp1.HeaderText = "ServeurFtp"
        Me.ServeurFtp1.Name = "ServeurFtp1"
        Me.ServeurFtp1.ReadOnly = True
        '
        'RepertoireFTP1
        '
        Me.RepertoireFTP1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RepertoireFTP1.HeaderText = "RepertoireFTP"
        Me.RepertoireFTP1.Name = "RepertoireFTP1"
        Me.RepertoireFTP1.ReadOnly = True
        '
        'UserFtp1
        '
        Me.UserFtp1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.UserFtp1.HeaderText = "UserFtp"
        Me.UserFtp1.Name = "UserFtp1"
        Me.UserFtp1.ReadOnly = True
        '
        'PwdFtp1
        '
        Me.PwdFtp1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PwdFtp1.HeaderText = "PwdFtp"
        Me.PwdFtp1.Name = "PwdFtp1"
        Me.PwdFtp1.ReadOnly = True
        '
        'Supprimer
        '
        Me.Supprimer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Supprimer.HeaderText = "Supprimer"
        Me.Supprimer.Name = "Supprimer"
        Me.Supprimer.Width = 50
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(770, 15)
        Me.Panel1.TabIndex = 9
        '
        'BT_Delete
        '
        Me.BT_Delete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BT_Delete.Image = Global.Import_Planifier_IM.My.Resources.Resources.criticalind_status
        Me.BT_Delete.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Delete.Location = New System.Drawing.Point(386, 13)
        Me.BT_Delete.Name = "BT_Delete"
        Me.BT_Delete.Size = New System.Drawing.Size(72, 23)
        Me.BT_Delete.TabIndex = 1
        Me.BT_Delete.Text = "Supprimer"
        Me.BT_Delete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Delete.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.Location = New System.Drawing.Point(288, 13)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(83, 23)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "&Modifier"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox4.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(770, 9)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        '
        'BT_Quit
        '
        Me.BT_Quit.Image = Global.Import_Planifier_IM.My.Resources.Resources.image034
        Me.BT_Quit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Quit.Location = New System.Drawing.Point(488, 13)
        Me.BT_Quit.Name = "BT_Quit"
        Me.BT_Quit.Size = New System.Drawing.Size(79, 23)
        Me.BT_Quit.TabIndex = 2
        Me.BT_Quit.Text = "&Quitter"
        Me.BT_Quit.UseVisualStyleBackColor = True
        '
        'BT_Save
        '
        Me.BT_Save.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_16
        Me.BT_Save.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Save.Location = New System.Drawing.Point(587, 13)
        Me.BT_Save.Name = "BT_Save"
        Me.BT_Save.Size = New System.Drawing.Size(79, 23)
        Me.BT_Save.TabIndex = 1
        Me.BT_Save.Text = "&Enregistrer"
        Me.BT_Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Save.UseVisualStyleBackColor = True
        '
        'Parametre_Piece_Commerciale
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 586)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Parametre_Piece_Commerciale"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parametrage de Creation des Pieces Commerciales"
        Me.TopMost = True
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
    Friend WithEvents FindFile As System.Windows.Forms.OpenFileDialog
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
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents basegescom1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents basecpta1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Piece1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Format1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IDDossier1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fournisseur1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dostraite1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dosdest1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents par2 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ServeurFtp1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RepertoireFTP1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UserFtp1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PwdFtp1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Supprimer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents basegescom As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents basecpta As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Piece As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Format As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents IDDossier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fournisseur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dostraite As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents par3 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents dosdest As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Par4 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ServeurFtp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Ping As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents RepertoireFTP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UserFtp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PwdFtp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
