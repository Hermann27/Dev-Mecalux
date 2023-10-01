<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormatintegrationStock
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormatintegrationStock))
        Me.SaveFileXml = New System.Windows.Forms.SaveFileDialog
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DataDispo = New System.Windows.Forms.DataGridView
        Me.ColDispos = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibelleDispos = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DisPositio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DisLongueur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Fiche = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_New = New System.Windows.Forms.Button
        Me.BT_DelForm = New System.Windows.Forms.Button
        Me.BT_SaveFormat = New System.Windows.Forms.Button
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.BT_Down = New System.Windows.Forms.Button
        Me.BT_DelDispo = New System.Windows.Forms.Button
        Me.BT_UP = New System.Windows.Forms.Button
        Me.BT_DelSelt = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DataSelect = New System.Windows.Forms.DataGridView
        Me.Selection = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Position = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Infos = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Fichier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Piece = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Article = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Defauts = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ValeurDefauts = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Libelles = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.NumUpDown = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.Txt_Chemin = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Cmb_Format = New System.Windows.Forms.ComboBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.CkPunitaire = New System.Windows.Forms.CheckBox
        Me.CbFichier = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Ckauto = New System.Windows.Forms.CheckBox
        Me.Cb_Date = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.CbMod = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Txtype = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DataDispo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.NumUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(891, 355)
        Me.Panel1.TabIndex = 47
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(891, 355)
        Me.SplitContainer1.SplitterDistance = 170
        Me.SplitContainer1.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DataDispo)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(170, 355)
        Me.GroupBox3.TabIndex = 47
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Colonnes Disponibles"
        '
        'DataDispo
        '
        Me.DataDispo.AllowUserToAddRows = False
        Me.DataDispo.AllowUserToDeleteRows = False
        Me.DataDispo.AllowUserToResizeRows = False
        Me.DataDispo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataDispo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDispos, Me.LibelleDispos, Me.DisPositio, Me.DisLongueur, Me.Fiche})
        Me.DataDispo.Cursor = System.Windows.Forms.Cursors.Hand
        Me.DataDispo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataDispo.Location = New System.Drawing.Point(3, 16)
        Me.DataDispo.MultiSelect = False
        Me.DataDispo.Name = "DataDispo"
        Me.DataDispo.RowHeadersVisible = False
        Me.DataDispo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataDispo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DataDispo.Size = New System.Drawing.Size(164, 336)
        Me.DataDispo.TabIndex = 10
        '
        'ColDispos
        '
        Me.ColDispos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ColDispos.HeaderText = "Colonnes disponibles"
        Me.ColDispos.Name = "ColDispos"
        Me.ColDispos.ReadOnly = True
        '
        'LibelleDispos
        '
        Me.LibelleDispos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.LibelleDispos.HeaderText = "LibelleDispo"
        Me.LibelleDispos.Name = "LibelleDispos"
        Me.LibelleDispos.ReadOnly = True
        Me.LibelleDispos.Visible = False
        '
        'DisPositio
        '
        DataGridViewCellStyle1.NullValue = "0"
        Me.DisPositio.DefaultCellStyle = DataGridViewCellStyle1
        Me.DisPositio.HeaderText = "Position"
        Me.DisPositio.Name = "DisPositio"
        Me.DisPositio.Visible = False
        '
        'DisLongueur
        '
        DataGridViewCellStyle2.NullValue = "0"
        Me.DisLongueur.DefaultCellStyle = DataGridViewCellStyle2
        Me.DisLongueur.HeaderText = "Longueur"
        Me.DisLongueur.Name = "DisLongueur"
        Me.DisLongueur.Visible = False
        '
        'Fiche
        '
        Me.Fiche.HeaderText = "Fiche"
        Me.Fiche.Name = "Fiche"
        Me.Fiche.Visible = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Size = New System.Drawing.Size(717, 355)
        Me.SplitContainer2.SplitterDistance = 89
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BT_New)
        Me.GroupBox1.Controls.Add(Me.BT_DelForm)
        Me.GroupBox1.Controls.Add(Me.BT_SaveFormat)
        Me.GroupBox1.Controls.Add(Me.GroupBox8)
        Me.GroupBox1.Controls.Add(Me.BT_Down)
        Me.GroupBox1.Controls.Add(Me.BT_DelDispo)
        Me.GroupBox1.Controls.Add(Me.BT_UP)
        Me.GroupBox1.Controls.Add(Me.BT_DelSelt)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(89, 355)
        Me.GroupBox1.TabIndex = 45
        Me.GroupBox1.TabStop = False
        '
        'BT_New
        '
        Me.BT_New.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_New.Image = Global.Import_Planifier_IM.My.Resources.Resources.image019
        Me.BT_New.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_New.Location = New System.Drawing.Point(1, 206)
        Me.BT_New.Name = "BT_New"
        Me.BT_New.Size = New System.Drawing.Size(78, 24)
        Me.BT_New.TabIndex = 48
        Me.BT_New.Text = "&Nouveau"
        Me.BT_New.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_New.UseVisualStyleBackColor = True
        '
        'BT_DelForm
        '
        Me.BT_DelForm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_DelForm.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.BT_DelForm.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_DelForm.Location = New System.Drawing.Point(0, 285)
        Me.BT_DelForm.Name = "BT_DelForm"
        Me.BT_DelForm.Size = New System.Drawing.Size(81, 23)
        Me.BT_DelForm.TabIndex = 2
        Me.BT_DelForm.Text = "&Supprimer"
        Me.BT_DelForm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_DelForm.UseVisualStyleBackColor = True
        '
        'BT_SaveFormat
        '
        Me.BT_SaveFormat.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_SaveFormat.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_16
        Me.BT_SaveFormat.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_SaveFormat.Location = New System.Drawing.Point(-1, 246)
        Me.BT_SaveFormat.Name = "BT_SaveFormat"
        Me.BT_SaveFormat.Size = New System.Drawing.Size(81, 23)
        Me.BT_SaveFormat.TabIndex = 0
        Me.BT_SaveFormat.Text = "&Ok"
        Me.BT_SaveFormat.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox8.Location = New System.Drawing.Point(3, 314)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(83, 38)
        Me.GroupBox8.TabIndex = 47
        Me.GroupBox8.TabStop = False
        '
        'BT_Down
        '
        Me.BT_Down.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowdown_16
        Me.BT_Down.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_Down.Location = New System.Drawing.Point(-1, 165)
        Me.BT_Down.Name = "BT_Down"
        Me.BT_Down.Size = New System.Drawing.Size(80, 24)
        Me.BT_Down.TabIndex = 45
        Me.BT_Down.Text = "&Descendre"
        Me.BT_Down.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_Down.UseVisualStyleBackColor = True
        '
        'BT_DelDispo
        '
        Me.BT_DelDispo.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowforward_16
        Me.BT_DelDispo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_DelDispo.Location = New System.Drawing.Point(-1, 50)
        Me.BT_DelDispo.Name = "BT_DelDispo"
        Me.BT_DelDispo.Size = New System.Drawing.Size(80, 23)
        Me.BT_DelDispo.TabIndex = 42
        Me.BT_DelDispo.Text = "&Ajouter"
        Me.BT_DelDispo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_DelDispo.UseVisualStyleBackColor = True
        '
        'BT_UP
        '
        Me.BT_UP.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowup_16
        Me.BT_UP.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_UP.Location = New System.Drawing.Point(0, 123)
        Me.BT_UP.Name = "BT_UP"
        Me.BT_UP.Size = New System.Drawing.Size(79, 23)
        Me.BT_UP.TabIndex = 44
        Me.BT_UP.Text = "&Monter"
        Me.BT_UP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_UP.UseVisualStyleBackColor = True
        '
        'BT_DelSelt
        '
        Me.BT_DelSelt.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowback_16
        Me.BT_DelSelt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_DelSelt.Location = New System.Drawing.Point(-1, 86)
        Me.BT_DelSelt.Name = "BT_DelSelt"
        Me.BT_DelSelt.Size = New System.Drawing.Size(80, 22)
        Me.BT_DelSelt.TabIndex = 43
        Me.BT_DelSelt.Text = "&Supprimer"
        Me.BT_DelSelt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BT_DelSelt.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataSelect)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 80)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(624, 275)
        Me.GroupBox2.TabIndex = 48
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Personnalisation de la Liste"
        '
        'DataSelect
        '
        Me.DataSelect.AllowUserToAddRows = False
        Me.DataSelect.AllowUserToDeleteRows = False
        Me.DataSelect.AllowUserToResizeRows = False
        Me.DataSelect.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataSelect.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Selection, Me.Position, Me.Infos, Me.Fichier, Me.Piece, Me.Article, Me.Defauts, Me.ValeurDefauts, Me.Libelles})
        Me.DataSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataSelect.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataSelect.EnableHeadersVisualStyles = False
        Me.DataSelect.Location = New System.Drawing.Point(3, 16)
        Me.DataSelect.MultiSelect = False
        Me.DataSelect.Name = "DataSelect"
        Me.DataSelect.RowHeadersVisible = False
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataSelect.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.DataSelect.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataSelect.Size = New System.Drawing.Size(618, 256)
        Me.DataSelect.TabIndex = 6
        '
        'Selection
        '
        Me.Selection.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.Format = "N0"
        DataGridViewCellStyle3.NullValue = Nothing
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.Selection.DefaultCellStyle = DataGridViewCellStyle3
        Me.Selection.HeaderText = "Colonnes Selectionnées"
        Me.Selection.Name = "Selection"
        Me.Selection.ReadOnly = True
        Me.Selection.Width = 168
        '
        'Position
        '
        Me.Position.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle4.Format = "N0"
        DataGridViewCellStyle4.NullValue = "0"
        Me.Position.DefaultCellStyle = DataGridViewCellStyle4
        Me.Position.FillWeight = 40.0!
        Me.Position.HeaderText = "Position"
        Me.Position.Name = "Position"
        Me.Position.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Position.Width = 60
        '
        'Infos
        '
        Me.Infos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Infos.HeaderText = "Info Libre"
        Me.Infos.Name = "Infos"
        Me.Infos.ReadOnly = True
        Me.Infos.Width = 60
        '
        'Fichier
        '
        Me.Fichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Fichier.HeaderText = "Entête/Ligne"
        Me.Fichier.Name = "Fichier"
        Me.Fichier.ReadOnly = True
        Me.Fichier.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Fichier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Fichier.Width = 90
        '
        'Piece
        '
        Me.Piece.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.Blue
        DataGridViewCellStyle5.NullValue = False
        Me.Piece.DefaultCellStyle = DataGridViewCellStyle5
        Me.Piece.HeaderText = "Pièce"
        Me.Piece.Name = "Piece"
        Me.Piece.Width = 38
        '
        'Article
        '
        Me.Article.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Article.HeaderText = "Article"
        Me.Article.Name = "Article"
        Me.Article.Width = 40
        '
        'Defauts
        '
        Me.Defauts.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Defauts.HeaderText = "Defaut"
        Me.Defauts.Name = "Defauts"
        Me.Defauts.Width = 50
        '
        'ValeurDefauts
        '
        Me.ValeurDefauts.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ValeurDefauts.HeaderText = "Valeur Defaut"
        Me.ValeurDefauts.Name = "ValeurDefauts"
        Me.ValeurDefauts.ReadOnly = True
        '
        'Libelles
        '
        Me.Libelles.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Libelles.HeaderText = "Libelle"
        Me.Libelles.Name = "Libelles"
        Me.Libelles.ReadOnly = True
        Me.Libelles.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.NumUpDown)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Txt_Chemin)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Cmb_Format)
        Me.Panel2.Controls.Add(Me.GroupBox4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(624, 80)
        Me.Panel2.TabIndex = 0
        '
        'NumUpDown
        '
        Me.NumUpDown.Location = New System.Drawing.Point(295, 33)
        Me.NumUpDown.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NumUpDown.Name = "NumUpDown"
        Me.NumUpDown.Size = New System.Drawing.Size(57, 20)
        Me.NumUpDown.TabIndex = 18
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(203, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Decalage entête"
        '
        'Txt_Chemin
        '
        Me.Txt_Chemin.BackColor = System.Drawing.SystemColors.Window
        Me.Txt_Chemin.Location = New System.Drawing.Point(195, 7)
        Me.Txt_Chemin.Name = "Txt_Chemin"
        Me.Txt_Chemin.ReadOnly = True
        Me.Txt_Chemin.Size = New System.Drawing.Size(429, 20)
        Me.Txt_Chemin.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(1, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Format"
        '
        'Cmb_Format
        '
        Me.Cmb_Format.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Cmb_Format.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Cmb_Format.FormattingEnabled = True
        Me.Cmb_Format.Location = New System.Drawing.Point(46, 6)
        Me.Cmb_Format.Name = "Cmb_Format"
        Me.Cmb_Format.Size = New System.Drawing.Size(147, 21)
        Me.Cmb_Format.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CkPunitaire)
        Me.GroupBox4.Controls.Add(Me.CbFichier)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.Ckauto)
        Me.GroupBox4.Controls.Add(Me.Cb_Date)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.CbMod)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Txtype)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(624, 80)
        Me.GroupBox4.TabIndex = 20
        Me.GroupBox4.TabStop = False
        '
        'CkPunitaire
        '
        Me.CkPunitaire.AutoSize = True
        Me.CkPunitaire.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CkPunitaire.Location = New System.Drawing.Point(414, 60)
        Me.CkPunitaire.Name = "CkPunitaire"
        Me.CkPunitaire.Size = New System.Drawing.Size(112, 17)
        Me.CkPunitaire.TabIndex = 21
        Me.CkPunitaire.Text = "P.U Par Défaut"
        Me.CkPunitaire.UseVisualStyleBackColor = True
        '
        'CbFichier
        '
        Me.CbFichier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CbFichier.FormattingEnabled = True
        Me.CbFichier.Items.AddRange(New Object() {"Document Entête", "Document Ligne"})
        Me.CbFichier.Location = New System.Drawing.Point(46, 32)
        Me.CbFichier.Name = "CbFichier"
        Me.CbFichier.Size = New System.Drawing.Size(147, 21)
        Me.CbFichier.TabIndex = 20
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(1, 37)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Fichier"
        '
        'Ckauto
        '
        Me.Ckauto.AutoSize = True
        Me.Ckauto.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ckauto.Location = New System.Drawing.Point(295, 58)
        Me.Ckauto.Name = "Ckauto"
        Me.Ckauto.Size = New System.Drawing.Size(88, 17)
        Me.Ckauto.TabIndex = 15
        Me.Ckauto.Text = "Pièce Auto"
        Me.Ckauto.UseVisualStyleBackColor = True
        '
        'Cb_Date
        '
        Me.Cb_Date.FormattingEnabled = True
        Me.Cb_Date.Items.AddRange(New Object() {"jj/mm/aa", "jj/mm/aaaa", "aammjj", "jjmmaa", "aaaammjj", "jjmmaaaa", "aa-mm-jj", "jj-mm-aa", "aaaa-mm-jj", "jj-mm-aaaa", "mmjjaa", "mmjjaaaa", "aajjmm", "aaaajjmm", "mm/jj/aa", "mm/jj/aaaa", "mm-jj-aa", "mm-jj-aaaa", "aa-jj-mm", "aaaa-jj-mm"})
        Me.Cb_Date.Location = New System.Drawing.Point(427, 33)
        Me.Cb_Date.Name = "Cb_Date"
        Me.Cb_Date.Size = New System.Drawing.Size(85, 21)
        Me.Cb_Date.TabIndex = 14
        Me.Cb_Date.Text = "jj/mm/aaaa"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(358, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(65, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Format Date"
        '
        'CbMod
        '
        Me.CbMod.FormattingEnabled = True
        Me.CbMod.Items.AddRange(New Object() {"Création", "Modification"})
        Me.CbMod.Location = New System.Drawing.Point(46, 55)
        Me.CbMod.Name = "CbMod"
        Me.CbMod.Size = New System.Drawing.Size(81, 21)
        Me.CbMod.TabIndex = 6
        Me.CbMod.Text = "Création"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Mode"
        '
        'Txtype
        '
        Me.Txtype.FormattingEnabled = True
        Me.Txtype.Items.AddRange(New Object() {"Point virgule", "Tabulation", "Excel", "Pipe"})
        Me.Txtype.Location = New System.Drawing.Point(195, 55)
        Me.Txtype.Name = "Txtype"
        Me.Txtype.Size = New System.Drawing.Size(94, 21)
        Me.Txtype.TabIndex = 4
        Me.Txtype.Text = "Point virgule"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(154, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Type"
        '
        'FormatintegrationStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(891, 355)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormatintegrationStock"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Formats d'integration des documents de Stock"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DataDispo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataSelect, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.NumUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SaveFileXml As System.Windows.Forms.SaveFileDialog
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataDispo As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BT_DelForm As System.Windows.Forms.Button
    Friend WithEvents BT_SaveFormat As System.Windows.Forms.Button
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents BT_Down As System.Windows.Forms.Button
    Friend WithEvents BT_DelDispo As System.Windows.Forms.Button
    Friend WithEvents BT_UP As System.Windows.Forms.Button
    Friend WithEvents BT_DelSelt As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DataSelect As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Txt_Chemin As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Cmb_Format As System.Windows.Forms.ComboBox
    Friend WithEvents BT_New As System.Windows.Forms.Button
    Friend WithEvents NumUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Txtype As System.Windows.Forms.ComboBox
    Friend WithEvents CbMod As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Cb_Date As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Ckauto As System.Windows.Forms.CheckBox
    Friend WithEvents CbFichier As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ColDispos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibelleDispos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DisPositio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DisLongueur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fiche As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Selection As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Position As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Infos As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Fichier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Piece As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Article As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Defauts As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ValeurDefauts As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Libelles As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CkPunitaire As System.Windows.Forms.CheckBox
End Class
