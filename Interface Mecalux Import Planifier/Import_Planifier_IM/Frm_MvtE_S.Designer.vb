<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_MvtE_S
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_MvtE_S))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Datagridaffiche = New System.Windows.Forms.DataGridView
        Me.C1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.C2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.C3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.C4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.C7 = New System.Windows.Forms.DataGridViewImageColumn
        Me.C5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.C6 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.C8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PicLigne = New System.Windows.Forms.PictureBox
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.WAREHOUSE_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PRODUCT_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LOT_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SERIAL_NO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EXP_DATE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.OWNER_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CONTAINER_NO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SIGN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.REASON = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DESCRIPTION = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SORDER_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LINE_NUMBER = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.REC_ORDER_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QUANTITY = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QUALITY = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BEST_BEFORE_DATE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PRODUCTION_DATE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SHELF_LIFE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DATE_DAYS_OF_LIFE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SIZE_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COLOR_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SOURCE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VERSION_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PRODUCTION_METHOD = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.POST_PRODUCTION_TREATMENT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LOT_COUNT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button3 = New System.Windows.Forms.Button
        Me.BT_integrer = New System.Windows.Forms.Button
        Me.btnview = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblType = New System.Windows.Forms.Label
        Me.ChkPieceAuto = New System.Windows.Forms.CheckBox
        Me.CheckFille = New System.Windows.Forms.CheckBox
        Me.BtnListe = New System.Windows.Forms.Button
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.ComboDate = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer7 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer8 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.infosExport = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ChEncapsuler = New System.Windows.Forms.CheckBox
        CType(Me.Datagridaffiche, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicLigne, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer7.Panel1.SuspendLayout()
        Me.SplitContainer7.Panel2.SuspendLayout()
        Me.SplitContainer7.SuspendLayout()
        Me.SplitContainer8.Panel2.SuspendLayout()
        Me.SplitContainer8.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1184, 10)
        Me.ProgressBar1.TabIndex = 3
        '
        'Datagridaffiche
        '
        Me.Datagridaffiche.AllowUserToAddRows = False
        Me.Datagridaffiche.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Datagridaffiche.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Datagridaffiche.BackgroundColor = System.Drawing.Color.Snow
        Me.Datagridaffiche.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.C1, Me.C2, Me.C3, Me.C4, Me.C7, Me.C5, Me.C6, Me.C8})
        Me.Datagridaffiche.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Datagridaffiche.EnableHeadersVisualStyles = False
        Me.Datagridaffiche.Location = New System.Drawing.Point(0, 0)
        Me.Datagridaffiche.Name = "Datagridaffiche"
        Me.Datagridaffiche.RowHeadersVisible = False
        Me.Datagridaffiche.Size = New System.Drawing.Size(1184, 146)
        Me.Datagridaffiche.TabIndex = 1
        '
        'C1
        '
        Me.C1.HeaderText = "Information Echangées"
        Me.C1.Name = "C1"
        '
        'C2
        '
        Me.C2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.C2.HeaderText = "Type d'Echange"
        Me.C2.Name = "C2"
        Me.C2.Width = 110
        '
        'C3
        '
        Me.C3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.C3.HeaderText = "Code du Type"
        Me.C3.Name = "C3"
        Me.C3.Width = 99
        '
        'C4
        '
        Me.C4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.C4.HeaderText = "Version"
        Me.C4.Name = "C4"
        Me.C4.Width = 67
        '
        'C7
        '
        Me.C7.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.C7.HeaderText = "Statut"
        Me.C7.Name = "C7"
        Me.C7.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.C7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.C7.Width = 60
        '
        'C5
        '
        Me.C5.HeaderText = "Fichier"
        Me.C5.Name = "C5"
        Me.C5.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.C5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'C6
        '
        Me.C6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.C6.HeaderText = "Choix"
        Me.C6.Name = "C6"
        Me.C6.Width = 39
        '
        'C8
        '
        Me.C8.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.C8.HeaderText = "Chemin"
        Me.C8.Name = "C8"
        Me.C8.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.C8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.C8.Visible = False
        '
        'PicLigne
        '
        Me.PicLigne.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PicLigne.Image = Global.Import_Planifier_IM.My.Resources.Resources.fleche
        Me.PicLigne.Location = New System.Drawing.Point(3, 3)
        Me.PicLigne.Name = "PicLigne"
        Me.PicLigne.Size = New System.Drawing.Size(24, 16)
        Me.PicLigne.TabIndex = 5
        Me.PicLigne.TabStop = False
        Me.PicLigne.Tag = "Ajouter Champs"
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataListeIntegrer.BackgroundColor = System.Drawing.Color.LightBlue
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.WAREHOUSE_CODE, Me.PRODUCT_CODE, Me.LOT_CODE, Me.SERIAL_NO, Me.EXP_DATE, Me.OWNER_CODE, Me.CONTAINER_NO, Me.SIGN, Me.REASON, Me.DESCRIPTION, Me.SORDER_CODE, Me.LINE_NUMBER, Me.REC_ORDER_CODE, Me.QUANTITY, Me.QUALITY, Me.BEST_BEFORE_DATE, Me.PRODUCTION_DATE, Me.SHELF_LIFE, Me.DATE_DAYS_OF_LIFE, Me.SIZE_CODE, Me.COLOR_CODE, Me.SOURCE, Me.VERSION_CODE, Me.PRODUCTION_METHOD, Me.POST_PRODUCTION_TREATMENT, Me.LOT_COUNT})
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.EnableHeadersVisualStyles = False
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        Me.DataListeIntegrer.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.DataListeIntegrer.Size = New System.Drawing.Size(1184, 289)
        Me.DataListeIntegrer.TabIndex = 2
        '
        'WAREHOUSE_CODE
        '
        Me.WAREHOUSE_CODE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.WAREHOUSE_CODE.HeaderText = "           Code du magasin auquel il appartient"
        Me.WAREHOUSE_CODE.Name = "WAREHOUSE_CODE"
        Me.WAREHOUSE_CODE.Width = 239
        '
        'PRODUCT_CODE
        '
        Me.PRODUCT_CODE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.PRODUCT_CODE.HeaderText = "Code de l'article"
        Me.PRODUCT_CODE.Name = "PRODUCT_CODE"
        Me.PRODUCT_CODE.Width = 107
        '
        'LOT_CODE
        '
        Me.LOT_CODE.HeaderText = "Code du lot"
        Me.LOT_CODE.Name = "LOT_CODE"
        '
        'SERIAL_NO
        '
        Me.SERIAL_NO.HeaderText = "Numéro de série"
        Me.SERIAL_NO.Name = "SERIAL_NO"
        '
        'EXP_DATE
        '
        Me.EXP_DATE.HeaderText = "Date limite d'utilisation"
        Me.EXP_DATE.Name = "EXP_DATE"
        Me.EXP_DATE.Visible = False
        '
        'OWNER_CODE
        '
        Me.OWNER_CODE.HeaderText = "Code du propriétaire"
        Me.OWNER_CODE.Name = "OWNER_CODE"
        Me.OWNER_CODE.Visible = False
        '
        'CONTAINER_NO
        '
        Me.CONTAINER_NO.HeaderText = "Code du conteneur"
        Me.CONTAINER_NO.Name = "CONTAINER_NO"
        '
        'SIGN
        '
        Me.SIGN.HeaderText = "Signe de la variation de stock"
        Me.SIGN.Name = "SIGN"
        Me.SIGN.Visible = False
        '
        'REASON
        '
        Me.REASON.HeaderText = "Cause de la variation"
        Me.REASON.Name = "REASON"
        Me.REASON.Visible = False
        '
        'DESCRIPTION
        '
        Me.DESCRIPTION.HeaderText = "Description"
        Me.DESCRIPTION.Name = "DESCRIPTION"
        Me.DESCRIPTION.Visible = False
        '
        'SORDER_CODE
        '
        Me.SORDER_CODE.HeaderText = "variation est due à une commande"
        Me.SORDER_CODE.Name = "SORDER_CODE"
        Me.SORDER_CODE.Visible = False
        '
        'LINE_NUMBER
        '
        Me.LINE_NUMBER.HeaderText = "Numéro de ligne de Cmd"
        Me.LINE_NUMBER.Name = "LINE_NUMBER"
        '
        'REC_ORDER_CODE
        '
        Me.REC_ORDER_CODE.HeaderText = "Code du préavis"
        Me.REC_ORDER_CODE.Name = "REC_ORDER_CODE"
        '
        'QUANTITY
        '
        Me.QUANTITY.HeaderText = "Quantité du stock"
        Me.QUANTITY.Name = "QUANTITY"
        Me.QUANTITY.Visible = False
        '
        'QUALITY
        '
        Me.QUALITY.HeaderText = "Qualité"
        Me.QUALITY.Name = "QUALITY"
        Me.QUALITY.Visible = False
        '
        'BEST_BEFORE_DATE
        '
        Me.BEST_BEFORE_DATE.HeaderText = "Date de consommation recommandée"
        Me.BEST_BEFORE_DATE.Name = "BEST_BEFORE_DATE"
        Me.BEST_BEFORE_DATE.Visible = False
        '
        'PRODUCTION_DATE
        '
        Me.PRODUCTION_DATE.HeaderText = "Date de fabrication"
        Me.PRODUCTION_DATE.Name = "PRODUCTION_DATE"
        Me.PRODUCTION_DATE.Visible = False
        '
        'SHELF_LIFE
        '
        Me.SHELF_LIFE.HeaderText = "Durée de vie du stock "
        Me.SHELF_LIFE.Name = "SHELF_LIFE"
        Me.SHELF_LIFE.Visible = False
        '
        'DATE_DAYS_OF_LIFE
        '
        Me.DATE_DAYS_OF_LIFE.HeaderText = "Durée de vie restante du stock"
        Me.DATE_DAYS_OF_LIFE.Name = "DATE_DAYS_OF_LIFE"
        Me.DATE_DAYS_OF_LIFE.Visible = False
        '
        'SIZE_CODE
        '
        Me.SIZE_CODE.HeaderText = "Calibre"
        Me.SIZE_CODE.Name = "SIZE_CODE"
        Me.SIZE_CODE.Visible = False
        '
        'COLOR_CODE
        '
        Me.COLOR_CODE.HeaderText = "Couleur"
        Me.COLOR_CODE.Name = "COLOR_CODE"
        Me.COLOR_CODE.Visible = False
        '
        'SOURCE
        '
        Me.SOURCE.HeaderText = "Source"
        Me.SOURCE.Name = "SOURCE"
        Me.SOURCE.Visible = False
        '
        'VERSION_CODE
        '
        Me.VERSION_CODE.HeaderText = "Version"
        Me.VERSION_CODE.Name = "VERSION_CODE"
        Me.VERSION_CODE.Visible = False
        '
        'PRODUCTION_METHOD
        '
        Me.PRODUCTION_METHOD.HeaderText = "Méthode de production"
        Me.PRODUCTION_METHOD.Name = "PRODUCTION_METHOD"
        Me.PRODUCTION_METHOD.Visible = False
        '
        'POST_PRODUCTION_TREATMENT
        '
        Me.POST_PRODUCTION_TREATMENT.HeaderText = "Traitement de post-production"
        Me.POST_PRODUCTION_TREATMENT.Name = "POST_PRODUCTION_TREATMENT"
        Me.POST_PRODUCTION_TREATMENT.Visible = False
        '
        'LOT_COUNT
        '
        Me.LOT_COUNT.HeaderText = "Nombre de lots associés"
        Me.LOT_COUNT.Name = "LOT_COUNT"
        Me.LOT_COUNT.Visible = False
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(336, 11)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(109, 32)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Quitter"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BT_integrer
        '
        Me.BT_integrer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_integrer.Image = Global.Import_Planifier_IM.My.Resources.Resources.AnalyzeWizard
        Me.BT_integrer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BT_integrer.Location = New System.Drawing.Point(469, 11)
        Me.BT_integrer.Name = "BT_integrer"
        Me.BT_integrer.Size = New System.Drawing.Size(109, 32)
        Me.BT_integrer.TabIndex = 1
        Me.BT_integrer.Text = "Executer"
        Me.BT_integrer.UseVisualStyleBackColor = True
        Me.BT_integrer.Visible = False
        '
        'btnview
        '
        Me.btnview.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnview.Image = Global.Import_Planifier_IM.My.Resources.Resources.forms
        Me.btnview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnview.Location = New System.Drawing.Point(51, 11)
        Me.btnview.Name = "btnview"
        Me.btnview.Size = New System.Drawing.Size(109, 32)
        Me.btnview.TabIndex = 0
        Me.btnview.Text = "Aperçu"
        Me.btnview.UseVisualStyleBackColor = True
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.CheckFille)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnListe)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChEncapsuler)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ComboDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ProgressBar1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1184, 552)
        Me.SplitContainer1.SplitterDistance = 53
        Me.SplitContainer1.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(844, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "Type de document :"
        Me.Label1.Visible = False
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.ForeColor = System.Drawing.Color.Red
        Me.lblType.Location = New System.Drawing.Point(944, 101)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(23, 15)
        Me.lblType.TabIndex = 33
        Me.lblType.Text = "...."
        Me.lblType.Visible = False
        '
        'ChkPieceAuto
        '
        Me.ChkPieceAuto.AutoSize = True
        Me.ChkPieceAuto.Checked = True
        Me.ChkPieceAuto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkPieceAuto.Location = New System.Drawing.Point(722, 101)
        Me.ChkPieceAuto.Name = "ChkPieceAuto"
        Me.ChkPieceAuto.Size = New System.Drawing.Size(123, 17)
        Me.ChkPieceAuto.TabIndex = 32
        Me.ChkPieceAuto.Text = "Piéce automatique ?"
        Me.ChkPieceAuto.UseVisualStyleBackColor = True
        Me.ChkPieceAuto.Visible = False
        '
        'CheckFille
        '
        Me.CheckFille.AutoSize = True
        Me.CheckFille.Location = New System.Drawing.Point(838, 19)
        Me.CheckFille.Name = "CheckFille"
        Me.CheckFille.Size = New System.Drawing.Size(254, 17)
        Me.CheckFille.TabIndex = 31
        Me.CheckFille.Text = "Déplacer le Fichier à la fin de chaque traitement "
        Me.CheckFille.UseVisualStyleBackColor = True
        '
        'BtnListe
        '
        Me.BtnListe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnListe.Image = Global.Import_Planifier_IM.My.Resources.Resources.export1
        Me.BtnListe.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnListe.Location = New System.Drawing.Point(6, 15)
        Me.BtnListe.Name = "BtnListe"
        Me.BtnListe.Size = New System.Drawing.Size(124, 23)
        Me.BtnListe.TabIndex = 30
        Me.BtnListe.Text = "   Affiche To Liste"
        Me.BtnListe.UseVisualStyleBackColor = True
        '
        'PictureBox3
        '
        Me.PictureBox3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox3.Image = Global.Import_Planifier_IM.My.Resources.Resources.btFermer22
        Me.PictureBox3.Location = New System.Drawing.Point(1117, 19)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(24, 19)
        Me.PictureBox3.TabIndex = 29
        Me.PictureBox3.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox4.Image = Global.Import_Planifier_IM.My.Resources.Resources.Checked
        Me.PictureBox4.Location = New System.Drawing.Point(1094, 19)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(24, 19)
        Me.PictureBox4.TabIndex = 28
        Me.PictureBox4.TabStop = False
        Me.PictureBox4.Tag = "Sélectionner tous"
        '
        'ComboDate
        '
        Me.ComboDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboDate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboDate.FormattingEnabled = True
        Me.ComboDate.Items.AddRange(New Object() {"jj/mm/aa", "jj/mm/aaaa", "aammjj", "jjmmaa", "aaaammjj", "jjmmaaaa", "aa-mm-jj", "jj-mm-aa", "aaaa-mm-jj", "jj-mm-aaaa", "mmaa", "aamm", "mmaaaa", "aaaamm", "mm/aa", "aa/mm", "mm/aaaa", "aaaa/mm", "mm-aa", "aa-mm", "mm-aaaa", "aaaa-mm", "mmjjaa", "mmjjaaaa", "aajjmm", "aaaajjmm", "mm/jj/aa", "mm/jj/aaaa", "mm-jj-aa", "mm-jj-aaaa", "aa-jj-mm", "aaaa-jj-mm"})
        Me.ComboDate.Location = New System.Drawing.Point(226, 16)
        Me.ComboDate.Name = "ComboDate"
        Me.ComboDate.Size = New System.Drawing.Size(180, 21)
        Me.ComboDate.TabIndex = 27
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(141, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 13)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Format de Date :"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lblType)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkPieceAuto)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Datagridaffiche)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(1184, 495)
        Me.SplitContainer2.SplitterDistance = 146
        Me.SplitContainer2.TabIndex = 0
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.PicLigne)
        Me.SplitContainer3.Panel1.Controls.Add(Me.DataListeIntegrer)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer7)
        Me.SplitContainer3.Size = New System.Drawing.Size(1184, 345)
        Me.SplitContainer3.SplitterDistance = 289
        Me.SplitContainer3.TabIndex = 0
        '
        'SplitContainer7
        '
        Me.SplitContainer7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer7.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer7.Name = "SplitContainer7"
        '
        'SplitContainer7.Panel1
        '
        Me.SplitContainer7.Panel1.Controls.Add(Me.SplitContainer8)
        '
        'SplitContainer7.Panel2
        '
        Me.SplitContainer7.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer7.Size = New System.Drawing.Size(1184, 52)
        Me.SplitContainer7.SplitterDistance = 560
        Me.SplitContainer7.TabIndex = 4
        '
        'SplitContainer8
        '
        Me.SplitContainer8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer8.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer8.Name = "SplitContainer8"
        '
        'SplitContainer8.Panel1
        '
        Me.SplitContainer8.Panel1.Tag = "DPHJ"
        '
        'SplitContainer8.Panel2
        '
        Me.SplitContainer8.Panel2.Controls.Add(Me.GroupBox3)
        Me.SplitContainer8.Size = New System.Drawing.Size(560, 52)
        Me.SplitContainer8.SplitterDistance = 25
        Me.SplitContainer8.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.infosExport)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(531, 52)
        Me.GroupBox3.TabIndex = 22
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Information d'indication "
        '
        'infosExport
        '
        Me.infosExport.AutoSize = True
        Me.infosExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.infosExport.ForeColor = System.Drawing.Color.Red
        Me.infosExport.Location = New System.Drawing.Point(45, 27)
        Me.infosExport.Name = "infosExport"
        Me.infosExport.Size = New System.Drawing.Size(20, 16)
        Me.infosExport.TabIndex = 19
        Me.infosExport.Text = "..."
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.BT_integrer)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.btnview)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(620, 52)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Action"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Image = Global.Import_Planifier_IM.My.Resources.Resources.export
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(184, 11)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 32)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Transformation"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ChEncapsuler
        '
        Me.ChEncapsuler.AutoSize = True
        Me.ChEncapsuler.Location = New System.Drawing.Point(588, 19)
        Me.ChEncapsuler.Name = "ChEncapsuler"
        Me.ChEncapsuler.Size = New System.Drawing.Size(231, 17)
        Me.ChEncapsuler.TabIndex = 27
        Me.ChEncapsuler.Text = "Encapsulation des documents dans un seul"
        Me.ChEncapsuler.UseVisualStyleBackColor = True
        '
        'Frm_MvtE_S
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(1184, 552)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Frm_MvtE_S"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Frm_FluxEntrantCritére"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Datagridaffiche, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicLigne, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer7.Panel1.ResumeLayout(False)
        Me.SplitContainer7.Panel2.ResumeLayout(False)
        Me.SplitContainer7.ResumeLayout(False)
        Me.SplitContainer8.Panel2.ResumeLayout(False)
        Me.SplitContainer8.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Datagridaffiche As System.Windows.Forms.DataGridView
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents PicLigne As System.Windows.Forms.PictureBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BT_integrer As System.Windows.Forms.Button
    Friend WithEvents btnview As System.Windows.Forms.Button
    Friend WithEvents C1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents C2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents C3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents C4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents C7 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents C5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents C6 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents C8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer7 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer8 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents infosExport As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents ChkPieceAuto As System.Windows.Forms.CheckBox
    Friend WithEvents CheckFille As System.Windows.Forms.CheckBox
    Friend WithEvents BtnListe As System.Windows.Forms.Button
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents ComboDate As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents WAREHOUSE_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PRODUCT_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LOT_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SERIAL_NO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EXP_DATE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OWNER_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CONTAINER_NO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SIGN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents REASON As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DESCRIPTION As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SORDER_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LINE_NUMBER As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents REC_ORDER_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QUANTITY As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QUALITY As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BEST_BEFORE_DATE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PRODUCTION_DATE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SHELF_LIFE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DATE_DAYS_OF_LIFE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SIZE_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COLOR_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SOURCE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VERSION_CODE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PRODUCTION_METHOD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents POST_PRODUCTION_TREATMENT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LOT_COUNT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ChEncapsuler As System.Windows.Forms.CheckBox
End Class
