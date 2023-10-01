<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmExtractionBCFournisseur
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmExtractionBCFournisseur))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.lblLignes = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblInfos = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RbtStock = New System.Windows.Forms.RadioButton
        Me.RbtAchat = New System.Windows.Forms.RadioButton
        Me.RbtVente = New System.Windows.Forms.RadioButton
        Me.lblsms = New System.Windows.Forms.Label
        Me.lblSne = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblligne = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblentete = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.CmbStatut = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.ChFlag = New System.Windows.Forms.CheckBox
        Me.lblinfosLibre = New System.Windows.Forms.Label
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
        Me.Choix = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Status = New System.Windows.Forms.DataGridViewImageColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.ListBox = New System.Windows.Forms.ListBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.BtnModif = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BackColor = System.Drawing.Color.AliceBlue
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1019, 522)
        Me.SplitContainer1.SplitterDistance = 210
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer3.Size = New System.Drawing.Size(1019, 210)
        Me.SplitContainer3.SplitterDistance = 672
        Me.SplitContainer3.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.lblLignes)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.lblInfos)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.lblsms)
        Me.GroupBox1.Controls.Add(Me.lblSne)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblligne)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.lblentete)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.lblinfosLibre)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(672, 210)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Traitement"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(387, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(219, 13)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Evolution du traitement des Lignes  encours :"
        '
        'lblLignes
        '
        Me.lblLignes.AutoSize = True
        Me.lblLignes.Location = New System.Drawing.Point(605, 128)
        Me.lblLignes.Name = "lblLignes"
        Me.lblLignes.Size = New System.Drawing.Size(39, 13)
        Me.lblLignes.TabIndex = 25
        Me.lblLignes.Text = "Label2"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(444, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(162, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Evolution du traitement encours :"
        '
        'lblInfos
        '
        Me.lblInfos.AutoSize = True
        Me.lblInfos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfos.ForeColor = System.Drawing.Color.Red
        Me.lblInfos.Location = New System.Drawing.Point(224, 149)
        Me.lblInfos.Name = "lblInfos"
        Me.lblInfos.Size = New System.Drawing.Size(39, 20)
        Me.lblInfos.TabIndex = 23
        Me.lblInfos.Text = "......"
        Me.lblInfos.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RbtStock)
        Me.GroupBox2.Controls.Add(Me.RbtAchat)
        Me.GroupBox2.Controls.Add(Me.RbtVente)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 42)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(654, 54)
        Me.GroupBox2.TabIndex = 16
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Domaine du document"
        '
        'RbtStock
        '
        Me.RbtStock.AutoSize = True
        Me.RbtStock.Location = New System.Drawing.Point(421, 19)
        Me.RbtStock.Name = "RbtStock"
        Me.RbtStock.Size = New System.Drawing.Size(128, 17)
        Me.RbtStock.TabIndex = 5
        Me.RbtStock.Text = "Achat(Bon de Retour)"
        Me.RbtStock.UseVisualStyleBackColor = True
        Me.RbtStock.Visible = False
        '
        'RbtAchat
        '
        Me.RbtAchat.AutoSize = True
        Me.RbtAchat.Checked = True
        Me.RbtAchat.Location = New System.Drawing.Point(239, 19)
        Me.RbtAchat.Name = "RbtAchat"
        Me.RbtAchat.Size = New System.Drawing.Size(169, 17)
        Me.RbtAchat.TabIndex = 4
        Me.RbtAchat.TabStop = True
        Me.RbtAchat.Text = "Achat(Commande Fournisseur)"
        Me.RbtAchat.UseVisualStyleBackColor = True
        '
        'RbtVente
        '
        Me.RbtVente.AutoSize = True
        Me.RbtVente.Location = New System.Drawing.Point(106, 19)
        Me.RbtVente.Name = "RbtVente"
        Me.RbtVente.Size = New System.Drawing.Size(120, 17)
        Me.RbtVente.TabIndex = 3
        Me.RbtVente.Text = "Vente(Retour Client)"
        Me.RbtVente.UseVisualStyleBackColor = True
        '
        'lblsms
        '
        Me.lblsms.AutoSize = True
        Me.lblsms.Location = New System.Drawing.Point(605, 108)
        Me.lblsms.Name = "lblsms"
        Me.lblsms.Size = New System.Drawing.Size(39, 13)
        Me.lblsms.TabIndex = 20
        Me.lblsms.Text = "Label2"
        '
        'lblSne
        '
        Me.lblSne.AutoSize = True
        Me.lblSne.Location = New System.Drawing.Point(199, 103)
        Me.lblSne.Name = "lblSne"
        Me.lblSne.Size = New System.Drawing.Size(25, 13)
        Me.lblSne.TabIndex = 15
        Me.lblSne.Text = "Null"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(14, 123)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Traitemen encours :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(113, 123)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(197, 13)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Extraction des Commandes Fournisseurs"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(14, 102)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(192, 13)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Enchainement du Senario d'execution :"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(142, 103)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(39, 13)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "Label1"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(225, 189)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(138, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Nombre d'infos libre Traiter :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 189)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Nombre de Ligne Traiter :"
        '
        'lblligne
        '
        Me.lblligne.AutoSize = True
        Me.lblligne.Location = New System.Drawing.Point(142, 192)
        Me.lblligne.Name = "lblligne"
        Me.lblligne.Size = New System.Drawing.Size(19, 13)
        Me.lblligne.TabIndex = 3
        Me.lblligne.Text = "00"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 168)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Nombre d'entete Traiter :"
        '
        'lblentete
        '
        Me.lblentete.AutoSize = True
        Me.lblentete.Location = New System.Drawing.Point(142, 169)
        Me.lblentete.Name = "lblentete"
        Me.lblentete.Size = New System.Drawing.Size(19, 13)
        Me.lblentete.TabIndex = 1
        Me.lblentete.Text = "00"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.CmbStatut)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.ChFlag)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(654, 29)
        Me.Panel1.TabIndex = 0
        '
        'CmbStatut
        '
        Me.CmbStatut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbStatut.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CmbStatut.FormattingEnabled = True
        Me.CmbStatut.Items.AddRange(New Object() {"Saisi", "Confirmé", "Réceptionné"})
        Me.CmbStatut.Location = New System.Drawing.Point(133, 4)
        Me.CmbStatut.Name = "CmbStatut"
        Me.CmbStatut.Size = New System.Drawing.Size(102, 21)
        Me.CmbStatut.TabIndex = 30
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(26, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Statut du Document"
        '
        'ChFlag
        '
        Me.ChFlag.AutoSize = True
        Me.ChFlag.Checked = True
        Me.ChFlag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChFlag.Location = New System.Drawing.Point(457, 5)
        Me.ChFlag.Name = "ChFlag"
        Me.ChFlag.Size = New System.Drawing.Size(117, 17)
        Me.ChFlag.TabIndex = 28
        Me.ChFlag.Text = "Flagage document "
        Me.ChFlag.UseVisualStyleBackColor = True
        '
        'lblinfosLibre
        '
        Me.lblinfosLibre.AutoSize = True
        Me.lblinfosLibre.Location = New System.Drawing.Point(366, 190)
        Me.lblinfosLibre.Name = "lblinfosLibre"
        Me.lblinfosLibre.Size = New System.Drawing.Size(19, 13)
        Me.lblinfosLibre.TabIndex = 6
        Me.lblinfosLibre.Text = "00"
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
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Societe1, Me.Type1, Me.Chemin1, Me.UserSage1, Me.PasseSage1, Me.Serveur1, Me.bdd1, Me.NomUtil, Me.Mot, Me.Choix, Me.Status})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.EnableHeadersVisualStyles = False
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        Me.DataListeIntegrer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowTemplate.Height = 24
        Me.DataListeIntegrer.Size = New System.Drawing.Size(343, 210)
        Me.DataListeIntegrer.TabIndex = 15
        '
        'Societe1
        '
        Me.Societe1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Societe1.HeaderText = "Société"
        Me.Societe1.Name = "Societe1"
        Me.Societe1.ReadOnly = True
        Me.Societe1.Width = 68
        '
        'Type1
        '
        Me.Type1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Type1.DefaultCellStyle = DataGridViewCellStyle1
        Me.Type1.HeaderText = "Type Base"
        Me.Type1.Name = "Type1"
        Me.Type1.ReadOnly = True
        Me.Type1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Type1.Width = 83
        '
        'Chemin1
        '
        Me.Chemin1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Chemin1.HeaderText = "Fichier Sage"
        Me.Chemin1.Name = "Chemin1"
        Me.Chemin1.ReadOnly = True
        Me.Chemin1.Visible = False
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
        Me.PasseSage1.Visible = False
        Me.PasseSage1.Width = 125
        '
        'Serveur1
        '
        Me.Serveur1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Serveur1.HeaderText = "Serveur SQL"
        Me.Serveur1.Name = "Serveur1"
        Me.Serveur1.Visible = False
        '
        'bdd1
        '
        Me.bdd1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.bdd1.HeaderText = "Base SQL"
        Me.bdd1.Name = "bdd1"
        Me.bdd1.Visible = False
        '
        'NomUtil
        '
        Me.NomUtil.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        DataGridViewCellStyle2.Format = "N0"
        Me.NomUtil.DefaultCellStyle = DataGridViewCellStyle2
        Me.NomUtil.FillWeight = 40.0!
        Me.NomUtil.HeaderText = "Nom SQL"
        Me.NomUtil.Name = "NomUtil"
        Me.NomUtil.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.NomUtil.Visible = False
        Me.NomUtil.Width = 80
        '
        'Mot
        '
        Me.Mot.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Mot.HeaderText = "Mot de Passe"
        Me.Mot.Name = "Mot"
        Me.Mot.Visible = False
        '
        'Choix
        '
        Me.Choix.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Choix.HeaderText = "Choix"
        Me.Choix.Name = "Choix"
        '
        'Status
        '
        Me.Status.HeaderText = "Status"
        Me.Status.Image = Global.Import_Planifier_IM.My.Resources.Resources.btFermer22
        Me.Status.Name = "Status"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.ListBox)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Button4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Button3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnModif)
        Me.SplitContainer2.Size = New System.Drawing.Size(1019, 308)
        Me.SplitContainer2.SplitterDistance = 273
        Me.SplitContainer2.TabIndex = 0
        '
        'ListBox
        '
        Me.ListBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox.FormattingEnabled = True
        Me.ListBox.Location = New System.Drawing.Point(0, 0)
        Me.ListBox.Name = "ListBox"
        Me.ListBox.Size = New System.Drawing.Size(1019, 264)
        Me.ListBox.TabIndex = 1
        '
        'Button4
        '
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Image = Global.Import_Planifier_IM.My.Resources.Resources.exportcsv11
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(591, 10)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(184, 27)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "Traitement en arriere plan"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(469, 10)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(110, 27)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "Quitter"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BtnModif
        '
        Me.BtnModif.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnModif.Image = Global.Import_Planifier_IM.My.Resources.Resources.Creer1
        Me.BtnModif.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModif.Location = New System.Drawing.Point(244, 10)
        Me.BtnModif.Name = "BtnModif"
        Me.BtnModif.Size = New System.Drawing.Size(175, 27)
        Me.BtnModif.TabIndex = 4
        Me.BtnModif.Text = "Lancer le Traitement"
        Me.BtnModif.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker4
        '
        '
        'FrmExtractionBCFournisseur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1019, 522)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmExtractionBCFournisseur"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Extraction des commande Fournisseur"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblligne As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblentete As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblinfosLibre As System.Windows.Forms.Label
    Friend WithEvents lblSne As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BtnModif As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RbtStock As System.Windows.Forms.RadioButton
    Friend WithEvents RbtAchat As System.Windows.Forms.RadioButton
    Friend WithEvents RbtVente As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblsms As System.Windows.Forms.Label
    Friend WithEvents lblInfos As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblLignes As System.Windows.Forms.Label
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents ChFlag As System.Windows.Forms.CheckBox
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
    Friend WithEvents Choix As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Status As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents CmbStatut As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
