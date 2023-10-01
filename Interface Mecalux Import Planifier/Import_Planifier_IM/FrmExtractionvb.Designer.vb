<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmExtraction
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmExtraction))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.lblSne = New System.Windows.Forms.Label
        Me.CheckSommeil = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.RbtSommeilOFF = New System.Windows.Forms.RadioButton
        Me.RbtSommeilON = New System.Windows.Forms.RadioButton
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.CheckInfosLibre = New System.Windows.Forms.CheckBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.lblsmss = New System.Windows.Forms.Label
        Me.txtinfosLibre = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.ComboInfosLibre = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lblsms = New System.Windows.Forms.Label
        Me.CheckSuivi = New System.Windows.Forms.CheckBox
        Me.ComboSuivi = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblligne = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblentete = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ChBoxEC_Qté = New System.Windows.Forms.CheckBox
        Me.Ckmodifier = New System.Windows.Forms.CheckBox
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
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
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
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1019, 742)
        Me.SplitContainer1.SplitterDistance = 414
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DataListeIntegrer)
        Me.SplitContainer2.Size = New System.Drawing.Size(1019, 414)
        Me.SplitContainer2.SplitterDistance = 603
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblSne)
        Me.GroupBox3.Controls.Add(Me.CheckSommeil)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.Panel2)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox3.Location = New System.Drawing.Point(0, 280)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(603, 118)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Mise en sommeil de l’article "
        '
        'lblSne
        '
        Me.lblSne.AutoSize = True
        Me.lblSne.Location = New System.Drawing.Point(197, 64)
        Me.lblSne.Name = "lblSne"
        Me.lblSne.Size = New System.Drawing.Size(25, 13)
        Me.lblSne.TabIndex = 10
        Me.lblSne.Text = "Null"
        '
        'CheckSommeil
        '
        Me.CheckSommeil.AutoSize = True
        Me.CheckSommeil.Checked = True
        Me.CheckSommeil.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckSommeil.Location = New System.Drawing.Point(457, 18)
        Me.CheckSommeil.Name = "CheckSommeil"
        Me.CheckSommeil.Size = New System.Drawing.Size(126, 17)
        Me.CheckSommeil.TabIndex = 9
        Me.CheckSommeil.Text = "Activer Filtre Sommeil"
        Me.CheckSommeil.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(390, 60)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Traitemen encours :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(489, 60)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(111, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Extraction des Articles"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 63)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(192, 13)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "Enchainement du Senario d'execution :"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(140, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(39, 13)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Label1"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.RbtSommeilOFF)
        Me.Panel2.Controls.Add(Me.RbtSommeilON)
        Me.Panel2.Location = New System.Drawing.Point(12, 18)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(370, 38)
        Me.Panel2.TabIndex = 1
        '
        'RbtSommeilOFF
        '
        Me.RbtSommeilOFF.AutoSize = True
        Me.RbtSommeilOFF.Checked = True
        Me.RbtSommeilOFF.Location = New System.Drawing.Point(187, 10)
        Me.RbtSommeilOFF.Name = "RbtSommeilOFF"
        Me.RbtSommeilOFF.Size = New System.Drawing.Size(45, 17)
        Me.RbtSommeilOFF.TabIndex = 1
        Me.RbtSommeilOFF.TabStop = True
        Me.RbtSommeilOFF.Text = "Non"
        Me.RbtSommeilOFF.UseVisualStyleBackColor = True
        '
        'RbtSommeilON
        '
        Me.RbtSommeilON.AutoSize = True
        Me.RbtSommeilON.Location = New System.Drawing.Point(110, 10)
        Me.RbtSommeilON.Name = "RbtSommeilON"
        Me.RbtSommeilON.Size = New System.Drawing.Size(41, 17)
        Me.RbtSommeilON.TabIndex = 0
        Me.RbtSommeilON.Text = "Oui"
        Me.RbtSommeilON.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckInfosLibre)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.lblsmss)
        Me.GroupBox4.Controls.Add(Me.txtinfosLibre)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.ComboInfosLibre)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox4.Location = New System.Drawing.Point(0, 186)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(603, 94)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Zone Libre"
        '
        'CheckInfosLibre
        '
        Me.CheckInfosLibre.AutoSize = True
        Me.CheckInfosLibre.Location = New System.Drawing.Point(457, 19)
        Me.CheckInfosLibre.Name = "CheckInfosLibre"
        Me.CheckInfosLibre.Size = New System.Drawing.Size(136, 17)
        Me.CheckInfosLibre.TabIndex = 12
        Me.CheckInfosLibre.Text = "Activer Filtre Infos Libre"
        Me.CheckInfosLibre.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(73, 23)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(162, 13)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Evolution du traitement encours :"
        '
        'lblsmss
        '
        Me.lblsmss.AutoSize = True
        Me.lblsmss.Location = New System.Drawing.Point(241, 23)
        Me.lblsmss.Name = "lblsmss"
        Me.lblsmss.Size = New System.Drawing.Size(16, 13)
        Me.lblsmss.TabIndex = 18
        Me.lblsmss.Text = "..."
        '
        'txtinfosLibre
        '
        Me.txtinfosLibre.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtinfosLibre.Location = New System.Drawing.Point(326, 49)
        Me.txtinfosLibre.Name = "txtinfosLibre"
        Me.txtinfosLibre.Size = New System.Drawing.Size(267, 20)
        Me.txtinfosLibre.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(283, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(43, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Valeur :"
        '
        'ComboInfosLibre
        '
        Me.ComboInfosLibre.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboInfosLibre.FormattingEnabled = True
        Me.ComboInfosLibre.Location = New System.Drawing.Point(93, 49)
        Me.ComboInfosLibre.Name = "ComboInfosLibre"
        Me.ComboInfosLibre.Size = New System.Drawing.Size(187, 21)
        Me.ComboInfosLibre.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Choix infos Libre :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblsms)
        Me.GroupBox2.Controls.Add(Me.CheckSuivi)
        Me.GroupBox2.Controls.Add(Me.ComboSuivi)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.Location = New System.Drawing.Point(0, 100)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(603, 86)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Type de suivi du stock "
        '
        'lblsms
        '
        Me.lblsms.AutoSize = True
        Me.lblsms.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblsms.ForeColor = System.Drawing.Color.Red
        Me.lblsms.Location = New System.Drawing.Point(89, 56)
        Me.lblsms.Name = "lblsms"
        Me.lblsms.Size = New System.Drawing.Size(134, 20)
        Me.lblsms.TabIndex = 11
        Me.lblsms.Text = "........................."
        Me.lblsms.Visible = False
        '
        'CheckSuivi
        '
        Me.CheckSuivi.AutoSize = True
        Me.CheckSuivi.Checked = True
        Me.CheckSuivi.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckSuivi.Location = New System.Drawing.Point(457, 26)
        Me.CheckSuivi.Name = "CheckSuivi"
        Me.CheckSuivi.Size = New System.Drawing.Size(141, 17)
        Me.CheckSuivi.TabIndex = 10
        Me.CheckSuivi.Text = "Activer Filtre Suivi Stock"
        Me.CheckSuivi.UseVisualStyleBackColor = True
        '
        'ComboSuivi
        '
        Me.ComboSuivi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboSuivi.FormattingEnabled = True
        Me.ComboSuivi.Items.AddRange(New Object() {"Aucun", "Sérialisé", "CMUP", "FIFO", "LIFO", "Par Lot"})
        Me.ComboSuivi.Location = New System.Drawing.Point(57, 26)
        Me.ComboSuivi.Name = "ComboSuivi"
        Me.ComboSuivi.Size = New System.Drawing.Size(220, 21)
        Me.ComboSuivi.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Choix  :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblligne)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.lblentete)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.lblinfosLibre)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(603, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Mouvement de l'article"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(284, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(138, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Nombre d'infos libre Traiter :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Nombre de Ligne Traiter :"
        '
        'lblligne
        '
        Me.lblligne.AutoSize = True
        Me.lblligne.Location = New System.Drawing.Point(135, 70)
        Me.lblligne.Name = "lblligne"
        Me.lblligne.Size = New System.Drawing.Size(19, 13)
        Me.lblligne.TabIndex = 3
        Me.lblligne.Text = "00"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Nombre d'entete Traiter :"
        '
        'lblentete
        '
        Me.lblentete.AutoSize = True
        Me.lblentete.Location = New System.Drawing.Point(135, 47)
        Me.lblentete.Name = "lblentete"
        Me.lblentete.Size = New System.Drawing.Size(19, 13)
        Me.lblentete.TabIndex = 1
        Me.lblentete.Text = "00"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ChBoxEC_Qté)
        Me.Panel1.Controls.Add(Me.Ckmodifier)
        Me.Panel1.Location = New System.Drawing.Point(9, 15)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(574, 28)
        Me.Panel1.TabIndex = 0
        '
        'ChBoxEC_Qté
        '
        Me.ChBoxEC_Qté.AutoSize = True
        Me.ChBoxEC_Qté.Checked = True
        Me.ChBoxEC_Qté.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChBoxEC_Qté.Location = New System.Drawing.Point(127, 3)
        Me.ChBoxEC_Qté.Name = "ChBoxEC_Qté"
        Me.ChBoxEC_Qté.Size = New System.Drawing.Size(278, 17)
        Me.ChBoxEC_Qté.TabIndex = 11
        Me.ChBoxEC_Qté.Text = "Ignorer les conditionnements dont la quantité égal à 1"
        Me.ChBoxEC_Qté.UseVisualStyleBackColor = True
        '
        'Ckmodifier
        '
        Me.Ckmodifier.AutoSize = True
        Me.Ckmodifier.Checked = True
        Me.Ckmodifier.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Ckmodifier.Location = New System.Drawing.Point(445, 3)
        Me.Ckmodifier.Name = "Ckmodifier"
        Me.Ckmodifier.Size = New System.Drawing.Size(119, 17)
        Me.Ckmodifier.TabIndex = 20
        Me.Ckmodifier.Text = "Récemment modifié"
        Me.Ckmodifier.UseVisualStyleBackColor = True
        '
        'lblinfosLibre
        '
        Me.lblinfosLibre.AutoSize = True
        Me.lblinfosLibre.Location = New System.Drawing.Point(419, 51)
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
        Me.DataListeIntegrer.Size = New System.Drawing.Size(412, 414)
        Me.DataListeIntegrer.TabIndex = 11
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
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Type1.DefaultCellStyle = DataGridViewCellStyle1
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
        Me.Chemin1.Visible = False
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
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.ListBox)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.Button4)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Button3)
        Me.SplitContainer3.Panel2.Controls.Add(Me.BtnModif)
        Me.SplitContainer3.Size = New System.Drawing.Size(1019, 324)
        Me.SplitContainer3.SplitterDistance = 274
        Me.SplitContainer3.TabIndex = 0
        '
        'ListBox
        '
        Me.ListBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox.FormattingEnabled = True
        Me.ListBox.Location = New System.Drawing.Point(0, 0)
        Me.ListBox.Name = "ListBox"
        Me.ListBox.Size = New System.Drawing.Size(1019, 264)
        Me.ListBox.TabIndex = 0
        '
        'Button4
        '
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Image = Global.Import_Planifier_IM.My.Resources.Resources.exportcsv1
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(591, 8)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(184, 27)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Traitement en arriere plan"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(455, 8)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(110, 27)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Quitter"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BtnModif
        '
        Me.BtnModif.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnModif.Image = Global.Import_Planifier_IM.My.Resources.Resources.Creer1
        Me.BtnModif.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModif.Location = New System.Drawing.Point(244, 8)
        Me.BtnModif.Name = "BtnModif"
        Me.BtnModif.Size = New System.Drawing.Size(175, 27)
        Me.BtnModif.TabIndex = 0
        Me.BtnModif.Text = "Lancer le Traitement"
        Me.BtnModif.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerSupportsCancellation = True
        '
        'BackgroundWorker3
        '
        Me.BackgroundWorker3.WorkerSupportsCancellation = True
        '
        'BackgroundWorker4
        '
        Me.BackgroundWorker4.WorkerSupportsCancellation = True
        '
        'FrmExtraction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1019, 742)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmExtraction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Extraction des articles"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ComboSuivi As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblligne As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblentete As System.Windows.Forms.Label
    Friend WithEvents lblinfosLibre As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtinfosLibre As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ComboInfosLibre As System.Windows.Forms.ComboBox
    Friend WithEvents CheckInfosLibre As System.Windows.Forms.CheckBox
    Friend WithEvents CheckSommeil As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents RbtSommeilOFF As System.Windows.Forms.RadioButton
    Friend WithEvents RbtSommeilON As System.Windows.Forms.RadioButton
    Friend WithEvents CheckSuivi As System.Windows.Forms.CheckBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BtnModif As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblSne As System.Windows.Forms.Label
    Friend WithEvents ChBoxEC_Qté As System.Windows.Forms.CheckBox
    Friend WithEvents Ckmodifier As System.Windows.Forms.CheckBox
    Friend WithEvents lblsms As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblsmss As System.Windows.Forms.Label
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
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
End Class
