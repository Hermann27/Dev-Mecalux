<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_FichierConfiguration
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_FichierConfiguration))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_FicCpta = New System.Windows.Forms.Button
        Me.CkConso = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPasw = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtBDCpta = New System.Windows.Forms.TextBox
        Me.TxtUserCpta = New System.Windows.Forms.TextBox
        Me.TxtUtilisateur = New System.Windows.Forms.TextBox
        Me.TxtPasword = New System.Windows.Forms.TextBox
        Me.TxtFichierCpta = New System.Windows.Forms.TextBox
        Me.Txtsql = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CheckLot = New System.Windows.Forms.CheckBox
        Me.ChekCodeEDI = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtFlagueArticle = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.BtnCheminMecalux = New System.Windows.Forms.Button
        Me.txtCheminMecalux = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.BtnOpenCheminErreur = New System.Windows.Forms.Button
        Me.txtCheminErreur = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.BtnCheminXfert = New System.Windows.Forms.Button
        Me.txtCheminXfert = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.BtnVST = New System.Windows.Forms.Button
        Me.txtFileVSTTempon = New System.Windows.Forms.TextBox
        Me.btnCRP = New System.Windows.Forms.Button
        Me.txtFileCRPTempon = New System.Windows.Forms.TextBox
        Me.BtnCSO = New System.Windows.Forms.Button
        Me.txtFileCSOTempon = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtFlag = New System.Windows.Forms.TextBox
        Me.Bt_tiers = New System.Windows.Forms.Button
        Me.Bt_Article = New System.Windows.Forms.Button
        Me.BT_Access = New System.Windows.Forms.Button
        Me.BT_FicJournal = New System.Windows.Forms.Button
        Me.BT_FicRep = New System.Windows.Forms.Button
        Me.Txtiers = New System.Windows.Forms.TextBox
        Me.TxtArticle = New System.Windows.Forms.TextBox
        Me.TxtAccess = New System.Windows.Forms.TextBox
        Me.TxtFilejr = New System.Windows.Forms.TextBox
        Me.Txt_Rep = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.RadioButton6 = New System.Windows.Forms.RadioButton
        Me.RadioButton5 = New System.Windows.Forms.RadioButton
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.OpenFileFicCpta = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepjournal = New System.Windows.Forms.FolderBrowserDialog
        Me.FolderRepsaving = New System.Windows.Forms.FolderBrowserDialog
        Me.FolderRepFact = New System.Windows.Forms.FolderBrowserDialog
        Me.FolderRepSave = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileAccess = New System.Windows.Forms.OpenFileDialog
        Me.OpenProgExterne = New System.Windows.Forms.OpenFileDialog
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.CmbStatut = New System.Windows.Forms.ComboBox
        Me.CmbStatutFrs = New System.Windows.Forms.ComboBox
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(8, 20)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(668, 498)
        Me.TabControl1.TabIndex = 62
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(660, 472)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Paramètre de Consolidation"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Khaki
        Me.GroupBox1.Controls.Add(Me.BT_FicCpta)
        Me.GroupBox1.Controls.Add(Me.CkConso)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.TxtPasw)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtBDCpta)
        Me.GroupBox1.Controls.Add(Me.TxtUserCpta)
        Me.GroupBox1.Controls.Add(Me.TxtUtilisateur)
        Me.GroupBox1.Controls.Add(Me.TxtPasword)
        Me.GroupBox1.Controls.Add(Me.TxtFichierCpta)
        Me.GroupBox1.Controls.Add(Me.Txtsql)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(21, 23)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(614, 432)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Paramétres de Consolidation"
        '
        'BT_FicCpta
        '
        Me.BT_FicCpta.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_FicCpta.Image = Global.Import_Planifier_IM.My.Resources.Resources.documents_16
        Me.BT_FicCpta.Location = New System.Drawing.Point(485, 148)
        Me.BT_FicCpta.Name = "BT_FicCpta"
        Me.BT_FicCpta.Size = New System.Drawing.Size(29, 23)
        Me.BT_FicCpta.TabIndex = 29
        Me.BT_FicCpta.TabStop = False
        Me.BT_FicCpta.UseVisualStyleBackColor = True
        '
        'CkConso
        '
        Me.CkConso.AutoSize = True
        Me.CkConso.Location = New System.Drawing.Point(19, 91)
        Me.CkConso.Name = "CkConso"
        Me.CkConso.Size = New System.Drawing.Size(151, 19)
        Me.CkConso.TabIndex = 0
        Me.CkConso.Text = "Tentative de Connexion"
        Me.CkConso.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 209)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(138, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Mot de Passe  Comptable"
        '
        'TxtPasw
        '
        Me.TxtPasw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtPasw.Location = New System.Drawing.Point(174, 322)
        Me.TxtPasw.Name = "TxtPasw"
        Me.TxtPasw.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtPasw.Size = New System.Drawing.Size(306, 23)
        Me.TxtPasw.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 238)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(151, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Base de Données SQL Server"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(10, 325)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(132, 13)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Mot de Passe SQL Server"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(10, 296)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(119, 13)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Utilisateur SQL Server "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 151)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(157, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Chemin du Fichier Comptable"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 180)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Utilisateur Comptable"
        '
        'TxtBDCpta
        '
        Me.TxtBDCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtBDCpta.Location = New System.Drawing.Point(174, 235)
        Me.TxtBDCpta.Name = "TxtBDCpta"
        Me.TxtBDCpta.Size = New System.Drawing.Size(306, 23)
        Me.TxtBDCpta.TabIndex = 3
        '
        'TxtUserCpta
        '
        Me.TxtUserCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtUserCpta.Location = New System.Drawing.Point(173, 177)
        Me.TxtUserCpta.Name = "TxtUserCpta"
        Me.TxtUserCpta.Size = New System.Drawing.Size(307, 23)
        Me.TxtUserCpta.TabIndex = 1
        '
        'TxtUtilisateur
        '
        Me.TxtUtilisateur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtUtilisateur.Location = New System.Drawing.Point(174, 293)
        Me.TxtUtilisateur.Name = "TxtUtilisateur"
        Me.TxtUtilisateur.Size = New System.Drawing.Size(306, 23)
        Me.TxtUtilisateur.TabIndex = 5
        '
        'TxtPasword
        '
        Me.TxtPasword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtPasword.Location = New System.Drawing.Point(173, 206)
        Me.TxtPasword.Name = "TxtPasword"
        Me.TxtPasword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtPasword.Size = New System.Drawing.Size(307, 23)
        Me.TxtPasword.TabIndex = 2
        '
        'TxtFichierCpta
        '
        Me.TxtFichierCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtFichierCpta.Location = New System.Drawing.Point(173, 148)
        Me.TxtFichierCpta.Name = "TxtFichierCpta"
        Me.TxtFichierCpta.Size = New System.Drawing.Size(306, 23)
        Me.TxtFichierCpta.TabIndex = 0
        '
        'Txtsql
        '
        Me.Txtsql.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txtsql.Location = New System.Drawing.Point(174, 264)
        Me.Txtsql.Name = "Txtsql"
        Me.Txtsql.Size = New System.Drawing.Size(306, 23)
        Me.Txtsql.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(10, 267)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "Serveur SQL"
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Image = Global.Import_Planifier_IM.My.Resources.Resources.k1
        Me.PictureBox1.Location = New System.Drawing.Point(3, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(608, 410)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 30
        Me.PictureBox1.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage2.Controls.Add(Me.GroupBox3)
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(660, 472)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Fichier Access/Journalisation"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox3.Controls.Add(Me.CheckLot)
        Me.GroupBox3.Controls.Add(Me.ChekCodeEDI)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox3.Location = New System.Drawing.Point(6, 397)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(648, 69)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Optional"
        '
        'CheckLot
        '
        Me.CheckLot.AutoSize = True
        Me.CheckLot.Location = New System.Drawing.Point(406, 23)
        Me.CheckLot.Name = "CheckLot"
        Me.CheckLot.Size = New System.Drawing.Size(211, 19)
        Me.CheckLot.TabIndex = 85
        Me.CheckLot.Text = "Gestion des Lot-->Commentaire ?"
        Me.CheckLot.UseVisualStyleBackColor = True
        '
        'ChekCodeEDI
        '
        Me.ChekCodeEDI.AutoSize = True
        Me.ChekCodeEDI.Location = New System.Drawing.Point(19, 23)
        Me.ChekCodeEDI.Name = "ChekCodeEDI"
        Me.ChekCodeEDI.Size = New System.Drawing.Size(365, 19)
        Me.ChekCodeEDI.TabIndex = 59
        Me.ChekCodeEDI.Text = "Utiliser le Code EDI comme code barre de l'article conditionné"
        Me.ChekCodeEDI.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Khaki
        Me.GroupBox2.Controls.Add(Me.txtFlagueArticle)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.BtnCheminMecalux)
        Me.GroupBox2.Controls.Add(Me.txtCheminMecalux)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.BtnOpenCheminErreur)
        Me.GroupBox2.Controls.Add(Me.txtCheminErreur)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.BtnCheminXfert)
        Me.GroupBox2.Controls.Add(Me.txtCheminXfert)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.BtnVST)
        Me.GroupBox2.Controls.Add(Me.txtFileVSTTempon)
        Me.GroupBox2.Controls.Add(Me.btnCRP)
        Me.GroupBox2.Controls.Add(Me.txtFileCRPTempon)
        Me.GroupBox2.Controls.Add(Me.BtnCSO)
        Me.GroupBox2.Controls.Add(Me.txtFileCSOTempon)
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Controls.Add(Me.TxtFlag)
        Me.GroupBox2.Controls.Add(Me.Bt_tiers)
        Me.GroupBox2.Controls.Add(Me.Bt_Article)
        Me.GroupBox2.Controls.Add(Me.BT_Access)
        Me.GroupBox2.Controls.Add(Me.BT_FicJournal)
        Me.GroupBox2.Controls.Add(Me.BT_FicRep)
        Me.GroupBox2.Controls.Add(Me.Txtiers)
        Me.GroupBox2.Controls.Add(Me.TxtArticle)
        Me.GroupBox2.Controls.Add(Me.TxtAccess)
        Me.GroupBox2.Controls.Add(Me.TxtFilejr)
        Me.GroupBox2.Controls.Add(Me.Txt_Rep)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(648, 385)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Autres Parametres"
        '
        'txtFlagueArticle
        '
        Me.txtFlagueArticle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFlagueArticle.Location = New System.Drawing.Point(523, 346)
        Me.txtFlagueArticle.Name = "txtFlagueArticle"
        Me.txtFlagueArticle.Size = New System.Drawing.Size(72, 21)
        Me.txtFlagueArticle.TabIndex = 85
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(436, 348)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(81, 15)
        Me.Label17.TabIndex = 84
        Me.Label17.Text = "Flague Article"
        '
        'BtnCheminMecalux
        '
        Me.BtnCheminMecalux.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnCheminMecalux.Image = CType(resources.GetObject("BtnCheminMecalux.Image"), System.Drawing.Image)
        Me.BtnCheminMecalux.Location = New System.Drawing.Point(602, 311)
        Me.BtnCheminMecalux.Name = "BtnCheminMecalux"
        Me.BtnCheminMecalux.Size = New System.Drawing.Size(29, 21)
        Me.BtnCheminMecalux.TabIndex = 83
        Me.BtnCheminMecalux.UseVisualStyleBackColor = True
        '
        'txtCheminMecalux
        '
        Me.txtCheminMecalux.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCheminMecalux.Location = New System.Drawing.Point(322, 311)
        Me.txtCheminMecalux.Name = "txtCheminMecalux"
        Me.txtCheminMecalux.Size = New System.Drawing.Size(273, 21)
        Me.txtCheminMecalux.TabIndex = 81
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(144, 314)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(172, 15)
        Me.Label21.TabIndex = 82
        Me.Label21.Text = "ARCHIVAGE Fichier (Mecalux)"
        '
        'BtnOpenCheminErreur
        '
        Me.BtnOpenCheminErreur.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnOpenCheminErreur.Image = CType(resources.GetObject("BtnOpenCheminErreur.Image"), System.Drawing.Image)
        Me.BtnOpenCheminErreur.Location = New System.Drawing.Point(602, 283)
        Me.BtnOpenCheminErreur.Name = "BtnOpenCheminErreur"
        Me.BtnOpenCheminErreur.Size = New System.Drawing.Size(29, 21)
        Me.BtnOpenCheminErreur.TabIndex = 80
        Me.BtnOpenCheminErreur.UseVisualStyleBackColor = True
        '
        'txtCheminErreur
        '
        Me.txtCheminErreur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCheminErreur.Location = New System.Drawing.Point(322, 283)
        Me.txtCheminErreur.Name = "txtCheminErreur"
        Me.txtCheminErreur.Size = New System.Drawing.Size(273, 21)
        Me.txtCheminErreur.TabIndex = 78
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(153, 287)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(163, 15)
        Me.Label20.TabIndex = 79
        Me.Label20.Text = "ARCHIVAGE CSV (ERREUR)"
        '
        'BtnCheminXfert
        '
        Me.BtnCheminXfert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnCheminXfert.Image = CType(resources.GetObject("BtnCheminXfert.Image"), System.Drawing.Image)
        Me.BtnCheminXfert.Location = New System.Drawing.Point(602, 252)
        Me.BtnCheminXfert.Name = "BtnCheminXfert"
        Me.BtnCheminXfert.Size = New System.Drawing.Size(29, 21)
        Me.BtnCheminXfert.TabIndex = 71
        Me.BtnCheminXfert.UseVisualStyleBackColor = True
        '
        'txtCheminXfert
        '
        Me.txtCheminXfert.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCheminXfert.Location = New System.Drawing.Point(322, 252)
        Me.txtCheminXfert.Name = "txtCheminXfert"
        Me.txtCheminXfert.Size = New System.Drawing.Size(273, 21)
        Me.txtCheminXfert.TabIndex = 69
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(3, 256)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(313, 15)
        Me.Label16.TabIndex = 70
        Me.Label16.Text = "Répertoire(Mvt Transfert) Temporaire (EasyWMS->ERP)"
        '
        'BtnVST
        '
        Me.BtnVST.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnVST.Image = CType(resources.GetObject("BtnVST.Image"), System.Drawing.Image)
        Me.BtnVST.Location = New System.Drawing.Point(602, 224)
        Me.BtnVST.Name = "BtnVST"
        Me.BtnVST.Size = New System.Drawing.Size(29, 21)
        Me.BtnVST.TabIndex = 68
        Me.BtnVST.UseVisualStyleBackColor = True
        '
        'txtFileVSTTempon
        '
        Me.txtFileVSTTempon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileVSTTempon.Location = New System.Drawing.Point(322, 224)
        Me.txtFileVSTTempon.Name = "txtFileVSTTempon"
        Me.txtFileVSTTempon.Size = New System.Drawing.Size(273, 21)
        Me.txtFileVSTTempon.TabIndex = 66
        '
        'btnCRP
        '
        Me.btnCRP.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCRP.Image = CType(resources.GetObject("btnCRP.Image"), System.Drawing.Image)
        Me.btnCRP.Location = New System.Drawing.Point(602, 196)
        Me.btnCRP.Name = "btnCRP"
        Me.btnCRP.Size = New System.Drawing.Size(29, 21)
        Me.btnCRP.TabIndex = 65
        Me.btnCRP.UseVisualStyleBackColor = True
        '
        'txtFileCRPTempon
        '
        Me.txtFileCRPTempon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileCRPTempon.Location = New System.Drawing.Point(322, 196)
        Me.txtFileCRPTempon.Name = "txtFileCRPTempon"
        Me.txtFileCRPTempon.Size = New System.Drawing.Size(273, 21)
        Me.txtFileCRPTempon.TabIndex = 63
        '
        'BtnCSO
        '
        Me.BtnCSO.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnCSO.Image = CType(resources.GetObject("BtnCSO.Image"), System.Drawing.Image)
        Me.BtnCSO.Location = New System.Drawing.Point(602, 168)
        Me.BtnCSO.Name = "BtnCSO"
        Me.BtnCSO.Size = New System.Drawing.Size(29, 21)
        Me.BtnCSO.TabIndex = 62
        Me.BtnCSO.UseVisualStyleBackColor = True
        '
        'txtFileCSOTempon
        '
        Me.txtFileCSOTempon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileCSOTempon.Location = New System.Drawing.Point(322, 168)
        Me.txtFileCSOTempon.Name = "txtFileCSOTempon"
        Me.txtFileCSOTempon.Size = New System.Drawing.Size(273, 21)
        Me.txtFileCSOTempon.TabIndex = 60
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(16, 348)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(228, 15)
        Me.Label18.TabIndex = 58
        Me.Label18.Text = "Zone de Flagage(Entête de documenet) "
        '
        'TxtFlag
        '
        Me.TxtFlag.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtFlag.Location = New System.Drawing.Point(322, 346)
        Me.TxtFlag.Name = "TxtFlag"
        Me.TxtFlag.Size = New System.Drawing.Size(108, 21)
        Me.TxtFlag.TabIndex = 57
        '
        'Bt_tiers
        '
        Me.Bt_tiers.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Bt_tiers.Image = CType(resources.GetObject("Bt_tiers.Image"), System.Drawing.Image)
        Me.Bt_tiers.Location = New System.Drawing.Point(602, 112)
        Me.Bt_tiers.Name = "Bt_tiers"
        Me.Bt_tiers.Size = New System.Drawing.Size(29, 21)
        Me.Bt_tiers.TabIndex = 53
        Me.Bt_tiers.UseVisualStyleBackColor = True
        '
        'Bt_Article
        '
        Me.Bt_Article.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Bt_Article.Image = CType(resources.GetObject("Bt_Article.Image"), System.Drawing.Image)
        Me.Bt_Article.Location = New System.Drawing.Point(602, 140)
        Me.Bt_Article.Name = "Bt_Article"
        Me.Bt_Article.Size = New System.Drawing.Size(29, 21)
        Me.Bt_Article.TabIndex = 54
        Me.Bt_Article.UseVisualStyleBackColor = True
        '
        'BT_Access
        '
        Me.BT_Access.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_Access.Image = CType(resources.GetObject("BT_Access.Image"), System.Drawing.Image)
        Me.BT_Access.Location = New System.Drawing.Point(602, 28)
        Me.BT_Access.Name = "BT_Access"
        Me.BT_Access.Size = New System.Drawing.Size(29, 21)
        Me.BT_Access.TabIndex = 50
        Me.BT_Access.UseVisualStyleBackColor = True
        '
        'BT_FicJournal
        '
        Me.BT_FicJournal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_FicJournal.Image = Global.Import_Planifier_IM.My.Resources.Resources.foldeopen_161
        Me.BT_FicJournal.Location = New System.Drawing.Point(602, 56)
        Me.BT_FicJournal.Name = "BT_FicJournal"
        Me.BT_FicJournal.Size = New System.Drawing.Size(29, 21)
        Me.BT_FicJournal.TabIndex = 51
        Me.BT_FicJournal.UseVisualStyleBackColor = True
        '
        'BT_FicRep
        '
        Me.BT_FicRep.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_FicRep.Image = CType(resources.GetObject("BT_FicRep.Image"), System.Drawing.Image)
        Me.BT_FicRep.Location = New System.Drawing.Point(602, 84)
        Me.BT_FicRep.Name = "BT_FicRep"
        Me.BT_FicRep.Size = New System.Drawing.Size(29, 21)
        Me.BT_FicRep.TabIndex = 52
        Me.BT_FicRep.UseVisualStyleBackColor = True
        '
        'Txtiers
        '
        Me.Txtiers.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txtiers.Location = New System.Drawing.Point(322, 112)
        Me.Txtiers.Name = "Txtiers"
        Me.Txtiers.Size = New System.Drawing.Size(273, 21)
        Me.Txtiers.TabIndex = 15
        '
        'TxtArticle
        '
        Me.TxtArticle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtArticle.Location = New System.Drawing.Point(322, 140)
        Me.TxtArticle.Name = "TxtArticle"
        Me.TxtArticle.Size = New System.Drawing.Size(273, 21)
        Me.TxtArticle.TabIndex = 16
        '
        'TxtAccess
        '
        Me.TxtAccess.BackColor = System.Drawing.SystemColors.Window
        Me.TxtAccess.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtAccess.Location = New System.Drawing.Point(322, 28)
        Me.TxtAccess.Name = "TxtAccess"
        Me.TxtAccess.ReadOnly = True
        Me.TxtAccess.Size = New System.Drawing.Size(273, 21)
        Me.TxtAccess.TabIndex = 12
        '
        'TxtFilejr
        '
        Me.TxtFilejr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtFilejr.Location = New System.Drawing.Point(322, 56)
        Me.TxtFilejr.Name = "TxtFilejr"
        Me.TxtFilejr.Size = New System.Drawing.Size(273, 21)
        Me.TxtFilejr.TabIndex = 13
        '
        'Txt_Rep
        '
        Me.Txt_Rep.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txt_Rep.Location = New System.Drawing.Point(322, 84)
        Me.Txt_Rep.Name = "Txt_Rep"
        Me.Txt_Rep.Size = New System.Drawing.Size(273, 21)
        Me.Txt_Rep.TabIndex = 14
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(21, 228)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(295, 15)
        Me.Label14.TabIndex = 67
        Me.Label14.Text = "Répertoire(Mvt Stock) Temporaire (EasyWMS->ERP)"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(34, 205)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(282, 15)
        Me.Label13.TabIndex = 64
        Me.Label13.Text = "Répertoire(VENTE) Temporaire (EasyWMS->ERP)"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(35, 172)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(281, 15)
        Me.Label12.TabIndex = 61
        Me.Label12.Text = "Répertoire(ACHAT) Temporaire (EasyWMS->ERP)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(106, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(210, 15)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Répertoire des fichiers extraire (ERP)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(68, 142)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(248, 15)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Répertoire des fichiers à intégre (EasyWMS)"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(165, 28)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(151, 15)
        Me.Label15.TabIndex = 42
        Me.Label15.Text = "Chemin du Fichier  Access"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(108, 56)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(208, 15)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Répertoire de Journalisation Fichiers"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(70, 84)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(246, 15)
        Me.Label11.TabIndex = 30
        Me.Label11.Text = "Répertoire de Sauvegarde Fichiers(CSV Ok)"
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage3.Controls.Add(Me.GroupBox4)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(660, 472)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Autres/"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Khaki
        Me.GroupBox4.Controls.Add(Me.CmbStatutFrs)
        Me.GroupBox4.Controls.Add(Me.CmbStatut)
        Me.GroupBox4.Controls.Add(Me.GroupBox5)
        Me.GroupBox4.Controls.Add(Me.Label22)
        Me.GroupBox4.Controls.Add(Me.Label19)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(654, 466)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "GroupBox4"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.RadioButton6)
        Me.GroupBox5.Controls.Add(Me.RadioButton5)
        Me.GroupBox5.Controls.Add(Me.RadioButton4)
        Me.GroupBox5.Controls.Add(Me.RadioButton3)
        Me.GroupBox5.Controls.Add(Me.RadioButton2)
        Me.GroupBox5.Controls.Add(Me.RadioButton1)
        Me.GroupBox5.Location = New System.Drawing.Point(19, 112)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(613, 328)
        Me.GroupBox5.TabIndex = 37
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "GroupBox5"
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(343, 178)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton6.TabIndex = 5
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "RadioButton6"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(343, 120)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton5.TabIndex = 4
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "RadioButton5"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(153, 120)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton4.TabIndex = 3
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "RadioButton4"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(153, 178)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "RadioButton3"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(343, 62)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "RadioButton2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(153, 62)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(31, 72)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(231, 13)
        Me.Label22.TabIndex = 35
        Me.Label22.Text = "Statut de document commande fournisseurs"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(63, 45)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(199, 13)
        Me.Label19.TabIndex = 33
        Me.Label19.Text = "Statut de document commande Client"
        '
        'FolderRepjournal
        '
        Me.FolderRepjournal.RootFolder = System.Environment.SpecialFolder.DesktopDirectory
        Me.FolderRepjournal.ShowNewFolderButton = False
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Image = Global.Import_Planifier_IM.My.Resources.Resources.delete_161
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(468, 528)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(82, 29)
        Me.Button2.TabIndex = 64
        Me.Button2.Text = "&Quitter"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_16
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(366, 528)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 29)
        Me.Button1.TabIndex = 63
        Me.Button1.Text = "&Valider"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CmbStatut
        '
        Me.CmbStatut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbStatut.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CmbStatut.FormattingEnabled = True
        Me.CmbStatut.Items.AddRange(New Object() {"Saisi", "Confirmé", "Réceptionné"})
        Me.CmbStatut.Location = New System.Drawing.Point(268, 37)
        Me.CmbStatut.Name = "CmbStatut"
        Me.CmbStatut.Size = New System.Drawing.Size(102, 21)
        Me.CmbStatut.TabIndex = 38
        '
        'CmbStatutFrs
        '
        Me.CmbStatutFrs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbStatutFrs.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CmbStatutFrs.FormattingEnabled = True
        Me.CmbStatutFrs.Items.AddRange(New Object() {"Saisi", "Confirmé", "Réceptionné"})
        Me.CmbStatutFrs.Location = New System.Drawing.Point(268, 64)
        Me.CmbStatutFrs.Name = "CmbStatutFrs"
        Me.CmbStatutFrs.Size = New System.Drawing.Size(102, 21)
        Me.CmbStatutFrs.TabIndex = 39
        '
        'Frm_FichierConfiguration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Honeydew
        Me.ClientSize = New System.Drawing.Size(699, 566)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Frm_FichierConfiguration"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fichier de configuration "
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CkConso As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPasw As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtBDCpta As System.Windows.Forms.TextBox
    Friend WithEvents TxtUserCpta As System.Windows.Forms.TextBox
    Friend WithEvents TxtUtilisateur As System.Windows.Forms.TextBox
    Friend WithEvents TxtPasword As System.Windows.Forms.TextBox
    Friend WithEvents TxtFichierCpta As System.Windows.Forms.TextBox
    Friend WithEvents Txtsql As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Bt_tiers As System.Windows.Forms.Button
    Friend WithEvents Bt_Article As System.Windows.Forms.Button
    Friend WithEvents BT_Access As System.Windows.Forms.Button
    Friend WithEvents BT_FicJournal As System.Windows.Forms.Button
    Friend WithEvents BT_FicRep As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Txtiers As System.Windows.Forms.TextBox
    Friend WithEvents TxtArticle As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtAccess As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtFilejr As System.Windows.Forms.TextBox
    Friend WithEvents Txt_Rep As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BT_FicCpta As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileFicCpta As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FolderRepjournal As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderRepsaving As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderRepFact As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderRepSave As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileAccess As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OpenProgExterne As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtFlag As System.Windows.Forms.TextBox
    Friend WithEvents ChekCodeEDI As System.Windows.Forms.CheckBox
    Friend WithEvents BtnCSO As System.Windows.Forms.Button
    Friend WithEvents txtFileCSOTempon As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BtnVST As System.Windows.Forms.Button
    Friend WithEvents txtFileVSTTempon As System.Windows.Forms.TextBox
    Friend WithEvents btnCRP As System.Windows.Forms.Button
    Friend WithEvents txtFileCRPTempon As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents BtnCheminXfert As System.Windows.Forms.Button
    Friend WithEvents txtCheminXfert As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BtnCheminMecalux As System.Windows.Forms.Button
    Friend WithEvents txtCheminMecalux As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents BtnOpenCheminErreur As System.Windows.Forms.Button
    Friend WithEvents txtCheminErreur As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckLot As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents txtFlagueArticle As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents CmbStatutFrs As System.Windows.Forms.ComboBox
    Friend WithEvents CmbStatut As System.Windows.Forms.ComboBox

End Class
