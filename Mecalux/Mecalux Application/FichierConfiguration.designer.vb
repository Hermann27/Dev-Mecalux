<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_FichierConfiguration
    Inherits Telerik.WinControls.UI.RadForm

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
        Me.TxtFichierCpta = New System.Windows.Forms.TextBox
        Me.TxtPasword = New System.Windows.Forms.TextBox
        Me.TxtUserCpta = New System.Windows.Forms.TextBox
        Me.TxtBDCpta = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BT_FicCpta = New Telerik.WinControls.UI.RadButton
        Me.CkConso = New System.Windows.Forms.CheckBox
        Me.TxtPasw = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtUtilisateur = New System.Windows.Forms.TextBox
        Me.Txtsql = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtAccess = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Txt_Rep = New System.Windows.Forms.TextBox
        Me.TxtFilejr = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.OpenFileFicCpta = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepjournal = New System.Windows.Forms.FolderBrowserDialog
        Me.FolderRepsaving = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileGesCom = New System.Windows.Forms.OpenFileDialog
        Me.FolderRepFact = New System.Windows.Forms.FolderBrowserDialog
        Me.FolderRepSave = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileAccess = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Bt_Article = New Telerik.WinControls.UI.RadButton
        Me.Bt_tiers = New Telerik.WinControls.UI.RadButton
        Me.BT_FicRep = New Telerik.WinControls.UI.RadButton
        Me.BT_FicJournal = New Telerik.WinControls.UI.RadButton
        Me.BT_Access = New Telerik.WinControls.UI.RadButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.Txtiers = New System.Windows.Forms.TextBox
        Me.TxtArticle = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.OpenProgExterne = New System.Windows.Forms.OpenFileDialog
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Windows8Theme1 = New Telerik.WinControls.Themes.Windows8Theme
        Me.TelerikMetroTheme1 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.TelerikMetroTheme2 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.Button2 = New Telerik.WinControls.UI.RadButton
        Me.Button1 = New Telerik.WinControls.UI.RadButton
        Me.GroupBox1.SuspendLayout()
        CType(Me.BT_FicCpta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.Bt_Article, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bt_tiers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_FicRep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_FicJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Access, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.Button2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Button1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtFichierCpta
        '
        Me.TxtFichierCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtFichierCpta.Location = New System.Drawing.Point(185, 58)
        Me.TxtFichierCpta.Name = "TxtFichierCpta"
        Me.TxtFichierCpta.Size = New System.Drawing.Size(306, 23)
        Me.TxtFichierCpta.TabIndex = 0
        '
        'TxtPasword
        '
        Me.TxtPasword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtPasword.Location = New System.Drawing.Point(185, 116)
        Me.TxtPasword.Name = "TxtPasword"
        Me.TxtPasword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtPasword.Size = New System.Drawing.Size(307, 23)
        Me.TxtPasword.TabIndex = 2
        '
        'TxtUserCpta
        '
        Me.TxtUserCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtUserCpta.Location = New System.Drawing.Point(185, 87)
        Me.TxtUserCpta.Name = "TxtUserCpta"
        Me.TxtUserCpta.Size = New System.Drawing.Size(307, 23)
        Me.TxtUserCpta.TabIndex = 1
        '
        'TxtBDCpta
        '
        Me.TxtBDCpta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtBDCpta.Location = New System.Drawing.Point(186, 145)
        Me.TxtBDCpta.Name = "TxtBDCpta"
        Me.TxtBDCpta.Size = New System.Drawing.Size(306, 23)
        Me.TxtBDCpta.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 61)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(157, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Chemin du Fichier Comptable"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(22, 119)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(138, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Mot de Passe  Comptable"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(22, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Utilisateur Comptable"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(22, 148)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(151, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Base de Données SQL Server"
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
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(31, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(541, 279)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Paramétres de Consolidation"
        '
        'BT_FicCpta
        '
        Me.BT_FicCpta.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.BT_FicCpta.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.BT_FicCpta.Location = New System.Drawing.Point(496, 58)
        Me.BT_FicCpta.Name = "BT_FicCpta"
        Me.BT_FicCpta.Size = New System.Drawing.Size(36, 21)
        Me.BT_FicCpta.TabIndex = 38
        Me.BT_FicCpta.ThemeName = "Windows8"
        '
        'CkConso
        '
        Me.CkConso.AutoSize = True
        Me.CkConso.Location = New System.Drawing.Point(235, 33)
        Me.CkConso.Name = "CkConso"
        Me.CkConso.Size = New System.Drawing.Size(151, 19)
        Me.CkConso.TabIndex = 0
        Me.CkConso.Text = "Tentative de Connexion"
        Me.CkConso.UseVisualStyleBackColor = True
        '
        'TxtPasw
        '
        Me.TxtPasw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtPasw.Location = New System.Drawing.Point(186, 232)
        Me.TxtPasw.Name = "TxtPasw"
        Me.TxtPasw.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtPasw.Size = New System.Drawing.Size(306, 23)
        Me.TxtPasw.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(22, 235)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(132, 13)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Mot de Passe SQL Server"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(22, 206)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(119, 13)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Utilisateur SQL Server "
        '
        'TxtUtilisateur
        '
        Me.TxtUtilisateur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtUtilisateur.Location = New System.Drawing.Point(186, 203)
        Me.TxtUtilisateur.Name = "TxtUtilisateur"
        Me.TxtUtilisateur.Size = New System.Drawing.Size(306, 23)
        Me.TxtUtilisateur.TabIndex = 5
        '
        'Txtsql
        '
        Me.Txtsql.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txtsql.Location = New System.Drawing.Point(186, 174)
        Me.Txtsql.Name = "Txtsql"
        Me.Txtsql.Size = New System.Drawing.Size(306, 23)
        Me.Txtsql.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI Semibold", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(22, 177)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "Serveur SQL"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(7, 54)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(151, 15)
        Me.Label15.TabIndex = 42
        Me.Label15.Text = "Chemin du Fichier  Access"
        '
        'TxtAccess
        '
        Me.TxtAccess.BackColor = System.Drawing.SystemColors.Window
        Me.TxtAccess.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtAccess.Location = New System.Drawing.Point(254, 54)
        Me.TxtAccess.Name = "TxtAccess"
        Me.TxtAccess.ReadOnly = True
        Me.TxtAccess.Size = New System.Drawing.Size(272, 21)
        Me.TxtAccess.TabIndex = 12
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(7, 116)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(197, 15)
        Me.Label11.TabIndex = 30
        Me.Label11.Text = "Répertoire de Sauvegarde Fichiers"
        '
        'Txt_Rep
        '
        Me.Txt_Rep.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txt_Rep.Location = New System.Drawing.Point(254, 116)
        Me.Txt_Rep.Name = "Txt_Rep"
        Me.Txt_Rep.Size = New System.Drawing.Size(273, 21)
        Me.Txt_Rep.TabIndex = 14
        '
        'TxtFilejr
        '
        Me.TxtFilejr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtFilejr.Location = New System.Drawing.Point(254, 85)
        Me.TxtFilejr.Name = "TxtFilejr"
        Me.TxtFilejr.Size = New System.Drawing.Size(273, 21)
        Me.TxtFilejr.TabIndex = 13
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(7, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(208, 15)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Répertoire de Journalisation Fichiers"
        '
        'FolderRepjournal
        '
        Me.FolderRepjournal.RootFolder = System.Environment.SpecialFolder.DesktopDirectory
        Me.FolderRepjournal.ShowNewFolderButton = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Khaki
        Me.GroupBox2.Controls.Add(Me.Bt_Article)
        Me.GroupBox2.Controls.Add(Me.Bt_tiers)
        Me.GroupBox2.Controls.Add(Me.BT_FicRep)
        Me.GroupBox2.Controls.Add(Me.BT_FicJournal)
        Me.GroupBox2.Controls.Add(Me.BT_Access)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Txtiers)
        Me.GroupBox2.Controls.Add(Me.TxtArticle)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.TxtAccess)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.TxtFilejr)
        Me.GroupBox2.Controls.Add(Me.Txt_Rep)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(15, 48)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(573, 239)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Autres Parametres"
        '
        'Bt_Article
        '
        Me.Bt_Article.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.Bt_Article.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.Bt_Article.Location = New System.Drawing.Point(531, 178)
        Me.Bt_Article.Name = "Bt_Article"
        Me.Bt_Article.Size = New System.Drawing.Size(36, 21)
        Me.Bt_Article.TabIndex = 64
        Me.Bt_Article.ThemeName = "Windows8"
        '
        'Bt_tiers
        '
        Me.Bt_tiers.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.Bt_tiers.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.Bt_tiers.Location = New System.Drawing.Point(531, 147)
        Me.Bt_tiers.Name = "Bt_tiers"
        Me.Bt_tiers.Size = New System.Drawing.Size(36, 21)
        Me.Bt_tiers.TabIndex = 63
        Me.Bt_tiers.ThemeName = "Windows8"
        '
        'BT_FicRep
        '
        Me.BT_FicRep.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.BT_FicRep.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.BT_FicRep.Location = New System.Drawing.Point(531, 116)
        Me.BT_FicRep.Name = "BT_FicRep"
        Me.BT_FicRep.Size = New System.Drawing.Size(36, 21)
        Me.BT_FicRep.TabIndex = 62
        Me.BT_FicRep.ThemeName = "Windows8"
        '
        'BT_FicJournal
        '
        Me.BT_FicJournal.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.BT_FicJournal.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.BT_FicJournal.Location = New System.Drawing.Point(531, 85)
        Me.BT_FicJournal.Name = "BT_FicJournal"
        Me.BT_FicJournal.Size = New System.Drawing.Size(36, 21)
        Me.BT_FicJournal.TabIndex = 61
        Me.BT_FicJournal.ThemeName = "Windows8"
        '
        'BT_Access
        '
        Me.BT_Access.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.BT_Access.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.BT_Access.Location = New System.Drawing.Point(531, 54)
        Me.BT_Access.Name = "BT_Access"
        Me.BT_Access.Size = New System.Drawing.Size(36, 21)
        Me.BT_Access.TabIndex = 60
        Me.BT_Access.ThemeName = "Windows8"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 147)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(210, 15)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Répertoire des fichiers extraire (ERP)"
        '
        'Txtiers
        '
        Me.Txtiers.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Txtiers.Location = New System.Drawing.Point(254, 147)
        Me.Txtiers.Name = "Txtiers"
        Me.Txtiers.Size = New System.Drawing.Size(273, 21)
        Me.Txtiers.TabIndex = 15
        '
        'TxtArticle
        '
        Me.TxtArticle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtArticle.Location = New System.Drawing.Point(254, 178)
        Me.TxtArticle.Name = "TxtArticle"
        Me.TxtArticle.Size = New System.Drawing.Size(273, 21)
        Me.TxtArticle.TabIndex = 16
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 178)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(248, 15)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Répertoire des fichiers à intégre (EasyWMS)"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(610, 361)
        Me.TabControl1.TabIndex = 61
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(602, 335)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Paramètre de Consolidation"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(602, 335)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Fichier Access/Journalisation"
        '
        'Button2
        '
        Me.Button2.Image = Global.Mecalux_Application.My.Resources.Resources.criticalind_status1
        Me.Button2.Location = New System.Drawing.Point(477, 379)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(110, 24)
        Me.Button2.TabIndex = 63
        Me.Button2.Text = "&Quitter"
        Me.Button2.ThemeName = "Windows8"
        '
        'Button1
        '
        Me.Button1.Image = Global.Mecalux_Application.My.Resources.Resources.btn_valider
        Me.Button1.Location = New System.Drawing.Point(359, 379)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(110, 24)
        Me.Button1.TabIndex = 62
        Me.Button1.Text = "&Enregistrer"
        Me.Button1.ThemeName = "Windows8"
        '
        'Frm_FichierConfiguration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(634, 409)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Frm_FichierConfiguration"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fichier de Configuration"
        Me.ThemeName = "TelerikMetro"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.BT_FicCpta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.Bt_Article, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bt_tiers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_FicRep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_FicJournal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Access, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.Button2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Button1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TxtFichierCpta As System.Windows.Forms.TextBox
    Friend WithEvents TxtPasword As System.Windows.Forms.TextBox
    Friend WithEvents TxtUserCpta As System.Windows.Forms.TextBox
    Friend WithEvents TxtBDCpta As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents OpenFileFicCpta As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Txtsql As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilejr As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtPasw As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtUtilisateur As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Txt_Rep As System.Windows.Forms.TextBox
    Friend WithEvents FolderRepjournal As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderRepsaving As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileGesCom As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FolderRepFact As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderRepSave As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtAccess As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileAccess As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Txtiers As System.Windows.Forms.TextBox
    Friend WithEvents TxtArticle As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CkConso As System.Windows.Forms.CheckBox
    Friend WithEvents OpenProgExterne As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As Telerik.WinControls.UI.RadButton
    Friend WithEvents Button2 As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_FicCpta As Telerik.WinControls.UI.RadButton
    Friend WithEvents Bt_Article As Telerik.WinControls.UI.RadButton
    Friend WithEvents Bt_tiers As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_FicRep As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_FicJournal As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_Access As Telerik.WinControls.UI.RadButton
   
   Friend WithEvents Windows8Theme1 As Telerik.WinControls.Themes.Windows8Theme
    Friend WithEvents TelerikMetroTheme1 As Telerik.WinControls.Themes.TelerikMetroTheme
    Friend WithEvents TelerikMetroTheme2 As Telerik.WinControls.Themes.TelerikMetroTheme
End Class
