<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCorrespondance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCorrespondance))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BT_FicCpta = New System.Windows.Forms.Button
        Me.ComboType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txttableSage = New System.Windows.Forms.TextBox
        Me.txtintulé = New System.Windows.Forms.TextBox
        Me.txtCde = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DataListeSchema = New System.Windows.Forms.DataGridView
        Me.Cols = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Format = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PositionG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DefaultValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.InfosLibre = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ChampSage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Entete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Ligne = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BtnQuitter = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1093, 500)
        Me.SplitContainer1.SplitterDistance = 72
        Me.SplitContainer1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.ComboType)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txttableSage)
        Me.GroupBox1.Controls.Add(Me.txtintulé)
        Me.GroupBox1.Controls.Add(Me.txtCde)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1093, 72)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Traitement Table de Correspondance"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.BtnDelete)
        Me.Panel1.Controls.Add(Me.BT_FicCpta)
        Me.Panel1.Location = New System.Drawing.Point(1013, 48)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(80, 25)
        Me.Panel1.TabIndex = 8
        '
        'BtnDelete
        '
        Me.BtnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnDelete.Image = Global.Import_Planifier_IM.My.Resources.Resources.criticalind_status
        Me.BtnDelete.Location = New System.Drawing.Point(41, 1)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(33, 23)
        Me.BtnDelete.TabIndex = 1
        Me.BtnDelete.Text = "-"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BT_FicCpta
        '
        Me.BT_FicCpta.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BT_FicCpta.Image = Global.Import_Planifier_IM.My.Resources.Resources.create
        Me.BT_FicCpta.Location = New System.Drawing.Point(3, 1)
        Me.BT_FicCpta.Name = "BT_FicCpta"
        Me.BT_FicCpta.Size = New System.Drawing.Size(33, 23)
        Me.BT_FicCpta.TabIndex = 0
        Me.BT_FicCpta.Text = "+"
        Me.BT_FicCpta.UseVisualStyleBackColor = True
        '
        'ComboType
        '
        Me.ComboType.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboType.FormattingEnabled = True
        Me.ComboType.Items.AddRange(New Object() {"IMPORT", "EXPORT"})
        Me.ComboType.Location = New System.Drawing.Point(873, 21)
        Me.ComboType.Name = "ComboType"
        Me.ComboType.Size = New System.Drawing.Size(162, 21)
        Me.ComboType.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(777, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Type d'echange :"
        '
        'txttableSage
        '
        Me.txttableSage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttableSage.Location = New System.Drawing.Point(635, 22)
        Me.txttableSage.Name = "txttableSage"
        Me.txttableSage.Size = New System.Drawing.Size(126, 20)
        Me.txttableSage.TabIndex = 5
        '
        'txtintulé
        '
        Me.txtintulé.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtintulé.Location = New System.Drawing.Point(299, 22)
        Me.txtintulé.Name = "txtintulé"
        Me.txtintulé.Size = New System.Drawing.Size(255, 20)
        Me.txtintulé.TabIndex = 4
        '
        'txtCde
        '
        Me.txtCde.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCde.Location = New System.Drawing.Point(79, 22)
        Me.txtCde.Name = "txtCde"
        Me.txtCde.Size = New System.Drawing.Size(123, 20)
        Me.txtCde.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(570, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Table Sage :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(219, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Intitulé Table :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code Table :"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.DataListeSchema)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnQuitter)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnSave)
        Me.SplitContainer2.Size = New System.Drawing.Size(1093, 424)
        Me.SplitContainer2.SplitterDistance = 371
        Me.SplitContainer2.TabIndex = 0
        '
        'DataListeSchema
        '
        Me.DataListeSchema.AllowUserToAddRows = False
        Me.DataListeSchema.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeSchema.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataListeSchema.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DataListeSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cols, Me.Desc, Me.Format, Me.PositionG, Me.DefaultValue, Me.InfosLibre, Me.ChampSage, Me.Entete, Me.Ligne})
        Me.DataListeSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeSchema.EnableHeadersVisualStyles = False
        Me.DataListeSchema.Location = New System.Drawing.Point(0, 0)
        Me.DataListeSchema.Name = "DataListeSchema"
        Me.DataListeSchema.RowHeadersVisible = False
        Me.DataListeSchema.Size = New System.Drawing.Size(1093, 371)
        Me.DataListeSchema.TabIndex = 2
        '
        'Cols
        '
        Me.Cols.HeaderText = "Colonne EasyWMS"
        Me.Cols.Name = "Cols"
        '
        'Desc
        '
        Me.Desc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.Desc.HeaderText = "Description de la colonne"
        Me.Desc.Name = "Desc"
        Me.Desc.Width = 152
        '
        'Format
        '
        Me.Format.HeaderText = "Format du Champ"
        Me.Format.Name = "Format"
        '
        'PositionG
        '
        Me.PositionG.HeaderText = "Position de Gauche"
        Me.PositionG.Name = "PositionG"
        '
        'DefaultValue
        '
        Me.DefaultValue.HeaderText = "Valeur par defaut"
        Me.DefaultValue.Name = "DefaultValue"
        '
        'InfosLibre
        '
        Me.InfosLibre.HeaderText = "Infos Libre"
        Me.InfosLibre.Name = "InfosLibre"
        '
        'ChampSage
        '
        Me.ChampSage.HeaderText = "Colonne Sage"
        Me.ChampSage.Name = "ChampSage"
        '
        'Entete
        '
        Me.Entete.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Entete.HeaderText = " Entete ?"
        Me.Entete.Name = "Entete"
        Me.Entete.Width = 56
        '
        'Ligne
        '
        Me.Ligne.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Ligne.HeaderText = "Ligne ?"
        Me.Ligne.Name = "Ligne"
        Me.Ligne.Width = 48
        '
        'BtnQuitter
        '
        Me.BtnQuitter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnQuitter.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.BtnQuitter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQuitter.Location = New System.Drawing.Point(962, 12)
        Me.BtnQuitter.Name = "BtnQuitter"
        Me.BtnQuitter.Size = New System.Drawing.Size(110, 23)
        Me.BtnQuitter.TabIndex = 1
        Me.BtnQuitter.Text = "Quitter"
        Me.BtnQuitter.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSave.Image = Global.Import_Planifier_IM.My.Resources.Resources.save_161
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(808, 12)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(110, 23)
        Me.BtnSave.TabIndex = 0
        Me.BtnSave.Text = "Enregistrer"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'FrmCorrespondance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1093, 500)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCorrespondance"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Création des correspondances"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataListeSchema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtintulé As System.Windows.Forms.TextBox
    Friend WithEvents txtCde As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BT_FicCpta As System.Windows.Forms.Button
    Friend WithEvents ComboType As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txttableSage As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataListeSchema As System.Windows.Forms.DataGridView
    Friend WithEvents Cols As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Format As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PositionG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DefaultValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InfosLibre As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ChampSage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Entete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Ligne As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BtnQuitter As System.Windows.Forms.Button
    Friend WithEvents BtnSave As System.Windows.Forms.Button
End Class
