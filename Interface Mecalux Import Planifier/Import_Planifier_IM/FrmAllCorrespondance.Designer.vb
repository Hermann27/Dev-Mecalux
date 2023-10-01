<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAllCorrespondance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAllCorrespondance))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.Cols = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Format = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PositionG = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DefaultValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChampSage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.InfosLibre = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CodeTbls = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Entete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Ligne = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Modifier = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Suppression = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Button3 = New System.Windows.Forms.Button
        Me.BtnSup = New System.Windows.Forms.Button
        Me.BtnModif = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataListeIntegrer)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnSup)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnModif)
        Me.SplitContainer1.Size = New System.Drawing.Size(1135, 517)
        Me.SplitContainer1.SplitterDistance = 456
        Me.SplitContainer1.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.SlateGray
        Me.PictureBox1.Image = Global.Import_Planifier_IM.My.Resources.Resources._44
        Me.PictureBox1.Location = New System.Drawing.Point(143, 48)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(849, 361)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataListeIntegrer.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cols, Me.description, Me.Format, Me.PositionG, Me.DefaultValue, Me.ChampSage, Me.InfosLibre, Me.CodeTbls, Me.Entete, Me.Ligne, Me.Modifier, Me.Suppression})
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.EnableHeadersVisualStyles = False
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        Me.DataListeIntegrer.Size = New System.Drawing.Size(1135, 456)
        Me.DataListeIntegrer.TabIndex = 3
        '
        'Cols
        '
        Me.Cols.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.Cols.HeaderText = "Colonne EasyWMS"
        Me.Cols.Name = "Cols"
        Me.Cols.Width = 124
        '
        'description
        '
        Me.description.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.description.HeaderText = "Description de la colonne"
        Me.description.Name = "description"
        Me.description.Width = 152
        '
        'Format
        '
        Me.Format.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.Format.HeaderText = "Format du Champ"
        Me.Format.Name = "Format"
        Me.Format.Width = 115
        '
        'PositionG
        '
        Me.PositionG.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.PositionG.HeaderText = "Position de Gauche"
        Me.PositionG.Name = "PositionG"
        Me.PositionG.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PositionG.Width = 125
        '
        'DefaultValue
        '
        Me.DefaultValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.DefaultValue.HeaderText = "Valeur par defaut"
        Me.DefaultValue.Name = "DefaultValue"
        Me.DefaultValue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DefaultValue.Width = 113
        '
        'ChampSage
        '
        Me.ChampSage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.ChampSage.HeaderText = "Colonne Sage"
        Me.ChampSage.Name = "ChampSage"
        Me.ChampSage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ChampSage.Width = 80
        '
        'InfosLibre
        '
        Me.InfosLibre.HeaderText = "Infos Libre"
        Me.InfosLibre.Name = "InfosLibre"
        '
        'CodeTbls
        '
        Me.CodeTbls.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.CodeTbls.HeaderText = "Code Table"
        Me.CodeTbls.Name = "CodeTbls"
        Me.CodeTbls.Width = 87
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
        'Modifier
        '
        Me.Modifier.HeaderText = "     Modifier"
        Me.Modifier.Name = "Modifier"
        '
        'Suppression
        '
        Me.Suppression.HeaderText = "Suppression"
        Me.Suppression.Name = "Suppression"
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Image = Global.Import_Planifier_IM.My.Resources.Resources.btSupprimer221
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(981, 15)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(110, 30)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Quitter"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BtnSup
        '
        Me.BtnSup.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSup.Image = Global.Import_Planifier_IM.My.Resources.Resources.btFermer22
        Me.BtnSup.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSup.Location = New System.Drawing.Point(865, 15)
        Me.BtnSup.Name = "BtnSup"
        Me.BtnSup.Size = New System.Drawing.Size(110, 30)
        Me.BtnSup.TabIndex = 1
        Me.BtnSup.Text = "Supprimer"
        Me.BtnSup.UseVisualStyleBackColor = True
        '
        'BtnModif
        '
        Me.BtnModif.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnModif.Image = Global.Import_Planifier_IM.My.Resources.Resources.AnalyzeWizard1
        Me.BtnModif.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModif.Location = New System.Drawing.Point(739, 15)
        Me.BtnModif.Name = "BtnModif"
        Me.BtnModif.Size = New System.Drawing.Size(110, 30)
        Me.BtnModif.TabIndex = 0
        Me.BtnModif.Text = "Modifier"
        Me.BtnModif.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmAllCorrespondance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1135, 517)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAllCorrespondance"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Formulaire de Correspondance"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BtnSup As System.Windows.Forms.Button
    Friend WithEvents BtnModif As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Cols As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Format As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PositionG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DefaultValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChampSage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InfosLibre As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CodeTbls As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Entete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Ligne As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Modifier As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Suppression As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
