<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_OptionGrid
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DGVCE = New System.Windows.Forms.DataGridView
        Me.C1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.P1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnSup = New System.Windows.Forms.Button
        Me.BtnAjouter = New System.Windows.Forms.Button
        Me.DGVCV = New System.Windows.Forms.DataGridView
        Me.C11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.P11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DGVCE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DGVCV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BackColor = System.Drawing.Color.AliceBlue
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DGVCE)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(631, 289)
        Me.SplitContainer1.SplitterDistance = 280
        Me.SplitContainer1.TabIndex = 0
        '
        'DGVCE
        '
        Me.DGVCE.AllowUserToAddRows = False
        Me.DGVCE.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DGVCE.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVCE.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DGVCE.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.C1, Me.P1})
        Me.DGVCE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVCE.EnableHeadersVisualStyles = False
        Me.DGVCE.Location = New System.Drawing.Point(0, 0)
        Me.DGVCE.Name = "DGVCE"
        Me.DGVCE.RowHeadersVisible = False
        Me.DGVCE.Size = New System.Drawing.Size(280, 289)
        Me.DGVCE.TabIndex = 3
        '
        'C1
        '
        Me.C1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.C1.HeaderText = "                              Colonnes Disponibles"
        Me.C1.Name = "C1"
        '
        'P1
        '
        Me.P1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.P1.HeaderText = "Position"
        Me.P1.Name = "P1"
        Me.P1.Visible = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DGVCV)
        Me.SplitContainer2.Size = New System.Drawing.Size(347, 289)
        Me.SplitContainer2.SplitterDistance = 86
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnSup)
        Me.GroupBox1.Controls.Add(Me.BtnAjouter)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(86, 289)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtnSup
        '
        Me.BtnSup.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSup.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowback_16
        Me.BtnSup.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSup.Location = New System.Drawing.Point(3, 133)
        Me.BtnSup.Name = "BtnSup"
        Me.BtnSup.Size = New System.Drawing.Size(80, 23)
        Me.BtnSup.TabIndex = 1
        Me.BtnSup.Text = "   Supprimer"
        Me.BtnSup.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSup.UseVisualStyleBackColor = True
        '
        'BtnAjouter
        '
        Me.BtnAjouter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAjouter.Image = Global.Import_Planifier_IM.My.Resources.Resources.arrowforward_161
        Me.BtnAjouter.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAjouter.Location = New System.Drawing.Point(3, 90)
        Me.BtnAjouter.Name = "BtnAjouter"
        Me.BtnAjouter.Size = New System.Drawing.Size(80, 23)
        Me.BtnAjouter.TabIndex = 0
        Me.BtnAjouter.Text = "Ajouter "
        Me.BtnAjouter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAjouter.UseVisualStyleBackColor = True
        '
        'DGVCV
        '
        Me.DGVCV.AllowUserToAddRows = False
        Me.DGVCV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DGVCV.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVCV.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DGVCV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.C11, Me.P11})
        Me.DGVCV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVCV.EnableHeadersVisualStyles = False
        Me.DGVCV.Location = New System.Drawing.Point(0, 0)
        Me.DGVCV.Name = "DGVCV"
        Me.DGVCV.RowHeadersVisible = False
        Me.DGVCV.Size = New System.Drawing.Size(257, 289)
        Me.DGVCV.TabIndex = 4
        '
        'C11
        '
        Me.C11.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.C11.HeaderText = "                            Colonnes Sélectionnées"
        Me.C11.Name = "C11"
        '
        'P11
        '
        Me.P11.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.P11.HeaderText = "Position"
        Me.P11.Name = "P11"
        Me.P11.Visible = False
        '
        'Frm_OptionGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(631, 289)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "Frm_OptionGrid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Formulaire d'Option sur la Grille"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DGVCE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DGVCV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DGVCE As System.Windows.Forms.DataGridView
    Friend WithEvents C1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents P1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DGVCV As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnSup As System.Windows.Forms.Button
    Friend WithEvents BtnAjouter As System.Windows.Forms.Button
    Friend WithEvents C11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents P11 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
