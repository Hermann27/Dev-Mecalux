<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_ToPosting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_ToPosting))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DataListeIntegrer = New System.Windows.Forms.DataGridView
        Me.CodeFo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleFo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ArticleDis1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button4 = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataListeIntegrer)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button4)
        Me.SplitContainer1.Size = New System.Drawing.Size(576, 90)
        Me.SplitContainer1.SplitterDistance = 58
        Me.SplitContainer1.TabIndex = 0
        '
        'DataListeIntegrer
        '
        Me.DataListeIntegrer.AllowUserToAddRows = False
        Me.DataListeIntegrer.AllowUserToDeleteRows = False
        Me.DataListeIntegrer.AllowUserToOrderColumns = True
        Me.DataListeIntegrer.AllowUserToResizeRows = False
        Me.DataListeIntegrer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataListeIntegrer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataListeIntegrer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeFo1, Me.ArticleFo1, Me.ArticleDis1})
        Me.DataListeIntegrer.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataListeIntegrer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataListeIntegrer.Location = New System.Drawing.Point(0, 0)
        Me.DataListeIntegrer.MultiSelect = False
        Me.DataListeIntegrer.Name = "DataListeIntegrer"
        Me.DataListeIntegrer.RowHeadersVisible = False
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataListeIntegrer.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.DataListeIntegrer.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataListeIntegrer.Size = New System.Drawing.Size(576, 58)
        Me.DataListeIntegrer.TabIndex = 47
        '
        'CodeFo1
        '
        Me.CodeFo1.HeaderText = "Fournisseur"
        Me.CodeFo1.Name = "CodeFo1"
        Me.CodeFo1.ReadOnly = True
        Me.CodeFo1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ArticleFo1
        '
        Me.ArticleFo1.FillWeight = 90.90909!
        Me.ArticleFo1.HeaderText = "Code EAN"
        Me.ArticleFo1.Name = "ArticleFo1"
        Me.ArticleFo1.ReadOnly = True
        '
        'ArticleDis1
        '
        Me.ArticleDis1.FillWeight = 90.90909!
        Me.ArticleDis1.HeaderText = "Code Article DistriServices"
        Me.ArticleDis1.Name = "ArticleDis1"
        '
        'Button4
        '
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button4.Location = New System.Drawing.Point(253, 3)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(83, 23)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "&Modifier"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Frm_ToPosting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(576, 90)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_ToPosting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Modification Article"
        Me.TopMost = True
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataListeIntegrer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataListeIntegrer As System.Windows.Forms.DataGridView
    Friend WithEvents CodeFo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleFo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArticleDis1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button4 As System.Windows.Forms.Button
End Class
