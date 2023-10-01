<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmChoixTraitement
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmChoixTraitement))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RbtTdepot = New System.Windows.Forms.RadioButton
        Me.RbtFC = New System.Windows.Forms.RadioButton
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.GhostWhite
        Me.GroupBox1.Controls.Add(Me.RbtTdepot)
        Me.GroupBox1.Controls.Add(Me.RbtFC)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(368, 58)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Type de Transformation"
        '
        'RbtTdepot
        '
        Me.RbtTdepot.AutoSize = True
        Me.RbtTdepot.Location = New System.Drawing.Point(221, 29)
        Me.RbtTdepot.Name = "RbtTdepot"
        Me.RbtTdepot.Size = New System.Drawing.Size(135, 17)
        Me.RbtTdepot.TabIndex = 9
        Me.RbtTdepot.Text = "Confirmation Réception"
        Me.RbtTdepot.UseVisualStyleBackColor = True
        '
        'RbtFC
        '
        Me.RbtFC.AutoSize = True
        Me.RbtFC.Location = New System.Drawing.Point(52, 29)
        Me.RbtFC.Name = "RbtFC"
        Me.RbtFC.Size = New System.Drawing.Size(139, 17)
        Me.RbtFC.TabIndex = 8
        Me.RbtFC.Text = "Confirmation Commande"
        Me.RbtFC.UseVisualStyleBackColor = True
        '
        'FrmChoixTraitement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(368, 58)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmChoixTraitement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Option Traitement Transformation"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RbtTdepot As System.Windows.Forms.RadioButton
    Friend WithEvents RbtFC As System.Windows.Forms.RadioButton
End Class
