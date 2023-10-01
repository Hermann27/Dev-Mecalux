<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmChoix
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RtbPseudo = New System.Windows.Forms.RadioButton
        Me.Button1 = New System.Windows.Forms.Button
        Me.RbtTdepot = New System.Windows.Forms.RadioButton
        Me.RbtMvt = New System.Windows.Forms.RadioButton
        Me.RbtFC = New System.Windows.Forms.RadioButton
        Me.RbtCF = New System.Windows.Forms.RadioButton
        Me.RbtCClt = New System.Windows.Forms.RadioButton
        Me.RbtFrss = New System.Windows.Forms.RadioButton
        Me.Rbtclt = New System.Windows.Forms.RadioButton
        Me.RbtArt = New System.Windows.Forms.RadioButton
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.AliceBlue
        Me.GroupBox1.Controls.Add(Me.RtbPseudo)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.RbtTdepot)
        Me.GroupBox1.Controls.Add(Me.RbtMvt)
        Me.GroupBox1.Controls.Add(Me.RbtFC)
        Me.GroupBox1.Controls.Add(Me.RbtCF)
        Me.GroupBox1.Controls.Add(Me.RbtCClt)
        Me.GroupBox1.Controls.Add(Me.RbtFrss)
        Me.GroupBox1.Controls.Add(Me.Rbtclt)
        Me.GroupBox1.Controls.Add(Me.RbtArt)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(192, 327)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Type de Correspondance"
        '
        'RtbPseudo
        '
        Me.RtbPseudo.AutoSize = True
        Me.RtbPseudo.Location = New System.Drawing.Point(27, 94)
        Me.RtbPseudo.Name = "RtbPseudo"
        Me.RtbPseudo.Size = New System.Drawing.Size(90, 17)
        Me.RtbPseudo.TabIndex = 8
        Me.RtbPseudo.Text = "Fiche Pseudo"
        Me.RtbPseudo.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(204, 354)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(10, 17)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RbtTdepot
        '
        Me.RbtTdepot.AutoSize = True
        Me.RbtTdepot.Location = New System.Drawing.Point(27, 292)
        Me.RbtTdepot.Name = "RbtTdepot"
        Me.RbtTdepot.Size = New System.Drawing.Size(135, 17)
        Me.RbtTdepot.TabIndex = 7
        Me.RbtTdepot.Text = "Confirmation Réception"
        Me.RbtTdepot.UseVisualStyleBackColor = True
        '
        'RbtMvt
        '
        Me.RbtMvt.AutoSize = True
        Me.RbtMvt.Location = New System.Drawing.Point(27, 259)
        Me.RbtMvt.Name = "RbtMvt"
        Me.RbtMvt.Size = New System.Drawing.Size(103, 17)
        Me.RbtMvt.TabIndex = 6
        Me.RbtMvt.Text = "Mouvement E/S"
        Me.RbtMvt.UseVisualStyleBackColor = True
        '
        'RbtFC
        '
        Me.RbtFC.AutoSize = True
        Me.RbtFC.Location = New System.Drawing.Point(27, 226)
        Me.RbtFC.Name = "RbtFC"
        Me.RbtFC.Size = New System.Drawing.Size(139, 17)
        Me.RbtFC.TabIndex = 5
        Me.RbtFC.Text = "Confirmation Commande"
        Me.RbtFC.UseVisualStyleBackColor = True
        '
        'RbtCF
        '
        Me.RbtCF.AutoSize = True
        Me.RbtCF.Location = New System.Drawing.Point(27, 193)
        Me.RbtCF.Name = "RbtCF"
        Me.RbtCF.Size = New System.Drawing.Size(135, 17)
        Me.RbtCF.TabIndex = 4
        Me.RbtCF.Text = "Commande Fournisseur"
        Me.RbtCF.UseVisualStyleBackColor = True
        '
        'RbtCClt
        '
        Me.RbtCClt.AutoSize = True
        Me.RbtCClt.Location = New System.Drawing.Point(27, 160)
        Me.RbtCClt.Name = "RbtCClt"
        Me.RbtCClt.Size = New System.Drawing.Size(107, 17)
        Me.RbtCClt.TabIndex = 3
        Me.RbtCClt.Text = "Commande Client"
        Me.RbtCClt.UseVisualStyleBackColor = True
        '
        'RbtFrss
        '
        Me.RbtFrss.AutoSize = True
        Me.RbtFrss.Location = New System.Drawing.Point(27, 127)
        Me.RbtFrss.Name = "RbtFrss"
        Me.RbtFrss.Size = New System.Drawing.Size(108, 17)
        Me.RbtFrss.TabIndex = 2
        Me.RbtFrss.Text = "Fiche Fournisseur"
        Me.RbtFrss.UseVisualStyleBackColor = True
        '
        'Rbtclt
        '
        Me.Rbtclt.AutoSize = True
        Me.Rbtclt.Location = New System.Drawing.Point(27, 61)
        Me.Rbtclt.Name = "Rbtclt"
        Me.Rbtclt.Size = New System.Drawing.Size(80, 17)
        Me.Rbtclt.TabIndex = 1
        Me.Rbtclt.Text = "Fiche Client"
        Me.Rbtclt.UseVisualStyleBackColor = True
        '
        'RbtArt
        '
        Me.RbtArt.AutoSize = True
        Me.RbtArt.Location = New System.Drawing.Point(27, 28)
        Me.RbtArt.Name = "RbtArt"
        Me.RbtArt.Size = New System.Drawing.Size(83, 17)
        Me.RbtArt.TabIndex = 1
        Me.RbtArt.Text = "Fiche Article"
        Me.RbtArt.UseVisualStyleBackColor = True
        '
        'FrmChoix
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(192, 327)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmChoix"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Choix "
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RbtTdepot As System.Windows.Forms.RadioButton
    Friend WithEvents RbtMvt As System.Windows.Forms.RadioButton
    Friend WithEvents RbtFC As System.Windows.Forms.RadioButton
    Friend WithEvents RbtCF As System.Windows.Forms.RadioButton
    Friend WithEvents RbtCClt As System.Windows.Forms.RadioButton
    Friend WithEvents RbtFrss As System.Windows.Forms.RadioButton
    Friend WithEvents Rbtclt As System.Windows.Forms.RadioButton
    Friend WithEvents RbtArt As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RtbPseudo As System.Windows.Forms.RadioButton
End Class
