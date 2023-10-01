<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RechercheCriteredocument
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RechercheCriteredocument))
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtFichier = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CbCritere = New System.Windows.Forms.ComboBox
        Me.TxtSage = New System.Windows.Forms.TextBox
        Me.TxtLibelle = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.BT_Update = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(16, 110)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Critère"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 84)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(38, 13)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Intitulé"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Libellé"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtFichier)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CbCritere)
        Me.GroupBox1.Controls.Add(Me.TxtSage)
        Me.GroupBox1.Controls.Add(Me.TxtLibelle)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Location = New System.Drawing.Point(2, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(376, 166)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Regarder dans"
        '
        'TxtFichier
        '
        Me.TxtFichier.Enabled = False
        Me.TxtFichier.Location = New System.Drawing.Point(125, 29)
        Me.TxtFichier.Name = "TxtFichier"
        Me.TxtFichier.Size = New System.Drawing.Size(217, 20)
        Me.TxtFichier.TabIndex = 35
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Fichier"
        '
        'CbCritere
        '
        Me.CbCritere.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CbCritere.FormattingEnabled = True
        Me.CbCritere.Location = New System.Drawing.Point(125, 107)
        Me.CbCritere.Name = "CbCritere"
        Me.CbCritere.Size = New System.Drawing.Size(217, 23)
        Me.CbCritere.TabIndex = 28
        '
        'TxtSage
        '
        Me.TxtSage.Enabled = False
        Me.TxtSage.Location = New System.Drawing.Point(125, 81)
        Me.TxtSage.Name = "TxtSage"
        Me.TxtSage.Size = New System.Drawing.Size(217, 20)
        Me.TxtSage.TabIndex = 27
        '
        'TxtLibelle
        '
        Me.TxtLibelle.Enabled = False
        Me.TxtLibelle.Location = New System.Drawing.Point(125, 55)
        Me.TxtLibelle.Name = "TxtLibelle"
        Me.TxtLibelle.Size = New System.Drawing.Size(217, 20)
        Me.TxtLibelle.TabIndex = 26
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(114, 170)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 21)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "&Quitter"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BT_Update
        '
        Me.BT_Update.Location = New System.Drawing.Point(208, 170)
        Me.BT_Update.Name = "BT_Update"
        Me.BT_Update.Size = New System.Drawing.Size(62, 21)
        Me.BT_Update.TabIndex = 27
        Me.BT_Update.Text = "&Modifier"
        Me.BT_Update.UseVisualStyleBackColor = True
        '
        'RechercheCriteredocument
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(398, 196)
        Me.Controls.Add(Me.BT_Update)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "RechercheCriteredocument"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Renseigner un champ de critère"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BT_Update As System.Windows.Forms.Button
    Friend WithEvents CbCritere As System.Windows.Forms.ComboBox
    Friend WithEvents TxtSage As System.Windows.Forms.TextBox
    Friend WithEvents TxtLibelle As System.Windows.Forms.TextBox
    Friend WithEvents TxtFichier As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
