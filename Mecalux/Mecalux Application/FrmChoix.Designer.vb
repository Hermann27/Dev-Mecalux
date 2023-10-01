<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmChoix
    Inherits Telerik.WinControls.UI.RadForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Windows8Theme1 = New Telerik.WinControls.Themes.Windows8Theme
        Me.TelerikMetroTheme1 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox
        Me.RbtMvt = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtTdepot = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtFfrss = New Telerik.WinControls.UI.RadRadioButton
        Me.Rbtclt = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtFrss = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtFC = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtCClt = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtCF = New Telerik.WinControls.UI.RadRadioButton
        Me.RbtArt = New Telerik.WinControls.UI.RadRadioButton
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.RbtMvt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtTdepot, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtFfrss, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Rbtclt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtFrss, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtFC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtCClt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtCF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RbtArt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Controls.Add(Me.RbtMvt)
        Me.RadGroupBox1.Controls.Add(Me.RbtTdepot)
        Me.RadGroupBox1.Controls.Add(Me.RbtFfrss)
        Me.RadGroupBox1.Controls.Add(Me.Rbtclt)
        Me.RadGroupBox1.Controls.Add(Me.RbtFrss)
        Me.RadGroupBox1.Controls.Add(Me.RbtFC)
        Me.RadGroupBox1.Controls.Add(Me.RbtCClt)
        Me.RadGroupBox1.Controls.Add(Me.RbtCF)
        Me.RadGroupBox1.Controls.Add(Me.RbtArt)
        Me.RadGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadGroupBox1.HeaderText = ""
        Me.RadGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Size = New System.Drawing.Size(200, 364)
        Me.RadGroupBox1.TabIndex = 0
        Me.RadGroupBox1.ThemeName = "Windows7"
        '
        'RbtMvt
        '
        Me.RbtMvt.Location = New System.Drawing.Point(31, 103)
        Me.RbtMvt.Name = "RbtMvt"
        Me.RbtMvt.Size = New System.Drawing.Size(100, 18)
        Me.RbtMvt.TabIndex = 6
        Me.RbtMvt.Text = "Mouvement E/S"
        '
        'RbtTdepot
        '
        Me.RbtTdepot.Location = New System.Drawing.Point(31, 278)
        Me.RbtTdepot.Name = "RbtTdepot"
        Me.RbtTdepot.Size = New System.Drawing.Size(118, 18)
        Me.RbtTdepot.TabIndex = 5
        Me.RbtTdepot.Text = "Transfert de dépôts"
        '
        'RbtFfrss
        '
        Me.RbtFfrss.Location = New System.Drawing.Point(31, 243)
        Me.RbtFfrss.Name = "RbtFfrss"
        Me.RbtFfrss.Size = New System.Drawing.Size(132, 18)
        Me.RbtFfrss.TabIndex = 4
        Me.RbtFfrss.Text = "BL/Facture fournisseur"
        '
        'Rbtclt
        '
        Me.Rbtclt.Location = New System.Drawing.Point(31, 68)
        Me.Rbtclt.Name = "Rbtclt"
        Me.Rbtclt.Size = New System.Drawing.Size(78, 18)
        Me.Rbtclt.TabIndex = 3
        Me.Rbtclt.Text = "Fiche Client"
        '
        'RbtFrss
        '
        Me.RbtFrss.Location = New System.Drawing.Point(31, 173)
        Me.RbtFrss.Name = "RbtFrss"
        Me.RbtFrss.Size = New System.Drawing.Size(107, 18)
        Me.RbtFrss.TabIndex = 2
        Me.RbtFrss.Text = "Fiche Fournisseur"
        '
        'RbtFC
        '
        Me.RbtFC.Location = New System.Drawing.Point(31, 138)
        Me.RbtFC.Name = "RbtFC"
        Me.RbtFC.Size = New System.Drawing.Size(103, 18)
        Me.RbtFC.TabIndex = 1
        Me.RbtFC.Text = "BL/Facture client"
        '
        'RbtCClt
        '
        Me.RbtCClt.Location = New System.Drawing.Point(31, 208)
        Me.RbtCClt.Name = "RbtCClt"
        Me.RbtCClt.Size = New System.Drawing.Size(110, 18)
        Me.RbtCClt.TabIndex = 1
        Me.RbtCClt.Text = "Commande Client"
        '
        'RbtCF
        '
        Me.RbtCF.Location = New System.Drawing.Point(31, 313)
        Me.RbtCF.Name = "RbtCF"
        Me.RbtCF.Size = New System.Drawing.Size(139, 18)
        Me.RbtCF.TabIndex = 1
        Me.RbtCF.Text = "Commande Fournisseur"
        '
        'RbtArt
        '
        Me.RbtArt.Location = New System.Drawing.Point(31, 33)
        Me.RbtArt.Name = "RbtArt"
        Me.RbtArt.Size = New System.Drawing.Size(81, 18)
        Me.RbtArt.TabIndex = 0
        Me.RbtArt.Text = "Fiche Article"
        '
        'FrmChoix
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(200, 364)
        Me.Controls.Add(Me.RadGroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmChoix"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Choix Correspondance"
        Me.ThemeName = "TelerikMetro"
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.RbtMvt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtTdepot, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtFfrss, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Rbtclt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtFrss, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtFC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtCClt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtCF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RbtArt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Windows8Theme1 As Telerik.WinControls.Themes.Windows8Theme
    Friend WithEvents TelerikMetroTheme1 As Telerik.WinControls.Themes.TelerikMetroTheme
    Friend WithEvents RadGroupBox1 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents Rbtclt As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtFrss As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtFC As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtCClt As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtCF As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtArt As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtMvt As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtTdepot As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents RbtFfrss As Telerik.WinControls.UI.RadRadioButton
End Class

