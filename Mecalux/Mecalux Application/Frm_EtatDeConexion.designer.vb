<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_EtatDeConexion
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_EtatDeConexion))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.DataConnexion = New System.Windows.Forms.DataGridView
        Me.Connexion = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Reussie = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RadButton1 = New Telerik.WinControls.UI.RadButton
        Me.Windows8Theme1 = New Telerik.WinControls.Themes.Windows8Theme
        Me.TelerikMetroTheme1 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataConnexion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 3500
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 25
        '
        'Timer3
        '
        Me.Timer3.Enabled = True
        Me.Timer3.Interval = 1000
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.ProgressBar1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataConnexion)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadButton1)
        Me.SplitContainer1.Size = New System.Drawing.Size(512, 357)
        Me.SplitContainer1.SplitterDistance = 322
        Me.SplitContainer1.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 297)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(512, 25)
        Me.ProgressBar1.TabIndex = 3
        '
        'DataConnexion
        '
        Me.DataConnexion.AllowUserToAddRows = False
        Me.DataConnexion.AllowUserToDeleteRows = False
        Me.DataConnexion.BackgroundColor = System.Drawing.Color.LightSlateGray
        Me.DataConnexion.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataConnexion.ColumnHeadersVisible = False
        Me.DataConnexion.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Connexion, Me.Reussie})
        Me.DataConnexion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataConnexion.GridColor = System.Drawing.SystemColors.Control
        Me.DataConnexion.Location = New System.Drawing.Point(0, 0)
        Me.DataConnexion.Name = "DataConnexion"
        Me.DataConnexion.RowHeadersVisible = False
        Me.DataConnexion.RowTemplate.Height = 24
        Me.DataConnexion.Size = New System.Drawing.Size(512, 322)
        Me.DataConnexion.TabIndex = 2
        '
        'Connexion
        '
        Me.Connexion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Connexion.HeaderText = "Connexion"
        Me.Connexion.Name = "Connexion"
        Me.Connexion.ReadOnly = True
        '
        'Reussie
        '
        Me.Reussie.HeaderText = "Reussie"
        Me.Reussie.Name = "Reussie"
        Me.Reussie.ReadOnly = True
        '
        'RadButton1
        '
        Me.RadButton1.ImageAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.RadButton1.Location = New System.Drawing.Point(185, 3)
        Me.RadButton1.Name = "RadButton1"
        Me.RadButton1.Size = New System.Drawing.Size(143, 24)
        Me.RadButton1.TabIndex = 3
        Me.RadButton1.Text = "&Quitter"
        Me.RadButton1.ThemeName = "Windows8"
        '
        'Frm_EtatDeConexion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 357)
        Me.ControlBox = False
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_EtatDeConexion"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Vérification des Connexions SQL pour chaque Societe"
        Me.ThemeName = "TelerikMetro"
        Me.TopMost = True
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataConnexion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Timer3 As System.Windows.Forms.Timer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataConnexion As System.Windows.Forms.DataGridView
    Friend WithEvents Connexion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Reussie As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents RadButton1 As Telerik.WinControls.UI.RadButton
    Friend WithEvents Windows8Theme1 As Telerik.WinControls.Themes.Windows8Theme
    Friend WithEvents TelerikMetroTheme1 As Telerik.WinControls.Themes.TelerikMetroTheme

End Class
