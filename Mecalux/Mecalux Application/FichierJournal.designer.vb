<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FichierJournal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FichierJournal))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DataJournal = New System.Windows.Forms.DataGridView
        Me.Fichier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Chemin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Selection = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BT_Open = New Telerik.WinControls.UI.RadButton
        Me.BT_Del = New Telerik.WinControls.UI.RadButton
        Me.BT_Select = New Telerik.WinControls.UI.RadButton
        Me.BT_Qui = New Telerik.WinControls.UI.RadButton
        Me.BT_Deselect = New Telerik.WinControls.UI.RadButton
        Me.TelerikMetroTheme1 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.Windows8Theme1 = New Telerik.WinControls.Themes.Windows8Theme
        Me.Windows8Theme2 = New Telerik.WinControls.Themes.Windows8Theme
        Me.TelerikMetroTheme2 = New Telerik.WinControls.Themes.TelerikMetroTheme
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Open, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Del, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Select, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Qui, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BT_Deselect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataJournal)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Open)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Del)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Select)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Qui)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BT_Deselect)
        Me.SplitContainer1.Size = New System.Drawing.Size(771, 532)
        Me.SplitContainer1.SplitterDistance = 498
        Me.SplitContainer1.TabIndex = 0
        '
        'DataJournal
        '
        Me.DataJournal.AllowUserToAddRows = False
        Me.DataJournal.AllowUserToDeleteRows = False
        Me.DataJournal.BackgroundColor = System.Drawing.Color.SlateGray
        Me.DataJournal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataJournal.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Fichier, Me.Chemin, Me.Selection})
        Me.DataJournal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataJournal.GridColor = System.Drawing.SystemColors.ControlLightLight
        Me.DataJournal.Location = New System.Drawing.Point(0, 0)
        Me.DataJournal.Name = "DataJournal"
        Me.DataJournal.RowHeadersVisible = False
        Me.DataJournal.Size = New System.Drawing.Size(771, 498)
        Me.DataJournal.TabIndex = 0
        '
        'Fichier
        '
        Me.Fichier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Fichier.HeaderText = "Fichier"
        Me.Fichier.Name = "Fichier"
        '
        'Chemin
        '
        Me.Chemin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Chemin.HeaderText = "Chemin"
        Me.Chemin.Name = "Chemin"
        Me.Chemin.Visible = False
        '
        'Selection
        '
        Me.Selection.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Selection.HeaderText = "Selection"
        Me.Selection.Name = "Selection"
        Me.Selection.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Selection.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Selection.Width = 120
        '
        'BT_Open
        '
        Me.BT_Open.Image = Global.Mecalux_Application.My.Resources.Resources.foldeopen_16
        Me.BT_Open.ImageAlignment = System.Drawing.ContentAlignment.BottomLeft
        Me.BT_Open.Location = New System.Drawing.Point(392, 3)
        Me.BT_Open.Name = "BT_Open"
        Me.BT_Open.Size = New System.Drawing.Size(110, 24)
        Me.BT_Open.TabIndex = 4
        Me.BT_Open.Text = "&Ouvrir"
        Me.BT_Open.ThemeName = "Windows8"
        '
        'BT_Del
        '
        Me.BT_Del.Image = Global.Mecalux_Application.My.Resources.Resources.btSupprimer221
        Me.BT_Del.ImageAlignment = System.Drawing.ContentAlignment.BottomLeft
        Me.BT_Del.Location = New System.Drawing.Point(514, 3)
        Me.BT_Del.Name = "BT_Del"
        Me.BT_Del.Size = New System.Drawing.Size(110, 24)
        Me.BT_Del.TabIndex = 5
        Me.BT_Del.Text = "&Supprimer"
        Me.BT_Del.ThemeName = "Windows8"
        '
        'BT_Select
        '
        Me.BT_Select.Image = Global.Mecalux_Application.My.Resources.Resources.image033
        Me.BT_Select.ImageAlignment = System.Drawing.ContentAlignment.BottomLeft
        Me.BT_Select.Location = New System.Drawing.Point(25, 3)
        Me.BT_Select.Name = "BT_Select"
        Me.BT_Select.Size = New System.Drawing.Size(167, 24)
        Me.BT_Select.TabIndex = 1
        Me.BT_Select.Text = "&Selectionner Tous"
        Me.BT_Select.ThemeName = "Windows8"
        '
        'BT_Qui
        '
        Me.BT_Qui.Image = Global.Mecalux_Application.My.Resources.Resources.btFermer22
        Me.BT_Qui.ImageAlignment = System.Drawing.ContentAlignment.BottomLeft
        Me.BT_Qui.Location = New System.Drawing.Point(636, 3)
        Me.BT_Qui.Name = "BT_Qui"
        Me.BT_Qui.Size = New System.Drawing.Size(110, 24)
        Me.BT_Qui.TabIndex = 3
        Me.BT_Qui.Text = "&Quitter"
        Me.BT_Qui.ThemeName = "Windows8"
        '
        'BT_Deselect
        '
        Me.BT_Deselect.Image = Global.Mecalux_Application.My.Resources.Resources.image019
        Me.BT_Deselect.ImageAlignment = System.Drawing.ContentAlignment.TopLeft
        Me.BT_Deselect.Location = New System.Drawing.Point(198, 3)
        Me.BT_Deselect.Name = "BT_Deselect"
        Me.BT_Deselect.Size = New System.Drawing.Size(182, 24)
        Me.BT_Deselect.TabIndex = 2
        Me.BT_Deselect.Text = "&Désélectionner Tous"
        Me.BT_Deselect.ThemeName = "Windows8"
        '
        'FichierJournal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(771, 532)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FichierJournal"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fichier de Journalisation"
        Me.ThemeName = "TelerikMetro"
        Me.TopMost = True
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataJournal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Open, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Del, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Select, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Qui, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BT_Deselect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DataJournal As System.Windows.Forms.DataGridView
    Friend WithEvents Fichier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Chemin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Selection As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BT_Open As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_Del As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_Qui As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_Deselect As Telerik.WinControls.UI.RadButton
    Friend WithEvents BT_Select As Telerik.WinControls.UI.RadButton
     Friend WithEvents TelerikMetroTheme1 As Telerik.WinControls.Themes.TelerikMetroTheme
      Friend WithEvents Windows8Theme1 As Telerik.WinControls.Themes.Windows8Theme
    Friend WithEvents Windows8Theme2 As Telerik.WinControls.Themes.Windows8Theme
    Friend WithEvents TelerikMetroTheme2 As Telerik.WinControls.Themes.TelerikMetroTheme
End Class
