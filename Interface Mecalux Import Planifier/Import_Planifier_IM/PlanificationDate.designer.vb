<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlanificationDate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PlanificationDate))
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TimeDebut = New System.Windows.Forms.DateTimePicker
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TimeFin = New System.Windows.Forms.DateTimePicker
        Me.Button1 = New System.Windows.Forms.Button
        Me.BTvalider = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(86, 52)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 13)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Date fin"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(84, 23)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 13)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Date debut"
        '
        'TimeDebut
        '
        Me.TimeDebut.CustomFormat = ""
        Me.TimeDebut.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TimeDebut.Location = New System.Drawing.Point(167, 19)
        Me.TimeDebut.Name = "TimeDebut"
        Me.TimeDebut.Size = New System.Drawing.Size(83, 20)
        Me.TimeDebut.TabIndex = 18
        Me.TimeDebut.Value = New Date(2008, 1, 1, 0, 0, 0, 0)
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TimeFin)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.TimeDebut)
        Me.GroupBox1.Location = New System.Drawing.Point(2, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(425, 80)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        '
        'TimeFin
        '
        Me.TimeFin.CustomFormat = ""
        Me.TimeFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TimeFin.Location = New System.Drawing.Point(167, 50)
        Me.TimeFin.Name = "TimeFin"
        Me.TimeFin.Size = New System.Drawing.Size(83, 20)
        Me.TimeFin.TabIndex = 25
        Me.TimeFin.Value = New Date(2008, 1, 1, 0, 0, 0, 0)
        '
        'Button1
        '
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(215, 84)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(65, 21)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "&Quitter"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BTvalider
        '
        Me.BTvalider.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTvalider.Location = New System.Drawing.Point(302, 84)
        Me.BTvalider.Name = "BTvalider"
        Me.BTvalider.Size = New System.Drawing.Size(65, 21)
        Me.BTvalider.TabIndex = 27
        Me.BTvalider.Text = "&Valider"
        Me.BTvalider.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTvalider.UseVisualStyleBackColor = True
        '
        'PlanificationDate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(439, 109)
        Me.Controls.Add(Me.BTvalider)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "PlanificationDate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Fourchette date de traitement"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TimeDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BTvalider As System.Windows.Forms.Button
    Friend WithEvents TimeFin As System.Windows.Forms.DateTimePicker
End Class
