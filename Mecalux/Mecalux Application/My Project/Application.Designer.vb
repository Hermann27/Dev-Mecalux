﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Ce code a été généré par un outil.
'     Version du runtime :2.0.50727.4927
'
'     Les modifications apportées à ce fichier peuvent provoquer un comportement incorrect et seront perdues si
'     le code est régénéré.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    'REMARQUE : ce fichier étant généré automatiquement, ne le modifiez pas directement. Pour apporter des modifications,
    ' ou si vous rencontrez des erreurs de génération dans ce fichier, accédez au Concepteur de projets
    ' (allez dans les propriétés du projet ou double-cliquez sur le noeud Mon projet dans
    ' l'Explorateur de solutions), puis apportez vos modifications sous l'onglet Application.
    '
    Partial Friend Class MyApplication
        
        <Global.System.Diagnostics.DebuggerStepThroughAttribute()>  _
        Public Sub New()
            MyBase.New(Global.Microsoft.VisualBasic.ApplicationServices.AuthenticationMode.Windows)
            Me.IsSingleInstance = false
            Me.EnableVisualStyles = true
            Me.SaveMySettingsOnExit = true
            Me.ShutDownStyle = Global.Microsoft.VisualBasic.ApplicationServices.ShutdownMode.AfterMainFormCloses
        End Sub
        
        <Global.System.Diagnostics.DebuggerStepThroughAttribute()>  _
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = Global.Mecalux_Application.MenuApplication1
        End Sub
    End Class
End Namespace
