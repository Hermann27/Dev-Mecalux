Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class MenuPrincipal
    Private intNRStyleOuvert As Integer = 1     ' Menu ouvert
    Private intNRStyleAOuvrir As Integer        ' Menu à ouvrir

    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Global.System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilisez My.Computer.Clipboard pour insérer les images ou le texte sélectionné dans le Presse-papiers
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilisez My.Computer.Clipboard pour insérer les images ou le texte sélectionné dans le Presse-papiers
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Utilisez My.Computer.Clipboard.GetText() ou My.Computer.Clipboard.GetData pour extraire les informations du Presse-papiers.
    End Sub
    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub MenuPrincipal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Hide()
            frmChargement.Hide()
            LirefichierConfig()
            ' Paramétrage par défaut
            Me.Table1.RowStyles(1).SizeType = SizeType.Percent
            Me.Table1.RowStyles(1).Height = 100
            Me.Table1.RowStyles(3).SizeType = SizeType.Percent
            Me.Table1.RowStyles(3).Height = 0
            Me.Table1.RowStyles(5).SizeType = SizeType.Percent
            Me.Table1.RowStyles(5).Height = 0
            'ActiverLeLancementDesTâcheToolStripMenuItem_Click(sender, e)
            Me.Text = "Interface Mecalux< Utilisateur Connecté:[" & Nom_Util & "]<>Connexion à la base < [" & NomBaseCpta & "]>"
            EffetPlanifier(sender, e)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 1
            Me.Timer1.Start()
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ' Réduire le menu
        If Me.Table1.RowStyles(intNRStyleOuvert).Height > 0 Then

            Me.Table1.RowStyles(intNRStyleOuvert).SizeType = SizeType.Percent
            Me.Table1.RowStyles(intNRStyleOuvert).Height -= 5

        End If

        ' Augmenter le menu
        If Me.Table1.RowStyles(intNRStyleAOuvrir).Height < 100 Then

            Me.Table1.RowStyles(intNRStyleAOuvrir).SizeType = SizeType.Percent
            Me.Table1.RowStyles(intNRStyleAOuvrir).Height += 5

        End If

        ' Arrêt du cycle
        If Me.Table1.RowStyles(intNRStyleAOuvrir).Height = 100 Then

            intNRStyleOuvert = intNRStyleAOuvrir
            Me.Timer1.Stop()

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 3
            Me.Timer1.Start()
        End If
    End Sub

    Public Sub lblExtractionART_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblExtractionART.Click
        FrmExtraction.Close()
        With FrmExtraction
            .MdiParent = Me
            '.SocieteCyble.Clear()
            '.BackgroundWorker1.CancelAsync()
            '.BackgroundWorker2.CancelAsync()
            '.SocieteCyble.Add("SEREBIS")
            '.SocieteCyble.Add("CYNOCARE")
            '.TachePlanifie = "Export Article"
            '.Ckmodifier.Checked = False
            '.FrmExtraction_Load(sender, e)
            '.BtnModif_Click(sender, e)
            .Show()
        End With
    End Sub

    Private Sub lblClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblClient.Click
        FrmExtractionClient.MdiParent = Me
        FrmExtractionClient.Show()
    End Sub

    Private Sub lblFrs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFrs.Click
        FrmExtractionFournisseurs.MdiParent = Me
        FrmExtractionFournisseurs.Show()
    End Sub

    Private Sub lblBR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblBR.Click
        With FrmExtractionBCFournisseur
            .MdiParent = Me
            .Show()
        End With
    End Sub

    Private Sub lblCdClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCdClient.Click
        With FrmExtractionBCClient
            .MdiParent = Me
            .Item = " Bon de Commande Client "
            .RbtVente.Checked = True
            .RbtVente.Enabled = True
            .Show()
        End With
    End Sub

    Private Sub lblBT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblBT.Click
        With FrmExtractionPseudo
            .MdiParent = Me
            .Show()
        End With
    End Sub
    Private Sub ToolStripMenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem29.Click
        FrmCorrespondance.MdiParent = Me
        FrmCorrespondance.Show()
    End Sub

    Private Sub ToolStripMenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem30.Click
        With FrmChoix
            .MdiParent = Me
            .choix = "Correspondace"
            .Show()
        End With
    End Sub

    Private Sub ToolStripMenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T2.Click
        Frm_FichierConfiguration.MdiParent = Me
        Frm_FichierConfiguration.Show()
    End Sub

    Private Sub ToolStripMenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T1.Click
        ParametreSocieteConsoleWaza.MdiParent = Me
        ParametreSocieteConsoleWaza.Show()
    End Sub

    Private Sub ToolStripMenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem26.Click
        FichierJournal.MdiParent = Me
        FichierJournal.Show()
    End Sub

    Private Sub ToolStripMenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Frm_FluxEntrant.MdiParent = Me
        'Frm_FluxEntrant.Show()
    End Sub
    Private Sub ToolStripMenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem31.Click
        Me.Close()
    End Sub

    Private Sub TemplateCorrespondanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TemplateCorrespondanceToolStripMenuItem.Click
        With FrmChoix
            .MdiParent = Me
            .choix = "Template"
            .Show()
        End With
    End Sub

    Private Sub PlanifierLesTâchesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanifierLesTâchesToolStripMenuItem.Click
        PlanificationSpecial.MdiParent = Me
        PlanificationSpecial.Show()
        'PlanificationTache.MdiParent = Me
        'PlanificationTache.Show()
    End Sub
    Public t As Date
    Public TMR As String = Nothing
    Public DOT As Integer = 0
    Public Clock As String = Nothing
    Public Situation As Boolean = True
    Private Sub ActiverLeLancementDesTâcheToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActiverLeLancementDesTâcheToolStripMenuItem.Click
        'If Situation = False Then
        '    Timer2.Enabled = True
        '    Situation = True
        '    tlbetat.Text = "Statut de la planification : Actif"
        'Else
        '    Timer2.Enabled = False
        '    Situation = False
        '    tlbetat.Text = "Statut de la planification : Sommeil"
        'End If
    End Sub
    'Public Sub EffetPlanifier(ByVal libelleTache)

    'End Sub
    Private Sub EffetPlanifier(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim CLK As String = Mid(TMR, 1, 8)
        'variables base de données
        Dim LibreOleAdaptater As OleDbDataAdapter
        Dim Libredataset As DataSet
        Dim Libredatabase As DataTable
        'Dim DateDebut As Date
        Dim HeurDebut As String = ""
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable

        Try
            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION where Lancer=true AND (IntituleTache='Import Document Vente' OR IntituleTache='Import Document Achat' OR IntituleTache='Import Mouvement de Transfert' OR IntituleTache='Import Mouvement de Variation des Stock') ORDER BY Rang ", OleConnenection)
            Libredataset = New DataSet
            LibreOleAdaptater.Fill(Libredataset)
            Libredatabase = Libredataset.Tables(0)
            If Libredatabase.Rows.Count = 0 Then
                Dim OleCommandEnreg As OleDbCommand
                Dim Insertion As String
                Try
                    Insertion = "UPDATE PLANIFICATION SET Lancer=true Where Rang IS NOT NULL AND (IntituleTache='Import Document Vente' OR IntituleTache='Import Document Achat' OR IntituleTache='Import Mouvement de Transfert' OR IntituleTache='Import Mouvement de Variation des Stock') "
                    OleCommandEnreg = New OleDbCommand(Insertion, OleConnenection)
                    OleCommandEnreg.ExecuteNonQuery()
                    LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION where Lancer=true AND (IntituleTache='Import Document Vente' OR IntituleTache='Import Document Achat' OR IntituleTache='Import Mouvement de Transfert' OR IntituleTache='Import Mouvement de Variation des Stock') ORDER BY Rang ", OleConnenection)
                    Libredataset = New DataSet
                    LibreOleAdaptater.Fill(Libredataset)
                    Libredatabase = Libredataset.Tables(0)
                Catch ex As Exception
                End Try
            End If
            If Libredatabase.Rows.Count <> 0 Then
                For i As Integer = 0 To Libredatabase.Rows.Count - 1
                    Select Case Libredatabase.Rows(i).Item("IntituleTache")
                        Case "Import Document Vente"
                            Try
                                If Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) >= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime) And Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) <= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime) Then
                                    OleAdaptaterschema = New OleDbDataAdapter("select Id,Intitule,Heure1,Heure2,Critere1,Critere2 from PLANIFICATION WHERE IntituleTache='Import Document Vente' AND Lancer=true", OleConnenection)
                                    OleSchemaDataset = New DataSet
                                    OleAdaptaterschema.Fill(OleSchemaDataset)
                                    OledatableSchema = OleSchemaDataset.Tables(0)
                                    If OledatableSchema.Rows.Count <> 0 Then

                                        Frm_ConfirmationReception.Critere = "CRP"
                                        Frm_ConfirmationReception.Achat = "VENTE"
                                        Frm_ConfirmationReception.Frm_FluxEntrantCritére_Load(sender, e)
                                        Frm_ConfirmationReception.BtnXFormationCRP_Click(sender, e)

                                        Frm_FluxEntrantCritére.Critere = "CSO"
                                        Frm_FluxEntrantCritére.Achat = "VENTE"
                                        Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                                        Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                                        Fr_ImportationMvtVente.Fr_ImportationMvtVente_Load(sender, e)
                                        Fr_ImportationMvtVente.BT_integrer_Click(sender, e)
                                        MiseàjourTachePlanifier("Import Document Vente", OledatableSchema.Rows(0).Item("Id"))
                                    End If
                                Else

                                End If

                            Catch ex As Exception
                            End Try
                            Continue For
                        Case "Import Document Achat"
                            Try
                                If Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) >= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime) And Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) <= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime) Then
                                    OleAdaptaterschema = New OleDbDataAdapter("select Id,Intitule,Heure1,Heure2,Critere1,Critere2 from PLANIFICATION WHERE IntituleTache='Import Document Achat' AND Lancer=true", OleConnenection)
                                    OleSchemaDataset = New DataSet
                                    OleAdaptaterschema.Fill(OleSchemaDataset)
                                    OledatableSchema = OleSchemaDataset.Tables(0)
                                    If OledatableSchema.Rows.Count <> 0 Then
                                        '
                                        Frm_ConfirmationReception.Critere = "CRP"
                                        Frm_ConfirmationReception.Achat = "ACHAT"
                                        Frm_ConfirmationReception.Frm_FluxEntrantCritére_Load(sender, e)
                                        Frm_ConfirmationReception.BtnXFormationCRP_Click(sender, e)

                                        Frm_FluxEntrantCritére.Critere = "CSO"
                                        Frm_FluxEntrantCritére.Achat = "ACHAT"
                                        Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                                        Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                                        Fr_ImportationMvtAchat.Fr_ImportationMvtAchat_Load(sender, e)
                                        Fr_ImportationMvtAchat.BT_integrer_Click(sender, e)
                                        MiseàjourTachePlanifier("Import Document Achat", OledatableSchema.Rows(0).Item("Id"))
                                    End If
                                Else

                                End If

                            Catch ex As Exception
                            End Try
                            Continue For
                        Case "Import Mouvement de Transfert"
                            Try
                                If Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) >= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime) And Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) <= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime) Then
                                    OleAdaptaterschema = New OleDbDataAdapter("select Id,Intitule,Heure1,Heure2,Critere1,Critere2 from PLANIFICATION WHERE IntituleTache='Import Mouvement de Transfert' AND Lancer=true", OleConnenection)
                                    OleSchemaDataset = New DataSet
                                    OleAdaptaterschema.Fill(OleSchemaDataset)
                                    OledatableSchema = OleSchemaDataset.Tables(0)
                                    If OledatableSchema.Rows.Count <> 0 Then
                                        '
                                        Frm_FluxEntrantCritére.Critere = "CSO"
                                        Frm_FluxEntrantCritére.Achat = "TRANSFERT"
                                        Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                                        Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                                        Fr_ImportationTransfert.Fr_ImportationTransfert_Load(sender, e)
                                        Fr_ImportationTransfert.BT_integrer_Click(sender, e)
                                        MiseàjourTachePlanifier("Import Mouvement de Transfert", OledatableSchema.Rows(0).Item("Id"))
                                    End If
                                Else

                                End If
                            Catch ex As Exception
                            End Try
                            Continue For
                        Case "Import Mouvement de Variation des Stock"
                            Try
                                If Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) >= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime) And Strings.FormatDateTime(DateAndTime.TimeString, DateFormat.LongTime) <= Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime) Then
                                    OleAdaptaterschema = New OleDbDataAdapter("select Id,Intitule,Heure1,Heure2,Critere1,Critere2 from PLANIFICATION WHERE IntituleTache='Import Mouvement de Variation des Stock' AND Lancer=true", OleConnenection)
                                    OleSchemaDataset = New DataSet
                                    OleAdaptaterschema.Fill(OleSchemaDataset)
                                    OledatableSchema = OleSchemaDataset.Tables(0)
                                    If OledatableSchema.Rows.Count <> 0 Then
                                        '
                                        With Frm_MvtE_S
                                            .Critere = "VST"
                                            .Frm_FluxEntrantCritére_Load(sender, e)
                                            .Button1_Click_1(sender, e)
                                            .Text = " Traitement Transformation Variation de Stock "
                                        End With
                                        Fr_ImportationStock.Fr_ImportationStock_Load(sender, e)
                                        Fr_ImportationStock.BT_integrer_Click(sender, e)
                                        MiseàjourTachePlanifier("Import Mouvement de Variation des Stock", OledatableSchema.Rows(0).Item("Id"))
                                    End If
                                Else

                                End If

                            Catch ex As Exception
                            End Try
                            Continue For
                    End Select
                Next
                Me.Close()
            End If
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub
    Private Sub MiseàjourTachePlanifier(ByVal libeletache As String, ByVal identifiant As Integer)
        Dim OleCommandEnreg As OleDbCommand
        Dim Insertion As String
        Try
            Insertion = "UPDATE PLANIFICATION SET LastExecution='" & Now & "',Lancer=False Where IntituleTache='" & Join(Split(Trim(libeletache), "'"), "''") & "' " 'AND ID= & identifiant
            OleCommandEnreg = New OleDbCommand(Insertion, OleConnenection)
            OleCommandEnreg.ExecuteNonQuery()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        '
        If DOT < 5 Then
            t = Now
            TMR = Format(t, "HH:mm:ss")
            Label1.Text = "Heure Système : " & TMR
        Else
            t = Now
            TMR = Format(t, "HH mm ss")
            Label1.Text = "Heure Système : " & TMR
        End If
        '
        DOT = (DOT + 1) Mod 10
        '
        If DOT >= 5 Then
            'Alarme(sender, e)
        End If
        '
    End Sub
    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Select Case e.Node.Name
            'Traitement Mouvement Achat/Vente
            Case "N1" 'parametrage commercial
                Parametre_Piece_Commerciale.MdiParent = Me
                Parametre_Piece_Commerciale.Show()

            Case "N2" 'format d'integration
                FormatDintegrationMvt.MdiParent = Me
                FormatDintegrationMvt.Show()
            Case "N21" 'Ajouter format d'integration
                AjouterUnFormatMvt.MdiParent = Me
                AjouterUnFormatMvt.Show()
            Case "N3" 'parametre d'integratiion
                SchematintegrerMvt.MdiParent = Me
                SchematintegrerMvt.Show()

            Case "N4" 'correspondance article
                Frm_CorArticle.MdiParent = Me
                Frm_CorArticle.Show()

            Case "N5" 'infos libre
                InfoLibreCommercial.ShowDialog()

            Case "N6" 'Transcodage des infos
                Transcodage.MdiParent = Me
                Transcodage.Show()

            Case "N7" ' Lancement Achat
                Frm_ConfirmationReception.Critere = "CRP"
                Frm_ConfirmationReception.Achat = "ACHAT"
                Frm_ConfirmationReception.Frm_FluxEntrantCritére_Load(sender, e)
                Frm_ConfirmationReception.BtnXFormationCRP_Click(sender, e)

                Frm_FluxEntrantCritére.Critere = "CSO"
                Frm_FluxEntrantCritére.Achat = "ACHAT"
                Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                Fr_ImportationMvtAchat.MdiParent = Me
                Fr_ImportationMvtAchat.Show()

            Case "N8" 'Lancement Vente
                Frm_ConfirmationReception.Critere = "CRP"
                Frm_ConfirmationReception.Achat = "VENTE"
                Frm_ConfirmationReception.Frm_FluxEntrantCritére_Load(sender, e)
                Frm_ConfirmationReception.BtnXFormationCRP_Click(sender, e)

                Frm_FluxEntrantCritére.Critere = "CSO"
                Frm_FluxEntrantCritére.Achat = "VENTE"
                Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                Fr_ImportationMvtVente.MdiParent = Me
                Fr_ImportationMvtVente.Show()

            Case "N9" 'Lancement Transformation
                With FrmChoixTraitement
                    .MdiParent = Me
                    .Show()
                End With

                'Traiment Mouvement de Stock
            Case "N10" 'Format d'integration
                FormatintegrationStock.MdiParent = Me
                FormatintegrationStock.Show()
            Case "N22" 'Ajouter Format d'integration
                AFormatStock.MdiParent = Me
                AFormatStock.Show()

            Case "N11" 'parametre d'integratiion
                SchematintegrerStock.MdiParent = Me
                SchematintegrerStock.Show()

            Case "N12" 'infos libre
                InfoLibreStock.ShowDialog()

            Case "N13" 'Transcodage des infos
                Transcodage.MdiParent = Me
                Transcodage.Show()

            Case "N14" ' Lancement Mvt Stock
                With Frm_MvtE_S
                    .Critere = "VST"
                    .Frm_FluxEntrantCritére_Load(sender, e)
                    .Button1_Click_1(sender, e)
                    .Text = " Traitement Transformation Variation de Stock "
                End With
                Fr_ImportationStock.MdiParent = Me
                Fr_ImportationStock.Show()

            Case "N15" ' Lancement Transformation
                With Frm_MvtE_S
                    .Critere = "VST"
                    .Frm_FluxEntrantCritére_Load(sender, e)
                    .Button1_Click_1(sender, e)
                    .Text = " Traitement Transformation Variation de Stock "
                End With
                'Traitement Mouvement de Transfert 
            Case "N16" 'format d'integration
                FormatintegrationTransfert.MdiParent = Me
                FormatintegrationTransfert.Show()
            Case "N23" 'format d'integration
                AFormatTransfert.MdiParent = Me
                AFormatTransfert.Show()
            Case "N17" 'parametre d'integratiion
                SchematintegrerTransfert.MdiParent = Me
                SchematintegrerTransfert.Show()

            Case "N18" 'infos libre
                InfoLibreTransfert.ShowDialog()

            Case "N19" 'Transcodage des infos
                Transcodage.MdiParent = Me
                Transcodage.Show()

            Case "N20" ' Lancement
                Frm_FluxEntrantCritére.Critere = "CSO"
                Frm_FluxEntrantCritére.Achat = "TRANSFERT"
                Frm_FluxEntrantCritére.Frm_FluxEntrantCritére_Load(sender, e)
                Frm_FluxEntrantCritére.BtnXFormation_Click(sender, e)

                Fr_ImportationTransfert.MdiParent = Me
                Fr_ImportationTransfert.Show()
        End Select
    End Sub
    Private Sub FormatDescriptifLongueurFixeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormatDescriptifLongueurFixeToolStripMenuItem.Click
        With FormVersionng
            .MdiParent = Me
            .Show()
        End With
    End Sub
End Class
