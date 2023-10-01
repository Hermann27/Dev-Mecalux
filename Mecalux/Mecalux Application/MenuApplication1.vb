Imports System.Windows.Forms

Public Class MenuApplication1
    Private intNRStyleOuvert As Integer = 1     ' Menu ouvert
    Private intNRStyleAOuvrir As Integer        ' Menu à ouvrir
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click
        ParametreSocieteConsoleWaza.MdiParent = Me
        ParametreSocieteConsoleWaza.Show()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Frm_FichierConfiguration.MdiParent = Me
        Frm_FichierConfiguration.Show()
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click

        'With Authentification
        '    .Show()
        '    Me.Close()
        'End With
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

    'Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolBarToolStripMenuItem.Click
    '    Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    'End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
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

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Fermez tous les formulaires enfants du parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer = 0

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Intégration.Click
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 1
            Me.Timer1.Start()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ' Réduire le menu
        If Me.TableLayoutPanel1.RowStyles(intNRStyleOuvert).Height > 0 Then

            Me.TableLayoutPanel1.RowStyles(intNRStyleOuvert).SizeType = SizeType.Percent
            Me.TableLayoutPanel1.RowStyles(intNRStyleOuvert).Height -= 5

        End If

        ' Augmenter le menu
        If Me.TableLayoutPanel1.RowStyles(intNRStyleAOuvrir).Height < 100 Then

            Me.TableLayoutPanel1.RowStyles(intNRStyleAOuvrir).SizeType = SizeType.Percent
            Me.TableLayoutPanel1.RowStyles(intNRStyleAOuvrir).Height += 5

        End If

        ' Arrêt du cycle
        If Me.TableLayoutPanel1.RowStyles(intNRStyleAOuvrir).Height = 100 Then

            intNRStyleOuvert = intNRStyleAOuvrir
            Me.Timer1.Stop()

        End If
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Exportation.Click
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 3
            Me.Timer1.Start()
        End If
    End Sub

    Private Sub RadButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 5
            Me.Timer1.Start()
        End If
    End Sub

    Private Sub RadButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 7
            Me.Timer1.Start()
        End If

    End Sub

    Private Sub RadButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 9
            Me.Timer1.Start()
        End If
    End Sub
    Sub WriteToRegister(ByVal regkey As String, ByVal regvalue As String)
        Dim regedit = CreateObject("WScript.Shell")
        regedit.regwrite(regkey, regvalue)
    End Sub
    Public Function ReadToRegister(ByVal clé As String) As String
        Dim regedit = CreateObject("WScript.Shell")
        Return regedit.regread(clé)
    End Function
    Public Sub MenuApplication1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
        ' Paramétrage par défaut
        Me.TableLayoutPanel1.RowStyles(1).SizeType = SizeType.Percent
        Me.TableLayoutPanel1.RowStyles(1).Height = 100
        Me.TableLayoutPanel1.RowStyles(3).SizeType = SizeType.Percent
        Me.TableLayoutPanel1.RowStyles(3).Height = 0
        Me.TableLayoutPanel1.RowStyles(5).SizeType = SizeType.Percent
        Me.TableLayoutPanel1.RowStyles(5).Height = 0
        Me.Text = "Interface Mecalux< Utilisateur Connecté:[" & Nom_Util & "]<>Connexion à la base < [" & NomBaseCpta & "]>"
    End Sub

    Private Sub RadButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Timer1.Enabled = False Then
            intNRStyleAOuvrir = 10
            Me.Timer1.Start()
        End If
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'FormEcriture.MdiParent = Me
        'FormEcriture.Show()
    End Sub

    Private Sub GestionDuMotDePasseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub UtilisateursToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'With ListeUtilisateurs
        '    .MdiParent = Me
        '    .Show()
        'End With
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        'About.MdiParent = Me
        'About.Show()
    End Sub
    Dim Etat As Boolean = True
    Dim Etat2 As Boolean = True


    'Private Sub btnCommercial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If Etat2 = True Then
    '        PanelA.Enabled = False
    '        PanelA.Visible = False
    '        'Panel3.Location = New Point(Panel3.Location.X, btnComptabilite.Location.Y + 16)
    '        Etat2 = False
    '    Else
    '        PanelA.Enabled = True
    '        PanelA.Visible = True
    '        Panel3.Location = New Point(Panel3.Location.X, PanelA.Location.Y + PanelA.Height + 16)
    '        Etat2 = True
    '    End If
    'End Sub

    Private Sub PictureBox1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox1.BackColor = Color.CornflowerBlue
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox1.BackColor = Color.Transparent
    End Sub

    Private Sub PictureBox2_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox2.BackColor = Color.CornflowerBlue
    End Sub

    Private Sub PictureBox2_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox2.BackColor = Color.Transparent
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'PlanificationTache.MdiParent = Me
        'PlanificationTache.Show()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        FichierJournal.MdiParent = Me
        FichierJournal.Show()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'FichierJournalSQL.MdiParent = Me
        'FichierJournalSQL.Show()
    End Sub

    Private Sub FaireUneCorrespondanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FaireUneCorrespondanceToolStripMenuItem.Click
        FrmCorrespondance.MdiParent = Me
        FrmCorrespondance.Show()
    End Sub

    Private Sub VisualiserLesCorrespondanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisualiserLesCorrespondanceToolStripMenuItem.Click
        FrmChoix.MdiParent = Me
        FrmChoix.Show()
    End Sub

    Private Sub lblExtractionART_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblExtractionART.Click
        FrmExtraction.MdiParent = Me
        FrmExtraction.Show()
    End Sub

    Private Sub lblFrs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFrs.Click
        FrmExtractionFournisseurs.MdiParent = Me
        FrmExtractionFournisseurs.Show()
    End Sub

    Private Sub lblClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblClient.Click
        FrmExtractionClient.MdiParent = Me
        FrmExtractionClient.Show()
    End Sub

    Private Sub lblCdClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCdClient.Click
        With FrmExtractionBCClient
            .MdiParent = Me
            .Item = " Bon de Commande Client "
            .RbtStock.Enabled = False
            .RbtVente.IsChecked = True
            .RbtVente.Enabled = True
            .RbtAchat.Enabled = False
            .Show()
        End With
    End Sub

    Private Sub lblBR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblBR.Click
        With FrmExtractionBCClient
            .MdiParent = Me
            .Item = "Bon de retour Fournisseur "
            .RbtStock.Enabled = False
            .RbtAchat.IsChecked = True
            .RbtAchat.Enabled = True
            .RbtVente.Enabled = False
            .Show()
        End With
    End Sub

    Private Sub lblBT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblBT.Click
        With FrmExtractionBCClient
            .MdiParent = Me
            .Item = "Transfért de Dépôt "
            .RbtVente.Enabled = False
            .RbtStock.IsChecked = True
            .RbtStock.Enabled = True
            .RbtAchat.Enabled = False
            .Show()
        End With
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Frm_FluxEntrant.MdiParent = Me
        Frm_FluxEntrant.Show()
    End Sub

    Private Sub lblCFCommande_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCFCommande.Click
        With Frm_FluxEntrantCritére
            .Critere = "CSO"
            .MdiParent = Me
            .Text = " Traitement " & lblCFCommande.Text
            .Show()
        End With
    End Sub

    Private Sub lblCFRReception_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCFRReception.Click
        With Frm_FluxEntrantCritére
            .MdiParent = Me
            .Critere = "CRP"
            .Text = " Traitement " & lblCFRReception.Text
            .Show()
        End With
    End Sub
End Class
