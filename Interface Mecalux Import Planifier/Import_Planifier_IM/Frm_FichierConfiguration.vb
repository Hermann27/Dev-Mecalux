Imports System.IO
Imports System.Net.NetworkInformation
Public Class Frm_FichierConfiguration
    Private Sub Frm_FichierConfiguration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call LirefichierConfig()
            TxtFichierCpta.Text = PathsBaseCpta
            TxtBDCpta.Text = NomBaseCpta
            TxtUserCpta.Text = Nom_Util
            TxtPasword.Text = Mot_Pas
            Txtsql.Text = NomServersql
            TxtFilejr.Text = Pathsfilejournal
            TxtUtilisateur.Text = Nom_Utilsql
            TxtPasw.Text = Mot_Passql
            Txt_Rep.Text = PathsfileSave
            TxtAccess.Text = PathsFileAccess
            Txtiers.Text = PathsFileFormatiers
            TxtArticle.Text = PathsFileFormatArticle
            TxtFlag.Text = Flagtampon
            txtFileCSOTempon.Text = PathsFileCSO
            txtFileCRPTempon.Text = PathsFileCRP
            txtFileVSTTempon.Text = PathsFileVST
            txtCheminXfert.Text = PathsFileXFERT
            txtCheminErreur.Text = PathsFileCSV_ERROR
            txtCheminMecalux.Text = PathsFileMECALUX
            txtFlagueArticle.Text = FlagtamponArticle
            CmbStatut.SelectedIndex = DO_StatutClient
            CmbStatutFrs.SelectedIndex = DO_StatutFournisseur
            'CodeEDI
            If StatutConsolider = "Oui" Then
                CkConso.Checked = True
            Else
                CkConso.Checked = False
            End If
            If CodeEDI = True Then
                ChekCodeEDI.Checked = True
            Else
                ChekCodeEDI.Checked = False
            End If
            If LOT = True Then
                CheckLot.Checked = True
            Else
                CheckLot.Checked = False
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub BT_FicCpta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_FicCpta.Click
        OpenFileFicCpta.Filter = "Fichier Compta (*.mae)|*.mae"
        OpenFileFicCpta.FileName = Nothing
        If OpenFileFicCpta.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtFichierCpta.Text = OpenFileFicCpta.FileName
            Txtsql.Text = LireChaine(Trim(OpenFileFicCpta.FileName), "CBASE", "ServeurSQL")
            If File.Exists(Trim(OpenFileFicCpta.FileName)) = True Then
                TxtBDCpta.Text = System.IO.Path.GetFileNameWithoutExtension(Trim(OpenFileFicCpta.FileName))
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Bool As Boolean
        If ChekCodeEDI.Checked Then
            Bool = WritePrivateProfileString("AUTRES", "CODE EDI", Trim(True), Pougoue_Fichier)
        Else
            Bool = WritePrivateProfileString("AUTRES", "CODE EDI", Trim(False), Pougoue_Fichier)
        End If
        If CheckLot.Checked Then
            Bool = WritePrivateProfileString("COMMENTAIRE", "LOT", Trim(True), Pougoue_Fichier)
        Else
            Bool = WritePrivateProfileString("COMMENTAIRE", "LOT", Trim(False), Pougoue_Fichier)
        End If
        If Trim(TxtFichierCpta.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "CHEMIN DU FICHIER COMPTA", Trim(TxtFichierCpta.Text), Pougoue_Fichier)
        End If
        If Trim(TxtBDCpta.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "BASE DE DONNEES COMPTA", Trim(TxtBDCpta.Text), Pougoue_Fichier)
        End If

        Bool = WritePrivateProfileString("CONNECTION", "FLAGTAMPON", Trim(TxtFlag.Text), Pougoue_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "FLAG", Trim(txtFlagueArticle.Text), Pougoue_Fichier)
        Select Case CmbStatut.Text
            Case "Saisi"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT CLIENT", Trim(0), Pougoue_Fichier)
            Case "Confirmé"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT CLIENT", Trim(1), Pougoue_Fichier)
            Case "Réceptionné"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT CLIENT", Trim(2), Pougoue_Fichier)
        End Select
        Select Case CmbStatutFrs.Text
            Case "Saisi"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT FOURNISSEUR", Trim(0), Pougoue_Fichier)
            Case "Confirmé"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT FOURNISSEUR", Trim(1), Pougoue_Fichier)
            Case "Réceptionné"
                Bool = WritePrivateProfileString("STATUT DU DOCUMENT", "DO_STATUT FOURNISSEUR", Trim(2), Pougoue_Fichier)
        End Select

        Bool = WritePrivateProfileString("CONNECTION", "UTILISATEUR SQL", Trim(TxtUtilisateur.Text), Pougoue_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "MOT DE PASSE SQL", Trim(TxtPasw.Text), Pougoue_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "UTILISATEUR", Trim(TxtUserCpta.Text), Pougoue_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "MOT DE PASSE", Trim(TxtPasword.Text), Pougoue_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "SERVEUR SQL", Trim(Txtsql.Text), Pougoue_Fichier)

        If Trim(TxtFilejr.Text) <> "" Then
            If Strings.Right(Trim(TxtFilejr.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE JOURNAL", Trim(TxtFilejr.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE JOURNAL", Trim(TxtFilejr.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(Txt_Rep.Text) <> "" Then
            If Strings.Right(Trim(Txt_Rep.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE SAUVEGARDE", Trim(Txt_Rep.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE SAUVEGARDE", Trim(Txt_Rep.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(Txtiers.Text) <> "" Then
            If Strings.Right(Trim(Txtiers.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT TIERS", Trim(Txtiers.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT TIERS", Trim(Txtiers.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(TxtArticle.Text) <> "" Then
            If Strings.Right(Trim(TxtArticle.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT ARTICLES", Trim(TxtArticle.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT ARTICLES", Trim(TxtArticle.Text) & "\", Pougoue_Fichier)
            End If
        End If

        If Trim(txtFileCSOTempon.Text) <> "" Then
            If Strings.Right(Trim(txtFileCSOTempon.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(ACHAT) TEMPORAIRE", Trim(txtFileCSOTempon.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(ACHAT) TEMPORAIRE", Trim(txtFileCSOTempon.Text) & "\", Pougoue_Fichier)
            End If
        End If

        If Trim(txtFileCRPTempon.Text) <> "" Then
            If Strings.Right(Trim(txtFileCRPTempon.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(VENTE) TEMPORAIRE", Trim(txtFileCRPTempon.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(VENTE) TEMPORAIRE", Trim(txtFileCRPTempon.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(txtFileVSTTempon.Text) <> "" Then
            If Strings.Right(Trim(txtFileVSTTempon.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(MVT STOCK) TEMPORAIRE", Trim(txtFileVSTTempon.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(MVT STOCK) TEMPORAIRE", Trim(txtFileVSTTempon.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(txtCheminXfert.Text) <> "" Then
            If Strings.Right(Trim(txtCheminXfert.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(TRANSFERT) TEMPORAIRE", Trim(txtCheminXfert.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(TRANSFERT) TEMPORAIRE", Trim(txtCheminXfert.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(txtCheminErreur.Text) <> "" Then
            If Strings.Right(Trim(txtCheminErreur.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(ERREUR TRANSFORMATION) TEMPORAIRE", Trim(txtCheminErreur.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(ERREUR TRANSFORMATION) TEMPORAIRE", Trim(txtCheminErreur.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(txtCheminMecalux.Text) <> "" Then
            If Strings.Right(Trim(txtCheminMecalux.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(SAUVEGADE FICHIERS MECALUX) TEMPORAIRE", Trim(txtCheminMecalux.Text), Pougoue_Fichier)
            Else
                Bool = WritePrivateProfileString("PARAMETRAGE", "REPERTOIRE(SAUVEGADE FICHIERS MECALUX) TEMPORAIRE", Trim(txtCheminMecalux.Text) & "\", Pougoue_Fichier)
            End If
        End If
        If Trim(TxtAccess.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "NOM FICHIER ACCESS", Trim(TxtAccess.Text), Pougoue_Fichier)
        End If
        If CkConso.Checked = True Then
            Bool = WritePrivateProfileString("CONNECTION", "Statut Connexion", "Oui", Pougoue_Fichier)
        Else
            Bool = WritePrivateProfileString("CONNECTION", "Statut Connexion", "Non", Pougoue_Fichier)
        End If
        MsgBox("Modification Terminée!", MsgBoxStyle.Information, "Modification Fichier Ini")
        Call LirefichierConfig()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub BT_Access_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Access.Click
        OpenFileAccess.Filter = "Fichier Access (*.accdb)|*.mdb"
        OpenFileAccess.FileName = Nothing
        If OpenFileAccess.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtAccess.Text = OpenFileAccess.FileName
        End If
    End Sub

    Private Sub BT_FicJournal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_FicJournal.Click
        FolderRepjournal.Description = "Repertoire de Journalisation"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtFilejr.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BT_FicRep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_FicRep.Click
        FolderRepjournal.Description = "Repertoire de Journalisation"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            Txt_Rep.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub Bt_tiers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_tiers.Click
        FolderRepjournal.Description = "Repertoire de Journalisation"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            Txtiers.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub Bt_Article_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_Article.Click
        FolderRepjournal.Description = "Repertoire de Journalisation"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtArticle.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BtnCSO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCSO.Click
        FolderRepjournal.Description = "Repertoire(CSO) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtFileCSOTempon.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub btnCRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCRP.Click
        FolderRepjournal.Description = "Repertoire(CRP) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtFileCRPTempon.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BtnVST_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnVST.Click
        FolderRepjournal.Description = "Repertoire(VST) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtFileVSTTempon.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BtnCheminXfert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCheminXfert.Click
        FolderRepjournal.Description = "Repertoire(Mvt Transfert) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtCheminXfert.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BtnOpenCheminErreur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOpenCheminErreur.Click
        FolderRepjournal.Description = "Repertoire(ERREUR TRANSFORMATION) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtCheminErreur.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub

    Private Sub BtnCheminMecalux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCheminMecalux.Click
        FolderRepjournal.Description = "Repertoire(Sauvegade Fichier Mecalux) Fichier Temporaire"
        FolderRepjournal.ShowNewFolderButton = True
        FolderRepjournal.SelectedPath = Nothing
        If FolderRepjournal.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtCheminMecalux.Text = FolderRepjournal.SelectedPath & "\"
        End If
    End Sub
End Class
