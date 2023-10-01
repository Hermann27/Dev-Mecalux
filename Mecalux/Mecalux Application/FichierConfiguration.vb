Imports System.IO
Imports System.Net.NetworkInformation
Public Class Frm_FichierConfiguration

    Private Sub Frm_FichierConfiguration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        If StatutConsolider = "Oui" Then
            CkConso.Checked = True
        Else
            CkConso.Checked = False
        End If
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Bool As Boolean
        If Trim(TxtFichierCpta.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "CHEMIN DU FICHIER COMPTA", Trim(TxtFichierCpta.Text), Pouliyou_Fichier)
        End If
        If Trim(TxtBDCpta.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "BASE DE DONNEES COMPTA", Trim(TxtBDCpta.Text), Pouliyou_Fichier)
        End If

        Bool = WritePrivateProfileString("CONNECTION", "UTILISATEUR SQL", Trim(TxtUtilisateur.Text), Pouliyou_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "MOT DE PASSE SQL", Trim(TxtPasw.Text), Pouliyou_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "UTILISATEUR", Trim(TxtUserCpta.Text), Pouliyou_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "MOT DE PASSE", Trim(TxtPasword.Text), Pouliyou_Fichier)
        Bool = WritePrivateProfileString("CONNECTION", "SERVEUR SQL", Trim(Txtsql.Text), Pouliyou_Fichier)
        If Trim(TxtFilejr.Text) <> "" Then
            If Strings.Right(Trim(TxtFilejr.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE JOURNAL", Trim(TxtFilejr.Text), Pouliyou_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE JOURNAL", Trim(TxtFilejr.Text) & "\", Pouliyou_Fichier)
            End If
        End If
        If Trim(Txt_Rep.Text) <> "" Then
            If Strings.Right(Trim(Txt_Rep.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE SAUVEGARDE", Trim(Txt_Rep.Text), Pouliyou_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE SAUVEGARDE", Trim(Txt_Rep.Text) & "\", Pouliyou_Fichier)
            End If
        End If
        If Trim(Txtiers.Text) <> "" Then
            If Strings.Right(Trim(Txtiers.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT TIERS", Trim(Txtiers.Text), Pouliyou_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT TIERS", Trim(Txtiers.Text) & "\", Pouliyou_Fichier)
            End If
        End If
        If Trim(TxtArticle.Text) <> "" Then
            If Strings.Right(Trim(TxtArticle.Text), 1) = "\" Then
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT ARTICLES", Trim(TxtArticle.Text), Pouliyou_Fichier)
            Else
                Bool = WritePrivateProfileString("CONNECTION", "REPERTOIRE FORMAT ARTICLES", Trim(TxtArticle.Text) & "\", Pouliyou_Fichier)
            End If
        End If

        If Trim(TxtAccess.Text) <> "" Then
            Bool = WritePrivateProfileString("CONNECTION", "NOM FICHIER ACCESS", Trim(TxtAccess.Text), Pouliyou_Fichier)
        End If
        If CkConso.Checked = True Then
            Bool = WritePrivateProfileString("CONNECTION", "Statut Connexion", "Oui", Pouliyou_Fichier)
        Else
            Bool = WritePrivateProfileString("CONNECTION", "Statut Connexion", "Non", Pouliyou_Fichier)
        End If
               MsgBox("Modification Terminée!", MsgBoxStyle.Information, "Modification Fichier Ini")
        Call LirefichierConfig()
    End Sub

    Private Sub RadButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
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

    Private Sub BT_Access_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Access.Click
        OpenFileAccess.Filter = "Fichier Access (*.mdb)|*.accdb"
        OpenFileAccess.FileName = Nothing
        If OpenFileAccess.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtAccess.Text = OpenFileAccess.FileName
        End If
    End Sub

    Private Sub BT_FicJournal2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_FicJournal.Click
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
End Class