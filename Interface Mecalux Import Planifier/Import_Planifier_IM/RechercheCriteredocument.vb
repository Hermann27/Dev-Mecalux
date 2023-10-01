Imports System.Data.OleDb
Public Class RechercheCriteredocument
    Private Sub RechercheCriteredocument_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Connected() = True Then
            CbCritere.Text = ""
            Affiche()
            AfficheCritère()
        End If
    End Sub
    Private Sub UpdateParametre()
        Dim OleUpdatAdaptater As OleDbDataAdapter
        Dim OleUpdatDataset As DataSet
        Dim OleDatable As DataTable
        Dim UpdateSociete As String
        Dim OleCommandUpdate As OleDbCommand
        Try
            OleUpdatAdaptater = New OleDbDataAdapter("select * From COLIMPMOUV where  ColDispo='" & Join(Split(Trim(TxtLibelle.Text), "'"), "''") & "' And Libelle='" & Trim(TxtSage.Text) & "' And Fichier='" & Trim(TxtFichier.Text) & "'", OleConnenection)
            OleUpdatDataset = New DataSet
            OleUpdatAdaptater.Fill(OleUpdatDataset)
            OleDatable = OleUpdatDataset.Tables(0)
            If OleDatable.Rows.Count <> 0 Then
                UpdateSociete = "Update COLIMPMOUV SET Champ='" & Trim(CbCritere.Text) & "' where   ColDispo='" & Join(Split(Trim(TxtLibelle.Text), "'"), "''") & "' And Libelle='" & Trim(TxtSage.Text) & "' And Fichier='" & Trim(TxtFichier.Text) & "'"
                OleCommandUpdate = New OleDbCommand(UpdateSociete)
                OleCommandUpdate.Connection = OleConnenection
                OleCommandUpdate.ExecuteNonQuery()
                MsgBox("Modification Effectuée avec Succès!", MsgBoxStyle.Information, "Modification Critère de Recherche")
            Else
                MsgBox("Aucune Modification Effectuée!", MsgBoxStyle.Information, "Modification Critère de Recherche")
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Affiche()
        Dim i As Integer
        Dim OleSocieteAdaptater As OleDbDataAdapter
        Dim OleSocieteDataset As DataSet
        Dim OledatableSociete As DataTable
        Try
            CbCritere.Items.Clear()
            OleSocieteAdaptater = New OleDbDataAdapter("select DISTINCT CHAMP From COLIMPMOUV", OleConnenection)
            OleSocieteDataset = New DataSet
            OleSocieteAdaptater.Fill(OleSocieteDataset)
            OledatableSociete = OleSocieteDataset.Tables(0)
            If OledatableSociete.Rows.Count <> 0 Then
                For i = 0 To OledatableSociete.Rows.Count - 1
                    CbCritere.Items.Add(OledatableSociete.Rows(i).Item("Champ"))
                Next i
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AfficheCritère()
        Dim OleSocieteAdaptater As OleDbDataAdapter
        Dim OleSocieteDataset As DataSet
        Dim OledatableSociete As DataTable
        Try
            OleSocieteAdaptater = New OleDbDataAdapter("select * From COLIMPMOUV where ColDispo='" & Join(Split(Trim(TxtLibelle.Text), "'"), "''") & "' And Libelle='" & Trim(TxtSage.Text) & "' And Fichier='" & Trim(TxtFichier.Text) & "'", OleConnenection)
            OleSocieteDataset = New DataSet
            OleSocieteAdaptater.Fill(OleSocieteDataset)
            OledatableSociete = OleSocieteDataset.Tables(0)
            If OledatableSociete.Rows.Count <> 0 Then
                CbCritere.Text = OledatableSociete.Rows(0).Item("Champ")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BT_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Update.Click
        UpdateParametre()
    End Sub
End Class