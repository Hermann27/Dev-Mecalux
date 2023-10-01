Imports System.Data.OleDb
Imports System.IO
Public Class PlanificationTache
    Private Function AfficheIdentifiant() As Double
        Dim OleSocieteAdaptat As OleDbDataAdapter
        Dim OleSocieteDatas As DataSet
        Dim OledatableSocie As DataTable
        OleSocieteAdaptat = New OleDbDataAdapter("select * from TACHEPLANIFIER ORDER BY IDTache DESC", OleConnenection)
        OleSocieteDatas = New DataSet
        OleSocieteAdaptat.Fill(OleSocieteDatas)
        OledatableSocie = OleSocieteDatas.Tables(0)
        If OledatableSocie.Rows.Count <> 0 Then
            AfficheIdentifiant = OledatableSocie.Rows(0).Item("IDTache") + 1
        Else
            AfficheIdentifiant = 1
        End If
    End Function
    Private Sub EnregistrerLatache()
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        If Trim(TxtIntitule.Text) <> "" Then
            If IsNumeric(Trim(TxtIDTache.Text)) = True And InStr(Trim(TxtIDTache.Text), ",") = 0 And InStr(Trim(TxtIDTache.Text), ".") = 0 Then
                OleAdaptaterEnreg = New OleDbDataAdapter("select * From TACHEPLANIFIER WHERE  Intitule='" & Join(Split(Trim(TxtIntitule.Text), "'"), "''") & "'", OleConnenection)
                OleEnregDataset = New DataSet
                OleAdaptaterEnreg.Fill(OleEnregDataset)
                OledatableEnreg = OleEnregDataset.Tables(0)
                If OledatableEnreg.Rows.Count <> 0 Then
                    MsgBox("Tache Planifiée  : " & Trim(TxtIntitule.Text) & " Existe dans la table de paramétrage", MsgBoxStyle.Information, "Planification de taches")
                Else
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From TACHEPLANIFIER WHERE  IDTache=" & CDbl(Trim(TxtIDTache.Text)) & "", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        MsgBox("L'ID de la Tache Planifiée  : " & Trim(TxtIDTache.Text) & " Existe dans la table de paramétrage", MsgBoxStyle.Information, "Planification de taches")
                    Else

                        If ChkLancer.Checked = True Then
                            Insertion = "Insert Into TACHEPLANIFIER (IDTache,Intitule,Lancer) VALUES ('" & CDbl(TxtIDTache.Text) & "','" & Join(Split(Trim(TxtIntitule.Text), "'"), "''") & "',True)"
                            OleCommandEnreg = New OleDbCommand(Insertion)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            Insert = True
                        Else
                            Insertion = "Insert Into TACHEPLANIFIER (IDTache,Intitule,Lancer) VALUES ('" & CDbl(TxtIDTache.Text) & "','" & Join(Split(Trim(TxtIntitule.Text), "'"), "''") & "',False)"
                            OleCommandEnreg = New OleDbCommand(Insertion)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            Insert = True
                        End If
                    End If
                End If
            Else
                MsgBox("L'ID de la tâche doit être un Entier : " & Trim(TxtIDTache.Text) & " Valeur Entière Obligatoire!", MsgBoxStyle.Information, "Planification de taches")
            End If
        Else
            MsgBox("Intitulé tâche Obligatoire !", MsgBoxStyle.Information, "Planification de taches")
        End If

        If Insert = True Then
            AfficheTacheplanifier()
            MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Planification de tâches")
        End If
    End Sub
    Private Sub MiseàjourTachePlanifier()
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Dim i As Integer
        Try
            For i = 0 To DataListeIntegrer.RowCount - 1
                If IsNumeric(Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value)) = True And InStr(Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value), ".") = 0 And InStr(Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value), ",") = 0 Then
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From TACHEPLANIFIER WHERE  Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        If DataListeIntegrer.Rows(i).Cells("Lancer").Value = True Then
                            Insertion = "UPDATE TACHEPLANIFIER SET IDTache='" & CDbl(Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value)) & "',Lancer=True Where Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'"
                            OleCommandEnreg = New OleDbCommand(Insertion)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            Insert = True
                        Else
                            Insertion = "UPDATE TACHEPLANIFIER SET IDTache='" & CDbl(Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value)) & "',Lancer=False Where Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'"
                            OleCommandEnreg = New OleDbCommand(Insertion)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            Insert = True
                        End If
                    End If
                Else
                    MsgBox("L'ID de la tâche doit être un Entier : " & Trim(DataListeIntegrer.Rows(i).Cells("Tache").Value) & " Valeur Entière Obligatoire!", MsgBoxStyle.Information, "Planification de taches")
                End If
            Next i
            If Insert = True Then
                AfficheTacheplanifier()
                MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour Planification de taches")
            End If
        Catch ex As Exception

        End Try

    End Sub
    Dim listeVal As String() = New String(5) {"Export article", "Import article", "Export client", "Import client", "Export fournisseur", "Import fournisseur"}
    Private Sub PlanificationTache_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TxtIntitule.AutoCompleteMode = AutoCompleteMode.Suggest
        TxtIntitule.AutoCompleteSource = AutoCompleteSource.CustomSource
        Dim m As New AutoCompleteStringCollection
        m.Add("Export article")
        m.Add("Import article")
        m.Add("Export client")
        m.Add("Import client")
        TxtIntitule.AutoCompleteCustomSource = m
        If Connected() = True Then
            DataListeIntegrer.Rows.Clear()
            AfficheTacheplanifier()
            Me.TxtIDTache.Focus()
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
    Private Sub AfficheTacheplanifier()
        Dim i As Integer
        DataListeIntegrer.Rows.Clear()
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        OleAdaptaterschema = New OleDbDataAdapter("select * from TACHEPLANIFIER", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        If OledatableSchema.Rows.Count <> 0 Then
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("Tache").Value = OledatableSchema.Rows(i).Item("IDTache")
                DataListeIntegrer.Rows(i).Cells("Intitule").Value = Trim(OledatableSchema.Rows(i).Item("Intitule"))
                DataListeIntegrer.Rows(i).Cells("Lancer").Value = Trim(OledatableSchema.Rows(i).Item("Lancer"))
            Next i
        End If
    End Sub
    Private Sub Delete_DataListeSch()
        Dim i As Integer
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim DelFile As String
        For i = 0 To DataListeIntegrer.RowCount - 1
            If DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True Then
                OleAdaptaterDelete = New OleDbDataAdapter("select * From TACHEPLANIFIER WHERE  Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'", OleConnenection)
                OleDeleteDataset = New DataSet
                OleAdaptaterDelete.Fill(OleDeleteDataset)
                OledatableDelete = OleDeleteDataset.Tables(0)
                If OledatableDelete.Rows.Count <> 0 Then
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From PLANIFICATION WHERE  IntituleTache='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        DelFile = "Delete From PLANIFICATION where IntituleTache='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'"
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                        DelFile = "Delete From TACHEPLANIFIER where Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'"
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                    Else
                        DelFile = "Delete From TACHEPLANIFIER where Intitule='" & Join(Split(Trim(DataListeIntegrer.Rows(i).Cells("Intitule").Value), "'"), "''") & "'"
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                    End If
                End If
            End If
        Next i
        AfficheTacheplanifier()
    End Sub
    Private Sub Bt_Sup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Delete_DataListeSch()
    End Sub
    Private Sub BT_SelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True
        Next i
    End Sub
    Private Sub BT_DelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Supprimer").Value = False
        Next i
    End Sub

    Private Sub Bt_New_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_New.Click
        TxtIDTache.Text = AfficheIdentifiant()
        TxtIntitule.Text = ""
        ChkLancer.Checked = False
    End Sub

    Private Sub Bt_Enregistrer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_Enregistrer.Click
        EnregistrerLatache()
    End Sub

    Private Sub BTupdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTupdate.Click
        MiseàjourTachePlanifier()
    End Sub

    Private Sub BTsup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTsup.Click
        Delete_DataListeSch()
    End Sub

    Private Sub DataListeIntegrer_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        If e.RowIndex >= 0 Then
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Traitement" Then
                'PlanificationSpecial.CbIntitule.Text = ""
                'PlanificationSpecial.CbIntitule.Text = DataListeIntegrer.Rows(e.RowIndex).Cells("Intitule").Value
                PlanificationSpecial.ShowDialog()
            End If
        End If
    End Sub

    Private Sub DataListeIntegrer_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellContentClick

    End Sub
End Class