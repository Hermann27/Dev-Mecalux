Imports System.Data.OleDb
Public Class Transcodage
    Public Num_Count As Integer
    Private Sub Transcodage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Connected() = True Then
            Me.WindowState = FormWindowState.Maximized
            Initialiser()
            InitialiserFichier()
        End If
    End Sub
    Private Sub UpdateParametre()
        Dim OleUpdatAdaptater As OleDbDataAdapter
        Dim OleUpdatDataset As DataSet
        Dim OleDatable As DataTable
        Dim Up As Boolean = False
        Dim UpdateSociete As String
        Dim i As Integer
        Dim OleCommandUpdate As OleDbCommand
        For i = 0 To Datacompte1.RowCount - 1
            If Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value) <> "" And Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) <> "" And (Datacompte1.Rows(i).Cells("Ligne1").Value = True Or Datacompte1.Rows(i).Cells("Entete1").Value = True) Then
                OleUpdatAdaptater = New OleDbDataAdapter("select * From TRANSCODAGEIMPORT where Menu='" & Trim(Datacompte1.Rows(i).Cells("Menu1").Value) & "' and IDDossier=" & CInt(Datacompte1.Rows(i).Cells("IDTraitement").Value) & " and Categorie='" & Trim(Datacompte1.Rows(i).Cells("Categorie1").Value) & "' And Concerne='" & Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) & "' and Valeurlue='" & Join(Split(Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value), "'"), "''") & "' And Ligne=" & Datacompte1.Rows(i).Cells("Ligne1").Value & " And Entete=" & Datacompte1.Rows(i).Cells("Entete1").Value & "", OleConnenection)
                OleUpdatDataset = New DataSet
                OleUpdatAdaptater.Fill(OleUpdatDataset)
                OleDatable = OleUpdatDataset.Tables(0)
                If OleDatable.Rows.Count <> 0 Then
                    UpdateSociete = "Update TRANSCODAGEIMPORT SET Correspond='" & Join(Split(Trim(Datacompte1.Rows(i).Cells("CompteImport1").Value), "'"), "''") & "' where  Menu='" & Trim(Datacompte1.Rows(i).Cells("Menu1").Value) & "' and IDDossier=" & CInt(Datacompte1.Rows(i).Cells("IDTraitement").Value) & " and Categorie='" & Trim(Datacompte1.Rows(i).Cells("Categorie1").Value) & "' And Concerne='" & Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) & "' and Valeurlue='" & Join(Split(Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value), "'"), "''") & "' And Ligne=" & Datacompte1.Rows(i).Cells("Ligne1").Value & " And Entete=" & Datacompte1.Rows(i).Cells("Entete1").Value & ""
                    OleCommandUpdate = New OleDbCommand(UpdateSociete)
                    OleCommandUpdate.Connection = OleConnenection
                    OleCommandUpdate.ExecuteNonQuery()
                    Up = True
                End If
            End If
        Next i
        Initialiser()
        If Up = True Then
            MsgBox("Modification Effectuée avec Succès!", MsgBoxStyle.Information, "Modification Transcodage")
        Else
            MsgBox("Aucune Modification Effectuée!", MsgBoxStyle.Information, "Modification Transcodage")
        End If
    End Sub
    Private Sub Creationperiode()
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim i As Integer
        Dim Insert As Boolean = False
        Dim Insertion As String
        If IsNumeric(Cbtraitement.Text) = True Then
            For i = 0 To DataCompte.RowCount - 1

                If Trim(DataCompte.Rows(i).Cells("CompteFichier").Value) <> "" And Trim(DataCompte.Rows(i).Cells("Fichier").Value) <> "" And (DataCompte.Rows(i).Cells("Ligne").Value = True Or DataCompte.Rows(i).Cells("Entete").Value = True) Then
                    Dim l = Replace(Trim(DataCompte.Rows(i).Cells("CompteFichier").Value), "'", "''")
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From TRANSCODAGEIMPORT where   Concerne='" & Trim(DataCompte.Rows(i).Cells("Fichier").Value) & "' and Valeurlue='" & Replace(Trim(DataCompte.Rows(i).Cells("CompteFichier").Value), "'", "''") & "' And Menu='" & DataCompte.Rows(i).Cells("Menu2").Value & "' And IDDossier=" & CInt(Cbtraitement.Text) & " And Categorie='" & Trim(CbCat.Text) & "' And Ligne=" & DataCompte.Rows(i).Cells("Ligne").Value & " And Entete=" & DataCompte.Rows(i).Cells("Entete").Value & "", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then

                    Else                                                                                                                                                                                                                                                                           'Trim(DataCompte.Rows(i).Cells("CompteImport").Value)
                        Insertion = "Insert Into TRANSCODAGEIMPORT (Concerne,Valeurlue,Correspond,Menu,Categorie,Ligne,Entete,IDDossier) VALUES ('" & Trim(DataCompte.Rows(i).Cells("Fichier").Value) & "','" & Replace(Trim(DataCompte.Rows(i).Cells("CompteFichier").Value), "'", "''") & "','" & Replace(Trim(DataCompte.Rows(i).Cells("CompteImport").Value), "'", "''") & "','" & Trim(DataCompte.Rows(i).Cells("Menu2").Value) & "','" & Trim(CbCat.Text) & "'," & DataCompte.Rows(i).Cells("Ligne").Value & "," & DataCompte.Rows(i).Cells("Entete").Value & "," & CInt(Cbtraitement.Text) & " ) "
                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    End If
                End If
            Next i
        Else
            MsgBox("ID Traitement n'est pas numérique", MsgBoxStyle.Information, "Insertion Transcodage")
        End If        
        Initialiser()
        If Insert = True Then
            MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Insertion Transcodage")
            DataCompte.Rows.Clear()
        Else
            MsgBox("Aucun Element créé", MsgBoxStyle.Information, "Insertion Transcodage")
        End If
    End Sub
    Private Sub supprimeperiode()
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim Supp As Boolean = False
        Dim i As Integer
        Dim DelFile As String
        For i = 0 To Datacompte1.RowCount - 1
            If Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value) <> "" And Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) <> "" Then
                If Datacompte1.Rows(i).Cells("Supprime1").Value = True Then
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From TRANSCODAGEIMPORT where Menu='" & Trim(Datacompte1.Rows(i).Cells("Menu1").Value) & "' and IDDossier=" & CInt(Datacompte1.Rows(i).Cells("IDTraitement").Value) & " and Categorie='" & Trim(Datacompte1.Rows(i).Cells("Categorie1").Value) & "' And Concerne='" & Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) & "' and Valeurlue='" & Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value) & "' And Ligne=" & Datacompte1.Rows(i).Cells("Ligne1").Value & " And Entete=" & Datacompte1.Rows(i).Cells("Entete1").Value & "", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        DelFile = "Delete From TRANSCODAGEIMPORT where Menu='" & Trim(Datacompte1.Rows(i).Cells("Menu1").Value) & "' and IDDossier=" & CInt(Datacompte1.Rows(i).Cells("IDTraitement").Value) & " and Categorie='" & Trim(Datacompte1.Rows(i).Cells("Categorie1").Value) & "' And Concerne='" & Trim(Datacompte1.Rows(i).Cells("Fichier1").Value) & "' and Valeurlue='" & Trim(Datacompte1.Rows(i).Cells("CompteFichier1").Value) & "' And Ligne=" & Datacompte1.Rows(i).Cells("Ligne1").Value & " And Entete=" & Datacompte1.Rows(i).Cells("Entete1").Value & ""
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                        Supp = True
                    End If
                End If
            End If
        Next i
        Initialiser()
    End Sub
    Private Sub Initialiser()
        Dim OleSocieteAdaptater As OleDbDataAdapter
        Dim OleSocieteDataset As DataSet
        Dim OledatableSociete As DataTable
        OleSocieteAdaptater = New OleDbDataAdapter("select * From TRANSCODAGEIMPORT", OleConnenection)
        OleSocieteDataset = New DataSet
        OleSocieteAdaptater.Fill(OleSocieteDataset)
        OledatableSociete = OleSocieteDataset.Tables(0)
        Dim i As Integer
        Datacompte1.Rows.Clear()
        If OledatableSociete.Rows.Count <> 0 Then
            Datacompte1.RowCount = OledatableSociete.Rows.Count
            For i = 0 To Datacompte1.RowCount - 1
                Datacompte1.Rows(i).Cells("Fichier1").Value = OledatableSociete.Rows(i).Item("Concerne")
                Datacompte1.Rows(i).Cells("CompteFichier1").Value = OledatableSociete.Rows(i).Item("Valeurlue")
                Datacompte1.Rows(i).Cells("CompteImport1").Value = OledatableSociete.Rows(i).Item("Correspond")
                Datacompte1.Rows(i).Cells("Categorie1").Value = OledatableSociete.Rows(i).Item("Categorie")
                Datacompte1.Rows(i).Cells("Menu1").Value = OledatableSociete.Rows(i).Item("Menu")
                Datacompte1.Rows(i).Cells("Ligne1").Value = OledatableSociete.Rows(i).Item("Ligne")
                Datacompte1.Rows(i).Cells("Entete1").Value = OledatableSociete.Rows(i).Item("Entete")
                Datacompte1.Rows(i).Cells("IDTraitement").Value = OledatableSociete.Rows(i).Item("IDDossier")
            Next i
        End If
    End Sub
    Private Sub InitialiserFichier()
        Dim OleSocieteAdaptater As OleDbDataAdapter
        Dim OleSocieteDataset As DataSet
        Dim OledatableSociete As DataTable
        OleSocieteAdaptater = New OleDbDataAdapter("select * From S_FICHIER", OleConnenection)
        OleSocieteDataset = New DataSet
        OleSocieteAdaptater.Fill(OleSocieteDataset)
        OledatableSociete = OleSocieteDataset.Tables(0)
        Dim i As Integer
        Fichier.Items.Clear()
        If OledatableSociete.Rows.Count <> 0 Then
            For i = 0 To OledatableSociete.Rows.Count - 1
                Fichier.Items.Add(Trim(OledatableSociete.Rows(i).Item("Fichier")))
            Next i
        End If
    End Sub
    Private Sub BT_Creer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Creer.Click
        Creationperiode()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Initialiser()
    End Sub

    Private Sub BT_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Update.Click
        UpdateParametre()
    End Sub

    Private Sub BT_Del_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Del.Click
        supprimeperiode()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub BT_DelRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelRow.Click
        Dim first As Integer
        Dim last As Integer
        first = DataCompte.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
        last = DataCompte.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        If last >= 0 Then
            If last - first >= 0 Then
                DataCompte.Rows.RemoveAt(DataCompte.CurrentRow.Index)
            End If
        End If
    End Sub

    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        Dim i As Integer = DataCompte.Rows.Add()
        DataCompte.Rows(i).Cells("Ligne").Value = False
        DataCompte.Rows(i).Cells("Entete").Value = False
    End Sub

    Private Sub DataCompte_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataCompte.CellClick
        Try
            If e.RowIndex >= 0 Then
                DataCompte.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataCompte_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataCompte.CellContentClick
        Try
            If e.RowIndex >= 0 Then
                DataCompte.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                DataCompte.EndEdit()
                If DataCompte.Columns(e.ColumnIndex).Name = "Entete" Then
                    If DataCompte.Rows(e.RowIndex).Cells("Ligne").Value = True Then
                        DataCompte.Rows(e.RowIndex).Cells("Entete").Value = False
                    End If
                End If
                If DataCompte.Columns(e.ColumnIndex).Name = "Ligne" Then
                    If DataCompte.Rows(e.RowIndex).Cells("Entete").Value = True Then
                        DataCompte.Rows(e.RowIndex).Cells("Ligne").Value = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataCompte_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataCompte.CellContentDoubleClick
        Try
            If e.RowIndex >= 0 Then
                DataCompte.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                DataCompte.EndEdit()
                If DataCompte.Columns(e.ColumnIndex).Name = "Entete" Then
                    If DataCompte.Rows(e.RowIndex).Cells("Ligne").Value = True Then
                        DataCompte.Rows(e.RowIndex).Cells("Entete").Value = False
                    End If
                End If
                If DataCompte.Columns(e.ColumnIndex).Name = "Ligne" Then
                    If DataCompte.Rows(e.RowIndex).Cells("Entete").Value = True Then
                        DataCompte.Rows(e.RowIndex).Cells("Ligne").Value = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataCompte_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataCompte.CellDoubleClick
        Try
            If e.RowIndex >= 0 Then
                DataCompte.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Datacompte1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Datacompte1.CellClick
        Try
            If e.RowIndex >= 0 Then
                Datacompte1.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Datacompte1_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Datacompte1.CellContentClick
        Try
            If e.RowIndex >= 0 Then
                Datacompte1.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                Datacompte1.EndEdit()
                If Datacompte1.Columns(e.ColumnIndex).Name = "Entete1" Then
                    If Datacompte1.Rows(e.RowIndex).Cells("Ligne1").Value = True Then
                        Datacompte1.Rows(e.RowIndex).Cells("Entete1").Value = False
                    End If
                End If
                If Datacompte1.Columns(e.ColumnIndex).Name = "Ligne1" Then
                    If Datacompte1.Rows(e.RowIndex).Cells("Entete1").Value = True Then
                        Datacompte1.Rows(e.RowIndex).Cells("Ligne1").Value = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Datacompte1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Datacompte1.CellContentDoubleClick
        Try
            If e.RowIndex >= 0 Then
                Datacompte1.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                Datacompte1.EndEdit()
                If Datacompte1.Columns(e.ColumnIndex).Name = "Entete1" Then
                    If Datacompte1.Rows(e.RowIndex).Cells("Ligne1").Value = True Then
                        Datacompte1.Rows(e.RowIndex).Cells("Entete1").Value = False
                    End If
                End If
                If Datacompte1.Columns(e.ColumnIndex).Name = "Ligne1" Then
                    If Datacompte1.Rows(e.RowIndex).Cells("Entete1").Value = True Then
                        Datacompte1.Rows(e.RowIndex).Cells("Ligne1").Value = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Datacompte1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Datacompte1.CellDoubleClick
        Try
            If e.RowIndex >= 0 Then
                Datacompte1.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CbCat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbCat.SelectedIndexChanged
        InitialiserIDtraitement(RenvoitableCategorie(CbCat.Text))
    End Sub
    Public Function RenvoitableCategorie(ByRef Categorietrait As String) As String
        If Trim(Categorietrait) = "Document" Then
            Return "SCHEMAS_IMPMOUV"
        Else
            If Trim(Categorietrait) = "Ecritures" Then
                Return "SCHEMASIE"
            Else
                If Trim(Categorietrait) = "Articles" Then
                    Return "SCHEMASIEART"
                Else
                    If Trim(Categorietrait) = "Tiers" Then
                        Return "SCHEMASI"
                    Else
                        If Trim(Categorietrait) = "CompteA" Then
                            Return "WICA_SCHEMA"
                        Else
                            If Trim(Categorietrait) = "Document Stock" Then
                                Return "WIS_SCHEMA"
                            Else
                                If Trim(Categorietrait) = "Document Transfert" Then
                                    Return "WIT_SCHEMA"
                                Else
                                    If Trim(Categorietrait) = "ModificationBL" Then
                                        Return "WICA_SCHEMA"
                                    Else
                                        Return Nothing
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Function
    Private Sub InitialiserIDtraitement(ByRef Categorietable As String)
        Cbtraitement.Items.Clear()
        If Trim(Categorietable) <> "" Then
            Dim OleSocieteAdaptater As OleDbDataAdapter
            Dim OleSocieteDataset As DataSet
            Dim OledatableSociete As DataTable
            OleSocieteAdaptater = New OleDbDataAdapter("select * From " & Trim(Categorietable) & "", OleConnenection)
            OleSocieteDataset = New DataSet
            OleSocieteAdaptater.Fill(OleSocieteDataset)
            OledatableSociete = OleSocieteDataset.Tables(0)
            If OledatableSociete.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSociete.Rows.Count - 1
                    Cbtraitement.Items.Add(Trim(OledatableSociete.Rows(i).Item("IDDossier")))
                Next i
            End If
        End If
    End Sub
End Class