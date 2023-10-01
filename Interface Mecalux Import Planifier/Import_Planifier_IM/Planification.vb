Imports System.Data.OleDb

Public Class Planification
    Dim i, j As Integer
    Public numLignList, numLignSelect As Integer
    Dim strDate As String
    Public Requete As String

    'variables base de données
    Dim LibreOleAdaptater As OleDbDataAdapter
    Dim Libredataset As DataSet
    Dim Libredatabase As DataTable
    Dim OleAdaptaterMag As OleDbDataAdapter
    Dim OleMagDataset As DataSet
    Dim OledatableMag As DataTable
    Dim accesscom As OleDbCommand

    Dim DatabaseCpta, ServeurCpta As String

    Private Sub Planification_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Connected() = True Then
            chargementTraitementEnregistre(Trim(CbIntitule.Text))
            chargementTache()
            chargeListeTraitement()
        End If
    End Sub

    'chargement des traitements dejà enregistrés
    Public Sub chargementTraitementEnregistre(ByRef TacheIntitule As String)
        Try
            dgvTraitementEnr.Rows.Clear()
            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION WHERE IntituleTache='" & Join(Split(Trim(TacheIntitule), "'"), "''") & "'order by Rang asc", OleConnenection)
            Libredataset = New DataSet
            LibreOleAdaptater.Fill(Libredataset)
            Libredatabase = Libredataset.Tables(0)
            If Libredatabase.Rows.Count <> 0 Then
                dgvTraitementEnr.RowCount = Libredatabase.Rows.Count
                For i = 0 To Libredatabase.Rows.Count - 1
                    dgvTraitementEnr.Rows(i).Cells("Intitule").Value = Libredatabase.Rows(i).Item("Intitule")
                    dgvTraitementEnr.Rows(i).Cells("IDDossier").Value = Libredatabase.Rows(i).Item("IDDossier")
                    dgvTraitementEnr.Rows(i).Cells("Rang").Value = Libredatabase.Rows(i).Item("Rang")
                    dgvTraitementEnr.Rows(i).Cells("Critere1").Value = Libredatabase.Rows(i).Item("Critere1")
                    dgvTraitementEnr.Rows(i).Cells("Critere2").Value = Libredatabase.Rows(i).Item("Critere2")
                    dgvTraitementEnr.Rows(i).Cells("supEnr").Value = False
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("LastExecution")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Execution").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("LastExecution"), DateFormat.GeneralDate)
                    End If
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("Heure1")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Heure1").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime) 'Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.ShortTime)
                    End If
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("Heure2")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Heure2").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime) ' Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.ShortTime)
                    End If
                Next i
            End If
        Catch ex As Exception
            MessageBox.Show("Erreur Chargement des Traitements Enregistrés: " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub chargementTache()
        Try
            CbIntitule.Items.Clear()
            LibreOleAdaptater = New OleDbDataAdapter("select * from TACHEPLANIFIER order by IDTache asc", OleConnenection)
            Libredataset = New DataSet
            LibreOleAdaptater.Fill(Libredataset)
            Libredatabase = Libredataset.Tables(0)
            If Libredatabase.Rows.Count <> 0 Then
                For i = 0 To Libredatabase.Rows.Count - 1
                    CbIntitule.Items.Add(Libredatabase.Rows(i).Item("Intitule"))
                Next i
            End If
        Catch ex As Exception
            MessageBox.Show("Erreur Chargement des Traitements Enregistrés: " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub dgvTraitementSelect_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            numLignSelect = e.RowIndex
        End If
    End Sub
    'vérification de la non existence d'un traitement
    Private Function verifTraitement(ByRef intitule As String, ByRef rang As Integer) As Boolean
        Try
            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION where Rang=" & rang, OleConnenection)
            Libredataset = New DataSet
            LibreOleAdaptater.Fill(Libredataset)
            Libredatabase = Libredataset.Tables(0)
            If Libredatabase.Rows.Count <> 0 Then
                MessageBox.Show("Il éxiste déjà un traitement en position '" & rang & "'" & Chr(13) & "Nom du Traitement présent : " & Libredatabase.Rows(0).Item("Intitule"), "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            Else
                LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION where Intitule='" & intitule & "'", OleConnenection)
                Libredataset = New DataSet
                LibreOleAdaptater.Fill(Libredataset)
                Libredatabase = Libredataset.Tables(0)
                If Libredatabase.Rows.Count <> 0 Then
                    MessageBox.Show("Le traitement '" & intitule & "' est déjà enregistré en position '" & Libredatabase.Rows(0).Item("Rang") & "'", "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Erreur Vérification des Traitements Enregistrés : " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub btnSupprimer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If dgvTraitementEnr.RowCount > 0 Then
            For i = dgvTraitementEnr.RowCount - 1 To 0 Step -1
                If dgvTraitementEnr.Rows(i).Cells("supEnr").Value = True Then
                    accesscom = New OleDbCommand("delete from PLANIFICATION where IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & " and Intitule='" & dgvTraitementEnr.Rows(i).Cells("Intitule").Value & "' and IntituleTache='" & dgvTraitementEnr.Rows(i).Cells("IntituleTache").Value & "'", OleConnenection)
                    accesscom.ExecuteNonQuery()
                End If
            Next i
            chargementTraitementEnregistre(Join(Split(Trim(CbIntitule.Text), "'"), "''"))
        End If
    End Sub

    Private Sub CbIntitule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbIntitule.SelectedIndexChanged
        chargementTraitementEnregistre(Trim(CbIntitule.SelectedItem))
    End Sub

    Private Sub CbIntitule_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CbIntitule.SelectedValueChanged
        chargementTraitementEnregistre(Trim(CbIntitule.Text))
    End Sub

    Private Sub dgvListeTrait_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvListeTrait.CellClick
        If e.RowIndex >= 0 Then
            numLignList = e.RowIndex

            If dgvListeTrait.Rows(numLignList).Cells("typeTrait").Value = "A" Then
                'PlanificationTraitement.choixRef.Enabled = False
            Else
                If dgvListeTrait.Rows(numLignList).Cells("typeTrait").Value = "B1" Then
                    'PlanificationTraitement.choixRef.Enabled = True
                End If
            End If
            If dgvListeTrait.Columns(e.ColumnIndex).Name = "Consulter" Then
                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Export Ecriture" Then
                    Requete = "Select * From WEE_SCHEMA WHERE IDDossier<>0"
                Else
                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Export Document" Then
                        Requete = "Select * From WED_SCHEMA WHERE IDDossier<>0"
                    Else
                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Ecriture" Then
                            Requete = "Select * From SCHEMASIE WHERE IDDossier<>0"
                        Else
                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Ecriture Readsoft" Then
                                Requete = "Select * From SOCIETEIMPORT_FICHEXML WHERE IDDossier<>0"
                            Else
                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Tiers" Then
                                    Requete = "Select * From SCHEMASI WHERE IDDossier<>0"
                                Else
                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Document Vente" Then
                                        Requete = "Select * From SCHEMAS_IMPMOUV WHERE IDDossier<>0"
                                    Else
                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Document Achat" Then
                                            Requete = "Select * From SCHEMAS_IMPMOUV WHERE IDDossier<>0"
                                        Else
                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Article" Then
                                                Requete = "Select * From SCHEMASIEART WHERE IDDossier<>0"
                                            Else
                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Devis" Then
                                                    Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='DEVIS' And IDDossier<>0"
                                                Else
                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Commande Vente" Then
                                                        Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='COMMANDE VENTE' And IDDossier<>0"
                                                    Else
                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Bon de Livraison" Then
                                                            Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='B.L' And IDDossier<>0"
                                                        Else
                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Facture Vente" Then
                                                                Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='FACTURE VENTE' And IDDossier<>0"
                                                            Else
                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Commande Achat" Then
                                                                    Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='COMMANDE ACHAT' And IDDossier<>0"
                                                                Else
                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Bon de Reception" Then
                                                                        Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='B. R. VALORISE' And IDDossier<>0"
                                                                    Else
                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EANCOM Facture Achat" Then
                                                                            Requete = "Select * From PARACOMMERCIAL Where Format='EANCOM' And Piece='FACTURE ACHAT' And IDDossier<>0"
                                                                        Else
                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Devis" Then
                                                                                Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='DEVIS'  And IDDossier<>0"
                                                                            Else
                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Commande Vente" Then
                                                                                    Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='COMMANDE VENTE'  And IDDossier<>0"
                                                                                Else
                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Bon de Livraison" Then
                                                                                        Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='B.L'  And IDDossier<>0"
                                                                                    Else
                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Facture Vente" Then
                                                                                            Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='FACTURE VENTE'  And IDDossier<>0"
                                                                                        Else
                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Commande Achat" Then
                                                                                                Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='COMMANDE ACHAT'  And IDDossier<>0"
                                                                                            Else
                                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Bon de Reception" Then
                                                                                                    Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='B. R. VALORISE'  And IDDossier<>0"
                                                                                                Else
                                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Sage 100 Facture Achat" Then
                                                                                                        Requete = "Select * From PARACOMMERCIAL Where Format='GESCOM100' And Piece='FACTURE ACHAT'  And IDDossier<>0"
                                                                                                    Else
                                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Devis" Then
                                                                                                            Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='DEVIS'  And IDDossier<>0"
                                                                                                        Else
                                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Commande Vente" Then
                                                                                                                Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='COMMANDE VENTE' And IDDossier<>0"
                                                                                                            Else
                                                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Bon de Livraison" Then
                                                                                                                    Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='B.L' And IDDossier<>0"
                                                                                                                Else
                                                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Facture Vente" Then
                                                                                                                        Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='FACTURE VENTE' And IDDossier<>0"
                                                                                                                    Else
                                                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Commande Achat" Then
                                                                                                                            Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='COMMANDE ACHAT' And IDDossier<>0"
                                                                                                                        Else
                                                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Bon de Reception" Then
                                                                                                                                Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='B. R. VALORISE' And IDDossier<>0"
                                                                                                                            Else
                                                                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import EAN96 Facture Achat" Then
                                                                                                                                    Requete = "Select * From PARACOMMERCIAL Where Format='EAN96' And Piece='FACTURE ACHAT' And IDDossier<>0"
                                                                                                                                Else
                                                                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Compte Analytique" Then
                                                                                                                                        Requete = "Select * From WICA_SCHEMA WHERE IDDossier<>0"
                                                                                                                                    Else
                                                                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Export Article" Then
                                                                                                                                            Requete = "Select * From WEA_SCHEMA WHERE IDDossier<>0"
                                                                                                                                        Else
                                                                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Export Tiers" Then
                                                                                                                                                Requete = "Select * From WET_SCHEMA WHERE IDDossier<>0"
                                                                                                                                            Else
                                                                                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Document Stock" Then
                                                                                                                                                    Requete = "Select * From WIS_SCHEMA WHERE IDDossier<>0"
                                                                                                                                                Else
                                                                                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import Document Transfert" Then
                                                                                                                                                        Requete = "Select * From WIT_SCHEMA WHERE IDDossier<>0"
                                                                                                                                                    Else
                                                                                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Import BL Modification" Then
                                                                                                                                                            Requete = "Select * From WI_SCHEMAS_IMPBL WHERE IDDossier<>0"
                                                                                                                                                        Else
                                                                                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Rafale Compte Général" Then
                                                                                                                                                                Requete = "select * from NomChemin where Statut='Master' And IDDossier<>0"
                                                                                                                                                            Else
                                                                                                                                                                If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Rafale Compte Tiers" Then
                                                                                                                                                                    Requete = "select * from NomChemin where Statut='Master' And IDDossier<>0"
                                                                                                                                                                Else
                                                                                                                                                                    If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Rafale Compte Analytique" Then
                                                                                                                                                                        Requete = "select * from NomChemin where Statut='Master' And IDDossier<>0"
                                                                                                                                                                    Else
                                                                                                                                                                        If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Rafale Compte Article" Then
                                                                                                                                                                            Requete = "select * from NomChemin where Statut='Master' And IDDossier<>0"
                                                                                                                                                                        Else
                                                                                                                                                                            If dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value = "Ecritures Intercos" Then
                                                                                                                                                                                Requete = "Select * From SOCIETEINTERCOS WHERE IDDossier<>0"
                                                                                                                                                                            Else

                                                                                                                                                                            End If
                                                                                                                                                                        End If
                                                                                                                                                                    End If
                                                                                                                                                                End If
                                                                                                                                                            End If
                                                                                                                                                        End If
                                                                                                                                                    End If
                                                                                                                                                End If
                                                                                                                                            End If
                                                                                                                                        End If
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                PlanificationTraitement.ShowDialog()
            End If
        End If
    End Sub
    'chargement de la liste des traitements
    Private Sub chargeListeTraitement()
        Try
            dgvListeTrait.Rows.Clear()
            dgvListeTrait.Rows.Add("Export Ecriture", "A")
            dgvListeTrait.Rows.Add("Export Document", "A")
            dgvListeTrait.Rows.Add("Export Article", "A")
            dgvListeTrait.Rows.Add("Export Tiers", "A")
            dgvListeTrait.Rows.Add("Import Ecriture", "B1")
            dgvListeTrait.Rows.Add("Import Compte Analytique", "B1")
            dgvListeTrait.Rows.Add("Import Ecriture Readsoft", "A")
            dgvListeTrait.Rows.Add("Import Tiers", "B1")
            dgvListeTrait.Rows.Add("Import Document Vente", "B1")
            dgvListeTrait.Rows.Add("Import Document Achat", "B1")
            dgvListeTrait.Rows.Add("Import Article", "A")
            dgvListeTrait.Rows.Add("Import Document Stock", "B1")
            dgvListeTrait.Rows.Add("Import Document Transfert", "B1")
            dgvListeTrait.Rows.Add("Import BL Modification", "B1")
            dgvListeTrait.Rows.Add("Import EANCOM Devis", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Commande Vente", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Bon de Livraison", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Facture Vente", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Commande Achat", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Bon de Reception", "A")
            dgvListeTrait.Rows.Add("Import EANCOM Facture Achat", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Devis", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Commande Vente", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Bon de Livraison", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Facture Vente", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Commande Achat", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Bon de Reception", "A")
            dgvListeTrait.Rows.Add("Import EAN96 Facture Achat", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Devis", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Commande Vente", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Bon de Livraison", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Facture Vente", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Commande Achat", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Bon de Reception", "A")
            dgvListeTrait.Rows.Add("Import Sage 100 Facture Achat", "A")
            dgvListeTrait.Rows.Add("Rafale Compte Général", "A")
            dgvListeTrait.Rows.Add("Rafale Compte Tiers", "A")
            dgvListeTrait.Rows.Add("Rafale Compte Analytique", "A")
            dgvListeTrait.Rows.Add("Rafale Compte Article", "A")
            dgvListeTrait.Rows.Add("Ecritures Intercos", "B1")
        Catch ex As Exception
            MessageBox.Show("Erreur Chargement de la Liste des Traitements : " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub dgvListeTrait_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvListeTrait.CellContentClick

    End Sub

    Private Sub BTsup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTsup.Click
        Delete_DataListeSch()
    End Sub
    Private Sub Delete_DataListeSch()
        Dim i As Integer
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim DelFile As String
        For i = 0 To dgvTraitementEnr.RowCount - 1
            If dgvTraitementEnr.Rows(i).Cells("supEnr").Value = True Then
                OleAdaptaterDelete = New OleDbDataAdapter("select * From PLANIFICATION WHERE  IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
                OleDeleteDataset = New DataSet
                OleAdaptaterDelete.Fill(OleDeleteDataset)
                OledatableDelete = OleDeleteDataset.Tables(0)
                If OledatableDelete.Rows.Count <> 0 Then
                    DelFile = "Delete From PLANIFICATION WHERE IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & ""
                    OleCommandDelete = New OleDbCommand(DelFile)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                End If
            End If
        Next i
        chargementTraitementEnregistre(Trim(CbIntitule.Text))
    End Sub

    Private Sub dgvTraitementEnr_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTraitementEnr.CellClick
        If e.RowIndex >= 0 Then
            Idexe = e.RowIndex
            If dgvTraitementEnr.Columns(e.ColumnIndex).Name = "Heure" Then
                PlanificationHeure.Text = "Planification"
                PlanificationHeure.ShowDialog()
            Else
                If dgvTraitementEnr.Columns(e.ColumnIndex).Name = "DateM" Then
                    PlanificationDate.Text = "Planification"
                    PlanificationDate.ShowDialog()
                End If
            End If
        End If
    End Sub

    Private Sub dgvTraitementEnr_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTraitementEnr.CellContentClick

    End Sub

    Private Sub BTupdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTupdate.Click
        MiseàjourTachePlanifier()
    End Sub
    Private Sub MiseàjourTachePlanifier()
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim Insert As Boolean = False
        Dim i As Integer
        Try
            For i = 0 To dgvTraitementEnr.RowCount - 1
                If IsNumeric(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value)) = True And InStr(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value), ".") = 0 And InStr(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value), ",") = 0 Then
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From PLANIFICATION WHERE  IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        If Trim(dgvTraitementEnr.Rows(i).Cells("Heure1").Value) <> "" And Trim(dgvTraitementEnr.Rows(i).Cells("Heure2").Value) <> "" Then
                            accesscom = New OleDbCommand("UPDATE  PLANIFICATION Set Rang=" & CInt(dgvTraitementEnr.Rows(i).Cells("Rang").Value) & ",Critere1='" & dgvTraitementEnr.Rows(i).Cells("Critere1").Value & "',Critere2='" & dgvTraitementEnr.Rows(i).Cells("Critere2").Value & "',Heure1='" & Strings.FormatDateTime(dgvTraitementEnr.Rows(i).Cells("Heure1").Value, DateFormat.LongTime) & "',Heure2='" & Strings.FormatDateTime(dgvTraitementEnr.Rows(i).Cells("Heure2").Value, DateFormat.LongTime) & "' WHERE IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
                            accesscom.ExecuteNonQuery()
                            Insert = True
                        Else
                            If Trim(dgvTraitementEnr.Rows(i).Cells("Heure1").Value) = "" And Trim(dgvTraitementEnr.Rows(i).Cells("Heure2").Value) = "" Then
                                accesscom = New OleDbCommand("UPDATE  PLANIFICATION Set Rang=" & CInt(dgvTraitementEnr.Rows(i).Cells("Rang").Value) & ",Critere1='" & dgvTraitementEnr.Rows(i).Cells("Critere1").Value & "',Critere2='" & dgvTraitementEnr.Rows(i).Cells("Critere2").Value & "',Heure1=NULL,Heure2=NULL WHERE IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
                                accesscom.ExecuteNonQuery()
                                Insert = True
                            Else
                                MsgBox("Les Valeurs Heures doivent être tous renseignées ou Null !", MsgBoxStyle.Information, "Création des traitement")
                            End If
                        End If
                    End If
                Else
                    MsgBox("Le Rang du traitement doit être un Entier : " & Trim(dgvTraitementEnr.Rows(i).Cells("Tache").Value) & " Valeur Entière Obligatoire!", MsgBoxStyle.Information, "Planification de taches")
                End If
            Next i
            If Insert = True Then
                MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour Planification de taches")
            End If
        Catch ex As Exception

        End Try
        chargementTraitementEnregistre(Trim(CbIntitule.Text))
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub
End Class