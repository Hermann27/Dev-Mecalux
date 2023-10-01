Imports System.Data.OleDb

Public Class PlanificationSpecial
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
            'chargementTache()
            chargeListeTraitement()
        End If
    End Sub

    'chargement des traitements dejà enregistrés
    Public Shared OledatableSchema As DataTable
    Dim listeVal As String() = New String(5) {"Export Article", "Import Article", "Export Client", "Import Client", "Export fournisseur", "Import fournisseur"}
    Public Sub chargementTraitementEnregistre(ByRef TacheIntitule As String)
        Try

            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            OledatableSchema = Nothing
            OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenection)
            'OleSchemaDataset.Tables.Clear()
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)

            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION WHERE IntituleTache='" & TacheIntitule & "' order by Rang asc", OleConnenection)
            Libredataset = New DataSet
            LibreOleAdaptater.Fill(Libredataset)
            Libredatabase = Libredataset.Tables(0)
            If Libredatabase.Rows.Count <> 0 Then
                dgvTraitementEnr.RowCount = 1 ' Libredatabase.Rows.Count
                For i = 0 To 1 - 1 'Libredatabase.Rows.Count - 1
                    dgvTraitementEnr.Rows(i).Cells("Intitule").Value = Libredatabase.Rows(i).Item("IntituleTache")
                    dgvTraitementEnr.Rows(i).Cells("Societe1").Value = Libredatabase.Rows(i).Item("Intitule")
                    dgvTraitementEnr.Rows(i).Cells("Heure1").Value = Libredatabase.Rows(i).Item("Heure1")
                    dgvTraitementEnr.Rows(i).Cells("Heure2").Value = Libredatabase.Rows(i).Item("Heure2")
                    dgvTraitementEnr.Rows(i).Cells("Rang").Value = Libredatabase.Rows(i).Item("Rang")
                    dgvTraitementEnr.Rows(i).Cells("Critere1").Value = Libredatabase.Rows(i).Item("Critere1")
                    dgvTraitementEnr.Rows(i).Cells("Critere2").Value = Libredatabase.Rows(i).Item("Critere2")
                    dgvTraitementEnr.Rows(i).Cells("Id").Value = Libredatabase.Rows(i).Item("Id")
                    dgvTraitementEnr.Rows(i).Cells("Lancer").Value = Libredatabase.Rows(i).Item("Lancer")
                    'Dim assignedToColint As DataGridViewComboBoxColumn = dgvTraitementEnr.Columns("IntituleTraits")
                    'assignedToColint.DataSource = listeVal


                    'For j As Integer = 0 To assignedToColint.Items.Count - 1
                    '    Dim p = assignedToColint.Items(j).ToString
                    '    Dim jk = Libredatabase.Rows(i).Item("IntituleTache")
                    '    If assignedToColint.Items(j).ToString.Equals(Libredatabase.Rows(i).Item("IntituleTache")) Then
                    '        dgvTraitementEnr.Rows(i).Cells("IntituleTraits").Value = assignedToColint.Items(j).ToString
                    '    End If
                    'Next
                    'Dim assignedToColumn As DataGridViewComboBoxColumn = dgvTraitementEnr.Columns("Societe")
                    'Dim l As Integer = OledatableSchema.Rows.Count
                    'assignedToColumn.DataSource = Nothing
                    'assignedToColumn.DataSource = OledatableSchema
                    'assignedToColumn.DisplayMember = "Societe"
                    'assignedToColumn.ValueMember = "Societe"
                    'For j As Integer = 0 To assignedToColumn.Items.Count - 1
                    '    Dim vt As DataRowView = assignedToColumn.Items(j)
                    '    If vt.Row(0).ToString.Equals(Libredatabase.Rows(i).Item("Intitule")) Then
                    '        dgvTraitementEnr.Rows(i).Cells("Societe").Value = vt.Row(0).ToString
                    '    End If
                    'Next
                    dgvTraitementEnr.Rows(i).Cells("Societe").Value = "Tous Les Sociétés"
                    dgvTraitementEnr.Rows(i).Cells("supEnr").Value = False
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("LastExecution")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Execution").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("LastExecution"), DateFormat.GeneralDate)
                        dgvTraitementEnr.Rows(i).Cells("Etat").Value = My.Resources.accepter
                    End If
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("Heure1")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Heure1").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure1"), DateFormat.LongTime)
                    End If
                    If Convert.IsDBNull(Libredatabase.Rows(i).Item("Heure2")) = False Then
                        dgvTraitementEnr.Rows(i).Cells("Heure2").Value = Strings.FormatDateTime(Libredatabase.Rows(i).Item("Heure2"), DateFormat.LongTime)
                    End If
                Next i
            Else
                dgvTraitementEnr.Rows.Clear()
            End If
        Catch ex As Exception
            MessageBox.Show("Erreur Chargement des Traitements Enregistrés: " & Chr(13) & ex.Message, "Interface Mecalux", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Private Sub chargementTache()
    '    Try
    '        CbIntitule.Items.Clear()
    '        LibreOleAdaptater = New OleDbDataAdapter("select * from TACHEPLANIFIER order by IDTache asc", OleConnenection)
    '        Libredataset = New DataSet
    '        LibreOleAdaptater.Fill(Libredataset)
    '        Libredatabase = Libredataset.Tables(0)
    '        If Libredatabase.Rows.Count <> 0 Then
    '            For i = 0 To Libredatabase.Rows.Count - 1
    '                CbIntitule.Items.Add(Libredatabase.Rows(i).Item("Intitule"))
    '            Next i
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show("Erreur Chargement des Traitements Enregistrés: " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
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
            ' chargementTraitementEnregistre(Join(Split(Trim(CbIntitule.Text), "'"), "''"))
        End If
    End Sub

    Private Sub dgvListeTrait_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvListeTrait.CellClick
        If e.RowIndex >= 0 Then
            Try
                chargementTraitementEnregistre(dgvListeTrait.Rows(e.RowIndex).Cells("intituleTrait").Value)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
        If e.ColumnIndex = dgvListeTrait.ColumnCount - 1 Then
            PlanificationTraitement.ShowDialog()
        End If
    End Sub
    'chargement de la liste des traitements
    Private Sub chargeListeTraitement()
        Try
            dgvListeTrait.Rows.Clear()
            dgvListeTrait.Rows.Add("Export Article", "A")
            dgvListeTrait.Rows.Add("Export Pseudos", "A")
            dgvListeTrait.Rows.Add("Export Client", "A")
            dgvListeTrait.Rows.Add("Export Fournisseur", "A")
            dgvListeTrait.Rows.Add("Export Commande Client", "A")
            dgvListeTrait.Rows.Add("Export Commande Fournisseur", "A")
            dgvListeTrait.Rows.Add("Import BL Client", "B1")
            dgvListeTrait.Rows.Add("Import BL Fournisseur", "B1")
            dgvListeTrait.Rows.Add("Import Mvt E/S", "B1")
        Catch ex As Exception
            MessageBox.Show("Erreur Chargement de la Liste des Traitements : " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub dgvListeTrait_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvListeTrait.CellContentClick

    End Sub

    Private Sub BTsup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTsup.Click
        Delete_DataListeSch()
        Planification_Load(sender, e)
    End Sub
    Private Sub Delete_DataListeSch()
        Try
            Dim i As Integer
            'Dim OleAdaptaterDelete As OleDbDataAdapter
            'Dim OleDeleteDataset As DataSet
            'Dim OledatableDelete As DataTable
            Dim OleCommandDelete As OleDbCommand
            'Dim DelFile As String
            For i = 0 To dgvTraitementEnr.RowCount - 1
                If dgvTraitementEnr.Rows(i).Cells("supEnr").Value = True Then
                    Dim d As Integer = dgvTraitementEnr.Rows(i).Cells("Id").Value
                    OleCommandDelete = New OleDbCommand("Delete From PLANIFICATION  WHERE  Id=" & dgvTraitementEnr.Rows(i).Cells("Id").Value)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                End If
            Next i
            dgvTraitementEnr.Rows.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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
        Dim Chaine As String = ""
        Dim Insert As Boolean = False
        Dim lancement As Boolean
        Dim i As Integer
        Try
            For i = 0 To dgvTraitementEnr.RowCount - 1

                If IsNumeric(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value)) = True And InStr(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value), ".") = 0 And InStr(Trim(dgvTraitementEnr.Rows(i).Cells("Rang").Value), ",") = 0 Then
                    If dgvTraitementEnr.Rows(i).Cells("Lancer").Value = True Then
                        lancement = True
                    Else
                        lancement = False
                    End If
                    Chaine = "UPDATE PLANIFICATION SET Intitule='" & Join(Split(dgvTraitementEnr.Rows(i).Cells("Societe").Value, "'"), "''") & "',Rang=" & CInt(dgvTraitementEnr.Rows(i).Cells("Rang").Value) & ",Critere1='" & Trim(dgvTraitementEnr.Rows(i).Cells("Critere1").Value) & "',Critere2='" & Trim(dgvTraitementEnr.Rows(i).Cells("Critere2").Value) & "',Heure1='" & dgvTraitementEnr.Rows(i).Cells("Heure1").Value & "',Heure2='" & dgvTraitementEnr.Rows(i).Cells("Heure2").Value & "',Lancer=" & lancement & " WHERE Id=" & CUInt(dgvTraitementEnr.Rows(i).Cells("Id").Value)
                    If Trim(dgvTraitementEnr.Rows(i).Cells("Heure1").Value) <> "" And Trim(dgvTraitementEnr.Rows(i).Cells("Heure2").Value) <> "" Then
                        accesscom = New OleDbCommand(Chaine, OleConnenection)
                        accesscom.ExecuteNonQuery()
                        Insert = True
                    Else
                        If Trim(dgvTraitementEnr.Rows(i).Cells("Heure1").Value) = "" And Trim(dgvTraitementEnr.Rows(i).Cells("Heure2").Value) = "" Then
                            'accesscom = New OleDbCommand("UPDATE  PLANIFICATION Set Rang=" & CInt(dgvTraitementEnr.Rows(i).Cells("Rang").Value) & ",Critere1='" & dgvTraitementEnr.Rows(i).Cells("Critere1").Value & "',Critere2='" & dgvTraitementEnr.Rows(i).Cells("Critere2").Value & "',Heure1=NULL,Heure2=NULL WHERE IntituleTache='" & Join(Split(Trim(CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(dgvTraitementEnr.Rows(i).Cells("Intitule").Value), "'"), "''") & "' And IDDossier=" & CInt(dgvTraitementEnr.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
                            'accesscom.ExecuteNonQuery()
                            'Insert = True
                        Else
                            MsgBox("Les Valeurs Heures doivent être tous renseignées ou Null !", MsgBoxStyle.Information, "Création des traitement")
                        End If
                        MsgBox("Les Valeurs Heures doivent être tous renseignées ou Null !", MsgBoxStyle.Information, "Création des traitement")
                    End If

                Else
                    MsgBox("Le Rang du traitement doit être un Entier : ", MsgBoxStyle.Information, "Planification de taches")
                End If
            Next i
            If Insert = True Then
                MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour Planification de taches")
            End If
        Catch ex As Exception
            MessageBox.Show("Erreur lors de la modification des des parametre planifié Erreur Retournée " & vbCrLf & ex.Message, "Traitement de modification de la planification ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

        'chargementTraitementEnregistre(Trim(CbIntitule.Text))
    End Sub
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Planification_Load(sender, e)
    End Sub
End Class