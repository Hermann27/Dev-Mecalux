Imports System.Data.OleDb
Imports System.IO
Public Class PlanificationTraitement
    Private Sub AfficheSchemasConso()
        Dim i, m As Integer
        m = 0
        lblInfos.Text = ""
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim EstVide, EstPlein As Integer
        Dim OleAdaptaterschemaVerif As OleDbDataAdapter
        Dim OleSchemaDatasetVerif As DataSet
        Dim OledatableSchemaVerif As DataTable
        BtValider.Enabled = True
        Try

            OleAdaptaterschemaVerif = New OleDbDataAdapter("SELECT distinct intitule FROM PLANIFICATION WHERE IntituleTache='" & Join(Split(PlanificationSpecial.dgvListeTrait.SelectedRows(0).Cells("intituleTrait").Value, "'"), "''") & "'", OleConnenection)
            OleSchemaDatasetVerif = New DataSet
            OleAdaptaterschemaVerif.Fill(OleSchemaDatasetVerif)
            OledatableSchemaVerif = OleSchemaDatasetVerif.Tables(0)

            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            EstPlein = OledatableSchema.Rows.Count

            For i = 0 To 1 - 1 ' OledatableSchema.Rows.Count - 1
                OleAdaptaterschemaVerif = New OleDbDataAdapter("SELECT distinct intitule FROM PLANIFICATION WHERE IntituleTache='" & Join(Split(PlanificationSpecial.dgvListeTrait.SelectedRows(0).Cells("intituleTrait").Value, "'"), "''") & "' AND intitule='" & OledatableSchema.Rows(i).Item("Societe").ToString & "'", OleConnenection)
                OleSchemaDatasetVerif = New DataSet
                OleAdaptaterschemaVerif.Fill(OleSchemaDatasetVerif)
                OledatableSchemaVerif = OleSchemaDatasetVerif.Tables(0)

                If OledatableSchemaVerif.Rows.Count = 0 Then
                    DataListeIntegrer.RowCount = m + 1
                    DataListeIntegrer.Rows(m).Cells("Societe1").Value = "Tous Les Sociétés" ' OledatableSchema.Rows(i).Item("Societe")
                    m += 1
                Else
                    EstVide = i + 1
                End If
            Next i
            If EstVide = EstPlein Then
                BtValider.Enabled = False
                lblInfos.Text = "Toutes les Société on déjà été Paramètrer sur la Planification " & PlanificationSpecial.dgvListeTrait.SelectedRows(0).Cells("intituleTrait").Value
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub PlanificationTraitement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
        If Connected() = True Then
            AfficheSchemasConso()
        End If
        'Dim LibreOleAdaptater As OleDbDataAdapter
        'Dim Libredataset As DataSet
        'Dim Libredatabase As DataTable
        'Try
        '    Dim Requete As String = Planification.Requete
        '    Dim i, j As Integer
        '    Dim OleAdaptaterschema As OleDbDataAdapter
        '    Dim OleSchemaDataset As DataSet
        '    Dim OledatableSchema As DataTable
        '    GrvTraitement.Rows.Clear()
        '    GrvTraitement.Columns.Clear()
        '    OleAdaptaterschema = New OleDbDataAdapter(Requete, OleConnenection)
        '    OleSchemaDataset = New DataSet
        '    OleAdaptaterschema.Fill(OleSchemaDataset)
        '    OledatableSchema = OleSchemaDataset.Tables(0)
        '    If OledatableSchema.Columns.Count <> 0 Then
        '        Dim ocolumn As New DataGridViewCheckBoxColumn
        '        With ocolumn
        '            .Name = "Rattacher"
        '            .HeaderText = "Rattacher"
        '            .Visible = True
        '            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        '            .Width = 60
        '            .ReadOnly = False
        '            .SortMode = DataGridViewColumnSortMode.NotSortable
        '        End With
        '        GrvTraitement.Columns.Add(ocolumn)
        '        Dim ocolumns As New DataGridViewTextBoxColumn
        '        With ocolumns
        '            .Name = "Rang"
        '            .HeaderText = "Rang"
        '            .Visible = True
        '            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        '            .Width = 60
        '            .ReadOnly = False
        '            .SortMode = DataGridViewColumnSortMode.NotSortable
        '        End With
        '        GrvTraitement.Columns.Add(ocolumns)
        '        CreationAutoCellule(GrvTraitement, "Heure 1", "Heure1", "")
        '        CreationAutoCellule(GrvTraitement, "Heure 2", "Heure2", "")
        '        Dim docolumn As New DataGridViewButtonColumn
        '        With docolumn
        '            .Name = "Heure"
        '            .HeaderText = "Heure"
        '            .Visible = True
        '            .Text = "..."
        '            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        '            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        '            .DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '            .UseColumnTextForButtonValue = True
        '            .ReadOnly = False
        '            .SortMode = DataGridViewColumnSortMode.NotSortable
        '            .Width = 50

        '        End With
        '        GrvTraitement.Columns.Add(docolumn)
        '        CreationAutoCellule(GrvTraitement, "Critere 1", "Critere1", "")
        '        CreationAutoCellule(GrvTraitement, "Critere 2", "Critere2", "")
        '        Dim docolumn1 As New DataGridViewButtonColumn
        '        With docolumn1
        '            .Name = "DateM"
        '            .HeaderText = "Date"
        '            .Visible = True
        '            .Text = "..."
        '            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        '            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        '            .DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '            .UseColumnTextForButtonValue = True
        '            .ReadOnly = False
        '            .SortMode = DataGridViewColumnSortMode.NotSortable
        '            .Width = 50

        '        End With
        '        GrvTraitement.Columns.Add(docolumn1)
        '        For i = OledatableSchema.Columns.Count - 1 To 0 Step -1
        '            If OledatableSchema.Columns(i).DataType.Name = "Boolean" Then
        '                CreationAutoCellule(GrvTraitement, OledatableSchema.Columns(i).ColumnName, OledatableSchema.Columns(i).ColumnName, "Boolean")
        '            Else
        '                CreationAutoCellule(GrvTraitement, OledatableSchema.Columns(i).ColumnName, OledatableSchema.Columns(i).ColumnName, "")
        '            End If

        '        Next i
        '    End If
        '    If OledatableSchema.Rows.Count <> 0 Then
        '        GrvTraitement.RowCount = OledatableSchema.Rows.Count
        '        Dim m As Integer = 0
        '        For i = 0 To OledatableSchema.Rows.Count - 1
        '            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION Where IntituleTache='" & Join(Split(Trim(Planification.CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value), "'"), "''") & "' And IDDossier=" & OledatableSchema.Rows(i).Item("IDDossier") & "", OleConnenection)
        '            Libredataset = New DataSet
        '            LibreOleAdaptater.Fill(Libredataset)
        '            Libredatabase = Libredataset.Tables(0)
        '            If Libredatabase.Rows.Count <> 0 Then
        '            Else
        '                LibreOleAdaptater = New OleDbDataAdapter("select Max(Rang) as Rang from PLANIFICATION Where IntituleTache='" & Join(Split(Trim(Planification.CbIntitule.Text), "'"), "''") & "'", OleConnenection)
        '                Libredataset = New DataSet
        '                LibreOleAdaptater.Fill(Libredataset)
        '                Libredatabase = Libredataset.Tables(0)
        '                If Libredatabase.Rows.Count <> 0 Then
        '                    If Convert.IsDBNull(Libredatabase.Rows(0).Item("Rang")) = False Then
        '                        GrvTraitement.Rows(i).Cells("Rang").Value = Libredatabase.Rows(0).Item("Rang") + m + 1
        '                    Else
        '                        GrvTraitement.Rows(i).Cells("Rang").Value = m + 1
        '                    End If
        '                Else
        '                    GrvTraitement.Rows(i).Cells("Rang").Value = m + 1
        '                End If
        '                m = m + 1
        '            End If
        '            For j = 0 To OledatableSchema.Columns.Count - 1
        '                GrvTraitement.Rows(i).Cells("" & Trim(OledatableSchema.Columns(j).ColumnName) & "").Value = OledatableSchema.Rows(i).Item(j)
        '            Next j
        '        Next i
        '    End If

        'Catch ex As Exception
        '    MsgBox("Message Systeme: " & ex.Message, MsgBoxStyle.Information, "Affichage des Traitement à Planifier")
        'End Try
    End Sub
    Public Sub CreationAutoCellule(ByRef Dataobject As DataGridView, ByRef Colname As String, ByRef HeaderName As String, ByRef TypeColonne As String)

        If Trim(TypeColonne) = "Boolean" Then
            Dim ocolumn As New DataGridViewCheckBoxColumn
            With ocolumn
                .Name = HeaderName
                .HeaderText = Colname
                .Width = 60
                .Visible = True
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .ReadOnly = True
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
            Dataobject.Columns.Add(ocolumn)
        Else
            Dim ocolumn As New DataGridViewTextBoxColumn
            With ocolumn
                .Name = HeaderName
                .HeaderText = Colname
                .Width = 100
                .Visible = True
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .ReadOnly = True
                .SortMode = DataGridViewColumnSortMode.Automatic
            End With
            Dataobject.Columns.Add(ocolumn)
        End If
    End Sub

    Private Sub GrvTraitement_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GrvTraitement.CellClick
        If e.RowIndex >= 0 Then
            Idexe = e.RowIndex
            If GrvTraitement.Columns(e.ColumnIndex).Name = "Heure" Then
                PlanificationHeure.Text = "PlanificationTraitement"
                PlanificationHeure.ShowDialog()
            Else
                If GrvTraitement.Columns(e.ColumnIndex).Name = "DateM" Then
                    PlanificationDate.Text = "PlanificationTraitement"
                    PlanificationDate.ShowDialog()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub GrvTraitement_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GrvTraitement.CellContentClick

    End Sub

    Private Sub BtValider_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtValider.Click
        Dim accesscom As OleDbCommand
        Dim LibreOleAdaptater As OleDbDataAdapter
        Dim Libredataset As DataSet
        Dim Insert As Boolean = False
        Dim Libredatabase As DataTable
        Dim etat As Integer = 0
        Try

            For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1

                If DataListeIntegrer.Rows(i).Cells("charge").Value = True Then

                    etat = 1
                    If DataListeIntegrer.Rows(i).Cells("Heure1").Value > DataListeIntegrer.Rows(i).Cells("Heure2").Value Then
                        Throw New Exception("L'heure de debut doit être inférieure à l'heure de fin ligne N° :" & i)
                    End If

                    If DataListeIntegrer.Rows(i).Cells("Critere1").Value > DataListeIntegrer.Rows(i).Cells("Critere2").Value Then
                        Throw New Exception("La date de debut doit être inférieure à la date de fin ligne N° :" & i)
                    End If
                    If Trim(DataListeIntegrer.Rows(i).Cells("Heure1").Value) <> "" And Trim(DataListeIntegrer.Rows(i).Cells("Heure2").Value) <> "" And Trim(DataListeIntegrer.Rows(i).Cells("Critere1").Value) <> "" And Trim(DataListeIntegrer.Rows(i).Cells("Critere2").Value) <> "" Then
                        'DateAndTime.Hour(Strings.FormatDateTime(GrvTraitement.Rows(i).Cells("Heure1").Value, DateFormat.ShortTime)) & ":" & DateAndTime.Minute(Strings.FormatDateTime(GrvTraitement.Rows(i).Cells("Heure1").Value, DateFormat.ShortTime))

                        Dim q As String = "insert into PLANIFICATION (IntituleTache,Intitule,Rang,Critere1,Critere2,Heure1,Heure2) values ('" & Join(Split(PlanificationSpecial.dgvListeTrait.SelectedRows(0).Cells("intituleTrait").Value, "'"), "''") & "','" & Join(Split(DataListeIntegrer.Rows(0).Cells("Societe1").Value, "'"), "''") & "'" & "," & CInt(DataListeIntegrer.Rows(i).Cells("Rang").Value) & ",'" & Trim(DataListeIntegrer.Rows(i).Cells("Critere1").Value) & "','" & Trim(DataListeIntegrer.Rows(i).Cells("Critere2").Value) & "','" & DataListeIntegrer.Rows(i).Cells("Heure1").Value & "','" & DataListeIntegrer.Rows(i).Cells("Heure2").Value & "')"
                        accesscom = New OleDbCommand(q, OleConnenection)
                        accesscom.ExecuteNonQuery()
                        accesscom = New OleDbCommand("SELECT MAX(Id) FROM Planification", OleConnenection)
                        Dim id As Integer = CUInt(accesscom.ExecuteScalar)
                        Insert = True
                        PlanificationSpecial.dgvTraitementEnr.Rows.Add()
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("intitule").Value = PlanificationSpecial.dgvListeTrait.SelectedRows(0).Cells("intituleTrait").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Heure1").Value = DataListeIntegrer.Rows(i).Cells("Heure1").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Heure2").Value = DataListeIntegrer.Rows(i).Cells("Heure2").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Critere1").Value = DataListeIntegrer.Rows(i).Cells("Critere1").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Critere2").Value = DataListeIntegrer.Rows(i).Cells("Critere2").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Rang").Value = DataListeIntegrer.Rows(i).Cells("Rang").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Societe1").Value = DataListeIntegrer.Rows(i).Cells("Societe1").Value
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Id").Value = id
                        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Societe").Value = "Tous les Sociétés"
                        'Dim assignedToColumn As DataGridViewComboBoxColumn = PlanificationSpecial.dgvTraitementEnr.Columns("Societe")
                        ''Dim vt As DataRowView = assignedToColumn.Items(0)
                        ''Planification.dgvTraitementEnr.Rows(i).Cells("Societe").Value = vt.Row(0).ToString
                        'assignedToColumn.DataSource = Nothing
                        'assignedToColumn.DataSource = PlanificationSpecial.OledatableSchema
                        'assignedToColumn.DisplayMember = "Societe"
                        'assignedToColumn.ValueMember = "Societe"
                        'For j As Integer = 0 To assignedToColumn.Items.Count - 1
                        '    Dim vt As DataRowView = assignedToColumn.Items(j)
                        '    Dim f As String = vt.Row(0).ToString
                        '    If vt.Row(0).ToString.Equals(DataListeIntegrer.Rows(i).Cells("Societe1").Value) Then
                        '        PlanificationSpecial.dgvTraitementEnr.Rows(PlanificationSpecial.dgvTraitementEnr.Rows.Count - 1).Cells("Societe").Value = vt.Row(0).ToString
                        '    End If
                        'Next
                    Else
                        MsgBox("Toutes les valeurs doivent être renseignées !", MsgBoxStyle.Information, "Création des traitement")
                    End If

                End If

            Next

            If etat = 0 Then
                MsgBox("Pour valider cochez la case Charger !", MsgBoxStyle.Information, "Création des traitement")
            Else
                etat = 0
            End If


        Catch ex As Exception
            MessageBox.Show("Erreur Enregistrement des Traitements : " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'Dim accesscom As OleDbCommand
        'Dim i As Integer
        'Dim LibreOleAdaptater As OleDbDataAdapter
        'Dim Libredataset As DataSet
        'Dim Insert As Boolean = False
        'Dim Libredatabase As DataTable
        'For i = 0 To GrvTraitement.RowCount - 1
        '    If GrvTraitement.Rows(i).Cells("Rattacher").Value = True Then
        '        Try
        '            LibreOleAdaptater = New OleDbDataAdapter("select * from PLANIFICATION Where IntituleTache='" & Join(Split(Trim(Planification.CbIntitule.Text), "'"), "''") & "' And Intitule='" & Join(Split(Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value), "'"), "''") & "' And IDDossier=" & CInt(GrvTraitement.Rows(i).Cells("IDDossier").Value) & "", OleConnenection)
        '            Libredataset = New DataSet
        '            LibreOleAdaptater.Fill(Libredataset)
        '            Libredatabase = Libredataset.Tables(0)
        '            If Libredatabase.Rows.Count <> 0 Then
        '                MsgBox("Traitement : " & Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value & " , ID du traitement : " & GrvTraitement.Rows(i).Cells("IDDossier").Value & " déja Existant dans la tache", MsgBoxStyle.Information, " Création Traitement")
        '            Else
        '                If Trim(Planification.CbIntitule.Text) <> "" And Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value) <> "" And Trim(GrvTraitement.Rows(i).Cells("IDDossier").Value) <> "" Then
        '                    If Trim(GrvTraitement.Rows(i).Cells("Heure1").Value) <> "" And Trim(GrvTraitement.Rows(i).Cells("Heure2").Value) <> "" Then
        '                        'DateAndTime.Hour(Strings.FormatDateTime(GrvTraitement.Rows(i).Cells("Heure1").Value, DateFormat.ShortTime)) & ":" & DateAndTime.Minute(Strings.FormatDateTime(GrvTraitement.Rows(i).Cells("Heure1").Value, DateFormat.ShortTime))
        '                        accesscom = New OleDbCommand("insert into PLANIFICATION (IntituleTache,Intitule,IDDossier,Rang,Critere1,Critere2,Heure1,Heure2) values ('" & Join(Split(Planification.CbIntitule.Text, "'"), "''") & "','" & Join(Split(Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value), "'"), "''") & "', " & CInt(GrvTraitement.Rows(i).Cells("IDDossier").Value) & "," & CInt(GrvTraitement.Rows(i).Cells("Rang").Value) & ",'" & Trim(GrvTraitement.Rows(i).Cells("Critere1").Value) & "','" & Trim(GrvTraitement.Rows(i).Cells("Critere2").Value) & "','" & GrvTraitement.Rows(i).Cells("Heure1").Value & "','" & GrvTraitement.Rows(i).Cells("Heure2").Value & "')", OleConnenection)
        '                        accesscom.ExecuteNonQuery()
        '                        Insert = True
        '                    Else
        '                        If Trim(GrvTraitement.Rows(i).Cells("Heure1").Value) = "" And Trim(GrvTraitement.Rows(i).Cells("Heure2").Value) = "" Then
        '                            'accesscom = New OleDbCommand("insert into PLANIFICATION (IntituleTache,Intitule,IDDossier,Rang,Critere1,Critere2) values ('" & Join(Split(Planification.CbIntitule.Text, "'"), "''") & "','" & Join(Split(Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value), "'"), "''") & "', " & CInt(GrvTraitement.Rows(i).Cells("IDDossier").Value) & "," & CInt(GrvTraitement.Rows(i).Cells("Rang").Value) & ",'" & Trim(GrvTraitement.Rows(i).Cells("Critere1").Value) & "','" & Trim(GrvTraitement.Rows(i).Cells("Critere2").Value) & "')", OleConnenection)
        '                            accesscom = New OleDbCommand("insert into PLANIFICATION (IntituleTache,Intitule,IDDossier,Rang,Critere1,Critere2,Heure1,Heure2) values ('" & Join(Split(Planification.CbIntitule.Text, "'"), "''") & "','" & Join(Split(Trim(Planification.dgvListeTrait.Rows(Planification.numLignList).Cells("intituleTrait").Value), "'"), "''") & "', " & CInt(GrvTraitement.Rows(i).Cells("IDDossier").Value) & "," & CInt(GrvTraitement.Rows(i).Cells("Rang").Value) & ",'" & Trim(GrvTraitement.Rows(i).Cells("Critere1").Value) & "','" & Trim(GrvTraitement.Rows(i).Cells("Critere2").Value) & "','" & GrvTraitement.Rows(i).Cells("Heure1").Value & "','" & GrvTraitement.Rows(i).Cells("Heure2").Value & "')", OleConnenection)
        '                            accesscom.ExecuteNonQuery()
        '                            Insert = True
        '                        Else
        '                            MsgBox("Les valeurs Heures doivent être tous renseignées ou Null !", MsgBoxStyle.Information, "Création des traitement")
        '                        End If
        '                    End If
        '                End If
        '            End If
        '        Catch ex As Exception
        '            MessageBox.Show("Erreur Enregistrement des Traitements : " & Chr(13) & ex.Message, "Console Waza", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        End Try
        '    End If
        'Next i
        'Planification.chargementTraitementEnregistre(Trim(Planification.CbIntitule.Text))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DataListeIntegrer_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        If e.RowIndex >= 0 Then
            Idexe = e.RowIndex
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Heure" Then
                PlanificationHeure.Text = "PlanificationTraitement"
                PlanificationHeure.ShowDialog()
            Else
                If DataListeIntegrer.Columns(e.ColumnIndex).Name = "DateM" Then
                    PlanificationDate.Text = "PlanificationTraitement"
                    PlanificationDate.ShowDialog()
                Else
                End If
            End If
        End If
    End Sub
End Class