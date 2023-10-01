Imports System.Data.OleDb

Public Class FrmAllCorrespondance
    Public Critère As String = ""
    Public OleAdaptaterschema As OleDbDataAdapter
    Public OleSchemaDataset As DataSet
    Public OledatableSchema As DataTable

    Private Sub AfficheSchemasConso()

        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='" & Critère & "'", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
    End Sub

    Private Sub FrmAllCorrespondance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
            If Connected() Then
                BackgroundWorker1.RunWorkerAsync()
            Else
                MsgBox("Erreur de connexion au fichier access")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        If Connected() Then
            MiseàjourLeSchema()
        Else
            MsgBox("Erreur de connexion au fichier access")
        End If
    End Sub
    Private Sub MiseàjourLeSchema()
        Dim n As Integer
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String = ""
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    If Trim(DataListeIntegrer.Rows(n).Cells("Modifier").Value) = "True" Then
                        If DataListeIntegrer.Rows(n).Cells("InfosLibre").Value = "True" Then
                            If DataListeIntegrer.Rows(n).Cells("DefaultValue").Value <> "" Or DataListeIntegrer.Rows(n).Cells("ChampSage").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            ElseIf DataListeIntegrer.Rows(n).Cells("ChampSage").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            ElseIf DataListeIntegrer.Rows(n).Cells("DefaultValue").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            Else
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            End If
                        Else
                            If DataListeIntegrer.Rows(n).Cells("DefaultValue").Value <> "" Or DataListeIntegrer.Rows(n).Cells("ChampSage").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            ElseIf DataListeIntegrer.Rows(n).Cells("ChampSage").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            ElseIf DataListeIntegrer.Rows(n).Cells("DefaultValue").Value <> "" Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            Else
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "',ChampSage='" & Join(Split(DataListeIntegrer.Rows(n).Cells("ChampSage").Value, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                            End If
                        End If
                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    End If
                Next n
                If Insert = True Then
                    MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à jour de la Paramétrage de la Colonne")
                    BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub SuppressionSchema()
        Dim n As Integer
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Dim Cpter As Integer = 0
        Dim Cpters As Integer = 0
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    If Trim(DataListeIntegrer.Rows(n).Cells("Suppression").Value) = "True" Then
                        Cpter += 1
                    End If
                Next
                If Cpter > 0 Then
                    If MessageBox.Show("Voulez-vous vraiment éffectuer cette Operation de Suppression ", "Information Suppresion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                        For n = 0 To DataListeIntegrer.RowCount - 1
                            If Trim(DataListeIntegrer.Rows(n).Cells("Suppression").Value) = "True" Then
                                Insertion = "DELETE FROM P_COLONNEST  where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value & "'"
                                OleCommandEnreg = New OleDbCommand(Insertion)
                                OleCommandEnreg.Connection = OleConnenection
                                OleCommandEnreg.ExecuteNonQuery()
                                Insert = True
                                Cpters += 1
                            End If
                            If Cpters = Cpter Then
                                Exit For
                            End If
                        Next n
                        If Insert = True Then
                            MsgBox("Suppression éffectuer avec Succès nombre element Supprimer ( " & Cpter & " )", MsgBoxStyle.Information, "Suppresion Colonne paramétrée")
                            BackgroundWorker1.RunWorkerAsync()
                        Else
                            MsgBox("Erreur de Suppression nombre element Supprimer ( " & Cpter & " )", MsgBoxStyle.Information, "Suppresion Colonne paramétrée")
                        End If
                    End If
                Else
                    MsgBox("Selectionner l'élément à Supprimé ", MsgBoxStyle.Information, "Suppresion Colonne paramétrée")
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BtnSup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSup.Click
        Try
            If Connected() Then
                SuppressionSchema()
            Else
                MsgBox("Erreur de connexion au fichier access")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton2.Click
        Me.Close()
    End Sub
    Delegate Sub Evenement()

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        AfficheSchemasConso()
    End Sub
    Public Sub p()
        Dim i As Integer
        Try
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("Cols").Value = OledatableSchema.Rows(i).Item("Cols")
                DataListeIntegrer.Rows(i).Cells("description").Value = OledatableSchema.Rows(i).Item("description")
                DataListeIntegrer.Rows(i).Cells("Format").Value = OledatableSchema.Rows(i).Item("Format")
                DataListeIntegrer.Rows(i).Cells("PositionG").Value = OledatableSchema.Rows(i).Item("PositionG")
                DataListeIntegrer.Rows(i).Cells("DefaultValue").Value = OledatableSchema.Rows(i).Item("DefaultValue")
                DataListeIntegrer.Rows(i).Cells("ChampSage").Value = OledatableSchema.Rows(i).Item("ChampSage")
                DataListeIntegrer.Rows(i).Cells("InfosLibre").Value = OledatableSchema.Rows(i).Item("InfosLibre")
                DataListeIntegrer.Rows(i).Cells("CodeTbls").Value = OledatableSchema.Rows(i).Item("CodeTbls")
                DataListeIntegrer.Rows(i).Cells("Entete").Value = OledatableSchema.Rows(i).Item("Entete")
                DataListeIntegrer.Rows(i).Cells("Ligne").Value = OledatableSchema.Rows(i).Item("Ligne")
                DataListeIntegrer.Rows(i).Cells("Modifier").Value = "False"
                DataListeIntegrer.Rows(i).Cells("Suppression").Value = "False"
            Next i
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If DataListeIntegrer.InvokeRequired Then
            Dim MonDelegate As New Evenement(AddressOf p)
            DataListeIntegrer.Invoke(MonDelegate)
        Else
            p()
        End If
        PictureBox1.Visible = False
    End Sub
End Class
