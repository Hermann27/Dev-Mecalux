Imports System.Data.OleDb
Public Class FrmAllCorrespondance
    Public Critère As String = ""
    Public OleAdaptaterschema As OleDbDataAdapter
    Public OleSchemaDataset As DataSet
    Public OledatableSchema As DataTable

    Private Sub AfficheSchemasConso()
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='" & Critère & "' ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            'DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmAllCorrespondance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
            Me.WindowState = FormWindowState.Maximized
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
        Dim ValeurDefault As String = ""
        Dim ChampSage As String = ""
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    If Trim(DataListeIntegrer.Rows(n).Cells("Modifier").Value) = "True" Then
                        If DataListeIntegrer.Rows(n).Cells("InfosLibre").Value = "True" Then
                            If IsDBNull(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value) = False Then
                                ValeurDefault = DataListeIntegrer.Rows(n).Cells("DefaultValue").Value
                            Else
                                ValeurDefault = DataListeIntegrer.Rows(n).Cells("DefaultValue").Value
                            End If
                            If IsDBNull(DataListeIntegrer.Rows(n).Cells("ChampSage").Value) = True Then
                                ChampSage = "" 'DataListeIntegrer.Rows(n).Cells("ChampSage").Value
                            Else
                                ChampSage = DataListeIntegrer.Rows(n).Cells("ChampSage").Value
                            End If

                            'ValeurDefault = DataListeIntegrer.Rows(n).Cells("DefaultValue").Value
                            'ChampSage = DataListeIntegrer.Rows(n).Cells("ChampSage").Value
                            If IsDBNull(ValeurDefault) = False Or IsDBNull(ChampSage) = False Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            ElseIf IsDBNull(ChampSage) = False Then
                                Insertion = "UPDATE P_COLONNEST SET ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            ElseIf IsDBNull(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value) = False Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            Else
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & True & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            End If
                        Else
                            If IsDBNull(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value) = False Then
                                ValeurDefault = DataListeIntegrer.Rows(n).Cells("DefaultValue").Value
                            Else
                                If DataListeIntegrer.Rows(n).Cells("DefaultValue").Value.ToString <> "" Then
                                    ValeurDefault = DataListeIntegrer.Rows(n).Cells("DefaultValue").Value
                                Else
                                    ValeurDefault = ""
                                End If
                            End If
                            If IsDBNull(DataListeIntegrer.Rows(n).Cells("ChampSage").Value) = True Then
                                ChampSage = "" 'DataListeIntegrer.Rows(n).Cells("ChampSage").Value
                            Else
                                If DataListeIntegrer.Rows(n).Cells("ChampSage").Value <> "" Then
                                    ChampSage = DataListeIntegrer.Rows(n).Cells("ChampSage").Value
                                Else
                                    ChampSage = ""
                                End If
                            End If

                            If IsDBNull(ValeurDefault) = False Or IsDBNull(ChampSage) = False Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            ElseIf IsDBNull(DataListeIntegrer.Rows(n).Cells("ChampSage").Value) = False Then
                                Insertion = "UPDATE P_COLONNEST SET ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            ElseIf IsDBNull(DataListeIntegrer.Rows(n).Cells("DefaultValue").Value) = False Then
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
                            Else
                                Insertion = "UPDATE P_COLONNEST SET DefaultValue='" & Join(Split(ValeurDefault, "'"), "''") & "',ChampSage='" & Join(Split(ChampSage, "'"), "''") & "',InfosLibre=" & False & " where Cols='" & DataListeIntegrer.Rows(n).Cells("Cols").Value.ToString & "' And CodeTbls ='" & DataListeIntegrer.Rows(n).Cells("CodeTbls").Value.ToString & "' AND Ligne=" & DataListeIntegrer.Rows(n).Cells("Ligne").Value.ToString & " And Entete=" & DataListeIntegrer.Rows(n).Cells("Entete").Value.ToString
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
            MsgBox(ex.Message)
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
    Delegate Sub Evenement()
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            AfficheSchemasConso()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub p()
        Dim i As Integer
        Try
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
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
        Try
            If DataListeIntegrer.InvokeRequired Then
                Dim MonDelegate As New Evenement(AddressOf p)
                DataListeIntegrer.Invoke(MonDelegate)
            Else
                p()
            End If
            PictureBox1.Visible = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub DataListeIntegrer_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataListeIntegrer.Sorted
        Try
        Catch ex As Exception
        End Try
    End Sub
End Class