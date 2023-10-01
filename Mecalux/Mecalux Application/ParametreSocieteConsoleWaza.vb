Imports System.Data.OleDb
Imports System.IO
Public Class ParametreSocieteConsoleWaza
    Private Sub AfficheSchemasConso()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()
        OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("Societe1").Value = OledatableSchema.Rows(i).Item("Societe")
            DataListeIntegrer.Rows(i).Cells("Chemin1").Value = OledatableSchema.Rows(i).Item("Chemin1")
            DataListeIntegrer.Rows(i).Cells("Type1").Value = OledatableSchema.Rows(i).Item("nomtype")
            DataListeIntegrer.Rows(i).Cells("UserSage1").Value = OledatableSchema.Rows(i).Item("UserSage")
            DataListeIntegrer.Rows(i).Cells("PasseSage1").Value = "********" 'OledatableSchema.Rows(i).Item("PasseSage")
            DataListeIntegrer.Rows(i).Cells("bdd1").Value = OledatableSchema.Rows(i).Item("BaseDonnee")
            DataListeIntegrer.Rows(i).Cells("Serveur1").Value = OledatableSchema.Rows(i).Item("Serveur")
            DataListeIntegrer.Rows(i).Cells("NomUtil").Value = OledatableSchema.Rows(i).Item("NomUser")
            DataListeIntegrer.Rows(i).Cells("Mot").Value = "********" 'OledatableSchema.Rows(i).Item("MotPas")
        Next i
    End Sub
    Private Sub AfficheSociete()
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim i As Integer
        DataListeSchema.Rows.Clear()
        Type.Items.Clear()
        Try
            OleAdaptater = New OleDbDataAdapter("select * from TypeSage ", OleConnenection)

            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            For i = 0 To Oledatable.Rows.Count - 1
                If Trim(Oledatable.Rows(i).Item("nomtype")) <> "" Then
                    Type.Items.AddRange(New String() {Oledatable.Rows(i).Item("nomtype")})
                End If
            Next i
        Catch ex As Exception
        End Try
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
                OleAdaptaterDelete = New OleDbDataAdapter("select * from PARAMETRE where Societe='" & DataListeIntegrer.Rows(i).Cells("Societe1").Value & "' and Chemin1 ='" & DataListeIntegrer.Rows(i).Cells("Chemin1").Value & "'", OleConnenection)
                OleDeleteDataset = New DataSet
                OleAdaptaterDelete.Fill(OleDeleteDataset)
                OledatableDelete = OleDeleteDataset.Tables(0)
                If OledatableDelete.Rows.Count <> 0 Then
                    DelFile = "Delete From PARAMETRE where Societe='" & DataListeIntegrer.Rows(i).Cells("Societe1").Value & "' and Chemin1 ='" & DataListeIntegrer.Rows(i).Cells("Chemin1").Value & "'"
                    OleCommandDelete = New OleDbCommand(DelFile)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                End If
            End If
        Next i
        AfficheSchemasConso()
    End Sub
    Private Sub EnregistrerLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand

        Dim Insert As Boolean = False
        Dim Insertion As String
        If DataListeSchema.RowCount >= 1 Then
            For n = 0 To DataListeSchema.RowCount - 1

                If Trim(DataListeSchema.Rows(n).Cells("Societe").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("type").Value) <> "" Then
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARAMETRE WHERE Societe='" & DataListeSchema.Rows(n).Cells("Societe").Value & "' And  nomtype ='" & DataListeSchema.Rows(n).Cells("type").Value & "'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        MsgBox("Cette Société Existe déja!", MsgBoxStyle.Information, "Creation Societe")

                    Else
                        If Trim(DataListeSchema.Rows(n).Cells("Societe").Value) <> "" Then
                            Insertion = "Insert Into PARAMETRE (Societe,nomtype,Chemin1,UserSage,PasseSage,BaseDonnee,Serveur,NomUser,MotPas) VALUES ('" & DataListeSchema.Rows(n).Cells("Societe").Value & "','" & DataListeSchema.Rows(n).Cells("Type").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("UserSage").Value & "','" & DataListeSchema.Rows(n).Cells("PasseSage").Value & "','" & DataListeSchema.Rows(n).Cells("bdd").Value & "','" & DataListeSchema.Rows(n).Cells("Serveur").Value & "','" & DataListeSchema.Rows(n).Cells("Utilisateur").Value & "','" & DataListeSchema.Rows(n).Cells("Passe").Value & "')"
                            OleCommandEnreg = New OleDbCommand(Insertion)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            Insert = True

                        End If
                    End If
                End If

            Next n
            If Insert = True Then
                MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Parametrage des Societes")
                DataListeSchema.Rows.Clear()
            End If
        End If
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Delete_DataListeSch()
    End Sub
    Private Sub BT_DelRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelRow.Click
        Dim first As Integer
        Dim last As Integer
        first = DataListeSchema.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
        last = DataListeSchema.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        If last >= 0 Then
            If last - first >= 0 Then
                DataListeSchema.Rows.RemoveAt(DataListeSchema.CurrentRow.Index)
            End If
        End If
    End Sub
    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        DataListeSchema.Rows.Add()
    End Sub

    Private Sub ParametreSocieteConsoleWaza_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
        If Connected() = True Then
            AfficheSchemasConso()
            AfficheSociete()
        End If
    End Sub

    Private Sub DataListeSchema_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellContentClick
        Try
            Dim i As Integer
            Dim test As Boolean = False
            If e.RowIndex >= 0 Then
                If DataListeSchema.Columns(e.ColumnIndex).Name = "find" Then
                    If DataListeSchema.Rows(e.RowIndex).Cells("Type").Value = "COMPTABILITE" Then
                        FindFile.Filter = "Fichier Base Sage Comptabilité (*.MAE)|*.MAE"
                    End If
                    If DataListeSchema.Rows(e.RowIndex).Cells("Type").Value = "COMMERCIAL" Then
                        FindFile.Filter = "Fichier Base Sage Commmercial (*.GCM)|*.GCM"
                    End If
                    If DataListeSchema.Rows(e.RowIndex).Cells("Type").Value = "" Then
                        MsgBox("Veuillez choisir d'abord le Type!", MsgBoxStyle.Information, "Selection Chemin Base")
                        Exit Sub
                    End If
                    FindFile.FileName = Nothing
                    If FindFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                        For i = DataListeSchema.Rows.Count - 1 To 0 Step -1
                            If DataListeSchema.Rows(i).Cells("Chemin").Value = Trim(FindFile.FileName) Then
                                test = True
                            End If
                        Next
                        If test = True Then
                            MessageBox.Show("Cette base est déjà listée, entrez-en une autre qui ne l'est pas S.V.P.", "Parametrage des bases")
                        Else
                            If File.Exists(Trim(FindFile.FileName)) = True Then
                                DataListeSchema.Rows(e.RowIndex).Cells("bdd").Value = System.IO.Path.GetFileNameWithoutExtension(Trim(FindFile.FileName))
                            End If
                            DataListeSchema.Rows(e.RowIndex).Cells("Serveur").Value = LireChaine(Trim(FindFile.FileName), "CBASE", "ServeurSQL")
                            DataListeSchema.Rows(e.RowIndex).Cells("Chemin").Value = FindFile.FileName
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub DataListeIntegrer_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellContentClick

    End Sub

    Private Sub MiseàjourLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARAMETRE where  Societe='" & DataListeIntegrer.Rows(n).Cells("Societe1").Value & "' And nomtype ='" & DataListeIntegrer.Rows(n).Cells("Type1").Value & "'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        If Trim(DataListeIntegrer.Rows(n).Cells("PasseSage1").Value) <> "********" Then
                            If Trim(DataListeIntegrer.Rows(n).Cells("Mot").Value) <> "********" Then
                                Insertion = "UPDATE PARAMETRE SET UserSage='" & DataListeIntegrer.Rows(n).Cells("UserSage1").Value & "',PasseSage='" & DataListeIntegrer.Rows(n).Cells("PasseSage1").Value & "',Serveur='" & DataListeIntegrer.Rows(n).Cells("Serveur1").Value & "',BaseDonnee='" & DataListeIntegrer.Rows(n).Cells("bdd1").Value & "',NomUser='" & DataListeIntegrer.Rows(n).Cells("NomUtil").Value & "',MotPas='" & DataListeIntegrer.Rows(n).Cells("Mot").Value & "' where Societe='" & DataListeIntegrer.Rows(n).Cells("Societe1").Value & "' And nomtype ='" & DataListeIntegrer.Rows(n).Cells("Type1").Value & "'"
                                OleCommandEnreg = New OleDbCommand(Insertion)
                                OleCommandEnreg.Connection = OleConnenection
                                OleCommandEnreg.ExecuteNonQuery()
                                Insert = True
                            Else
                                Insertion = "UPDATE PARAMETRE SET UserSage='" & DataListeIntegrer.Rows(n).Cells("UserSage1").Value & "',PasseSage='" & DataListeIntegrer.Rows(n).Cells("PasseSage1").Value & "',Serveur='" & DataListeIntegrer.Rows(n).Cells("Serveur1").Value & "',BaseDonnee='" & DataListeIntegrer.Rows(n).Cells("bdd1").Value & "',NomUser='" & DataListeIntegrer.Rows(n).Cells("NomUtil").Value & "' where Societe='" & DataListeIntegrer.Rows(n).Cells("Societe1").Value & "' And nomtype ='" & DataListeIntegrer.Rows(n).Cells("Type1").Value & "'"
                                OleCommandEnreg = New OleDbCommand(Insertion)
                                OleCommandEnreg.Connection = OleConnenection
                                OleCommandEnreg.ExecuteNonQuery()
                                Insert = True
                            End If
                        Else
                            If Trim(DataListeIntegrer.Rows(n).Cells("Mot").Value) <> "********" Then
                                Insertion = "UPDATE PARAMETRE SET UserSage='" & DataListeIntegrer.Rows(n).Cells("UserSage1").Value & "',Serveur='" & DataListeIntegrer.Rows(n).Cells("Serveur1").Value & "',BaseDonnee='" & DataListeIntegrer.Rows(n).Cells("bdd1").Value & "',NomUser='" & DataListeIntegrer.Rows(n).Cells("NomUtil").Value & "',MotPas='" & DataListeIntegrer.Rows(n).Cells("Mot").Value & "' where Societe='" & DataListeIntegrer.Rows(n).Cells("Societe1").Value & "' And nomtype ='" & DataListeIntegrer.Rows(n).Cells("Type1").Value & "'"
                                OleCommandEnreg = New OleDbCommand(Insertion)
                                OleCommandEnreg.Connection = OleConnenection
                                OleCommandEnreg.ExecuteNonQuery()
                                Insert = True
                            End If
                        End If
                    Else
                    End If
                Next n
                If Insert = True Then
                    MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour des Conditionnement")
                End If
            End If
            AfficheSchemasConso()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        EnregistrerLeSchema()
        AfficheSchemasConso()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        MiseàjourLeSchema()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Delete_DataListeSch()
    End Sub

    Private Sub BT_Quit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub

    Private Sub BtnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTest.Click
        Frm_EtatDeConexion.Show()
    End Sub
End Class